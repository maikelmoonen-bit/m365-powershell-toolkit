#Requires -Version 7.4
#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Tenant-brede SharePoint / Microsoft 365 owner cleanup met verplichte dry-run.

.DESCRIPTION
    Verwijdert een of meer opgegeven accounts (bijvoorbeeld vertrokken
    medewerkers of oude migratie-accounts) als eigenaar uit alle sites van een
    tenant, zonder dat er sites of groepen zonder eigenaar achterblijven.

    Standaard voert dit script GEEN wijzigingen uit.
    Alleen wanneer -Execute wordt meegegeven worden wijzigingen uitgevoerd.

    Het script controleert:
      1. Site Collection Administrators.
      2. De Associated SharePoint Owners-groep van root sites en subsites.
      3. Owners van gekoppelde Microsoft 365-groepen.

    Regels voor de backup-owner (-BackupOwnerAccount):
      - Als het verwijderen van targets ertoe zou leiden dat een owner-scope
        geen owners meer overhoudt, wordt de backup-owner EERST toegevoegd.
      - Als de backup-owner al owner is en er na cleanup minimaal een andere,
        niet te verwijderen owner overblijft, wordt de backup-owner verwijderd.
      - De backup-owner blijft dus alleen staan wanneer deze noodzakelijk is om
        een owner-loze site/groep te voorkomen.

.PARAMETER TargetAccount
    Een of meer accounts die als owner verwijderd moeten worden. Een waarde met
    '@' wordt gezocht als UPN of e-mailadres (ook guest-UPN's met #EXT#);
    een waarde zonder '@' als accountnaam.

.PARAMETER BackupOwnerAccount
    Account dat tijdelijk owner wordt wanneer een scope anders zonder owner
    zou komen te staan. UPN/e-mail of accountnaam, zoals bij -TargetAccount.

.NOTES
    Performance:
      * Subsites hergebruiken de connectie van de site collection; de
        owners-groep wordt via CSOM op de bestaande context gewijzigd.
      * Microsoft 365 Group owners worden vooraf in een enkele gepagineerde
        Graph-query opgehaald. Groepen met exact 20 owners worden alsnog
        individueel opgehaald, omdat Microsoft Graph bij $expand maximaal 20
        gerelateerde objecten teruggeeft.
      * Get-PnPTenantSite wordt zonder -Detailed aangeroepen. Gebruik
        -DetailedSiteEnumeration als GroupId in de tenant niet gevuld blijkt
        te zijn (het script waarschuwt hier expliciet voor).
      * De site-loop draait parallel via ForEach-Object -Parallel: standaard
        8 threads in dry-run, 5 bij -Execute. Hoger dan ~8 levert bij
        SharePoint Online app-only doorgaans 429-throttling op.

    Vereiste app-only permissies:

    SharePoint Application:
      - Sites.FullControl.All

    Microsoft Graph Application:
      - User.Read.All
      - Group.ReadWrite.All

    Admin consent is vereist.

.EXAMPLE
    $connectie = @{
        ClientId              = '<app-id>'
        TenantId              = '<tenant-id>'
        TenantDomain          = 'contoso.onmicrosoft.com'
        TenantAdminUrl        = 'https://contoso-admin.sharepoint.com'
        CertificateThumbprint = '<thumbprint>'
    }
    .\Remove-SharePointOwnersTenantWide.ps1 @connectie -TargetAccount 'oud.account@contoso.com', 'svc-migratie' -BackupOwnerAccount 'svc-spo-backup@contoso.com'

    Alleen dry-run.

.EXAMPLE
    .\Remove-SharePointOwnersTenantWide.ps1 @connectie -TargetAccount 'oud.account@contoso.com' -BackupOwnerAccount 'svc-spo-backup@contoso.com' -Execute

    Voert de wijzigingen daadwerkelijk uit.

.EXAMPLE
    .\Remove-SharePointOwnersTenantWide.ps1 @connectie -TargetAccount 'oud.account@contoso.com' -BackupOwnerAccount 'svc-spo-backup@contoso.com' -ThrottleLimit 12 -SkipSubSites

    Dry-run met 12 gelijktijdige threads en zonder subsite-enumeratie.
    Aanzienlijk sneller wanneer de tenant vrijwel uitsluitend moderne sites
    zonder subwebs bevat.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$TargetAccount,

    [Parameter(Mandatory)]
    [string]$BackupOwnerAccount,

    [switch]$Execute,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$TenantDomain,

    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory)]
    [string]$CertificateThumbprint,

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 0,

    [switch]$SkipSubSites,

    [switch]$DetailedSiteEnumeration,

    [switch]$DisableGroupOwnerCache,

    [string]$ReportPath = (Join-Path -Path $PWD -ChildPath ("SharePoint_Owner_Cleanup_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===========================================================================
# Configuratie
# ===========================================================================

function New-AccountDefinition {
    param([Parameter(Mandatory)][string]$Value)

    # Met '@' zoeken op UPN/mail (incl. #EXT#-guests), anders op accountnaam.
    [pscustomobject]@{
        Label       = $Value
        SearchValue = $Value
        SearchType  = if ($Value -like "*@*") { "UPNOrMail" } else { "AccountName" }
    }
}

$TargetDefinitions = @($TargetAccount | ForEach-Object { New-AccountDefinition -Value $_ })

$BackupOwnerDefinition = New-AccountDefinition -Value $BackupOwnerAccount

$IncludedTemplates = @(
    "SITEPAGEPUBLISHING#0",
    "STS#3",
    "STS#0",
    "PROJECTSITE#0",
    "ENTERWIKI#0",
    "BDR#0",
    "APPCATALOG#0",
    "COMMUNITY#0",
    "CMSPUBLISHING#0",
    "BLANKINTERNETCONTAINER#0",
    "EDISC#0",
    "EDISC#1",
    "PWA#0",
    "PWS#0",

    # Microsoft 365 group-connected Team Sites
    "GROUP#0"
)

# ===========================================================================
# Thread-safe gedeelde state
# ===========================================================================

# ConcurrentBag in plaats van List: wordt vanuit meerdere runspaces gevuld.
$Results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

# ConcurrentDictionary als thread-safe HashSet voor groep-deduplicatie.
$ProcessedM365Groups = [System.Collections.Concurrent.ConcurrentDictionary[string, byte]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# Vooraf opgehaalde Microsoft 365 Group owners.
$GroupOwnerCache = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

$UseGroupOwnerCache = -not $DisableGroupOwnerCache
$ProcessSubSites = -not $SkipSubSites

# Voortgangsteller over alle threads heen.
$Counter = [hashtable]::Synchronized(@{ Done = 0 })

$ResolvedTargets = [System.Collections.Generic.List[object]]::new()
$BackupOwner = $null

if ($ThrottleLimit -le 0) {
    # Bij daadwerkelijke wijzigingen conservatiever, omdat schrijfacties
    # zwaarder meetellen in de SharePoint Online throttling-berekening.
    $ThrottleLimit = if ($Execute) { 5 } else { 8 }
}

# ===========================================================================
# Gedeelde functies
#
# Deze worden zowel in de hoofdthread als in elke parallelle runspace
# gedefinieerd. Functies uit de hoofdthread zijn NIET zichtbaar binnen
# ForEach-Object -Parallel, vandaar de here-string die per runspace
# gedot-sourcet wordt.
# ===========================================================================

$SharedFunctions = @'
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "OK", "WARN", "ERROR", "ACTION", "DRYRUN")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"

    switch ($Level) {
        "OK"     { $color = "Green" }
        "WARN"   { $color = "Yellow" }
        "ERROR"  { $color = "Red" }
        "ACTION" { $color = "Magenta" }
        "DRYRUN" { $color = "Cyan" }
        default  { $color = "Gray" }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Add-Result {
    param(
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$Template,
        [string]$ObjectType,
        [string]$ObjectName,
        [string]$TargetDisplayName,
        [string]$TargetUPN,
        [string]$TargetObjectId,
        [string]$DetectedAs,
        [string]$Action,
        [string]$Status,
        [string]$Details
    )

    $Results.Add(
        [pscustomobject]@{
            Timestamp         = (Get-Date).ToString("s")
            Mode              = if ($Execute) { "EXECUTE" } else { "DRY-RUN" }
            SiteUrl           = $SiteUrl
            WebUrl            = $WebUrl
            Template          = $Template
            ObjectType        = $ObjectType
            ObjectName        = $ObjectName
            TargetDisplayName = $TargetDisplayName
            TargetUPN         = $TargetUPN
            TargetObjectId    = $TargetObjectId
            DetectedAs        = $DetectedAs
            Action            = $Action
            Status            = $Status
            Details           = $Details
        }
    )
}

function New-AppOnlyConnection {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    Connect-PnPOnline `
        -Url $Url `
        -ClientId $ClientId `
        -Tenant $TenantDomain `
        -Thumbprint $CertificateThumbprint `
        -ReturnConnection `
        -ErrorAction Stop
}

function Get-AdminConnection {
    # De admin-connectie wordt per runspace eenmalig opgezet en hergebruikt.
    # Runspaces worden door ForEach-Object -Parallel hergebruikt over meerdere
    # items heen, dus dit scheelt een authenticatie per site.
    $cached = Get-Variable -Name "PnPAdminConnectionCache" -Scope Global -ErrorAction SilentlyContinue

    if ($cached -and $cached.Value) {
        return $cached.Value
    }

    $connection = New-AppOnlyConnection -Url $TenantAdminUrl
    Set-Variable -Name "PnPAdminConnectionCache" -Scope Global -Value $connection

    $connection
}

function Get-NextSiteIndex {
    [System.Threading.Monitor]::Enter($Counter.SyncRoot)

    try {
        $Counter.Done = $Counter.Done + 1
        return $Counter.Done
    }
    finally {
        [System.Threading.Monitor]::Exit($Counter.SyncRoot)
    }
}

function Escape-ODataString {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $Value.Replace("'", "''")
}

function Get-UniqueUsersById {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [object[]]$Users
    )

    if (-not $Users -or $Users.Count -eq 0) {
        return @()
    }

    $byId = @{}

    foreach ($user in $Users) {
        if (-not $user) {
            continue
        }

        $id = [string]$user.Id

        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $byId[$id] = $user
        }
    }

    @($byId.Values)
}

# ---------------------------------------------------------------------------
# Identity matching
# ---------------------------------------------------------------------------

function Get-NormalizedSharePointIdentity {
    param(
        [Parameter(Mandatory)]
        $Principal
    )

    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($propertyName in @("Email", "LoginName", "UserPrincipalName", "Title")) {
        # CSOM-objecten gooien een uitzondering wanneer een property niet is
        # geladen. Dat mag de matching niet onderbreken.
        $value = $null

        try {
            if ($Principal.PSObject.Properties.Name -contains $propertyName) {
                $value = [string]$Principal.$propertyName
            }
        }
        catch {
            $value = $null
        }

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            [void]$values.Add($value.Trim())

            # Claims login zoals:
            # i:0#.f|membership|user@domain
            if ($value.Contains("|")) {
                $lastPart = ($value -split "\|")[-1]

                if (-not [string]::IsNullOrWhiteSpace($lastPart)) {
                    [void]$values.Add($lastPart.Trim())
                }
            }
        }
    }

    @($values)
}

function Test-SharePointPrincipalMatchesUser {
    param(
        [Parameter(Mandatory)]
        $Principal,

        [Parameter(Mandatory)]
        $User
    )

    $principalValues = @(Get-NormalizedSharePointIdentity -Principal $Principal)

    $userValues = @(
        [string]$User.UserPrincipalName,
        [string]$User.Mail,
        [string]$User.RequestedValue
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($principalValue in $principalValues) {
        foreach ($userValue in $userValues) {
            if ($principalValue.Equals($userValue, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }

    return $false
}

function Find-TargetForSharePointPrincipal {
    param(
        [Parameter(Mandatory)]
        $Principal
    )

    foreach ($target in $ResolvedTargets) {
        if (Test-SharePointPrincipalMatchesUser -Principal $Principal -User $target) {
            return $target
        }
    }

    return $null
}

function Test-IsBackupSharePointPrincipal {
    param(
        [Parameter(Mandatory)]
        $Principal
    )

    Test-SharePointPrincipalMatchesUser -Principal $Principal -User $BackupOwner
}

function Find-TargetForEntraOwner {
    param(
        [Parameter(Mandatory)]
        $Owner
    )

    $ownerId = [string]$Owner.Id

    foreach ($target in $ResolvedTargets) {
        if (
            -not [string]::IsNullOrWhiteSpace($ownerId) -and
            -not [string]::IsNullOrWhiteSpace([string]$target.ObjectId) -and
            $ownerId.Equals([string]$target.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            return $target
        }
    }

    return $null
}

function Test-IsBackupEntraOwner {
    param(
        [Parameter(Mandatory)]
        $Owner
    )

    $ownerId = [string]$Owner.Id

    if ([string]::IsNullOrWhiteSpace($ownerId)) {
        return $false
    }

    $ownerId.Equals($BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
}

# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------

function Get-WebTemplateName {
    param(
        [Parameter(Mandatory)]
        $Web
    )

    $webTemplate = [string]$Web.WebTemplate
    $configuration = [string]$Web.Configuration

    if ([string]::IsNullOrWhiteSpace($webTemplate)) {
        return $null
    }

    "$webTemplate#$configuration"
}

function Test-TemplateIncluded {
    param(
        [string]$Template
    )

    if ([string]::IsNullOrWhiteSpace($Template)) {
        return $false
    }

    $IncludedTemplates -contains $Template.ToUpperInvariant()
}

# ---------------------------------------------------------------------------
# Microsoft 365 Group owner helpers
# ---------------------------------------------------------------------------

function Get-M365GroupOwnerList {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    # Eerst de vooraf opgehaalde cache. Alleen wanneer die geen bruikbaar
    # antwoord heeft, valt het script terug op een individuele Graph-call.
    if ($UseGroupOwnerCache) {
        $cached = $null

        if ($GroupOwnerCache.TryGetValue($GroupId, [ref]$cached)) {
            return @($cached)
        }
    }

    @(
        Get-PnPEntraIDGroupOwner `
            -Identity $GroupId `
            -Connection $Connection `
            -ErrorAction Stop
    )
}

function Add-M365GroupOwner {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$UserObjectId,

        [Parameter(Mandatory)]
        $Connection
    )

    $graphUrl = "groups/$GroupId/owners/`$ref"

    $body = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserObjectId"
    }

    Invoke-PnPGraphMethod `
        -Url $graphUrl `
        -Method Post `
        -Content $body `
        -Connection $Connection `
        -ErrorAction Stop | Out-Null
}

function Remove-M365GroupOwner {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$UserObjectId,

        [Parameter(Mandatory)]
        $Connection
    )

    $graphUrl = "groups/$GroupId/owners/$UserObjectId/`$ref"

    Invoke-PnPGraphMethod `
        -Url $graphUrl `
        -Method Delete `
        -Connection $Connection `
        -ErrorAction Stop | Out-Null
}

function Process-M365GroupOwners {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$Template,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    # TryAdd geeft $false wanneer de groep al door een andere thread is opgepakt.
    if (-not $ProcessedM365Groups.TryAdd($GroupId, [byte]1)) {
        return
    }

    Write-Log "Microsoft 365-groep controleren: $GroupId" "INFO"

    try {
        $owners = @(Get-M365GroupOwnerList -GroupId $GroupId -Connection $Connection)
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Template $Template `
            -ObjectType "Microsoft365Group" `
            -ObjectName $GroupId `
            -DetectedAs "GroupOwnerScan" `
            -Action "ReadOwners" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Owners van Microsoft 365-groep $GroupId konden niet worden gelezen: $($_.Exception.Message)" "ERROR"
        return
    }

    $targetMatches = @()
    $backupOwnerEntry = $null
    $survivingOtherOwners = @()

    foreach ($owner in $owners) {
        $target = Find-TargetForEntraOwner -Owner $owner

        if ($target) {
            $targetMatches += [pscustomobject]@{
                Owner  = $owner
                Target = $target
            }
            continue
        }

        if (Test-IsBackupEntraOwner -Owner $owner) {
            $backupOwnerEntry = $owner
            continue
        }

        $survivingOtherOwners += $owner
    }

    $mustAddBackup = ($targetMatches.Count -gt 0 -and $survivingOtherOwners.Count -eq 0 -and -not $backupOwnerEntry)
    $mustRemoveBackup = ($backupOwnerEntry -and $survivingOtherOwners.Count -gt 0)

    # -----------------------
    # Backup owner toevoegen
    # -----------------------
    if ($mustAddBackup) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup Owner" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou eerst als Microsoft 365 Group owner worden toegevoegd om 0 owners te voorkomen."

            Write-Log "DRY-RUN: backup-owner toevoegen aan M365 groep | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                Add-M365GroupOwner `
                    -GroupId $GroupId `
                    -UserObjectId $BackupOwner.ObjectId `
                    -Connection $Connection

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Owner" `
                    -Action "Add" `
                    -Status "ADDED" `
                    -Details "Backup-owner toegevoegd als Microsoft 365 Group owner voordat targets worden verwijderd."

                Write-Log "Backup-owner toegevoegd aan M365 groep | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Owner" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner toevoegen aan M365 groep mislukt; targets worden op deze groep NIET verwijderd: $SiteUrl | $($_.Exception.Message)" "ERROR"
                return
            }
        }
    }

    # -----------------------
    # Verwijdertargets
    # -----------------------
    foreach ($match in $targetMatches) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $match.Target.DisplayName `
                -TargetUPN $match.Target.UserPrincipalName `
                -TargetObjectId $match.Target.ObjectId `
                -DetectedAs "Microsoft 365 Group Owner" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: verwijdertarget zou uit de gekoppelde Microsoft 365-groep worden verwijderd."

            Write-Log "DRY-RUN: M365 owner verwijderen: $($match.Target.UserPrincipalName) | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                Remove-M365GroupOwner `
                    -GroupId $GroupId `
                    -UserObjectId $match.Target.ObjectId `
                    -Connection $Connection

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $match.Target.DisplayName `
                    -TargetUPN $match.Target.UserPrincipalName `
                    -TargetObjectId $match.Target.ObjectId `
                    -DetectedAs "Microsoft 365 Group Owner" `
                    -Action "Remove" `
                    -Status "REMOVED" `
                    -Details "Owner verwijderd uit gekoppelde Microsoft 365-groep."

                Write-Log "M365 owner verwijderd: $($match.Target.UserPrincipalName) | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $match.Target.DisplayName `
                    -TargetUPN $match.Target.UserPrincipalName `
                    -TargetObjectId $match.Target.ObjectId `
                    -DetectedAs "Microsoft 365 Group Owner" `
                    -Action "Remove" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "M365 owner verwijderen mislukt: $($match.Target.UserPrincipalName) | $SiteUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }

    # -----------------------
    # Overbodige backup-owner verwijderen
    # -----------------------
    if ($mustRemoveBackup) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup Owner" `
                -Action "WouldRemoveBackup" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner is niet nodig omdat minimaal een andere blijvende owner aanwezig is."

            Write-Log "DRY-RUN: overbodige backup-owner verwijderen uit M365 groep | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                Remove-M365GroupOwner `
                    -GroupId $GroupId `
                    -UserObjectId $BackupOwner.ObjectId `
                    -Connection $Connection

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Owner" `
                    -Action "RemoveBackup" `
                    -Status "REMOVED" `
                    -Details "Backup-owner verwijderd omdat een andere blijvende owner aanwezig is."

                Write-Log "Overbodige backup-owner verwijderd uit M365 groep | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Owner" `
                    -Action "RemoveBackup" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner verwijderen uit M365 groep mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Site Collection Administrators
# ---------------------------------------------------------------------------

function Process-SiteCollectionAdmins {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$Template,

        [Parameter(Mandatory)]
        $Connection
    )

    try {
        $admins = @(
            Get-PnPSiteCollectionAdmin `
                -Connection $Connection `
                -ErrorAction Stop
        )
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Template $Template `
            -ObjectType "SiteCollection" `
            -ObjectName $SiteUrl `
            -DetectedAs "SiteCollectionAdminScan" `
            -Action "ReadAdmins" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Site Collection Admins konden niet worden gelezen: $SiteUrl | $($_.Exception.Message)" "ERROR"
        return
    }

    $targetMatches = @()
    $backupAdmin = $null
    $survivingOtherAdmins = @()

    foreach ($admin in $admins) {
        $target = Find-TargetForSharePointPrincipal -Principal $admin

        if ($target) {
            $targetMatches += [pscustomobject]@{
                Principal = $admin
                Target    = $target
            }
            continue
        }

        if (Test-IsBackupSharePointPrincipal -Principal $admin) {
            $backupAdmin = $admin
            continue
        }

        $survivingOtherAdmins += $admin
    }

    $mustAddBackup = ($targetMatches.Count -gt 0 -and $survivingOtherAdmins.Count -eq 0 -and -not $backupAdmin)
    $mustRemoveBackup = ($backupAdmin -and $survivingOtherAdmins.Count -gt 0)

    if ($mustAddBackup) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "SiteCollectionAdmin" `
                -ObjectName $SiteUrl `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup Site Collection Admin" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou eerst als Site Collection Admin worden toegevoegd om 0 admins te voorkomen."

            Write-Log "DRY-RUN: backup Site Collection Admin toevoegen | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                # Via het tenant admin endpoint toevoegen. Dit is betrouwbaarder
                # in app-only context dan afhankelijk zijn van een reeds aanwezige
                # Site Collection Admin-context op de doel-site.
                # De admin-connectie wordt per runspace hergebruikt.
                Set-PnPTenantSite `
                    -Identity $SiteUrl `
                    -Owners $BackupOwner.UserPrincipalName `
                    -Connection (Get-AdminConnection) `
                    -ErrorAction Stop

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $SiteUrl `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Site Collection Admin" `
                    -Action "Add" `
                    -Status "ADDED" `
                    -Details "Backup-owner toegevoegd voordat verwijdertargets worden verwijderd."

                Write-Log "Backup Site Collection Admin toegevoegd | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $SiteUrl `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Site Collection Admin" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup Site Collection Admin toevoegen mislukt; targets worden hier NIET verwijderd: $SiteUrl | $($_.Exception.Message)" "ERROR"
                return
            }
        }
    }

    foreach ($match in $targetMatches) {
        $admin = $match.Principal
        $target = $match.Target

        $ownerIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$admin.LoginName)) {
            [string]$admin.LoginName
        }
        else {
            $target.UserPrincipalName
        }

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "SiteCollectionAdmin" `
                -ObjectName $ownerIdentity `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -DetectedAs "Site Collection Administrator" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: Site Collection Admin-recht zou worden verwijderd."

            Write-Log "DRY-RUN: Site Collection Admin verwijderen: $($target.UserPrincipalName) | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                Remove-PnPSiteCollectionAdmin `
                    -Owners $ownerIdentity `
                    -Connection $Connection `
                    -ErrorAction Stop

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $ownerIdentity `
                    -TargetDisplayName $target.DisplayName `
                    -TargetUPN $target.UserPrincipalName `
                    -TargetObjectId $target.ObjectId `
                    -DetectedAs "Site Collection Administrator" `
                    -Action "Remove" `
                    -Status "REMOVED" `
                    -Details "Site Collection Admin-recht verwijderd."

                Write-Log "Site Collection Admin verwijderd: $($target.UserPrincipalName) | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $ownerIdentity `
                    -TargetDisplayName $target.DisplayName `
                    -TargetUPN $target.UserPrincipalName `
                    -TargetObjectId $target.ObjectId `
                    -DetectedAs "Site Collection Administrator" `
                    -Action "Remove" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Site Collection Admin verwijderen mislukt: $($target.UserPrincipalName) | $SiteUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }

    if ($mustRemoveBackup) {
        $backupIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$backupAdmin.LoginName)) {
            [string]$backupAdmin.LoginName
        }
        else {
            $BackupOwner.UserPrincipalName
        }

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Template $Template `
                -ObjectType "SiteCollectionAdmin" `
                -ObjectName $backupIdentity `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup Site Collection Admin" `
                -Action "WouldRemoveBackup" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou als Site Collection Admin worden verwijderd omdat minimaal een andere blijvende admin aanwezig is."

            Write-Log "DRY-RUN: overbodige backup Site Collection Admin verwijderen | $SiteUrl" "DRYRUN"
        }
        else {
            try {
                Remove-PnPSiteCollectionAdmin `
                    -Owners $backupIdentity `
                    -Connection $Connection `
                    -ErrorAction Stop

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $backupIdentity `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Site Collection Admin" `
                    -Action "RemoveBackup" `
                    -Status "REMOVED" `
                    -Details "Backup-owner verwijderd omdat minimaal een andere blijvende Site Collection Admin aanwezig is."

                Write-Log "Overbodige backup Site Collection Admin verwijderd | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Template $Template `
                    -ObjectType "SiteCollectionAdmin" `
                    -ObjectName $backupIdentity `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup Site Collection Admin" `
                    -Action "RemoveBackup" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup Site Collection Admin verwijderen mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# SharePoint Associated Owners Group
#
# Werkt volledig op de CSOM-context van de site collection. Daardoor is er
# geen aparte authenticatie per subsite meer nodig.
# ---------------------------------------------------------------------------

function Process-SharePointOwnersGroup {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$WebUrl,

        [Parameter(Mandatory)]
        [string]$Template,

        [Parameter(Mandatory)]
        $Web
    )

    $context = $Web.Context

    try {
        $ownersGroup = $Web.AssociatedOwnerGroup

        $context.Load($ownersGroup)
        $context.Load($ownersGroup.Users)
        $context.ExecuteQuery()

        if ($null -eq $ownersGroup -or $ownersGroup.ServerObjectIsNull -eq $true) {
            return
        }

        $members = @($ownersGroup.Users)
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $WebUrl `
            -Template $Template `
            -ObjectType "SharePointOwnersGroup" `
            -ObjectName "Associated Owners Group" `
            -DetectedAs "OwnersGroupScan" `
            -Action "ReadMembers" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Owners-groep kon niet worden gelezen: $WebUrl | $($_.Exception.Message)" "ERROR"
        return
    }

    $targetMatches = @()
    $backupMember = $null
    $survivingOtherMembers = @()

    foreach ($member in $members) {
        $target = Find-TargetForSharePointPrincipal -Principal $member

        if ($target) {
            $targetMatches += [pscustomobject]@{
                Principal = $member
                Target    = $target
            }
            continue
        }

        if (Test-IsBackupSharePointPrincipal -Principal $member) {
            $backupMember = $member
            continue
        }

        $survivingOtherMembers += $member
    }

    $groupName = [string]$ownersGroup.Title
    $mustAddBackup = ($targetMatches.Count -gt 0 -and $survivingOtherMembers.Count -eq 0 -and -not $backupMember)
    $mustRemoveBackup = ($backupMember -and $survivingOtherMembers.Count -gt 0)

    if ($mustAddBackup) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Template $Template `
                -ObjectType "SharePointOwnersGroupMember" `
                -ObjectName $groupName `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup SharePoint Owner" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou eerst aan SharePoint Owners-groep '$groupName' worden toegevoegd om 0 owners te voorkomen."

            Write-Log "DRY-RUN: backup-owner toevoegen aan '$groupName' | $WebUrl" "DRYRUN"
        }
        else {
            try {
                $backupUser = $Web.EnsureUser($BackupOwner.UserPrincipalName)
                $ownersGroup.Users.AddUser($backupUser) | Out-Null
                $context.ExecuteQuery()

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup SharePoint Owner" `
                    -Action "Add" `
                    -Status "ADDED" `
                    -Details "Backup-owner toegevoegd aan SharePoint Owners-groep '$groupName' voordat targets worden verwijderd."

                Write-Log "Backup-owner toegevoegd aan '$groupName' | $WebUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup SharePoint Owner" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner toevoegen aan '$groupName' mislukt; targets worden hier NIET verwijderd: $WebUrl | $($_.Exception.Message)" "ERROR"
                return
            }
        }
    }

    foreach ($match in $targetMatches) {
        $member = $match.Principal
        $target = $match.Target
        $loginName = [string]$member.LoginName

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Template $Template `
                -ObjectType "SharePointOwnersGroupMember" `
                -ObjectName $groupName `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -DetectedAs "SharePoint Associated Owners Group" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: gebruiker zou uit SharePoint Owners-groep '$groupName' worden verwijderd."

            Write-Log "DRY-RUN: uit Owners-groep verwijderen: $($target.UserPrincipalName) | $WebUrl | $groupName" "DRYRUN"
        }
        else {
            try {
                $ownersGroup.Users.RemoveByLoginName($loginName)
                $context.ExecuteQuery()

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $target.DisplayName `
                    -TargetUPN $target.UserPrincipalName `
                    -TargetObjectId $target.ObjectId `
                    -DetectedAs "SharePoint Associated Owners Group" `
                    -Action "Remove" `
                    -Status "REMOVED" `
                    -Details "Gebruiker verwijderd uit SharePoint Owners-groep '$groupName'."

                Write-Log "Uit Owners-groep verwijderd: $($target.UserPrincipalName) | $WebUrl | $groupName" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $target.DisplayName `
                    -TargetUPN $target.UserPrincipalName `
                    -TargetObjectId $target.ObjectId `
                    -DetectedAs "SharePoint Associated Owners Group" `
                    -Action "Remove" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Verwijderen uit Owners-groep mislukt: $($target.UserPrincipalName) | $WebUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }

    if ($mustRemoveBackup) {
        $backupLoginName = [string]$backupMember.LoginName

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Template $Template `
                -ObjectType "SharePointOwnersGroupMember" `
                -ObjectName $groupName `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -DetectedAs "Backup SharePoint Owner" `
                -Action "WouldRemoveBackup" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou uit SharePoint Owners-groep '$groupName' worden verwijderd omdat minimaal een andere blijvende owner aanwezig is."

            Write-Log "DRY-RUN: overbodige backup-owner verwijderen uit '$groupName' | $WebUrl" "DRYRUN"
        }
        else {
            try {
                $ownersGroup.Users.RemoveByLoginName($backupLoginName)
                $context.ExecuteQuery()

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup SharePoint Owner" `
                    -Action "RemoveBackup" `
                    -Status "REMOVED" `
                    -Details "Backup-owner verwijderd uit SharePoint Owners-groep '$groupName' omdat een andere blijvende owner aanwezig is."

                Write-Log "Overbodige backup-owner verwijderd uit '$groupName' | $WebUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Template $Template `
                    -ObjectType "SharePointOwnersGroupMember" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -DetectedAs "Backup SharePoint Owner" `
                    -Action "RemoveBackup" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner verwijderen uit '$groupName' mislukt: $WebUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Verwerking van een volledige site collection
# ---------------------------------------------------------------------------

function Invoke-SiteProcessing {
    param(
        [Parameter(Mandatory)]
        $Site
    )

    $siteUrl = [string]$Site.Url
    $siteTemplate = ([string]$Site.Template).ToUpperInvariant()
    $index = Get-NextSiteIndex

    Write-Log "[$index/$TotalSites] Site verwerken: $siteUrl [$siteTemplate]" "INFO"

    try {
        $siteConnection = New-AppOnlyConnection -Url $siteUrl
    }
    catch {
        Add-Result `
            -SiteUrl $siteUrl `
            -WebUrl $siteUrl `
            -Template $siteTemplate `
            -ObjectType "SiteCollection" `
            -ObjectName $siteUrl `
            -DetectedAs "Connection" `
            -Action "Connect" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Verbinding met site mislukt: $siteUrl | $($_.Exception.Message)" "ERROR"
        return
    }

    # 1. Site Collection Administrators
    Process-SiteCollectionAdmins `
        -SiteUrl $siteUrl `
        -Template $siteTemplate `
        -Connection $siteConnection

    # 2. Root SharePoint Associated Owners Group
    try {
        $context = Get-PnPContext -Connection $siteConnection
        $rootWeb = $context.Web

        Process-SharePointOwnersGroup `
            -SiteUrl $siteUrl `
            -WebUrl $siteUrl `
            -Template $siteTemplate `
            -Web $rootWeb
    }
    catch {
        Add-Result `
            -SiteUrl $siteUrl `
            -WebUrl $siteUrl `
            -Template $siteTemplate `
            -ObjectType "SharePointOwnersGroup" `
            -ObjectName "Associated Owners Group" `
            -DetectedAs "OwnersGroupScan" `
            -Action "ReadMembers" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Root owners-groep kon niet worden verwerkt: $siteUrl | $($_.Exception.Message)" "ERROR"
    }

    # 3. Microsoft 365 Group owners indien group-connected
    $groupId = $null

    if ($Site.PSObject.Properties.Name -contains "GroupId") {
        $groupId = [string]$Site.GroupId
    }

    if (
        -not [string]::IsNullOrWhiteSpace($groupId) -and
        $groupId -ne "00000000-0000-0000-0000-000000000000"
    ) {
        Process-M365GroupOwners `
            -SiteUrl $siteUrl `
            -Template $siteTemplate `
            -GroupId $groupId `
            -Connection $siteConnection
    }

    # 4. Subsites recursief, op dezelfde connectie
    if (-not $ProcessSubSites) {
        return
    }

    try {
        $subWebs = @(
            Get-PnPSubWeb `
                -Recurse `
                -Includes WebTemplate,Configuration `
                -Connection $siteConnection `
                -ErrorAction Stop
        )
    }
    catch {
        Add-Result `
            -SiteUrl $siteUrl `
            -WebUrl $siteUrl `
            -Template $siteTemplate `
            -ObjectType "SubWebScan" `
            -ObjectName $siteUrl `
            -DetectedAs "SubWebEnumeration" `
            -Action "Read" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Subsites konden niet worden uitgelezen: $siteUrl | $($_.Exception.Message)" "ERROR"
        $subWebs = @()
    }

    foreach ($subWeb in $subWebs) {
        $subWebUrl = [string]$subWeb.Url
        $subWebTemplate = Get-WebTemplateName -Web $subWeb

        if (-not (Test-TemplateIncluded -Template $subWebTemplate)) {
            continue
        }

        Write-Log "Subsite verwerken: $subWebUrl [$subWebTemplate]" "INFO"

        # Geen nieuwe authenticatie: het subweb-object hangt aan dezelfde
        # ClientContext als de site collection.
        Process-SharePointOwnersGroup `
            -SiteUrl $siteUrl `
            -WebUrl $subWebUrl `
            -Template $subWebTemplate `
            -Web $subWeb
    }
}
'@

# Functies ook in de hoofdthread beschikbaar maken.
. ([scriptblock]::Create($SharedFunctions))

# ===========================================================================
# Entra ID identity-resolutie (blijft serieel; slechts zes lookups)
# ===========================================================================

function Resolve-EntraUserDefinition {
    param(
        [Parameter(Mandatory)]
        $Definition,

        [Parameter(Mandatory)]
        $Connection,

        [switch]$AllowUnresolved
    )

    $value = [string]$Definition.SearchValue
    $escapedValue = Escape-ODataString -Value $value
    $matches = @()

    switch ($Definition.SearchType) {
        "UPNOrMail" {
            # Eerst exact als UPN proberen. Als dit lukt is een aanvullende
            # mail-filter niet nodig en voorkomen we problemen met speciale
            # tekens zoals #EXT# in guest-UPN's.
            $direct = $null
            $directLookupError = $null

            try {
                $direct = Get-PnPEntraIDUser `
                    -Identity $value `
                    -Connection $Connection `
                    -ErrorAction Stop
            }
            catch {
                $directLookupError = $_.Exception.Message
            }

            if ($direct) {
                $matches += $direct
            }
            elseif ($value -match '(?i)#EXT#') {
                # Guest UPN's bevatten #EXT#. Een '#' mag niet ongecodeerd in
                # een URL terechtkomen, omdat alles erna dan als URL-fragment
                # kan worden geinterpreteerd. Daarom gebruiken we als fallback
                # Microsoft Graph rechtstreeks met een URL-gecodeerde UPN.
                try {
                    $encodedUpn = [System.Uri]::EscapeDataString($value)
                    $graphUrl = "users/{0}?`$select=id,displayName,userPrincipalName,mail" -f $encodedUpn

                    $graphUser = Invoke-PnPGraphMethod `
                        -Url $graphUrl `
                        -Method Get `
                        -Connection $Connection `
                        -ErrorAction Stop

                    if ($graphUser) {
                        $matches += $graphUser
                    }
                }
                catch {
                    $extra = if ($directLookupError) {
                        " De normale Identity-lookup gaf: $directLookupError"
                    }
                    else {
                        ""
                    }

                    throw "Guest-account '$value' kon niet exact via Microsoft Graph worden opgelost.$extra Graph fout: $($_.Exception.Message)"
                }
            }
            else {
                # Alleen wanneer de waarde GEEN #EXT# guest-UPN is en de
                # exacte Identity-lookup niets oplevert, zoeken we aanvullend
                # op het mail-attribuut.
                try {
                    $mailMatches = @(
                        Get-PnPEntraIDUser `
                            -Filter "mail eq '$escapedValue'" `
                            -Connection $Connection `
                            -ErrorAction Stop
                    )

                    $matches += $mailMatches
                }
                catch {
                    $extra = if ($directLookupError) {
                        " De normale Identity-lookup gaf: $directLookupError"
                    }
                    else {
                        ""
                    }

                    throw "Zoeken naar mailadres '$value' is mislukt.$extra Filter fout: $($_.Exception.Message)"
                }
            }
        }

        "DisplayName" {
            try {
                $matches = @(
                    Get-PnPEntraIDUser `
                        -Filter "displayName eq '$escapedValue'" `
                        -Connection $Connection `
                        -ErrorAction Stop
                )
            }
            catch {
                throw "Zoeken naar DisplayName '$value' is mislukt: $($_.Exception.Message)"
            }
        }

        "AccountName" {
            # Een username zonder domein wordt veilig opgelost via:
            # 1. exact DisplayName;
            # 2. UPN die begint met "<account>@";
            # 3. exact mailNickname.
            try {
                $displayMatches = @(
                    Get-PnPEntraIDUser `
                        -Filter "displayName eq '$escapedValue'" `
                        -Connection $Connection `
                        -ErrorAction Stop
                )
                $matches += $displayMatches
            }
            catch {
                throw "Zoeken naar account '$value' op DisplayName is mislukt: $($_.Exception.Message)"
            }

            try {
                $upnPrefix = Escape-ODataString -Value ("{0}@" -f $value)
                $upnMatches = @(
                    Get-PnPEntraIDUser `
                        -Filter "startswith(userPrincipalName,'$upnPrefix')" `
                        -Connection $Connection `
                        -ErrorAction Stop
                )
                $matches += $upnMatches
            }
            catch {
                throw "Zoeken naar account '$value' op UPN is mislukt: $($_.Exception.Message)"
            }

            try {
                $nickMatches = @(
                    Get-PnPEntraIDUser `
                        -Filter "mailNickname eq '$escapedValue'" `
                        -Select @("Id","DisplayName","UserPrincipalName","Mail","MailNickname","AccountEnabled") `
                        -Connection $Connection `
                        -ErrorAction Stop
                )
                $matches += $nickMatches
            }
            catch {
                # mailNickname filtering kan in sommige tenants/querycombinaties
                # niet beschikbaar zijn. De DisplayName/UPN zoekslagen blijven geldig.
                Write-Log "mailNickname lookup voor '$value' kon niet worden uitgevoerd; DisplayName/UPN resultaten worden gebruikt." "WARN"
            }
        }

        default {
            throw "Onbekend SearchType '$($Definition.SearchType)'."
        }
    }

    $matches = @(Get-UniqueUsersById -Users $matches)

    if ($matches.Count -eq 0) {
        if ($AllowUnresolved) {
            Write-Log "Geen uniek Entra-object gevonden voor '$($Definition.Label)' [$value]. SharePoint wordt wel op deze opgegeven identiteit/naam gecontroleerd; Microsoft 365 Group-ownerwijzigingen voor dit target worden overgeslagen zolang geen ObjectId beschikbaar is." "WARN"

            return [pscustomobject]@{
                Label             = [string]$Definition.Label
                RequestedValue    = $value
                SearchType        = [string]$Definition.SearchType
                ObjectId          = ""
                DisplayName       = if ($Definition.SearchType -eq "DisplayName") { $value } else { [string]$Definition.Label }
                UserPrincipalName = if ($value -match "@") { $value } else { "" }
                Mail              = if ($value -match "@") { $value } else { "" }
                ResolutionStatus  = "UNRESOLVED"
            }
        }

        throw "Geen Entra ID gebruiker gevonden voor '$($Definition.Label)' met zoekwaarde '$value'."
    }

    if ($matches.Count -gt 1) {
        $matchText = ($matches | ForEach-Object {
            "$($_.DisplayName) <$($_.UserPrincipalName)> [$($_.Id)]"
        }) -join "; "

        throw "Meerdere Entra ID gebruikers gevonden voor '$($Definition.Label)' ('$value'). Geen wijzigingen uitgevoerd. Matches: $matchText"
    }

    $user = $matches[0]

    [pscustomobject]@{
        Label             = [string]$Definition.Label
        RequestedValue    = $value
        SearchType        = [string]$Definition.SearchType
        ObjectId          = [string]$user.Id
        DisplayName       = [string]$user.DisplayName
        UserPrincipalName = [string]$user.UserPrincipalName
        Mail              = [string]$user.Mail
        ResolutionStatus  = "RESOLVED"
    }
}

function Resolve-AllIdentities {
    param(
        [Parameter(Mandatory)]
        $Connection
    )

    foreach ($definition in $TargetDefinitions) {
        Write-Log "Verwijdertarget oplossen: $($definition.Label) [$($definition.SearchValue)]" "INFO"

        $resolved = Resolve-EntraUserDefinition `
            -Definition $definition `
            -Connection $Connection `
            -AllowUnresolved

        $ResolvedTargets.Add($resolved)

        if ($resolved.ResolutionStatus -eq "RESOLVED") {
            Write-Log "Target gevonden: $($resolved.DisplayName) <$($resolved.UserPrincipalName)> [$($resolved.ObjectId)]" "OK"
        }
        else {
            Write-Log "Target niet in Entra opgelost; SharePoint-scan gaat door op opgegeven waarde: $($resolved.RequestedValue)" "WARN"
        }
    }

    Write-Log "Backup-owner oplossen: $($BackupOwnerDefinition.SearchValue)" "INFO"

    $script:BackupOwner = Resolve-EntraUserDefinition `
        -Definition $BackupOwnerDefinition `
        -Connection $Connection

    Write-Log "Backup-owner gevonden: $($BackupOwner.DisplayName) <$($BackupOwner.UserPrincipalName)> [$($BackupOwner.ObjectId)]" "OK"

    # Beveiliging: backup-owner mag nooit tevens een verwijdertarget zijn.
    foreach ($target in $ResolvedTargets) {
        if (
            -not [string]::IsNullOrWhiteSpace([string]$target.ObjectId) -and
            ([string]$target.ObjectId).Equals([string]$BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            throw "Backup-owner resolveert naar hetzelfde object als verwijdertarget '$($target.Label)'. Script afgebroken."
        }
    }
}

# ===========================================================================
# Microsoft 365 Group owner prefetch
# ===========================================================================

function Initialize-GroupOwnerCache {
    param(
        [Parameter(Mandatory)]
        $Connection
    )

    Write-Log "Microsoft 365 Group owners vooraf ophalen via Microsoft Graph..." "INFO"

    $truncatedGroupIds = [System.Collections.Generic.List[string]]::new()
    $pageCount = 0
    $url = 'groups?$select=id&$expand=owners&$top=999'

    try {
        do {
            $page = Invoke-PnPGraphMethod `
                -Url $url `
                -Method Get `
                -Connection $Connection `
                -ErrorAction Stop

            $pageCount++

            if ($page.PSObject.Properties.Name -contains "value") {
                foreach ($group in $page.value) {
                    $groupId = [string]$group.id

                    if ([string]::IsNullOrWhiteSpace($groupId)) {
                        continue
                    }

                    $owners = @()

                    if ($group.PSObject.Properties.Name -contains "owners" -and $group.owners) {
                        $owners = @($group.owners)
                    }

                    # Microsoft Graph geeft bij $expand maximaal 20 gerelateerde
                    # objecten terug. Bij exact 20 owners is het resultaat
                    # mogelijk afgekapt en halen we de groep alsnog los op.
                    if ($owners.Count -ge 20) {
                        $truncatedGroupIds.Add($groupId)
                        continue
                    }

                    [void]$GroupOwnerCache.TryAdd($groupId, $owners)
                }
            }

            $url = $null

            if ($page.PSObject.Properties.Name -contains "@odata.nextLink") {
                $url = [string]$page.'@odata.nextLink'
            }
        } while (-not [string]::IsNullOrWhiteSpace($url))
    }
    catch {
        Write-Log "Vooraf ophalen van group owners is mislukt; het script valt terug op een losse Graph-call per groep. Fout: $($_.Exception.Message)" "WARN"
        $script:UseGroupOwnerCache = $false
        return
    }

    foreach ($groupId in $truncatedGroupIds) {
        try {
            $owners = @(
                Get-PnPEntraIDGroupOwner `
                    -Identity $groupId `
                    -Connection $Connection `
                    -ErrorAction Stop
            )

            [void]$GroupOwnerCache.TryAdd($groupId, $owners)
        }
        catch {
            Write-Log "Owners van groep $groupId konden niet volledig worden opgehaald; deze groep wordt tijdens de scan individueel benaderd. Fout: $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "Group owner cache gevuld: $($GroupOwnerCache.Count) groepen in $pageCount pagina('s), waarvan $($truncatedGroupIds.Count) individueel nagevraagd." "OK"
}

# ===========================================================================
# Rapportage
# ===========================================================================

function Export-Results {
    $directory = Split-Path -Path $ReportPath -Parent

    if (
        -not [string]::IsNullOrWhiteSpace($directory) -and
        -not (Test-Path -LiteralPath $directory)
    ) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    if ($Results.Count -gt 0) {
        $Results |
            Sort-Object SiteUrl, WebUrl, ObjectType, TargetUPN, Action |
            Export-Csv -Path $ReportPath -NoTypeInformation -Encoding utf8BOM
    }
    else {
        [pscustomobject]@{
            Timestamp         = (Get-Date).ToString("s")
            Mode              = if ($Execute) { "EXECUTE" } else { "DRY-RUN" }
            SiteUrl           = ""
            WebUrl            = ""
            Template          = ""
            ObjectType        = ""
            ObjectName        = ""
            TargetDisplayName = ""
            TargetUPN         = ""
            TargetObjectId    = ""
            DetectedAs        = ""
            Action            = ""
            Status            = "NO_MATCHES"
            Details           = "Geen acties of matching owner-relaties gevonden."
        } | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding utf8BOM
    }
}

# ===========================================================================
# Main
# ===========================================================================

$modeText = if ($Execute) {
    "EXECUTE - WIJZIGINGEN WORDEN UITGEVOERD"
}
else {
    "DRY-RUN - GEEN WIJZIGINGEN"
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host " SharePoint / Microsoft 365 Owner Cleanup" -ForegroundColor Cyan
Write-Host " Mode: $modeText" -ForegroundColor $(if ($Execute) { "Magenta" } else { "Cyan" })
Write-Host " Parallelle threads: $ThrottleLimit" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host ""

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$TotalSites = 0

try {
    if ($TenantAdminUrl -notmatch "^https://.+-admin\.sharepoint\.com/?$") {
        throw "TenantAdminUrl lijkt ongeldig: '$TenantAdminUrl'. Verwacht bijvoorbeeld https://contoso-admin.sharepoint.com"
    }

    Write-Log "Verbinden met SharePoint Admin Center: $TenantAdminUrl" "INFO"

    $adminConnection = New-AppOnlyConnection -Url $TenantAdminUrl

    Write-Log "Tenant-ID valideren..." "INFO"

    $connectedTenantId = [string](
        Get-PnPTenantId `
            -Connection $adminConnection `
            -ErrorAction Stop
    )

    if (-not $connectedTenantId.Equals($TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Tenant mismatch. Verwacht TenantId '$TenantId', maar verbinding hoort bij '$connectedTenantId'."
    }

    Write-Log "Tenant-ID klopt: $connectedTenantId" "OK"

    Resolve-AllIdentities -Connection $adminConnection

    if ($UseGroupOwnerCache) {
        Initialize-GroupOwnerCache -Connection $adminConnection
    }
    else {
        Write-Log "Group owner cache uitgeschakeld via -DisableGroupOwnerCache; owners worden per groep opgehaald." "WARN"
    }

    Write-Log "SharePoint site collections ophalen..." "INFO"

    # -Detailed is bewust weggelaten: alleen Url, Template en GroupId zijn nodig.
    $allSites = if ($DetailedSiteEnumeration) {
        @(Get-PnPTenantSite -Detailed -Connection $adminConnection -ErrorAction Stop)
    }
    else {
        @(Get-PnPTenantSite -Connection $adminConnection -ErrorAction Stop)
    }

    $sites = @(
        $allSites |
            Where-Object {
                $_.Url -and
                $_.Url -notmatch "-my\.sharepoint\.com/personal/" -and
                (Test-TemplateIncluded -Template ([string]$_.Template))
            } |
            Sort-Object Url |
            ForEach-Object {
                $groupId = ""

                if ($_.PSObject.Properties.Name -contains "GroupId") {
                    $groupId = [string]$_.GroupId
                }

                if ($groupId -eq "00000000-0000-0000-0000-000000000000") {
                    $groupId = ""
                }

                # Bewust omzetten naar een eenvoudig object. Dat scheelt
                # serialisatie richting de runspaces en maakt het mogelijk
                # om GroupId hieronder alsnog aan te vullen.
                [pscustomobject]@{
                    Url      = [string]$_.Url
                    Template = ([string]$_.Template).ToUpperInvariant()
                    GroupId  = $groupId
                }
            }
    )

    $TotalSites = $sites.Count

    Write-Log "Site collections binnen opgegeven templates: $TotalSites" "OK"

    # Zonder -Detailed is GroupId in vrijwel alle gevallen gevuld, maar niet
    # gegarandeerd. De uitzonderingen worden hier individueel nagevraagd,
    # zodat Microsoft 365 Group owners niet stilzwijgend worden overgeslagen.
    if (-not $DetailedSiteEnumeration) {
        $groupSitesWithoutId = @(
            $sites | Where-Object {
                $_.Template -eq "GROUP#0" -and
                [string]::IsNullOrWhiteSpace($_.GroupId)
            }
        )

        if ($groupSitesWithoutId.Count -gt 0) {
            Write-Log "Van $($groupSitesWithoutId.Count) group-connected site(s) ontbreekt GroupId; deze worden individueel nagevraagd." "WARN"

            foreach ($siteEntry in $groupSitesWithoutId) {
                try {
                    $detail = Get-PnPTenantSite `
                        -Identity $siteEntry.Url `
                        -Detailed `
                        -Connection $adminConnection `
                        -ErrorAction Stop

                    $detailGroupId = ""

                    if ($detail -and $detail.PSObject.Properties.Name -contains "GroupId") {
                        $detailGroupId = [string]$detail.GroupId
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace($detailGroupId) -and
                        $detailGroupId -ne "00000000-0000-0000-0000-000000000000"
                    ) {
                        $siteEntry.GroupId = $detailGroupId

                        Write-Log "GroupId aangevuld voor $($siteEntry.Url): $detailGroupId" "OK"
                    }
                    else {
                        # Komt voor bij sites waarvan de gekoppelde groep is
                        # verwijderd of nog niet volledig is uitgerold. De site
                        # zelf wordt gewoon gescand; alleen de groep-ownercheck
                        # vervalt.
                        Add-Result `
                            -SiteUrl $siteEntry.Url `
                            -WebUrl $siteEntry.Url `
                            -Template $siteEntry.Template `
                            -ObjectType "Microsoft365Group" `
                            -ObjectName $siteEntry.Url `
                            -DetectedAs "GroupIdLookup" `
                            -Action "Read" `
                            -Status "WARN" `
                            -Details "Site heeft template GROUP#0 maar geen gekoppelde GroupId. Site Collection Admins en de SharePoint Owners-groep zijn wel gecontroleerd; de Microsoft 365 Group-ownercheck is overgeslagen."

                        Write-Log "Geen GroupId gevonden voor $($siteEntry.Url); alleen de SharePoint-scopes worden gecontroleerd." "WARN"
                    }
                }
                catch {
                    Add-Result `
                        -SiteUrl $siteEntry.Url `
                        -WebUrl $siteEntry.Url `
                        -Template $siteEntry.Template `
                        -ObjectType "Microsoft365Group" `
                        -ObjectName $siteEntry.Url `
                        -DetectedAs "GroupIdLookup" `
                        -Action "Read" `
                        -Status "ERROR" `
                        -Details $_.Exception.Message

                    Write-Log "GroupId opvragen mislukt voor $($siteEntry.Url): $($_.Exception.Message)" "ERROR"
                }
            }
        }
    }

    if ($TotalSites -gt 0) {
        Write-Log "Start parallelle verwerking met $ThrottleLimit threads..." "INFO"

        # Waarden die in elke runspace nodig zijn.
        $ResolvedTargetsArray = @($ResolvedTargets)
        $BackupOwnerResolved = $BackupOwner
        $ExecuteMode = [bool]$Execute
        $SharedFunctionText = $SharedFunctions

        $sites | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            if (-not (Get-Module -Name PnP.PowerShell)) {
                Import-Module PnP.PowerShell -DisableNameChecking -ErrorAction Stop
            }

            Set-StrictMode -Version Latest
            $ErrorActionPreference = "Stop"

            # Context voor de gedeelde functies.
            $ClientId              = $using:ClientId
            $TenantDomain          = $using:TenantDomain
            $CertificateThumbprint = $using:CertificateThumbprint
            $TenantAdminUrl        = $using:TenantAdminUrl
            $IncludedTemplates     = $using:IncludedTemplates
            $Execute               = $using:ExecuteMode
            $Results               = $using:Results
            $ResolvedTargets       = $using:ResolvedTargetsArray
            $BackupOwner           = $using:BackupOwnerResolved
            $ProcessedM365Groups   = $using:ProcessedM365Groups
            $GroupOwnerCache       = $using:GroupOwnerCache
            $UseGroupOwnerCache    = $using:UseGroupOwnerCache
            $ProcessSubSites       = $using:ProcessSubSites
            $Counter               = $using:Counter
            $TotalSites            = $using:TotalSites

            . ([scriptblock]::Create($using:SharedFunctionText))

            # $_ verwijst binnen een catch-blok naar de errorrecord, daarom
            # wordt de site hier eerst vastgelegd.
            $currentSite = $_
            $currentSiteUrl = [string]$currentSite.Url

            try {
                Invoke-SiteProcessing -Site $currentSite
            }
            catch {
                Add-Result `
                    -SiteUrl $currentSiteUrl `
                    -WebUrl $currentSiteUrl `
                    -ObjectType "SiteCollection" `
                    -ObjectName $currentSiteUrl `
                    -DetectedAs "SiteProcessing" `
                    -Action "Process" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Onverwachte fout bij siteverwerking: $currentSiteUrl | $($_.Exception.Message)" "ERROR"
            }
        }
    }
}
catch {
    Write-Log $_.Exception.Message "ERROR"

    Add-Result `
        -ObjectType "Fatal" `
        -ObjectName "Script" `
        -DetectedAs "FatalError" `
        -Action "Abort" `
        -Status "ERROR" `
        -Details $_.Exception.Message
}
finally {
    $stopwatch.Stop()

    try {
        Export-Results
        Write-Host ""
        Write-Log "Rapport opgeslagen: $ReportPath" "OK"
    }
    catch {
        Write-Log "Rapport kon niet worden opgeslagen: $($_.Exception.Message)" "ERROR"
    }

    $foundCount = @($Results | Where-Object { $_.Status -eq "FOUND" }).Count
    $addedCount = @($Results | Where-Object { $_.Status -eq "ADDED" }).Count
    $removedCount = @($Results | Where-Object { $_.Status -eq "REMOVED" }).Count
    $errorCount = @($Results | Where-Object { $_.Status -eq "ERROR" }).Count

    Write-Host ""
    Write-Host "Samenvatting" -ForegroundColor Cyan
    Write-Host "-----------"
    Write-Host "Mode        : $(if ($Execute) { 'EXECUTE' } else { 'DRY-RUN' })"
    Write-Host "Threads     : $ThrottleLimit"
    Write-Host "Sites       : $TotalSites"
    Write-Host "M365 groepen: $($ProcessedM365Groups.Count)"
    Write-Host "Dry-run hits: $foundCount"
    Write-Host "Added       : $addedCount"
    Write-Host "Removed     : $removedCount"
    Write-Host "Errors      : $errorCount"
    Write-Host "Doorlooptijd: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
    Write-Host "Report      : $ReportPath"
    Write-Host ""
}
