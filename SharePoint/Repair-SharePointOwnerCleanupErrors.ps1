#Requires -Version 7.4
#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Herstelt de openstaande fouten uit een run van
    Remove-SharePointOwnersTenantWide.ps1.

.DESCRIPTION
    Dit script leest het CSV-rapport van de hoofdrun, selecteert de sites met
    status ERROR of WARN en verwerkt uitsluitend die sites opnieuw. Het draait
    bewust SERIEEL, met retries, omdat een deel van de fouten hoogstwaarschijnlijk
    veroorzaakt is door gelijktijdigheid.

    Standaard voert dit script GEEN wijzigingen uit. Alleen met -Execute
    worden wijzigingen doorgevoerd.

    FOUTCATEGORIEEN EN AANPAK

    1. "Nullable object must have a value."
       Gezien bij Connect, SubWebEnumeration en OwnersGroupScan.
       Dit is geen eigenschap van de site maar een transiente fout uit
       PnP.PowerShell / CSOM. De aanroepen zijn hier voorzien van retries met
       oplopende wachttijd, en de subsite-enumeratie heeft een CSOM-fallback
       die Get-PnPSubWeb volledig omzeilt.

    2. "The group must have at least one owner, hence this owner cannot be removed."
       Dit treedt op wanneer de te verwijderen gebruiker de laatste owner is.
       De hoofdrun voegde Backup-owner eerst toe, maar Entra ID repliceert die
       toevoeging niet altijd snel genoeg om enkele milliseconden later de
       verwijdering toe te staan.
       Aanpak: owners vers uitlezen, backup toevoegen, actief WACHTEN tot de
       toevoeging bevestigd terugkomt uit Graph, daarna pas verwijderen, en
       targets een voor een verwijderen met een herlezing tussendoor.

    3. "Not Found (404): Resource '<guid>' does not exist"
       De SharePoint-site verwijst naar een Microsoft 365-groep die niet meer
       bestaat. Het script controleert of de groep in de prullenbak van Entra ID
       staat en rapporteert de site als weesobject. Er is geen owner om te
       verwijderen, dus dit is geen actiepunt maar governance-opruimwerk.

    4. "AssociatedOwnerGroup is null"
       De site heeft geen gekoppelde SharePoint Owners-groep. Bij GROUP#0-sites
       is de Microsoft 365-groep leidend, dus dit is doorgaans onschadelijk.
       Het script rapporteert dit en controleert wel de overige scopes.

    UITGANGSPUNT
    Waar het verwijderen van een target ertoe zou leiden dat een scope geen
    owner meer overhoudt, wordt de backup-owner (-BackupOwnerUpn) eerst
    toegevoegd en geverifieerd voordat er iets wordt verwijderd. Lukt de
    toevoeging niet, dan wordt er op die scope NIETS verwijderd.

.PARAMETER TargetAccount
    UPN('s) van de accounts die als owner verwijderd moeten worden. Gebruik
    dezelfde accounts als in de hoofdrun.

.PARAMETER BackupOwnerUpn
    UPN van het account dat owner wordt wanneer een scope anders zonder owner
    zou komen te staan.

.PARAMETER ReportPath
    Pad naar het CSV-rapport van de hoofdrun. Alle regels met status ERROR of
    WARN bepalen welke sites opnieuw worden verwerkt.

.PARAMETER SiteUrl
    Optioneel. Een of meer site-URL's om te verwerken in plaats van het rapport.

.EXAMPLE
    $connectie = @{
        ClientId              = '<app-id>'
        TenantId              = '<tenant-id>'
        TenantDomain          = 'contoso.onmicrosoft.com'
        TenantAdminUrl        = 'https://contoso-admin.sharepoint.com'
        CertificateThumbprint = '<thumbprint>'
        TargetAccount         = 'oud.account@contoso.com'
        BackupOwnerUpn        = 'svc-spo-backup@contoso.com'
    }
    .\Repair-SharePointOwnerCleanupErrors.ps1 @connectie -ReportPath .\SharePoint_Owner_Cleanup_20260810_164500.csv

    Dry-run over alleen de sites met fouten.

.EXAMPLE
    .\Repair-SharePointOwnerCleanupErrors.ps1 @connectie -ReportPath .\SharePoint_Owner_Cleanup_20260810_164500.csv -Execute

    Voert het herstel daadwerkelijk uit.

.EXAMPLE
    .\Repair-SharePointOwnerCleanupErrors.ps1 @connectie -SiteUrl "https://contoso.sharepoint.com/sites/allcompany" -Execute

    Herstelt een enkele site.
#>

[CmdletBinding(DefaultParameterSetName = "FromReport")]
param(
    [Parameter(Mandatory, ParameterSetName = "FromReport")]
    [string]$ReportPath,

    [Parameter(Mandatory, ParameterSetName = "FromUrl")]
    [string[]]$SiteUrl,

    [Parameter(Mandatory)]
    [string[]]$TargetAccount,

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

    [Parameter(Mandatory)]
    [string]$BackupOwnerUpn,

    [ValidateRange(1, 10)]
    [int]$MaxRetries = 4,

    # Maximale tijd om te wachten tot een toegevoegde owner bevestigd
    # terugkomt uit Microsoft Graph.
    [ValidateRange(5, 300)]
    [int]$OwnerConfirmTimeoutSeconds = 90,

    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath ("SharePoint_Owner_Repair_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===========================================================================
# Configuratie
# ===========================================================================

$TargetDefinitions = @(
    $TargetAccount | ForEach-Object {
        [pscustomobject]@{
            Label       = $_
            SearchValue = $_
            SearchType  = "UPNOrMail"
        }
    }
)

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
    "GROUP#0"
)

$Results = [System.Collections.Generic.List[object]]::new()
$ResolvedTargets = [System.Collections.Generic.List[object]]::new()
$BackupOwner = $null
$AdminConnection = $null

# ===========================================================================
# Basisfuncties
# ===========================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "OK", "WARN", "ERROR", "ACTION", "DRYRUN", "RETRY")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"

    switch ($Level) {
        "OK"     { $color = "Green" }
        "WARN"   { $color = "Yellow" }
        "ERROR"  { $color = "Red" }
        "ACTION" { $color = "Magenta" }
        "DRYRUN" { $color = "Cyan" }
        "RETRY"  { $color = "DarkYellow" }
        default  { $color = "Gray" }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Add-Result {
    param(
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$Scope,
        [string]$ObjectName,
        [string]$TargetDisplayName,
        [string]$TargetUPN,
        [string]$TargetObjectId,
        [string]$Category,
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
            Scope             = $Scope
            ObjectName        = $ObjectName
            TargetDisplayName = $TargetDisplayName
            TargetUPN         = $TargetUPN
            TargetObjectId    = $TargetObjectId
            Category          = $Category
            Action            = $Action
            Status            = $Status
            Details           = $Details
        }
    )
}

function Invoke-WithRetry {
    <#
        Voert een scriptblock uit met retries en oplopende wachttijd.
        Bedoeld voor transiente fouten zoals "Nullable object must have a value",
        socket-fouten en HTTP 429/503 vanuit SharePoint Online.
    #>
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$Description,

        [int]$Retries = $MaxRetries
    )

    $attempt = 0
    $lastError = $null

    while ($attempt -lt $Retries) {
        $attempt++

        try {
            return & $ScriptBlock
        }
        catch {
            $lastError = $_
            $message = [string]$_.Exception.Message

            # Fouten die met zekerheid niet transient zijn: direct doorgooien.
            $isPermanent = (
                $message -match "does not exist or one of its queried reference-property objects" -or
                $message -match "must have at least one owner" -or
                $message -match "Access denied" -or
                $message -match "Attempted to perform an unauthorized operation"
            )

            if ($isPermanent -or $attempt -ge $Retries) {
                throw
            }

            $waitSeconds = [Math]::Min(30, [Math]::Pow(2, $attempt))

            Write-Log "Poging $attempt van $Retries mislukt voor '$Description': $message. Opnieuw over $waitSeconds seconden." "RETRY"
            Start-Sleep -Seconds $waitSeconds
        }
    }

    throw $lastError
}

function New-AppOnlyConnection {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    Invoke-WithRetry -Description "verbinden met $Url" -ScriptBlock {
        Connect-PnPOnline `
            -Url $Url `
            -ClientId $ClientId `
            -Tenant $TenantDomain `
            -Thumbprint $CertificateThumbprint `
            -ReturnConnection `
            -ErrorAction Stop
    }
}

# ===========================================================================
# Identity-resolutie
# ===========================================================================

function Resolve-UserByUpn {
    param(
        [Parameter(Mandatory)]
        [string]$Upn,

        [switch]$AllowUnresolved
    )

    try {
        $encoded = [System.Uri]::EscapeDataString($Upn)
        $graphUrl = "users/{0}?`$select=id,displayName,userPrincipalName,mail" -f $encoded

        $user = Invoke-WithRetry -Description "gebruiker $Upn opzoeken" -ScriptBlock {
            Invoke-PnPGraphMethod `
                -Url $graphUrl `
                -Method Get `
                -Connection $AdminConnection `
                -ErrorAction Stop
        }

        return [pscustomobject]@{
            RequestedValue    = $Upn
            ObjectId          = [string]$user.id
            DisplayName       = [string]$user.displayName
            UserPrincipalName = [string]$user.userPrincipalName
            Mail              = if ($user.PSObject.Properties.Name -contains "mail") { [string]$user.mail } else { "" }
            ResolutionStatus  = "RESOLVED"
        }
    }
    catch {
        if ($AllowUnresolved) {
            Write-Log "Gebruiker '$Upn' kon niet in Entra ID worden gevonden: $($_.Exception.Message)" "WARN"

            return [pscustomobject]@{
                RequestedValue    = $Upn
                ObjectId          = ""
                DisplayName       = $Upn
                UserPrincipalName = $Upn
                Mail              = $Upn
                ResolutionStatus  = "UNRESOLVED"
            }
        }

        throw
    }
}

function Resolve-AllIdentities {
    foreach ($definition in $TargetDefinitions) {
        Write-Log "Verwijdertarget oplossen: $($definition.Label) [$($definition.SearchValue)]" "INFO"

        $resolved = Resolve-UserByUpn -Upn ([string]$definition.SearchValue) -AllowUnresolved
        $resolved | Add-Member -NotePropertyName "Label" -NotePropertyValue ([string]$definition.Label) -Force

        $ResolvedTargets.Add($resolved)

        if ($resolved.ResolutionStatus -eq "RESOLVED") {
            Write-Log "Target gevonden: $($resolved.DisplayName) <$($resolved.UserPrincipalName)> [$($resolved.ObjectId)]" "OK"
        }
    }

    Write-Log "Backup-owner oplossen: $BackupOwnerUpn" "INFO"

    $script:BackupOwner = Resolve-UserByUpn -Upn $BackupOwnerUpn
    $BackupOwner | Add-Member -NotePropertyName "Label" -NotePropertyValue "Backup-owner" -Force

    Write-Log "Backup-owner gevonden: $($BackupOwner.DisplayName) <$($BackupOwner.UserPrincipalName)> [$($BackupOwner.ObjectId)]" "OK"

    foreach ($target in $ResolvedTargets) {
        if (
            -not [string]::IsNullOrWhiteSpace([string]$target.ObjectId) -and
            ([string]$target.ObjectId).Equals([string]$BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            throw "De backup-owner resolveert naar hetzelfde object als verwijdertarget '$($target.Label)'. Script afgebroken."
        }
    }
}

# ===========================================================================
# Matching
# ===========================================================================

function Get-NormalizedSharePointIdentity {
    param(
        [Parameter(Mandatory)]
        $Principal
    )

    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($propertyName in @("Email", "LoginName", "UserPrincipalName", "Title")) {
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

function Find-TargetByObjectId {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$ObjectId
    )

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $null
    }

    foreach ($target in $ResolvedTargets) {
        if (
            -not [string]::IsNullOrWhiteSpace([string]$target.ObjectId) -and
            $ObjectId.Equals([string]$target.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            return $target
        }
    }

    return $null
}

# ===========================================================================
# Microsoft 365 Group owners
# ===========================================================================

function Get-M365GroupOwnersFresh {
    <#
        Haalt de owners rechtstreeks en ongecacht op. Gooit een specifieke
        uitzondering wanneer de groep niet bestaat.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    $owners = [System.Collections.Generic.List[object]]::new()
    $url = "groups/$GroupId/owners?`$select=id,displayName,userPrincipalName&`$top=100"

    do {
        $page = Invoke-WithRetry -Description "owners ophalen van groep $GroupId" -ScriptBlock {
            Invoke-PnPGraphMethod `
                -Url $url `
                -Method Get `
                -Connection $Connection `
                -ErrorAction Stop
        }

        if ($page.PSObject.Properties.Name -contains "value") {
            foreach ($owner in $page.value) {
                $owners.Add($owner)
            }
        }

        $url = $null

        if ($page.PSObject.Properties.Name -contains "@odata.nextLink") {
            $url = [string]$page.'@odata.nextLink'
        }
    } while (-not [string]::IsNullOrWhiteSpace($url))

    @($owners)
}

function Test-M365GroupExists {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    try {
        $group = Invoke-PnPGraphMethod `
            -Url "groups/$GroupId`?`$select=id,displayName" `
            -Method Get `
            -Connection $Connection `
            -ErrorAction Stop

        return [pscustomobject]@{
            Exists      = $true
            Deleted     = $false
            DisplayName = [string]$group.displayName
        }
    }
    catch {
        # Mogelijk soft-deleted. Dat maakt uit voor de beoordeling: een
        # soft-deleted groep kan binnen 30 dagen worden hersteld en zou dan
        # alsnog een ongewenste owner bevatten.
        try {
            $deleted = Invoke-PnPGraphMethod `
                -Url "directory/deletedItems/microsoft.graph.group/$GroupId`?`$select=id,displayName,deletedDateTime" `
                -Method Get `
                -Connection $Connection `
                -ErrorAction Stop

            return [pscustomobject]@{
                Exists      = $false
                Deleted     = $true
                DisplayName = [string]$deleted.displayName
            }
        }
        catch {
            return [pscustomobject]@{
                Exists      = $false
                Deleted     = $false
                DisplayName = ""
            }
        }
    }
}

function Add-M365GroupOwnerConfirmed {
    <#
        Voegt de backup-owner toe en wacht tot Microsoft Graph de toevoeging
        bevestigt. Dit voorkomt de race condition waarbij een direct volgende
        verwijdering afketst op "the group must have at least one owner".
    #>
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    $body = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($BackupOwner.ObjectId)"
    }

    try {
        Invoke-WithRetry -Description "backup-owner toevoegen aan groep $GroupId" -ScriptBlock {
            Invoke-PnPGraphMethod `
                -Url "groups/$GroupId/owners/`$ref" `
                -Method Post `
                -Content $body `
                -Connection $Connection `
                -ErrorAction Stop | Out-Null
        }
    }
    catch {
        # Wanneer de owner er al staat geeft Graph een conflictfout terug.
        # Dat is geen probleem; de verificatie hieronder bepaalt de uitkomst.
        if ($_.Exception.Message -notmatch "already exist|One or more added object references already exist") {
            throw
        }
    }

    $deadline = (Get-Date).AddSeconds($OwnerConfirmTimeoutSeconds)
    $waitSeconds = 2

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $waitSeconds

        $owners = @(Get-M365GroupOwnersFresh -GroupId $GroupId -Connection $Connection)

        $confirmed = $owners | Where-Object {
            ([string]$_.id).Equals($BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
        }

        if ($confirmed) {
            Write-Log "Backup-owner bevestigd aanwezig in groep $GroupId." "OK"
            return $true
        }

        Write-Log "Backup-owner nog niet zichtbaar in groep $GroupId; wachten op replicatie." "RETRY"
        $waitSeconds = [Math]::Min(15, $waitSeconds * 2)
    }

    return $false
}

function Remove-M365GroupOwnerById {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$UserObjectId,

        [Parameter(Mandatory)]
        $Connection
    )

    Invoke-WithRetry -Description "owner $UserObjectId verwijderen uit groep $GroupId" -ScriptBlock {
        Invoke-PnPGraphMethod `
            -Url "groups/$GroupId/owners/$UserObjectId/`$ref" `
            -Method Delete `
            -Connection $Connection `
            -ErrorAction Stop | Out-Null
    }
}

function Repair-M365GroupOwners {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    # 1. Bestaat de groep nog?
    try {
        $owners = @(Get-M365GroupOwnersFresh -GroupId $GroupId -Connection $Connection)
    }
    catch {
        if ($_.Exception.Message -match "does not exist or one of its queried reference-property objects") {
            $groupState = Test-M365GroupExists -GroupId $GroupId -Connection $Connection

            if ($groupState.Deleted) {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "Microsoft365Group" `
                    -ObjectName $GroupId `
                    -Category "OrphanedGroup" `
                    -Action "None" `
                    -Status "SOFT_DELETED" `
                    -Details "De gekoppelde Microsoft 365-groep '$($groupState.DisplayName)' staat in de prullenbak van Entra ID. Owners zijn niet benaderbaar. Let op: bij herstel binnen de bewaartermijn keert de oorspronkelijke ownerlijst terug, inclusief eventuele verwijdertargets."

                Write-Log "Groep $GroupId is soft-deleted; geen ownerwijziging mogelijk | $SiteUrl" "WARN"
            }
            else {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "Microsoft365Group" `
                    -ObjectName $GroupId `
                    -Category "OrphanedGroup" `
                    -Action "None" `
                    -Status "NOT_FOUND" `
                    -Details "De SharePoint-site verwijst naar een Microsoft 365-groep die niet meer bestaat in Entra ID. Er zijn geen groep-owners om op te schonen. De site zelf is een weesobject en hoort thuis op de opruimlijst."

                Write-Log "Groep $GroupId bestaat niet meer; site is een weesobject | $SiteUrl" "WARN"
            }

            return
        }

        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365Group" `
            -ObjectName $GroupId `
            -Category "GroupOwnerScan" `
            -Action "ReadOwners" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Owners van groep $GroupId konden niet worden gelezen: $($_.Exception.Message)" "ERROR"
        return
    }

    # 2. Owners indelen
    $targetOwners = @()
    $backupPresent = $false
    $survivors = @()

    foreach ($owner in $owners) {
        $ownerId = [string]$owner.id
        $target = Find-TargetByObjectId -ObjectId $ownerId

        if ($target) {
            $targetOwners += [pscustomobject]@{
                Owner  = $owner
                Target = $target
            }
            continue
        }

        if ($ownerId.Equals($BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)) {
            $backupPresent = $true
            continue
        }

        $survivors += $owner
    }

    if ($targetOwners.Count -eq 0) {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365Group" `
            -ObjectName $GroupId `
            -Category "GroupOwnerScan" `
            -Action "None" `
            -Status "ALREADY_CLEAN" `
            -Details "Geen verwijdertargets meer aanwezig als owner van deze groep. Aantal owners: $($owners.Count)."

        Write-Log "Groep $GroupId bevat geen targets meer | $SiteUrl" "OK"

        # Ook hier geldt de opruimregel voor een overbodige backup-owner.
        if ($backupPresent -and $survivors.Count -gt 0) {
            Remove-RedundantBackupOwner -SiteUrl $SiteUrl -GroupId $GroupId -Connection $Connection
        }

        return
    }

    # 3. Backup toevoegen wanneer er anders geen owner overblijft
    $needsBackup = ($survivors.Count -eq 0 -and -not $backupPresent)

    if ($needsBackup) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -Category "BackupOwner" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou als owner worden toegevoegd en bevestigd voordat de laatste target-owner wordt verwijderd."

            Write-Log "DRY-RUN: backup-owner toevoegen aan groep $GroupId | $SiteUrl" "DRYRUN"
        }
        else {
            $confirmed = $false

            try {
                $confirmed = Add-M365GroupOwnerConfirmed -GroupId $GroupId -Connection $Connection
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner toevoegen mislukt voor groep $GroupId; er wordt hier NIETS verwijderd: $($_.Exception.Message)" "ERROR"
                return
            }

            if (-not $confirmed) {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "Microsoft365GroupOwner" `
                    -ObjectName $GroupId `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "NOT_CONFIRMED" `
                    -Details "Backup-owner is toegevoegd maar kwam binnen $OwnerConfirmTimeoutSeconds seconden niet bevestigd terug uit Microsoft Graph. Er is niets verwijderd. Draai dit script later opnieuw voor deze groep."

                Write-Log "Backup-owner niet bevestigd binnen de wachttijd voor groep $GroupId; niets verwijderd." "WARN"
                return
            }

            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -Category "BackupOwner" `
                -Action "Add" `
                -Status "ADDED" `
                -Details "Backup-owner toegevoegd en bevestigd als owner voordat targets werden verwijderd."

            Write-Log "Backup-owner toegevoegd en bevestigd voor groep $GroupId" "ACTION"
            $backupPresent = $true
        }
    }

    # 4. Targets een voor een verwijderen
    foreach ($entry in $targetOwners) {
        $target = $entry.Target

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "GroupOwner" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: owner zou uit de Microsoft 365-groep worden verwijderd."

            Write-Log "DRY-RUN: owner verwijderen uit groep $GroupId : $($target.UserPrincipalName)" "DRYRUN"
            continue
        }

        # Veiligheidscontrole vlak voor de wijziging: blijft er een owner over?
        $currentOwners = @(Get-M365GroupOwnersFresh -GroupId $GroupId -Connection $Connection)

        $remainingAfterRemoval = @(
            $currentOwners | Where-Object {
                -not ([string]$_.id).Equals([string]$target.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
            }
        )

        if ($remainingAfterRemoval.Count -eq 0) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "GroupOwner" `
                -Action "Remove" `
                -Status "SKIPPED" `
                -Details "Verwijdering overgeslagen: deze owner is nog altijd de enige owner en de backup-owner is niet bevestigd aanwezig. De groep zou anders zonder owner achterblijven."

            Write-Log "Verwijdering overgeslagen voor groep $GroupId : laatste owner zonder backup." "WARN"
            continue
        }

        try {
            Remove-M365GroupOwnerById `
                -GroupId $GroupId `
                -UserObjectId ([string]$target.ObjectId) `
                -Connection $Connection

            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "GroupOwner" `
                -Action "Remove" `
                -Status "REMOVED" `
                -Details "Owner verwijderd uit de Microsoft 365-groep. Resterende owners: $($remainingAfterRemoval.Count)."

            Write-Log "Owner verwijderd uit groep $GroupId : $($target.UserPrincipalName)" "ACTION"
        }
        catch {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365GroupOwner" `
                -ObjectName $GroupId `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "GroupOwner" `
                -Action "Remove" `
                -Status "ERROR" `
                -Details $_.Exception.Message

            Write-Log "Owner verwijderen mislukt uit groep $GroupId : $($_.Exception.Message)" "ERROR"
        }
    }

    # 5. Overbodige backup opruimen
    if ($Execute) {
        $finalOwners = @(Get-M365GroupOwnersFresh -GroupId $GroupId -Connection $Connection)

        $finalSurvivors = @(
            $finalOwners | Where-Object {
                -not ([string]$_.id).Equals($BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase) -and
                -not (Find-TargetByObjectId -ObjectId ([string]$_.id))
            }
        )

        $finalBackupPresent = @(
            $finalOwners | Where-Object {
                ([string]$_.id).Equals($BackupOwner.ObjectId, [System.StringComparison]::OrdinalIgnoreCase)
            }
        ).Count -gt 0

        if ($finalBackupPresent -and $finalSurvivors.Count -gt 0) {
            Remove-RedundantBackupOwner -SiteUrl $SiteUrl -GroupId $GroupId -Connection $Connection
        }
    }
}

function Remove-RedundantBackupOwner {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        $Connection
    )

    if (-not $Execute) {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365GroupOwner" `
            -ObjectName $GroupId `
            -TargetDisplayName $BackupOwner.DisplayName `
            -TargetUPN $BackupOwner.UserPrincipalName `
            -TargetObjectId $BackupOwner.ObjectId `
            -Category "BackupOwner" `
            -Action "WouldRemoveBackup" `
            -Status "FOUND" `
            -Details "Dry-run: Backup-owner is niet nodig omdat er een andere blijvende owner is."

        return
    }

    try {
        Remove-M365GroupOwnerById `
            -GroupId $GroupId `
            -UserObjectId $BackupOwner.ObjectId `
            -Connection $Connection

        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365GroupOwner" `
            -ObjectName $GroupId `
            -TargetDisplayName $BackupOwner.DisplayName `
            -TargetUPN $BackupOwner.UserPrincipalName `
            -TargetObjectId $BackupOwner.ObjectId `
            -Category "BackupOwner" `
            -Action "RemoveBackup" `
            -Status "REMOVED" `
            -Details "Backup-owner verwijderd omdat er een andere blijvende owner aanwezig is."

        Write-Log "Overbodige backup-owner verwijderd uit groep $GroupId" "ACTION"
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365GroupOwner" `
            -ObjectName $GroupId `
            -TargetDisplayName $BackupOwner.DisplayName `
            -TargetUPN $BackupOwner.UserPrincipalName `
            -TargetObjectId $BackupOwner.ObjectId `
            -Category "BackupOwner" `
            -Action "RemoveBackup" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Overbodige backup-owner verwijderen mislukt uit groep $GroupId : $($_.Exception.Message)" "ERROR"
    }
}

# ===========================================================================
# Site Collection Administrators
# ===========================================================================

function Repair-SiteCollectionAdmins {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        $Connection
    )

    try {
        $admins = @(
            Invoke-WithRetry -Description "Site Collection Admins lezen van $SiteUrl" -ScriptBlock {
                Get-PnPSiteCollectionAdmin -Connection $Connection -ErrorAction Stop
            }
        )
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "SiteCollectionAdmin" `
            -ObjectName $SiteUrl `
            -Category "SiteCollectionAdminScan" `
            -Action "ReadAdmins" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Site Collection Admins lezen mislukt na $MaxRetries pogingen: $SiteUrl" "ERROR"
        return
    }

    $targetMatches = @()
    $backupAdmin = $null
    $survivors = @()

    foreach ($admin in $admins) {
        $target = Find-TargetForSharePointPrincipal -Principal $admin

        if ($target) {
            $targetMatches += [pscustomobject]@{ Principal = $admin; Target = $target }
            continue
        }

        if (Test-IsBackupSharePointPrincipal -Principal $admin) {
            $backupAdmin = $admin
            continue
        }

        $survivors += $admin
    }

    if ($targetMatches.Count -eq 0) {
        if ($backupAdmin -and $survivors.Count -gt 0 -and $Execute) {
            try {
                Remove-PnPSiteCollectionAdmin `
                    -Owners ([string]$backupAdmin.LoginName) `
                    -Connection $Connection `
                    -ErrorAction Stop

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "SiteCollectionAdmin" `
                    -ObjectName $SiteUrl `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "RemoveBackup" `
                    -Status "REMOVED" `
                    -Details "Backup-owner verwijderd als Site Collection Admin omdat er andere blijvende admins zijn."
            }
            catch {
                Write-Log "Overbodige backup Site Collection Admin verwijderen mislukt: $SiteUrl | $($_.Exception.Message)" "WARN"
            }
        }

        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "SiteCollectionAdmin" `
            -ObjectName $SiteUrl `
            -Category "SiteCollectionAdminScan" `
            -Action "None" `
            -Status "ALREADY_CLEAN" `
            -Details "Geen verwijdertargets aanwezig als Site Collection Admin."

        return
    }

    if ($survivors.Count -eq 0 -and -not $backupAdmin) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "SiteCollectionAdmin" `
                -ObjectName $SiteUrl `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -Category "BackupOwner" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou als Site Collection Admin worden toegevoegd om 0 admins te voorkomen."
        }
        else {
            try {
                Invoke-WithRetry -Description "backup Site Collection Admin toevoegen aan $SiteUrl" -ScriptBlock {
                    Set-PnPTenantSite `
                        -Identity $SiteUrl `
                        -Owners $BackupOwner.UserPrincipalName `
                        -Connection $AdminConnection `
                        -ErrorAction Stop
                }

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "SiteCollectionAdmin" `
                    -ObjectName $SiteUrl `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "ADDED" `
                    -Details "Backup-owner toegevoegd als Site Collection Admin voordat targets werden verwijderd."

                Write-Log "Backup Site Collection Admin toegevoegd | $SiteUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $SiteUrl `
                    -Scope "SiteCollectionAdmin" `
                    -ObjectName $SiteUrl `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup Site Collection Admin toevoegen mislukt; niets verwijderd | $SiteUrl" "ERROR"
                return
            }
        }
    }

    foreach ($match in $targetMatches) {
        $target = $match.Target
        $identity = [string]$match.Principal.LoginName

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $identity = [string]$target.UserPrincipalName
        }

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "SiteCollectionAdmin" `
                -ObjectName $identity `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "SiteCollectionAdmin" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: Site Collection Admin-recht zou worden verwijderd."

            continue
        }

        try {
            Invoke-WithRetry -Description "Site Collection Admin verwijderen van $SiteUrl" -ScriptBlock {
                Remove-PnPSiteCollectionAdmin -Owners $identity -Connection $Connection -ErrorAction Stop
            }

            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "SiteCollectionAdmin" `
                -ObjectName $identity `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "SiteCollectionAdmin" `
                -Action "Remove" `
                -Status "REMOVED" `
                -Details "Site Collection Admin-recht verwijderd."

            Write-Log "Site Collection Admin verwijderd: $($target.UserPrincipalName) | $SiteUrl" "ACTION"
        }
        catch {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "SiteCollectionAdmin" `
                -ObjectName $identity `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "SiteCollectionAdmin" `
                -Action "Remove" `
                -Status "ERROR" `
                -Details $_.Exception.Message

            Write-Log "Site Collection Admin verwijderen mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
        }
    }
}

# ===========================================================================
# SharePoint Owners-groep
# ===========================================================================

function Get-AssociatedOwnerGroupSafe {
    <#
        Retourneert de Associated Owners Group, of $null wanneer die niet
        bestaat. Een ontbrekende AssociatedOwnerGroup is een geldige toestand
        en mag niet als fout worden behandeld.
    #>
    param(
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
            return $null
        }

        return $ownersGroup
    }
    catch {
        $message = [string]$_.Exception.Message

        if (
            $message -match "AssociatedOwnerGroup" -or
            $message -match "null-object" -or
            $message -match "null object"
        ) {
            return $null
        }

        throw
    }
}

function Repair-SharePointOwnersGroup {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$WebUrl,

        [Parameter(Mandatory)]
        $Web
    )

    $context = $Web.Context

    try {
        $ownersGroup = Invoke-WithRetry -Description "Owners-groep lezen van $WebUrl" -ScriptBlock {
            Get-AssociatedOwnerGroupSafe -Web $Web
        }
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $WebUrl `
            -Scope "SharePointOwnersGroup" `
            -ObjectName "Associated Owners Group" `
            -Category "OwnersGroupScan" `
            -Action "ReadMembers" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Owners-groep lezen mislukt na $MaxRetries pogingen: $WebUrl" "ERROR"
        return
    }

    if ($null -eq $ownersGroup) {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $WebUrl `
            -Scope "SharePointOwnersGroup" `
            -ObjectName "Associated Owners Group" `
            -Category "NoOwnersGroup" `
            -Action "None" `
            -Status "NOT_APPLICABLE" `
            -Details "Deze web heeft geen gekoppelde SharePoint Owners-groep. Bij group-connected sites is de Microsoft 365-groep leidend voor eigenaarschap; die is apart gecontroleerd."

        Write-Log "Geen Associated Owners Group aanwezig: $WebUrl" "WARN"
        return
    }

    $groupName = [string]$ownersGroup.Title
    $members = @($ownersGroup.Users)

    $targetMatches = @()
    $backupMember = $null
    $survivors = @()

    foreach ($member in $members) {
        $target = Find-TargetForSharePointPrincipal -Principal $member

        if ($target) {
            $targetMatches += [pscustomobject]@{ Principal = $member; Target = $target }
            continue
        }

        if (Test-IsBackupSharePointPrincipal -Principal $member) {
            $backupMember = $member
            continue
        }

        $survivors += $member
    }

    if ($targetMatches.Count -eq 0) {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $WebUrl `
            -Scope "SharePointOwnersGroup" `
            -ObjectName $groupName `
            -Category "OwnersGroupScan" `
            -Action "None" `
            -Status "ALREADY_CLEAN" `
            -Details "Geen verwijdertargets aanwezig in Owners-groep '$groupName'. Aantal leden: $($members.Count)."

        if ($backupMember -and $survivors.Count -gt 0 -and $Execute) {
            try {
                $ownersGroup.Users.RemoveByLoginName([string]$backupMember.LoginName)
                $context.ExecuteQuery()

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Scope "SharePointOwnersGroup" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "RemoveBackup" `
                    -Status "REMOVED" `
                    -Details "Backup-owner verwijderd uit '$groupName' omdat er andere blijvende owners zijn."
            }
            catch {
                Write-Log "Overbodige backup-owner verwijderen mislukt uit '$groupName': $($_.Exception.Message)" "WARN"
            }
        }

        return
    }

    if ($survivors.Count -eq 0 -and -not $backupMember) {
        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Scope "SharePointOwnersGroup" `
                -ObjectName $groupName `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -Category "BackupOwner" `
                -Action "WouldAdd" `
                -Status "FOUND" `
                -Details "Dry-run: Backup-owner zou aan Owners-groep '$groupName' worden toegevoegd om 0 owners te voorkomen."
        }
        else {
            try {
                Invoke-WithRetry -Description "backup-owner toevoegen aan $groupName op $WebUrl" -ScriptBlock {
                    $backupUser = $Web.EnsureUser($BackupOwner.UserPrincipalName)
                    $ownersGroup.Users.AddUser($backupUser) | Out-Null
                    $context.ExecuteQuery()
                }

                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Scope "SharePointOwnersGroup" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "ADDED" `
                    -Details "Backup-owner toegevoegd aan Owners-groep '$groupName' voordat targets werden verwijderd."

                Write-Log "Backup-owner toegevoegd aan '$groupName' | $WebUrl" "ACTION"
            }
            catch {
                Add-Result `
                    -SiteUrl $SiteUrl `
                    -WebUrl $WebUrl `
                    -Scope "SharePointOwnersGroup" `
                    -ObjectName $groupName `
                    -TargetDisplayName $BackupOwner.DisplayName `
                    -TargetUPN $BackupOwner.UserPrincipalName `
                    -TargetObjectId $BackupOwner.ObjectId `
                    -Category "BackupOwner" `
                    -Action "Add" `
                    -Status "ERROR" `
                    -Details $_.Exception.Message

                Write-Log "Backup-owner toevoegen mislukt aan '$groupName'; niets verwijderd | $WebUrl" "ERROR"
                return
            }
        }
    }

    foreach ($match in $targetMatches) {
        $target = $match.Target
        $loginName = [string]$match.Principal.LoginName

        if (-not $Execute) {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Scope "SharePointOwnersGroup" `
                -ObjectName $groupName `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "OwnersGroupMember" `
                -Action "WouldRemove" `
                -Status "FOUND" `
                -Details "Dry-run: lid zou uit Owners-groep '$groupName' worden verwijderd."

            continue
        }

        try {
            Invoke-WithRetry -Description "lid verwijderen uit $groupName op $WebUrl" -ScriptBlock {
                $ownersGroup.Users.RemoveByLoginName($loginName)
                $context.ExecuteQuery()
            }

            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Scope "SharePointOwnersGroup" `
                -ObjectName $groupName `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "OwnersGroupMember" `
                -Action "Remove" `
                -Status "REMOVED" `
                -Details "Lid verwijderd uit Owners-groep '$groupName'."

            Write-Log "Uit Owners-groep verwijderd: $($target.UserPrincipalName) | $WebUrl" "ACTION"
        }
        catch {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Scope "SharePointOwnersGroup" `
                -ObjectName $groupName `
                -TargetDisplayName $target.DisplayName `
                -TargetUPN $target.UserPrincipalName `
                -TargetObjectId $target.ObjectId `
                -Category "OwnersGroupMember" `
                -Action "Remove" `
                -Status "ERROR" `
                -Details $_.Exception.Message

            Write-Log "Verwijderen uit Owners-groep mislukt: $WebUrl | $($_.Exception.Message)" "ERROR"
        }
    }

    if ($backupMember -and $survivors.Count -gt 0 -and $Execute) {
        try {
            $ownersGroup.Users.RemoveByLoginName([string]$backupMember.LoginName)
            $context.ExecuteQuery()

            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $WebUrl `
                -Scope "SharePointOwnersGroup" `
                -ObjectName $groupName `
                -TargetDisplayName $BackupOwner.DisplayName `
                -TargetUPN $BackupOwner.UserPrincipalName `
                -TargetObjectId $BackupOwner.ObjectId `
                -Category "BackupOwner" `
                -Action "RemoveBackup" `
                -Status "REMOVED" `
                -Details "Backup-owner verwijderd uit '$groupName' omdat er een andere blijvende owner is."
        }
        catch {
            Write-Log "Overbodige backup-owner verwijderen mislukt uit '$groupName': $($_.Exception.Message)" "WARN"
        }
    }
}

# ===========================================================================
# Subsites
# ===========================================================================

function Get-SubWebsCsom {
    <#
        Eigen recursieve subweb-enumeratie via CSOM. Get-PnPSubWeb -Recurse
        gaf in de hoofdrun op enkele sites "Nullable object must have a value";
        deze route omzeilt die codepath volledig.
    #>
    param(
        [Parameter(Mandatory)]
        $Web
    )

    $context = $Web.Context
    $collected = [System.Collections.Generic.List[object]]::new()

    $context.Load($Web.Webs)
    $context.ExecuteQuery()

    foreach ($child in $Web.Webs) {
        $collected.Add($child)

        foreach ($grandChild in (Get-SubWebsCsom -Web $child)) {
            $collected.Add($grandChild)
        }
    }

    @($collected)
}

function Get-WebTemplateSafe {
    param(
        [Parameter(Mandatory)]
        $Web
    )

    try {
        $template = [string]$Web.WebTemplate
        $configuration = [string]$Web.Configuration

        if ([string]::IsNullOrWhiteSpace($template)) {
            return $null
        }

        return ("$template#$configuration").ToUpperInvariant()
    }
    catch {
        return $null
    }
}

# ===========================================================================
# Verwerking per site
# ===========================================================================

function Repair-Site {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    Write-Log "Site herstellen: $SiteUrl" "INFO"

    try {
        $siteConnection = New-AppOnlyConnection -Url $SiteUrl
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "SiteCollection" `
            -ObjectName $SiteUrl `
            -Category "Connection" `
            -Action "Connect" `
            -Status "ERROR" `
            -Details "Verbinden mislukt na $MaxRetries pogingen: $($_.Exception.Message)"

        Write-Log "Verbinden mislukt na $MaxRetries pogingen: $SiteUrl" "ERROR"
        return
    }

    # 1. Site Collection Administrators
    Repair-SiteCollectionAdmins -SiteUrl $SiteUrl -Connection $siteConnection

    # 2. Owners-groep van de root web
    try {
        $context = Get-PnPContext -Connection $siteConnection
        $rootWeb = $context.Web

        Repair-SharePointOwnersGroup -SiteUrl $SiteUrl -WebUrl $SiteUrl -Web $rootWeb
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "SharePointOwnersGroup" `
            -ObjectName "Associated Owners Group" `
            -Category "OwnersGroupScan" `
            -Action "ReadMembers" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "Root owners-groep verwerken mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
    }

    # 3. Microsoft 365-groep
    try {
        $siteProperties = Invoke-WithRetry -Description "site-eigenschappen ophalen van $SiteUrl" -ScriptBlock {
            Get-PnPTenantSite -Identity $SiteUrl -Detailed -Connection $AdminConnection -ErrorAction Stop
        }

        $groupId = ""

        if ($siteProperties.PSObject.Properties.Name -contains "GroupId") {
            $groupId = [string]$siteProperties.GroupId
        }

        if (
            -not [string]::IsNullOrWhiteSpace($groupId) -and
            $groupId -ne "00000000-0000-0000-0000-000000000000"
        ) {
            Repair-M365GroupOwners -SiteUrl $SiteUrl -GroupId $groupId -Connection $siteConnection
        }
        elseif (([string]$siteProperties.Template).ToUpperInvariant() -eq "GROUP#0") {
            Add-Result `
                -SiteUrl $SiteUrl `
                -WebUrl $SiteUrl `
                -Scope "Microsoft365Group" `
                -ObjectName $SiteUrl `
                -Category "OrphanedGroup" `
                -Action "None" `
                -Status "NO_GROUP_ID" `
                -Details "Site heeft template GROUP#0 maar geen GroupId. De gekoppelde groep is waarschijnlijk verwijderd. De SharePoint-scopes zijn wel gecontroleerd."
        }
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "Microsoft365Group" `
            -ObjectName $SiteUrl `
            -Category "GroupIdLookup" `
            -Action "Read" `
            -Status "ERROR" `
            -Details $_.Exception.Message

        Write-Log "GroupId bepalen mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
    }

    # 4. Subsites
    try {
        $subWebs = @(
            Invoke-WithRetry -Description "subsites ophalen van $SiteUrl" -ScriptBlock {
                $context = Get-PnPContext -Connection $siteConnection
                Get-SubWebsCsom -Web $context.Web
            }
        )
    }
    catch {
        Add-Result `
            -SiteUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -Scope "SubWebScan" `
            -ObjectName $SiteUrl `
            -Category "SubWebEnumeration" `
            -Action "Read" `
            -Status "ERROR" `
            -Details "Subsite-enumeratie mislukt na $MaxRetries pogingen, ook via de CSOM-fallback: $($_.Exception.Message)"

        Write-Log "Subsites ophalen mislukt: $SiteUrl | $($_.Exception.Message)" "ERROR"
        $subWebs = @()
    }

    foreach ($subWeb in $subWebs) {
        $subWebUrl = [string]$subWeb.Url
        $subWebTemplate = Get-WebTemplateSafe -Web $subWeb

        if ($null -ne $subWebTemplate -and $IncludedTemplates -notcontains $subWebTemplate) {
            continue
        }

        if ($null -eq $subWebTemplate) {
            Write-Log "Template van subsite $subWebUrl kon niet worden bepaald; de subsite wordt voor de zekerheid wel gecontroleerd." "WARN"
        }

        Repair-SharePointOwnersGroup -SiteUrl $SiteUrl -WebUrl $subWebUrl -Web $subWeb
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
Write-Host " SharePoint Owner Cleanup - Foutherstel" -ForegroundColor Cyan
Write-Host " Mode: $modeText" -ForegroundColor $(if ($Execute) { "Magenta" } else { "Cyan" })
Write-Host " Verwerking: serieel, met $MaxRetries pogingen per bewerking" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host ""

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$sitesToProcess = @()

try {
    # 1. Sites bepalen
    if ($PSCmdlet.ParameterSetName -eq "FromReport") {
        if (-not (Test-Path -LiteralPath $ReportPath)) {
            throw "Rapport niet gevonden: $ReportPath"
        }

        $report = @(Import-Csv -LiteralPath $ReportPath)

        if ($report.Count -eq 0) {
            throw "Rapport bevat geen regels: $ReportPath"
        }

        if ($report[0].PSObject.Properties.Name -notcontains "Status") {
            throw "Rapport bevat geen kolom 'Status'. Is dit het rapport van Remove-SharePointOwnersTenantWide?"
        }

        $sitesToProcess = @(
            $report |
                Where-Object { $_.Status -in @("ERROR", "WARN") } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.SiteUrl) } |
                Select-Object -ExpandProperty SiteUrl -Unique |
                Sort-Object
        )

        Write-Log "Rapport ingelezen: $($report.Count) regels, $($sitesToProcess.Count) unieke site(s) met fouten of waarschuwingen." "OK"
    }
    else {
        $sitesToProcess = @($SiteUrl | Sort-Object -Unique)
        Write-Log "Handmatige sitelijst: $($sitesToProcess.Count) site(s)." "OK"
    }

    if ($sitesToProcess.Count -eq 0) {
        Write-Log "Geen sites om te verwerken. Klaar." "OK"
        return
    }

    # 2. Verbinden en valideren
    Write-Log "Verbinden met SharePoint Admin Center: $TenantAdminUrl" "INFO"

    $AdminConnection = New-AppOnlyConnection -Url $TenantAdminUrl

    $connectedTenantId = [string](Get-PnPTenantId -Connection $AdminConnection -ErrorAction Stop)

    if (-not $connectedTenantId.Equals($TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Tenant mismatch. Verwacht '$TenantId', verbonden met '$connectedTenantId'."
    }

    Write-Log "Tenant-ID klopt: $connectedTenantId" "OK"

    # 3. Identiteiten oplossen
    Resolve-AllIdentities

    # 4. Sites verwerken
    $index = 0

    foreach ($site in $sitesToProcess) {
        $index++
        Write-Host ""
        Write-Log "[$index/$($sitesToProcess.Count)] $site" "INFO"

        try {
            Repair-Site -SiteUrl $site
        }
        catch {
            Add-Result `
                -SiteUrl $site `
                -WebUrl $site `
                -Scope "SiteCollection" `
                -ObjectName $site `
                -Category "SiteProcessing" `
                -Action "Process" `
                -Status "ERROR" `
                -Details $_.Exception.Message

            Write-Log "Onverwachte fout bij $site : $($_.Exception.Message)" "ERROR"
        }
    }
}
catch {
    Write-Log $_.Exception.Message "ERROR"

    Add-Result `
        -Scope "Fatal" `
        -ObjectName "Script" `
        -Category "FatalError" `
        -Action "Abort" `
        -Status "ERROR" `
        -Details $_.Exception.Message
}
finally {
    $stopwatch.Stop()

    try {
        $directory = Split-Path -Path $OutputPath -Parent

        if (
            -not [string]::IsNullOrWhiteSpace($directory) -and
            -not (Test-Path -LiteralPath $directory)
        ) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        if ($Results.Count -gt 0) {
            $Results |
                Sort-Object SiteUrl, WebUrl, Scope, TargetUPN, Action |
                Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8BOM

            Write-Host ""
            Write-Log "Rapport opgeslagen: $OutputPath" "OK"
        }
        else {
            Write-Log "Geen resultaten om te rapporteren." "WARN"
        }
    }
    catch {
        Write-Log "Rapport kon niet worden opgeslagen: $($_.Exception.Message)" "ERROR"
    }

    $statusSummary = $Results |
        Group-Object Status |
        Sort-Object Name |
        ForEach-Object { "  {0,-16} {1}" -f $_.Name, $_.Count }

    Write-Host ""
    Write-Host "Samenvatting" -ForegroundColor Cyan
    Write-Host "-----------"
    Write-Host "Mode        : $(if ($Execute) { 'EXECUTE' } else { 'DRY-RUN' })"
    Write-Host "Sites       : $($sitesToProcess.Count)"
    Write-Host "Doorlooptijd: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
    Write-Host ""
    Write-Host "Status per categorie:" -ForegroundColor Cyan

    foreach ($line in $statusSummary) {
        Write-Host $line
    }

    Write-Host ""
    Write-Host "Report      : $OutputPath"
    Write-Host ""
}
