<#
.SYNOPSIS
    Kloont een bestaand (admin) account naar een nieuw cloud-only account, inclusief
    alle kopieerbare Entra ID groepslidmaatschappen en directory-rollen.

.DESCRIPTION
    Dit script maakt een nieuw cloud-only gebruikersaccount aan en kopieert:
      - Groepslidmaatschappen van het bronaccount (alleen groepen die via Graph
        beheerbaar zijn: assigned security groups en Microsoft 365 groups).
      - Directory-rollen via de unified role assignments (incl. directory scope).

    Groepen die NIET via Graph gevuld kunnen worden, worden overgeslagen en gelogd:
      - Dynamische groepen           (lidmaatschap wordt door regel bepaald)
      - On-premises gesynchroniseerde groepen (beheer via AD, niet in de cloud)
      - Mail-enabled security groups / distributielijsten (beheer via Exchange Online)

    LET OP: PIM-eligible role assignments worden NIET gekopieerd (alleen actieve
    assignments). Dit is bewust; eligible toekenningen horen via PIM te lopen.

.PARAMETER SourceUserUPN
    UserPrincipalName van het bronaccount dat gekopieerd wordt.

.PARAMETER NewDisplayName
    Display name van het nieuwe account, bv. "Admin Jan de Vries (Cloud Only)".

.PARAMETER NewUserPrincipalName
    UPN van het nieuwe cloud-only account, bv. adm.jan.de.vries@contoso.onmicrosoft.com

.PARAMETER Execute
    Zonder deze switch draait het script in DRY-RUN modus en wordt er niets aangemaakt/
    gewijzigd. Met -Execute worden de wijzigingen daadwerkelijk doorgevoerd.

.PARAMETER LogPath
    Pad voor de CSV-log (semicolon-delimited, UTF-8). Default: naast het script.

.EXAMPLE
    # Dry-run: laat zien wat er zou gebeuren
    .\Copy-AdminAccount.ps1 -SourceUserUPN adm.bron@contoso.com -NewDisplayName "Admin Jan de Vries (Cloud Only)" -NewUserPrincipalName adm.jan.de.vries@contoso.com

.EXAMPLE
    # Daadwerkelijk uitvoeren
    .\Copy-AdminAccount.ps1 -SourceUserUPN adm.bron@contoso.com -NewDisplayName "Admin Jan de Vries (Cloud Only)" -NewUserPrincipalName adm.jan.de.vries@contoso.com -Execute

.NOTES
    Vereiste modules : Microsoft.Graph (Users, Groups, Identity.DirectoryManagement, Identity.Governance)
    Vereiste scopes  : User.ReadWrite.All, Group.ReadWrite.All,
                       RoleManagement.ReadWrite.Directory, Directory.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceUserUPN,

    [Parameter(Mandatory = $true)]
    [string]$NewDisplayName,

    [Parameter(Mandatory = $true)]
    [string]$NewUserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$MailNickname,

    [Parameter(Mandatory = $false)]
    [switch]$ForceChangePassword = $true,

    [Parameter(Mandatory = $false)]
    [switch]$Execute,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path $PSScriptRoot ("CopyAdminAccount_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)))
)

#region Helpers ----------------------------------------------------------------

$script:LogEntries = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param(
        [string]$Type,       # INFO / ACTIE / SKIP / FOUT / DRYRUN
        [string]$Category,   # bv. Groep, Rol, Account
        [string]$Item,       # naam van het object
        [string]$Message
    )
    $color = switch ($Type) {
        'FOUT'   { 'Red' }
        'SKIP'   { 'Yellow' }
        'ACTIE'  { 'Green' }
        'DRYRUN' { 'Cyan' }
        default  { 'Gray' }
    }
    Write-Host ("[{0,-6}] {1,-8} {2} {3}" -f $Type, $Category, $Item, $Message) -ForegroundColor $color

    $script:LogEntries.Add([pscustomobject]@{
        Tijdstip = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Type     = $Type
        Categorie = $Category
        Item     = $Item
        Bericht  = $Message
    })
}

function New-RandomPassword {
    param([int]$Length = 24)
    # Zorg voor complexiteit: minimaal 1 uit elke set
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lower   = 'abcdefghijkmnpqrstuvwxyz'
    $digit   = '23456789'
    $special = '!@#$%^&*()-_=+'
    $all     = $upper + $lower + $digit + $special

    $chars = @(
        $upper[(Get-Random -Max $upper.Length)]
        $lower[(Get-Random -Max $lower.Length)]
        $digit[(Get-Random -Max $digit.Length)]
        $special[(Get-Random -Max $special.Length)]
    )
    for ($i = $chars.Count; $i -lt $Length; $i++) {
        $chars += $all[(Get-Random -Max $all.Length)]
    }
    # Shuffle
    -join ($chars | Sort-Object { Get-Random })
}

#endregion ---------------------------------------------------------------------

#region Verbinding -------------------------------------------------------------

Write-Log -Type 'INFO' -Category 'Verbinding' -Item '' -Message 'Verbinden met Microsoft Graph...'

$requiredScopes = @(
    'User.ReadWrite.All',
    'Group.ReadWrite.All',
    'RoleManagement.ReadWrite.Directory',
    'Directory.Read.All'
)

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
    } else {
        # Controleer of alle benodigde scopes aanwezig zijn
        $missing = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
        if ($missing) {
            Write-Log -Type 'INFO' -Category 'Verbinding' -Item '' -Message "Ontbrekende scopes ($($missing -join ', ')), opnieuw verbinden..."
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        }
    }
    $ctx = Get-MgContext
    Write-Log -Type 'INFO' -Category 'Verbinding' -Item '' -Message "Verbonden met tenant $($ctx.TenantId) als $($ctx.Account)"
}
catch {
    Write-Log -Type 'FOUT' -Category 'Verbinding' -Item '' -Message "Verbinden mislukt: $($_.Exception.Message)"
    return
}

#endregion ---------------------------------------------------------------------

#region Bronaccount ophalen ----------------------------------------------------

try {
    $source = Get-MgUser -UserId $SourceUserUPN -Property Id,DisplayName,UserPrincipalName,UsageLocation,Department,JobTitle,CompanyName -ErrorAction Stop
    Write-Log -Type 'INFO' -Category 'Bron' -Item $source.UserPrincipalName -Message "Bronaccount gevonden (Id: $($source.Id))"
}
catch {
    Write-Log -Type 'FOUT' -Category 'Bron' -Item $SourceUserUPN -Message "Bronaccount niet gevonden: $($_.Exception.Message)"
    return
}

# Controleer of het doel-UPN al bestaat
$existing = Get-MgUser -Filter "userPrincipalName eq '$NewUserPrincipalName'" -ErrorAction SilentlyContinue
if ($existing) {
    Write-Log -Type 'FOUT' -Category 'Doel' -Item $NewUserPrincipalName -Message 'Er bestaat al een account met deze UPN. Afgebroken.'
    return
}

if (-not $MailNickname) {
    $MailNickname = ($NewUserPrincipalName -split '@')[0]
}

#endregion ---------------------------------------------------------------------

#region Groepen inventariseren -------------------------------------------------

Write-Log -Type 'INFO' -Category 'Groep' -Item '' -Message 'Groepslidmaatschappen ophalen...'

$memberOf = Get-MgUserMemberOf -UserId $source.Id -All -ErrorAction Stop

$copyableGroups = [System.Collections.Generic.List[object]]::new()

foreach ($obj in $memberOf) {
    $odataType = $obj.AdditionalProperties['@odata.type']
    if ($odataType -ne '#microsoft.graph.group') { continue }   # rollen apart afgehandeld

    # Volledige groepdetails ophalen voor correcte classificatie
    $g = Get-MgGroup -GroupId $obj.Id -Property Id,DisplayName,GroupTypes,MailEnabled,SecurityEnabled,OnPremisesSyncEnabled -ErrorAction SilentlyContinue
    if (-not $g) { continue }

    $isDynamic  = $g.GroupTypes -contains 'DynamicMembership'
    $isOnPrem   = $g.OnPremisesSyncEnabled -eq $true
    $isMailOnly = ($g.MailEnabled -eq $true -and $g.SecurityEnabled -eq $false)   # distributielijst
    $isMailSec  = ($g.MailEnabled -eq $true -and $g.SecurityEnabled -eq $true)    # mail-enabled security group

    if     ($isDynamic)  { Write-Log -Type 'SKIP' -Category 'Groep' -Item $g.DisplayName -Message 'Dynamische groep (lidmaatschap via regel) - overgeslagen'; continue }
    elseif ($isOnPrem)   { Write-Log -Type 'SKIP' -Category 'Groep' -Item $g.DisplayName -Message 'On-prem gesynchroniseerd - beheer via AD - overgeslagen'; continue }
    elseif ($isMailOnly) { Write-Log -Type 'SKIP' -Category 'Groep' -Item $g.DisplayName -Message 'Distributielijst - beheer via Exchange Online - overgeslagen'; continue }
    elseif ($isMailSec)  { Write-Log -Type 'SKIP' -Category 'Groep' -Item $g.DisplayName -Message 'Mail-enabled security group - beheer via Exchange Online - overgeslagen'; continue }
    else {
        $copyableGroups.Add($g)
    }
}

Write-Log -Type 'INFO' -Category 'Groep' -Item '' -Message "$($copyableGroups.Count) kopieerbare groep(en) gevonden."

#endregion ---------------------------------------------------------------------

#region Rollen inventariseren --------------------------------------------------

Write-Log -Type 'INFO' -Category 'Rol' -Item '' -Message 'Directory-rollen (actieve assignments) ophalen...'

$roleAssignments = Get-MgRoleManagementDirectoryRoleAssignment `
    -Filter "principalId eq '$($source.Id)'" `
    -ExpandProperty RoleDefinition -All -ErrorAction Stop

foreach ($ra in $roleAssignments) {
    $scopeLabel = if ($ra.DirectoryScopeId -eq '/') { 'Tenant-breed' } else { "Scope: $($ra.DirectoryScopeId)" }
    Write-Log -Type 'INFO' -Category 'Rol' -Item $ra.RoleDefinition.DisplayName -Message $scopeLabel
}

Write-Log -Type 'INFO' -Category 'Rol' -Item '' -Message "$($roleAssignments.Count) actieve rol-assignment(s) gevonden."

#endregion ---------------------------------------------------------------------

#region Uitvoeren / Dry-run ----------------------------------------------------

if (-not $Execute) {
    Write-Host ""
    Write-Log -Type 'DRYRUN' -Category 'Account' -Item $NewUserPrincipalName -Message "Zou aangemaakt worden met display name '$NewDisplayName'"
    Write-Log -Type 'DRYRUN' -Category 'Groep'   -Item '' -Message "Zou aan $($copyableGroups.Count) groep(en) toegevoegd worden"
    Write-Log -Type 'DRYRUN' -Category 'Rol'     -Item '' -Message "Zou $($roleAssignments.Count) rol-assignment(s) krijgen"
    Write-Host ""
    Write-Host "DRY-RUN: er is niets gewijzigd. Draai met -Execute om daadwerkelijk aan te maken." -ForegroundColor Cyan
}
else {
    # --- 1. Account aanmaken ---
    $plainPassword = New-RandomPassword -Length 24
    $passwordProfile = @{
        Password                      = $plainPassword
        ForceChangePasswordNextSignIn = [bool]$ForceChangePassword
    }

    $userParams = @{
        DisplayName       = $NewDisplayName
        UserPrincipalName = $NewUserPrincipalName
        MailNickname      = $MailNickname
        AccountEnabled    = $true
        PasswordProfile   = $passwordProfile
    }
    if ($source.UsageLocation) { $userParams['UsageLocation'] = $source.UsageLocation }

    try {
        $newUser = New-MgUser -BodyParameter $userParams -ErrorAction Stop
        Write-Log -Type 'ACTIE' -Category 'Account' -Item $NewUserPrincipalName -Message "Aangemaakt (Id: $($newUser.Id))"
    }
    catch {
        Write-Log -Type 'FOUT' -Category 'Account' -Item $NewUserPrincipalName -Message "Aanmaken mislukt: $($_.Exception.Message)"
        return
    }

    # Korte wachttijd i.v.m. directory-replicatie voordat we lidmaatschappen toevoegen
    Start-Sleep -Seconds 10

    # --- 2. Groepen toevoegen ---
    foreach ($g in $copyableGroups) {
        try {
            New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $newUser.Id -ErrorAction Stop
            Write-Log -Type 'ACTIE' -Category 'Groep' -Item $g.DisplayName -Message 'Toegevoegd'
        }
        catch {
            # 'added object references already exist' = al lid, geen echte fout
            if ($_.Exception.Message -match 'already exist') {
                Write-Log -Type 'SKIP' -Category 'Groep' -Item $g.DisplayName -Message 'Was al lid'
            } else {
                Write-Log -Type 'FOUT' -Category 'Groep' -Item $g.DisplayName -Message $_.Exception.Message
            }
        }
    }

    # --- 3. Rollen toewijzen ---
    foreach ($ra in $roleAssignments) {
        try {
            $params = @{
                RoleDefinitionId = $ra.RoleDefinitionId
                PrincipalId      = $newUser.Id
                DirectoryScopeId = $ra.DirectoryScopeId
            }
            New-MgRoleManagementDirectoryRoleAssignment -BodyParameter $params -ErrorAction Stop | Out-Null
            Write-Log -Type 'ACTIE' -Category 'Rol' -Item $ra.RoleDefinition.DisplayName -Message 'Toegewezen'
        }
        catch {
            if ($_.Exception.Message -match 'already exist|conflicting') {
                Write-Log -Type 'SKIP' -Category 'Rol' -Item $ra.RoleDefinition.DisplayName -Message 'Was al toegewezen'
            } else {
                Write-Log -Type 'FOUT' -Category 'Rol' -Item $ra.RoleDefinition.DisplayName -Message $_.Exception.Message
            }
        }
    }

    # --- 4. Wachtwoord tonen ---
    Write-Host ""
    Write-Host "==================== NIEUW ACCOUNT ====================" -ForegroundColor Green
    Write-Host " UPN         : $NewUserPrincipalName"
    Write-Host " DisplayName : $NewDisplayName"
    Write-Host " Wachtwoord  : $plainPassword"
    Write-Host " (Wachtwoord wijzigen bij eerste login: $([bool]$ForceChangePassword))"
    Write-Host "=======================================================" -ForegroundColor Green
    Write-Host " Bewaar dit wachtwoord veilig - het wordt niet gelogd." -ForegroundColor Yellow
    Write-Host ""
}

#endregion ---------------------------------------------------------------------

#region Log wegschrijven -------------------------------------------------------

try {
    $script:LogEntries | Export-Csv -Path $LogPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    Write-Host "Logbestand opgeslagen: $LogPath" -ForegroundColor Gray
}
catch {
    Write-Warning "Kon logbestand niet wegschrijven: $($_.Exception.Message)"
}

#endregion ---------------------------------------------------------------------