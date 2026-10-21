<#
.SYNOPSIS
    Zet de Single Sign-On modus van een Enterprise Application terug naar 'Disabled'
    en geeft de vastzittende Identifier (Entity ID) vrij.

.DESCRIPTION
    Entra weigert het wijzigen van preferredSingleSignOnMode van 'saml' naar null met:

        "Modifying 'preferredSingleSignOnMode' property on the service principal from
         'saml' to 'null' is not allowed when the application contains identifierUris
         that do not use a verified domain of the organization or its subdomain, or
         end with a trailing slash"

    Dit treft vrijwel elke handmatig geconfigureerde SAML-app, omdat de Identifier
    doorgaans een URL van de leverancier is (niet-geverifieerd domein) en/of eindigt
    op een slash, bijvoorbeeld https://app.leverancier.nl/.

    De oplossing is de identifierUris op de APP-REGISTRATIE eerst te vervangen door
    api://<appId>. Dat formaat wordt altijd geaccepteerd omdat het geen
    domeinverificatie vereist en geen trailing slash heeft. Pas daarna laat Entra de
    SSO-modus omlaag zetten.

    Uitvoeringsvolgorde (dwingend):
      1. identifierUris        -> @("api://<appId>")   op de applicatie
      2. preferredSingleSignOnMode -> null              op de service principal
      3. servicePrincipalNames -> @(<appId>)            geeft de Entity ID vrij
      4. samlSingleSignOnSettings / loginUrl / replyUrls opruimen
      5. keyCredentials (SAML signing certificaten) verwijderen
      6. identifierUris -> @() volledig leegmaken

    Stap 6 is optioneel: laat je api://<appId> staan, dan is de oorspronkelijke
    Entity ID evengoed vrijgegeven voor hergebruik in een andere applicatie.

.PARAMETER KeepApiIdentifier
    Laat api://<appId> staan in plaats van identifierUris volledig leeg te maken.
    Aanbevolen wanneer de app-registratie later nog hergebruikt wordt.

.EXAMPLE
    .\Reset-EnterpriseAppSsoMode.ps1 -AppId '<appId>' -All
    Simulatie van de volledige opschoning.

.EXAMPLE
    .\Reset-EnterpriseAppSsoMode.ps1 -AppId '<appId>' -All -Execute
    Voert de opschoning uit.

.NOTES
    Benodigde permissies : Application.ReadWrite.All + rol Application Administrator
                           of Cloud Application Administrator.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'ByName')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string]$DisplayName,

    [Parameter(Mandatory, ParameterSetName = 'ByAppId')]
    [string]$AppId,

    [Parameter(Mandatory, ParameterSetName = 'ByObjectId')]
    [string]$ServicePrincipalId,

    [switch]$ClearIdentifiers,
    [switch]$ClearReplyUrls,
    [switch]$RemoveSamlCertificates,
    [switch]$KeepApiIdentifier,
    [switch]$All,

    [switch]$Execute,

    [string]$LogPath = (Join-Path $PSScriptRoot 'Logs')
)

$ErrorActionPreference = 'Stop'

if ($All) {
    $ClearIdentifiers       = $true
    $ClearReplyUrls         = $true
    $RemoveSamlCertificates = $true
}

if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }
$script:LogFile   = Join-Path $LogPath ("Reset-SsoMode_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
$script:Resultaat = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Bericht,
        [ValidateSet('INFO','OK','WAARSCHUWING','FOUT','ACTIE','SIMULATIE','DEBUG')]
        [string]$Niveau = 'INFO'
    )
    $regel = "[{0:yyyy-MM-dd HH:mm:ss}] [{1,-11}] {2}" -f (Get-Date), $Niveau, $Bericht
    $kleur = switch ($Niveau) {
        'OK' { 'Green' } 'WAARSCHUWING' { 'Yellow' } 'FOUT' { 'Red' }
        'ACTIE' { 'Cyan' } 'SIMULATIE' { 'Magenta' } 'DEBUG' { 'DarkGray' }
        default { 'Gray' }
    }
    Write-Host $regel -ForegroundColor $kleur
    Add-Content -Path $script:LogFile -Value $regel -Encoding UTF8
}

function Invoke-GraphPatch {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Body,
        [Parameter(Mandatory)][string]$Stap
    )

    $json = $Body | ConvertTo-Json -Depth 6 -Compress

    if (-not $Execute) {
        Write-Log ("SIMULATIE [{0}] PATCH {1}" -f $Stap, $Uri) -Niveau SIMULATIE
        Write-Log ("SIMULATIE [{0}] Body: {1}" -f $Stap, $json) -Niveau SIMULATIE
        $script:Resultaat.Add([PSCustomObject]@{ Stap = $Stap; Status = 'SIMULATIE'; Detail = $json })
        return $true
    }

    if (-not $PSCmdlet.ShouldProcess($Uri, $Stap)) {
        $script:Resultaat.Add([PSCustomObject]@{ Stap = $Stap; Status = 'OVERGESLAGEN'; Detail = '' })
        return $false
    }

    Write-Log ("Uitvoeren [{0}] Body: {1}" -f $Stap, $json) -Niveau DEBUG

    $statusCode = $null
    try {
        $respons = Invoke-MgGraphRequest -Method PATCH -Uri $Uri `
                        -Body $json -ContentType 'application/json' `
                        -SkipHttpErrorCheck -StatusCode ([ref]$statusCode) -ErrorAction Stop
    } catch {
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Log ("[{0}] MISLUKT: {1}" -f $Stap, $detail) -Niveau FOUT
        $script:Resultaat.Add([PSCustomObject]@{ Stap = $Stap; Status = 'FOUT'; Detail = $detail })
        return $false
    }

    if ($null -eq $statusCode -or ($statusCode -ge 200 -and $statusCode -lt 300)) {
        if ($respons.error) {
            Write-Log ("[{0}] MISLUKT: {1} - {2}" -f $Stap, $respons.error.code, $respons.error.message) -Niveau FOUT
            $script:Resultaat.Add([PSCustomObject]@{ Stap = $Stap; Status = 'FOUT'; Detail = $respons.error.message })
            return $false
        }
        Write-Log ("[{0}] Gelukt." -f $Stap) -Niveau OK
        $script:Resultaat.Add([PSCustomObject]@{ Stap = $Stap; Status = 'OK'; Detail = "HTTP $statusCode" })
        return $true
    }

    Write-Log ("[{0}] MISLUKT met HTTP {1}." -f $Stap, $statusCode) -Niveau FOUT
    if ($respons.error.code)                    { Write-Log ("    code       : {0}" -f $respons.error.code) -Niveau FOUT }
    if ($respons.error.message)                 { Write-Log ("    message    : {0}" -f $respons.error.message) -Niveau FOUT }
    if ($respons.error.innerError.'request-id') { Write-Log ("    request-id : {0}" -f $respons.error.innerError.'request-id') -Niveau FOUT }

    $script:Resultaat.Add([PSCustomObject]@{
        Stap = $Stap; Status = 'FOUT'
        Detail = "HTTP $statusCode | $($respons.error.code) | $($respons.error.message)"
    })
    return $false
}

# ------------------------------------------------------------------ Start

Write-Log "=================================================================" -Niveau INFO
Write-Log "Reset Enterprise Application SSO-modus naar 'Disabled'" -Niveau INFO
Write-Log ("Modus: {0}" -f $(if ($Execute) { 'UITVOEREN' } else { 'DRY-RUN (simulatie)' })) -Niveau INFO
Write-Log "=================================================================" -Niveau INFO

foreach ($module in 'Microsoft.Graph.Authentication','Microsoft.Graph.Applications') {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "Module '$module' ontbreekt. Voer uit: Install-Module $module -Scope CurrentUser" -Niveau FOUT
        return
    }
    Import-Module $module -ErrorAction Stop
}

$context = Get-MgContext
if (-not $context -or $context.Scopes -notcontains 'Application.ReadWrite.All') {
    Connect-MgGraph -Scopes 'Application.ReadWrite.All' -NoWelcome
    $context = Get-MgContext
}
Write-Log ("Verbonden met tenant '{0}' als '{1}'." -f $context.TenantId, $context.Account) -Niveau OK

# --- Objecten ophalen ---------------------------------------------------------
$sp = switch ($PSCmdlet.ParameterSetName) {
    'ByObjectId' { Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -ErrorAction SilentlyContinue }
    'ByAppId'    { Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue }
    'ByName'     { @(Get-MgServicePrincipal -Filter "displayName eq '$($DisplayName -replace "'","''")'" -All -ErrorAction SilentlyContinue) }
}

if (-not $sp) { Write-Log "Geen service principal gevonden." -Niveau FOUT; return }
if (@($sp).Count -gt 1) {
    Write-Log "Meerdere service principals gevonden. Specificeer -AppId of -ServicePrincipalId:" -Niveau FOUT
    $sp | Select-Object DisplayName, AppId, Id | Format-Table -AutoSize | Out-String | Write-Host
    return
}
$sp = @($sp)[0]
$app = Get-MgApplication -Filter "appId eq '$($sp.AppId)'" -ErrorAction SilentlyContinue | Select-Object -First 1

Write-Log ("Applicatie : {0}" -f $sp.DisplayName) -Niveau OK
Write-Log ("  AppId                 : {0}" -f $sp.AppId)
Write-Log ("  ObjectId (SP)         : {0}" -f $sp.Id)
Write-Log ("  SSO-modus             : {0}" -f $(if ([string]::IsNullOrEmpty($sp.PreferredSingleSignOnMode)) { '<Disabled>' } else { $sp.PreferredSingleSignOnMode }))
Write-Log ("  servicePrincipalNames : {0}" -f ($sp.ServicePrincipalNames -join ' | '))
Write-Log ("  replyUrls             : {0}" -f $(if ($sp.ReplyUrls) { $sp.ReplyUrls -join ' | ' } else { '<geen>' }))
Write-Log ("  loginUrl              : {0}" -f $(if ($sp.LoginUrl) { $sp.LoginUrl } else { '<geen>' }))
Write-Log ("  keyCredentials        : {0}" -f @($sp.KeyCredentials).Count)

if ($app) {
    Write-Log ("  App-registratie       : {0} (ObjectId {1})" -f $app.DisplayName, $app.Id)
    Write-Log ("  identifierUris        : {0}" -f $(if ($app.IdentifierUris) { $app.IdentifierUris -join ' | ' } else { '<geen>' }))
} else {
    Write-Log "  Geen app-registratie in deze tenant (galerij-/multi-tenant app)." -Niveau WAARSCHUWING
}

# --- Backup -------------------------------------------------------------------
$backupBestand = Join-Path $LogPath ("Backup_{0}_{1:yyyyMMdd_HHmmss}.json" -f ($sp.DisplayName -replace '[^\w\-]','_'), (Get-Date))
[PSCustomObject]@{
    Tijdstip = Get-Date; ServicePrincipalId = $sp.Id; ApplicationId = $app.Id
    AppId = $sp.AppId; DisplayName = $sp.DisplayName
    PreferredSsoMode = $sp.PreferredSingleSignOnMode
    ServicePrincipalNames = $sp.ServicePrincipalNames
    ReplyUrls = $sp.ReplyUrls; LoginUrl = $sp.LoginUrl
    RelayState = $sp.SamlSingleSignOnSettings.RelayState
    IdentifierUris = $app.IdentifierUris
    KeyCredentials = $sp.KeyCredentials | Select-Object KeyId, DisplayName, Usage, Type, EndDateTime
} | ConvertTo-Json -Depth 6 | Set-Content -Path $backupBestand -Encoding UTF8
Write-Log ("Backup: {0}" -f $backupBestand) -Niveau OK

Write-Log "-----------------------------------------------------------------" -Niveau INFO

$spUri  = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)"
$appUri = if ($app) { "https://graph.microsoft.com/v1.0/applications/$($app.Id)" } else { $null }
$neutraleUri = "api://$($sp.AppId)"

# --- STAP 1: identifierUris neutraliseren -------------------------------------
# Dit MOET voor de SSO-modus. api://<appId> vereist geen domeinverificatie en
# heeft geen trailing slash, dus Entra accepteert het altijd.
$stap1Gelukt = $true
if ($app -and $app.IdentifierUris -and ($app.IdentifierUris -notcontains $neutraleUri -or @($app.IdentifierUris).Count -gt 1)) {
    Write-Log ("Huidige identifierUris blokkeren de SSO-wijziging: {0}" -f ($app.IdentifierUris -join ', ')) -Niveau ACTIE
    $stap1Gelukt = Invoke-GraphPatch -Uri $appUri -Stap "1. identifierUris -> $neutraleUri" `
                        -Body @{ identifierUris = @($neutraleUri) }
} elseif (-not $app) {
    Write-Log "Geen app-registratie beschikbaar; stap 1 overgeslagen." -Niveau WAARSCHUWING
} else {
    Write-Log "identifierUris bevat al alleen de neutrale waarde." -Niveau INFO
}

if (-not $stap1Gelukt -and $Execute) {
    Write-Log "Stap 1 mislukt. De SSO-modus kan niet gewijzigd worden. Zie foutmelding hierboven." -Niveau FOUT
} else {

    # --- STAP 2: SSO-modus op Disabled ----------------------------------------
    $ssoGelukt = Invoke-GraphPatch -Uri $spUri -Stap '2. preferredSingleSignOnMode -> null' `
                    -Body @{ preferredSingleSignOnMode = $null }

    if (-not $ssoGelukt -and $Execute) {
        Write-Log "Stap 2 mislukt; vervolgstappen overgeslagen." -Niveau WAARSCHUWING
    } else {

        # --- STAP 3: Entity ID vrijgeven op de service principal --------------
        if ($ClearIdentifiers) {
            $teVerwijderen = @($sp.ServicePrincipalNames) | Where-Object { $_ -ne $sp.AppId }
            if ($teVerwijderen) {
                Write-Log ("Vrij te geven identifiers: {0}" -f ($teVerwijderen -join ', ')) -Niveau ACTIE
                Invoke-GraphPatch -Uri $spUri -Stap '3. servicePrincipalNames -> alleen appId' `
                    -Body @{ servicePrincipalNames = @($sp.AppId) } | Out-Null
            } else {
                Write-Log "servicePrincipalNames bevat al alleen de appId." -Niveau INFO
            }
        }

        # --- STAP 4: SAML-restanten opruimen ----------------------------------
        if ($ClearReplyUrls) {
            if ($sp.SamlSingleSignOnSettings) {
                Invoke-GraphPatch -Uri $spUri -Stap '4a. samlSingleSignOnSettings -> null' `
                    -Body @{ samlSingleSignOnSettings = $null } | Out-Null
            }
            if ($sp.LoginUrl) {
                Invoke-GraphPatch -Uri $spUri -Stap '4b. loginUrl -> null' `
                    -Body @{ loginUrl = $null } | Out-Null
            }
            if ($sp.ReplyUrls) {
                Invoke-GraphPatch -Uri $spUri -Stap '4c. replyUrls -> leeg' `
                    -Body @{ replyUrls = @() } | Out-Null
            }
        }

        # --- STAP 5: SAML signing certificaten --------------------------------
        if ($RemoveSamlCertificates -and @($sp.KeyCredentials).Count -gt 0) {
            Invoke-GraphPatch -Uri $spUri -Stap '5. keyCredentials -> leeg' `
                -Body @{ keyCredentials = @(); passwordCredentials = @() } | Out-Null
        }

        # --- STAP 6: identifierUris volledig leegmaken ------------------------
        if ($ClearIdentifiers -and $app -and -not $KeepApiIdentifier) {
            Invoke-GraphPatch -Uri $appUri -Stap '6. identifierUris -> leeg' `
                -Body @{ identifierUris = @() } | Out-Null
        } elseif ($KeepApiIdentifier) {
            Write-Log ("identifierUris blijft staan op '{0}' (-KeepApiIdentifier)." -f $neutraleUri) -Niveau INFO
        }
    }
}

# --- Verificatie --------------------------------------------------------------
if ($Execute) {
    Write-Log "-----------------------------------------------------------------" -Niveau INFO
    Write-Log "Verificatie na propagatie..." -Niveau INFO
    Start-Sleep -Seconds 5

    $controle = Get-MgServicePrincipal -ServicePrincipalId $sp.Id
    $ssoLeeg  = [string]::IsNullOrEmpty($controle.PreferredSingleSignOnMode)
    Write-Log ("  SSO-modus             : {0}" -f $(if ($ssoLeeg) { '<Disabled>' } else { $controle.PreferredSingleSignOnMode })) -Niveau $(if ($ssoLeeg) { 'OK' } else { 'WAARSCHUWING' })
    Write-Log ("  servicePrincipalNames : {0}" -f ($controle.ServicePrincipalNames -join ' | '))
    Write-Log ("  replyUrls             : {0}" -f $(if ($controle.ReplyUrls) { $controle.ReplyUrls -join ' | ' } else { '<geen>' }))
    Write-Log ("  loginUrl              : {0}" -f $(if ($controle.LoginUrl) { $controle.LoginUrl } else { '<geen>' }))
    Write-Log ("  keyCredentials        : {0}" -f @($controle.KeyCredentials).Count)
    if ($app) {
        $controleApp = Get-MgApplication -ApplicationId $app.Id
        Write-Log ("  identifierUris        : {0}" -f $(if ($controleApp.IdentifierUris) { $controleApp.IdentifierUris -join ' | ' } else { '<geen>' }))
    }
}

# --- Samenvatting -------------------------------------------------------------
Write-Log "-----------------------------------------------------------------" -Niveau INFO
Write-Log "Samenvatting per stap:" -Niveau INFO
foreach ($r in $script:Resultaat) {
    $niveau = switch ($r.Status) { 'OK' { 'OK' } 'FOUT' { 'FOUT' } 'SIMULATIE' { 'SIMULATIE' } default { 'WAARSCHUWING' } }
    Write-Log ("  {0,-45} {1}" -f $r.Stap, $r.Status) -Niveau $niveau
}

$csvBestand = Join-Path $LogPath ("Resultaat_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
$script:Resultaat | Export-Csv -Path $csvBestand -Delimiter ';' -Encoding UTF8 -NoTypeInformation
Write-Log ("Resultaat-CSV: {0}" -f $csvBestand) -Niveau OK

Write-Log "=================================================================" -Niveau INFO
Write-Log ("Gereed. Logbestand: {0}" -f $script:LogFile) -Niveau OK
if (-not $Execute) {
    Write-Log "Dit was een simulatie. Voeg -Execute toe om door te voeren." -Niveau WAARSCHUWING
}
