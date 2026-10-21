#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Analyseert Entra ID sign-in logs op MFA-gebruik over een opgegeven periode (standaard 30 dagen).

.DESCRIPTION
    Dit script is volledig read-only: er worden geen wijzigingen in de tenant doorgevoerd.

    Het script:
      1. Bepaalt de tenant-baseline (Security Defaults aan/uit, aantal Conditional Access policies).
      2. Haalt de interactieve sign-in logs op via Microsoft Graph.
      3. Classificeert elke succesvolle login als single-factor of multi-factor, en maakt
         onderscheid tussen "MFA daadwerkelijk uitgevoerd" en "MFA overgenomen uit token/claim".
      4. Aggregeert per gebruiker: wie gebruikt wel MFA, wie gedeeltelijk, wie nooit.
      5. Inventariseert optioneel de geregistreerde authenticatiemethoden per gebruiker
         (werkt ook zonder Entra ID P1/P2).

    LET OP - licentievereiste:
    Het ophalen van sign-in logs via Graph (/auditLogs/signIns) vereist Entra ID P1 of P2.
    Tenants die op Security Defaults draaien zijn vaak licentievrij. Als de call faalt met een
    licentiefout schakelt het script automatisch over op de methode-inventarisatie, die wel
    beschikbaar is op de gratis tier.

.PARAMETER TenantId
    Tenant-id of domeinnaam van de klant. Optioneel wanneer er al een Graph-sessie actief is.

.PARAMETER DagenTerug
    Aantal dagen terug te analyseren. Maximaal 30 (retentie van de sign-in logs bij P1/P2).

.PARAMETER OutputPad
    Map waarin de CSV-rapportages worden weggeschreven.

.PARAMETER IncludeNonInteractive
    Neemt ook non-interactieve sign-ins mee (service principals namens gebruikers, tokenrefresh).
    Levert veel volume op en vertekent het MFA-percentage; standaard uit.

.PARAMETER MethodInventory
    Inventariseert per gebruiker de geregistreerde authenticatiemethoden. Wordt automatisch
    geactiveerd wanneer de sign-in logs niet beschikbaar zijn.

.PARAMETER SkipConnect
    Gebruikt de bestaande Graph-sessie zonder opnieuw te verbinden.

.EXAMPLE
    .\Get-EntraIdMfaSignInReport.ps1 -TenantId klant.onmicrosoft.com

.EXAMPLE
    .\Get-EntraIdMfaSignInReport.ps1 -DagenTerug 14 -MethodInventory -OutputPad 'C:\Rapportages\Contoso'

.NOTES
    Benodigde Graph-permissies (delegated of application):
        AuditLog.Read.All
        Directory.Read.All
        Policy.Read.All
        UserAuthenticationMethod.Read.All   (alleen voor -MethodInventory)
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [ValidateRange(1, 30)]
    [int]$DagenTerug = 30,

    [Parameter()]
    [string]$OutputPad = (Join-Path -Path $PWD -ChildPath 'EntraID-MFA-Rapportage'),

    [Parameter()]
    [switch]$IncludeNonInteractive,

    [Parameter()]
    [switch]$MethodInventory,

    [Parameter()]
    [switch]$UseBeta,

    [Parameter()]
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Hulpfuncties ---------------------------------------------------------------------

$script:LogPad = $null

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Bericht,
        [ValidateSet('INFO', 'OK', 'WAARSCHUWING', 'FOUT', 'STAP')][string]$Niveau = 'INFO'
    )

    $tijd = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $regel = '[{0}] [{1,-12}] {2}' -f $tijd, $Niveau, $Bericht

    $kleur = switch ($Niveau) {
        'OK'           { 'Green' }
        'WAARSCHUWING' { 'Yellow' }
        'FOUT'         { 'Red' }
        'STAP'         { 'Cyan' }
        default        { 'Gray' }
    }

    Write-Host $regel -ForegroundColor $kleur

    if ($script:LogPad) {
        Add-Content -Path $script:LogPad -Value $regel -Encoding utf8
    }
}

function Get-Eigenschap {
    <#  Veilig uitlezen van een (geneste) eigenschap die kan ontbreken in de Graph-respons. #>
    param(
        [Parameter()]$Object,
        [Parameter(Mandatory)][string]$Naam,
        [Parameter()]$Standaard = $null
    )

    if ($null -eq $Object) { return $Standaard }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Naam)) { return $Object[$Naam] }
        return $Standaard
    }

    $eigenschap = $Object.PSObject.Properties[$Naam]
    if ($null -eq $eigenschap -or $null -eq $eigenschap.Value) { return $Standaard }
    return $eigenschap.Value
}

function Invoke-GraphMetRetry {
    <#  Graph-call met afhandeling van throttling (429) en tijdelijke serverfouten (5xx). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxPogingen = 5
    )

    $poging = 0
    while ($true) {
        $poging++
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = $null }

            $herstelbaar = $statusCode -in @(429, 500, 502, 503, 504)

            if (-not $herstelbaar -or $poging -ge $MaxPogingen) {
                throw
            }

            $wachttijd = [math]::Min([math]::Pow(2, $poging), 60)
            try {
                $retryAfter = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds
                if ($retryAfter -gt 0) { $wachttijd = $retryAfter }
            }
            catch { }

            Write-Log -Niveau WAARSCHUWING -Bericht ("Graph gaf status {0}. Poging {1}/{2}, opnieuw over {3} seconden." -f $statusCode, $poging, $MaxPogingen, [int]$wachttijd)
            Start-Sleep -Seconds ([int]$wachttijd)
        }
    }
}

function Get-GraphCollectie {
    <#  Haalt een volledige Graph-collectie op inclusief paging. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Omschrijving = 'records'
    )

    $resultaat = [System.Collections.Generic.List[object]]::new()
    $volgende = $Uri
    $pagina = 0

    while ($volgende) {
        $pagina++
        $respons = Invoke-GraphMetRetry -Uri $volgende

        $waarden = Get-Eigenschap -Object $respons -Naam 'value'
        if ($waarden) {
            foreach ($item in @($waarden)) { $resultaat.Add($item) }
        }

        $volgende = Get-Eigenschap -Object $respons -Naam '@odata.nextLink'

        if ($volgende) {
            Write-Log -Bericht ("Pagina {0} verwerkt - {1} {2} opgehaald tot nu toe." -f $pagina, $resultaat.Count, $Omschrijving)
        }
    }

    return $resultaat
}

function Test-LicentieFout {
    param([Parameter(Mandatory)]$FoutRecord)

    $tekst = $FoutRecord.ToString()
    return ($tekst -match 'Authentication_RequestFromUnsupportedUserRole|premium|license|Tenant is not a B2C|RequestBodyRead|Feature is not available')
}

#endregion -------------------------------------------------------------------------------

#region Voorbereiding --------------------------------------------------------------------

$startTijd = Get-Date
$stempel = $startTijd.ToString('yyyyMMdd-HHmmss')

if (-not (Test-Path -Path $OutputPad)) {
    New-Item -Path $OutputPad -ItemType Directory -Force | Out-Null
}
$script:LogPad = Join-Path -Path $OutputPad -ChildPath ("MFA-analyse-{0}.log" -f $stempel)

Write-Log -Niveau STAP -Bericht '================================================================'
Write-Log -Niveau STAP -Bericht ' Entra ID MFA sign-in analyse (read-only, geen wijzigingen)'
Write-Log -Niveau STAP -Bericht '================================================================'
Write-Log -Bericht ("Analyseperiode: afgelopen {0} dagen." -f $DagenTerug)
Write-Log -Bericht ("Uitvoermap    : {0}" -f $OutputPad)

$benodigdeScopes = @(
    'AuditLog.Read.All'
    'Directory.Read.All'
    'Policy.Read.All'
)
if ($MethodInventory) { $benodigdeScopes += 'UserAuthenticationMethod.Read.All' }

if (-not $SkipConnect) {
    Write-Log -Niveau STAP -Bericht 'Verbinden met Microsoft Graph...'
    $connectParams = @{ Scopes = $benodigdeScopes; NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    try {
        Connect-MgGraph @connectParams
    }
    catch {
        Write-Log -Niveau FOUT -Bericht ("Verbinden met Graph is mislukt: {0}" -f $_.Exception.Message)
        throw
    }
}

$context = Get-MgContext
if (-not $context) {
    Write-Log -Niveau FOUT -Bericht 'Geen actieve Graph-sessie gevonden. Start het script zonder -SkipConnect.'
    throw 'Geen Graph-context.'
}

Write-Log -Niveau OK -Bericht ("Verbonden met tenant {0} als {1}." -f $context.TenantId, ($context.Account ?? $context.ClientId))

$ontbrekendeScopes = @($benodigdeScopes | Where-Object { $_ -notin $context.Scopes })
if ($ontbrekendeScopes.Count -gt 0) {
    Write-Log -Niveau WAARSCHUWING -Bericht ("Mogelijk ontbrekende permissies: {0}" -f ($ontbrekendeScopes -join ', '))
}

#endregion -------------------------------------------------------------------------------

#region Tenant-baseline ------------------------------------------------------------------

Write-Log -Niveau STAP -Bericht 'Stap 1 - Tenant-baseline bepalen.'

$securityDefaultsAan = $null
try {
    $sd = Invoke-GraphMetRetry -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
    $securityDefaultsAan = [bool](Get-Eigenschap -Object $sd -Naam 'isEnabled' -Standaard $false)
    if ($securityDefaultsAan) {
        Write-Log -Niveau OK -Bericht 'Security Defaults: INGESCHAKELD.'
    }
    else {
        Write-Log -Niveau WAARSCHUWING -Bericht 'Security Defaults: UITGESCHAKELD.'
    }
}
catch {
    Write-Log -Niveau WAARSCHUWING -Bericht ("Status Security Defaults kon niet worden opgehaald: {0}" -f $_.Exception.Message)
}

$caPolicies = @()
try {
    $caPolicies = @(Get-GraphCollectie -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Omschrijving 'CA-policies')
    $caActief = @($caPolicies | Where-Object { (Get-Eigenschap -Object $_ -Naam 'state') -eq 'enabled' })
    Write-Log -Bericht ("Conditional Access: {0} policies gevonden, waarvan {1} actief." -f $caPolicies.Count, $caActief.Count)
}
catch {
    Write-Log -Niveau WAARSCHUWING -Bericht 'Conditional Access policies niet opvraagbaar (waarschijnlijk geen Entra ID P1/P2).'
}

# --- CA-policy configuratie uitwerken -------------------------------------------------
# Microsoft-managed policies (createdBy = Microsoft) kunnen door beheerders worden
# aangepast. Uitzonderingen daarin zijn de meest voorkomende oorzaak van "MFA wordt
# niet afgedwongen". Daarom worden include/exclude scopes expliciet opgelost naar namen.

$script:NaamCache = @{}

function Resolve-DirectoryObject {
    param([Parameter(Mandatory)][string]$Id)

    if ([string]::IsNullOrWhiteSpace($Id)) { return '' }
    if ($Id -in @('All', 'None', 'GuestsOrExternalUsers')) { return $Id }
    if ($script:NaamCache.ContainsKey($Id)) { return $script:NaamCache[$Id] }

    $naam = $Id
    foreach ($pad in @('directoryObjects', 'directoryRoleTemplates')) {
        try {
            $obj = Invoke-GraphMetRetry -Uri ("https://graph.microsoft.com/v1.0/{0}/{1}" -f $pad, $Id)
            $weergave = Get-Eigenschap -Object $obj -Naam 'displayName' -Standaard $null
            $upn      = Get-Eigenschap -Object $obj -Naam 'userPrincipalName' -Standaard $null
            if ($upn)      { $naam = $upn; break }
            if ($weergave) { $naam = $weergave; break }
        }
        catch { continue }
    }

    $script:NaamCache[$Id] = $naam
    return $naam
}

function Format-Scope {
    param($Waarden)

    $lijst = @(@($Waarden) | Where-Object { $_ })
    if ($lijst.Count -eq 0) { return '' }
    return (($lijst | ForEach-Object { Resolve-DirectoryObject -Id ([string]$_) }) -join ', ')
}

$caConfiguratie = [System.Collections.Generic.List[object]]::new()

foreach ($p in $caPolicies) {

    $condities    = Get-Eigenschap -Object $p -Naam 'conditions'
    $gebruikers   = Get-Eigenschap -Object $condities -Naam 'users'
    $applicaties  = Get-Eigenschap -Object $condities -Naam 'applications'
    $grant        = Get-Eigenschap -Object $p -Naam 'grantControls'
    $grantLijst   = @(Get-Eigenschap -Object $grant -Naam 'builtInControls' -Standaard @())

    $uitgeslotenGebruikers = Format-Scope -Waarden (Get-Eigenschap -Object $gebruikers -Naam 'excludeUsers' -Standaard @())
    $uitgeslotenGroepen    = Format-Scope -Waarden (Get-Eigenschap -Object $gebruikers -Naam 'excludeGroups' -Standaard @())
    $uitgeslotenRollen     = Format-Scope -Waarden (Get-Eigenschap -Object $gebruikers -Naam 'excludeRoles'  -Standaard @())

    $aantalUitzonderingen = @(
        @(Get-Eigenschap -Object $gebruikers -Naam 'excludeUsers'  -Standaard @()).Count
        @(Get-Eigenschap -Object $gebruikers -Naam 'excludeGroups' -Standaard @()).Count
        @(Get-Eigenschap -Object $gebruikers -Naam 'excludeRoles'  -Standaard @()).Count
    ) | Measure-Object -Sum | Select-Object -ExpandProperty Sum

    $caConfiguratie.Add([pscustomobject]@{
        PolicyNaam             = Get-Eigenschap -Object $p -Naam 'displayName' -Standaard ''
        PolicyId               = Get-Eigenschap -Object $p -Naam 'id' -Standaard ''
        Status                 = Get-Eigenschap -Object $p -Naam 'state' -Standaard ''
        VereistMFA             = [bool](@($grantLijst | Where-Object { $_ -match 'mfa' }).Count)
        GrantControls          = ($grantLijst -join ', ')
        GrantOperator          = Get-Eigenschap -Object $grant -Naam 'operator' -Standaard ''
        IncludeGebruikers      = ((@(Get-Eigenschap -Object $gebruikers -Naam 'includeUsers'  -Standaard @())) -join ', ')
        IncludeGroepen         = (Format-Scope -Waarden (Get-Eigenschap -Object $gebruikers -Naam 'includeGroups' -Standaard @()))
        IncludeRollen          = (Format-Scope -Waarden (Get-Eigenschap -Object $gebruikers -Naam 'includeRoles'  -Standaard @()))
        AantalUitzonderingen   = [int]$aantalUitzonderingen
        UitgeslotenGebruikers  = $uitgeslotenGebruikers
        UitgeslotenGroepen     = $uitgeslotenGroepen
        UitgeslotenRollen      = $uitgeslotenRollen
        IncludeApplicaties     = ((@(Get-Eigenschap -Object $applicaties -Naam 'includeApplications' -Standaard @())) -join ', ')
        ExcludeApplicaties     = ((@(Get-Eigenschap -Object $applicaties -Naam 'excludeApplications' -Standaard @())) -join ', ')
        ClientAppTypes         = ((@(Get-Eigenschap -Object $condities -Naam 'clientAppTypes' -Standaard @())) -join ', ')
        GebruikersRisico       = ((@(Get-Eigenschap -Object $condities -Naam 'userRiskLevels' -Standaard @())) -join ', ')
        SignInRisico           = ((@(Get-Eigenschap -Object $condities -Naam 'signInRiskLevels' -Standaard @())) -join ', ')
        Aangemaakt             = Get-Eigenschap -Object $p -Naam 'createdDateTime' -Standaard ''
        Gewijzigd              = Get-Eigenschap -Object $p -Naam 'modifiedDateTime' -Standaard ''
    })
}

foreach ($c in ($caConfiguratie | Where-Object { $_.VereistMFA -and $_.Status -eq 'enabled' -and $_.AantalUitzonderingen -gt 0 })) {
    Write-Log -Niveau WAARSCHUWING -Bericht ("Policy '{0}' vereist MFA maar kent {1} uitzondering(en)." -f $c.PolicyNaam, $c.AantalUitzonderingen)
}

foreach ($c in ($caConfiguratie | Where-Object { $_.VereistMFA -and $_.Status -eq 'enabledForReportingButNotEnforced' })) {
    Write-Log -Niveau WAARSCHUWING -Bericht ("Policy '{0}' staat in report-only en dwingt dus niets af." -f $c.PolicyNaam)
}

#endregion -------------------------------------------------------------------------------

#region Sign-in logs ophalen -------------------------------------------------------------

Write-Log -Niveau STAP -Bericht 'Stap 2 - Sign-in logs ophalen.'

$vanaf = (Get-Date).ToUniversalTime().AddDays(-$DagenTerug)
$vanafTekst = $vanaf.ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Log -Bericht ("Filter vanaf: {0} (UTC)." -f $vanafTekst)

$signIns = [System.Collections.Generic.List[object]]::new()
$signInsBeschikbaar = $true

$filters = [ordered]@{
    'interactieve' = "createdDateTime ge $vanafTekst"
}
if ($IncludeNonInteractive) {
    $filters['non-interactieve'] = "createdDateTime ge $vanafTekst and signInEventTypes/any(t: t eq 'nonInteractiveUser')"
}

$apiVersie = if ($UseBeta) { 'beta' } else { 'v1.0' }
Write-Log -Bericht ("Graph endpoint: {0}/auditLogs/signIns" -f $apiVersie)

# Het signIns-endpoint geeft in een collectie niet alle eigenschappen standaard terug.
# authenticationRequirement en authenticationDetails zijn daar de belangrijkste
# slachtoffers van; zonder expliciete $select komen ze leeg binnen en wordt elke login
# ten onrechte als single-factor geclassificeerd.
$selectVelden = @(
    'id', 'createdDateTime', 'userPrincipalName', 'userDisplayName', 'userId'
    'appDisplayName', 'appId', 'resourceDisplayName', 'clientAppUsed', 'correlationId'
    'ipAddress', 'status', 'deviceDetail', 'location'
    'conditionalAccessStatus', 'appliedConditionalAccessPolicies'
    'authenticationRequirement', 'authenticationDetails'
    'riskLevelDuringSignIn', 'riskState', 'isInteractive'
) -join ','

foreach ($sleutel in $filters.Keys) {
    $filter = [uri]::EscapeDataString($filters[$sleutel])
    $basisUri = "https://graph.microsoft.com/$apiVersie/auditLogs/signIns?`$filter=$filter&`$top=1000"
    $uri = "$basisUri&`$select=$selectVelden"

    try {
        Write-Log -Bericht ("Ophalen van {0} sign-ins (met expliciete selectlijst)..." -f $sleutel)
        $batch = Get-GraphCollectie -Uri $uri -Omschrijving 'sign-ins'
        foreach ($item in $batch) { $signIns.Add($item) }
        Write-Log -Niveau OK -Bericht ("{0} {1} sign-ins opgehaald." -f $batch.Count, $sleutel)
    }
    catch {
        if (Test-LicentieFout -FoutRecord $_) {
            Write-Log -Niveau WAARSCHUWING -Bericht 'Sign-in logs zijn niet opvraagbaar via Graph. Dit vereist Entra ID P1 of P2.'
            Write-Log -Niveau WAARSCHUWING -Bericht 'Er wordt overgeschakeld op de inventarisatie van geregistreerde MFA-methoden.'
            $signInsBeschikbaar = $false
            $MethodInventory = $true
            break
        }

        # Een afgewezen $select mag de analyse niet blokkeren: opnieuw zonder select.
        Write-Log -Niveau WAARSCHUWING -Bericht ("Query met selectlijst is afgewezen ({0}). Opnieuw zonder select." -f $_.Exception.Message)
        try {
            $batch = Get-GraphCollectie -Uri $basisUri -Omschrijving 'sign-ins'
            foreach ($item in $batch) { $signIns.Add($item) }
            Write-Log -Niveau OK -Bericht ("{0} {1} sign-ins opgehaald zonder select." -f $batch.Count, $sleutel)
        }
        catch {
            Write-Log -Niveau FOUT -Bericht ("Ophalen sign-ins mislukt: {0}" -f $_.Exception.Message)
            throw
        }
    }
}

# Diagnostiek: laat zien welke eigenschappen Graph daadwerkelijk teruggeeft.
if ($signIns.Count -gt 0) {
    $aanwezig = @($signIns[0].PSObject.Properties.Name | Sort-Object)
    Write-Log -Bericht ("Eigenschappen in respons: {0}" -f ($aanwezig -join ', '))

    foreach ($vereist in @('authenticationRequirement', 'authenticationDetails')) {
        if ($vereist -notin $aanwezig) {
            Write-Log -Niveau WAARSCHUWING -Bericht ("Eigenschap '{0}' ontbreekt in de respons. Classificatie valt terug op de overige velden." -f $vereist)
        }
    }
}

#endregion -------------------------------------------------------------------------------

#region Sign-ins classificeren -----------------------------------------------------------

$signInRecords = [System.Collections.Generic.List[object]]::new()

if ($signInsBeschikbaar -and $signIns.Count -gt 0) {

    Write-Log -Niveau STAP -Bericht 'Stap 3 - Sign-ins classificeren op MFA-gebruik.'

    foreach ($s in $signIns) {

        $status      = Get-Eigenschap -Object $s -Naam 'status'
        $foutcode    = [int](Get-Eigenschap -Object $status -Naam 'errorCode' -Standaard 0)
        $succesvol   = ($foutcode -eq 0)

        $details     = @(Get-Eigenschap -Object $s -Naam 'authenticationDetails' -Standaard @())

        # Stappen waarin daadwerkelijk een tweede factor is uitgevoerd (dus geen claim uit token).
        $mfaStappen = @(
            $details | Where-Object {
                $req    = Get-Eigenschap -Object $_ -Naam 'authenticationStepRequirement' -Standaard ''
                $gelukt = Get-Eigenschap -Object $_ -Naam 'succeeded' -Standaard $false
                $req -match 'Multifactor|MFA' -and $gelukt -eq $true
            }
        )

        # MFA-bepaling uit meerdere bronnen, in volgorde van betrouwbaarheid:
        #   1. authenticationRequirement  - het expliciete oordeel van Entra ID
        #   2. authenticationDetails      - een geslaagde multifactor-stap in de sign-in
        #   3. conditionalAccessStatus    - een MFA-policy die met succes is toegepast
        # Zonder deze terugval werd elke login als single-factor geclassificeerd zodra
        # Graph het eerste veld niet meelevert.
        $authReq = Get-Eigenschap -Object $s -Naam 'authenticationRequirement' -Standaard ''
        $caStatus = Get-Eigenschap -Object $s -Naam 'conditionalAccessStatus' -Standaard ''

        if ($authReq -eq 'multiFactorAuthentication') {
            $mfaGebruikt = $true
            $mfaBron     = 'authenticationRequirement'
        }
        elseif ($authReq -eq 'singleFactorAuthentication') {
            $mfaGebruikt = $false
            $mfaBron     = 'authenticationRequirement'
        }
        elseif ($mfaStappen.Count -gt 0) {
            $mfaGebruikt = $true
            $mfaBron     = 'authenticationDetails'
        }
        elseif ($details.Count -gt 0) {
            $mfaGebruikt = $false
            $mfaBron     = 'authenticationDetails'
        }
        else {
            $mfaGebruikt = $false
            $mfaBron     = 'onbepaald'
        }

        if ([string]::IsNullOrWhiteSpace($authReq)) { $authReq = 'niet geleverd door Graph' }

        $viaToken = $false
        $verseMfa = $false
        foreach ($stap in $mfaStappen) {
            $resultaat = Get-Eigenschap -Object $stap -Naam 'authenticationStepResultDetail' -Standaard ''
            if ($resultaat -match 'satisfied by claim|Previously satisfied|satisfied by token') {
                $viaToken = $true
            }
            else {
                $verseMfa = $true
            }
        }

        $methoden = @(
            $mfaStappen |
                ForEach-Object { Get-Eigenschap -Object $_ -Naam 'authenticationMethod' -Standaard $null } |
                Where-Object { $_ } |
                Sort-Object -Unique
        )

        $locatie   = Get-Eigenschap -Object $s -Naam 'location'
        $apparaat  = Get-Eigenschap -Object $s -Naam 'deviceDetail'

        # Welke CA policies zijn op deze sign-in toegepast, en met welk resultaat?
        # result-waarden: success | failure | notApplied | notEnabled | reportOnly* | unknown
        $caToegepast = @(Get-Eigenschap -Object $s -Naam 'appliedConditionalAccessPolicies' -Standaard @())

        $caMfaPolicies = @(
            $caToegepast | Where-Object {
                $controls = @(Get-Eigenschap -Object $_ -Naam 'enforcedGrantControls' -Standaard @())
                @($controls | Where-Object { $_ -match 'Mfa|MultiFactor' }).Count -gt 0
            }
        )

        $caMfaGeslaagd = @($caMfaPolicies | Where-Object { (Get-Eigenschap -Object $_ -Naam 'result' -Standaard '') -eq 'success' })
        $caMfaNietToegepast = @($caMfaPolicies | Where-Object { (Get-Eigenschap -Object $_ -Naam 'result' -Standaard '') -in @('notApplied', 'notEnabled') })
        $caReportOnly = @($caMfaPolicies | Where-Object { (Get-Eigenschap -Object $_ -Naam 'result' -Standaard '') -match '^reportOnly' })

        $caResultaten = @(
            $caToegepast | ForEach-Object {
                '{0}={1}' -f (Get-Eigenschap -Object $_ -Naam 'displayName' -Standaard '?'), (Get-Eigenschap -Object $_ -Naam 'result' -Standaard '?')
            }
        )

        # Verklaring waarom een succesvolle login toch single-factor was.
        $redenGeenMfa = ''
        if ($succesvol -and -not $mfaGebruikt) {
            $redenGeenMfa = if ($caMfaPolicies.Count -eq 0) {
                'Geen enkele MFA-policy geldt voor deze app/gebruiker'
            }
            elseif ($caMfaNietToegepast.Count -eq $caMfaPolicies.Count) {
                'MFA-policy(s) niet toegepast: uitgesloten of buiten scope'
            }
            elseif ($caReportOnly.Count -gt 0) {
                'MFA-policy staat in report-only'
            }
            else {
                'Onbepaald - zie kolom CAResultaten'
            }
        }

        $tijdstipU = Get-Eigenschap -Object $s -Naam 'createdDateTime'
        $tijdstip  = if ($tijdstipU) { ([datetime]$tijdstipU).ToLocalTime() } else { $null }

        $signInRecords.Add([pscustomobject]@{
            Tijdstip                  = $tijdstip
            UPN                       = Get-Eigenschap -Object $s -Naam 'userPrincipalName' -Standaard ''
            Weergavenaam              = Get-Eigenschap -Object $s -Naam 'userDisplayName' -Standaard ''
            GebruikerId               = Get-Eigenschap -Object $s -Naam 'userId' -Standaard ''
            Applicatie                = Get-Eigenschap -Object $s -Naam 'appDisplayName' -Standaard ''
            Resource                  = Get-Eigenschap -Object $s -Naam 'resourceDisplayName' -Standaard ''
            ClientApp                 = Get-Eigenschap -Object $s -Naam 'clientAppUsed' -Standaard ''
            Succesvol                 = $succesvol
            Foutcode                  = $foutcode
            Foutomschrijving          = Get-Eigenschap -Object $status -Naam 'failureReason' -Standaard ''
            AuthenticatieVereiste     = $authReq
            MFABron                   = $mfaBron
            MFAGebruikt               = $mfaGebruikt
            MFAUitgevoerdDezeSessie   = $verseMfa
            MFAUitTokenOfClaim        = $viaToken
            MFAMethoden               = ($methoden -join ', ')
            ConditionalAccessStatus   = Get-Eigenschap -Object $s -Naam 'conditionalAccessStatus' -Standaard ''
            CAPolicyCount             = $caToegepast.Count
            CAMfaPolicyCount          = $caMfaPolicies.Count
            CAMfaPolicyGeslaagd       = $caMfaGeslaagd.Count
            CAMfaPolicyNietToegepast  = $caMfaNietToegepast.Count
            CAResultaten              = ($caResultaten -join ' | ')
            RedenGeenMFA              = $redenGeenMfa
            RisicoNiveau              = Get-Eigenschap -Object $s -Naam 'riskLevelDuringSignIn' -Standaard ''
            IPAdres                   = Get-Eigenschap -Object $s -Naam 'ipAddress' -Standaard ''
            Land                      = Get-Eigenschap -Object $locatie -Naam 'countryOrRegion' -Standaard ''
            Stad                      = Get-Eigenschap -Object $locatie -Naam 'city' -Standaard ''
            Besturingssysteem         = Get-Eigenschap -Object $apparaat -Naam 'operatingSystem' -Standaard ''
            Browser                   = Get-Eigenschap -Object $apparaat -Naam 'browser' -Standaard ''
            ApparaatCompliant         = Get-Eigenschap -Object $apparaat -Naam 'isCompliant' -Standaard ''
            SignInType                = ((@(Get-Eigenschap -Object $s -Naam 'signInEventTypes' -Standaard @('interactiveUser'))) -join ', ')
            CorrelatieId              = Get-Eigenschap -Object $s -Naam 'correlationId' -Standaard ''
        })
    }

    Write-Log -Niveau OK -Bericht ("{0} sign-ins geclassificeerd." -f $signInRecords.Count)
}
elseif ($signInsBeschikbaar) {
    Write-Log -Niveau WAARSCHUWING -Bericht 'Er zijn geen sign-ins gevonden in de opgegeven periode.'
}

#endregion -------------------------------------------------------------------------------

#region Aggregatie per gebruiker ---------------------------------------------------------

$gebruikerSamenvatting = [System.Collections.Generic.List[object]]::new()

if ($signInRecords.Count -gt 0) {

    Write-Log -Niveau STAP -Bericht 'Stap 4 - Aggregeren per gebruiker.'

    $gesorteerd = $signInRecords | Where-Object { $_.UPN } | Group-Object -Property UPN

    foreach ($groep in $gesorteerd) {

        $alle      = @($groep.Group)
        $succes    = @($alle | Where-Object { $_.Succesvol })
        $metMfa    = @($succes | Where-Object { $_.MFAGebruikt })
        $zonderMfa = @($succes | Where-Object { -not $_.MFAGebruikt })

        $percentage = if ($succes.Count -gt 0) {
            [math]::Round(($metMfa.Count / $succes.Count) * 100, 1)
        }
        else { 0 }

        $classificatie = if ($succes.Count -eq 0) { 'Geen succesvolle logins' }
                         elseif ($metMfa.Count -eq 0) { 'NOOIT MFA' }
                         elseif ($zonderMfa.Count -eq 0) { 'Altijd MFA' }
                         else { 'Gedeeltelijk MFA' }

        $methoden = @(
            $metMfa | ForEach-Object { $_.MFAMethoden -split ',\s*' } |
                Where-Object { $_ } | Sort-Object -Unique
        )

        $appsZonderMfa = @(
            $zonderMfa | Group-Object -Property Applicatie |
                Sort-Object -Property Count -Descending |
                Select-Object -First 5 |
                ForEach-Object { '{0} ({1}x)' -f $_.Name, $_.Count }
        )

        # Onder Set-StrictMode levert een property-lookup op $null een terminating error op.
        # Gebruikers zonder enige MFA-login geven een lege collectie, dus expliciet afvangen.
        $recordLaatsteMfa    = $metMfa | Sort-Object -Property Tijdstip -Descending | Select-Object -First 1
        $recordEersteLogin   = $alle   | Sort-Object -Property Tijdstip            | Select-Object -First 1
        $recordLaatsteLogin  = $alle   | Sort-Object -Property Tijdstip -Descending | Select-Object -First 1

        $laatsteMfa    = if ($recordLaatsteMfa)   { $recordLaatsteMfa.Tijdstip }   else { $null }
        $eersteLogin   = if ($recordEersteLogin)  { $recordEersteLogin.Tijdstip }  else { $null }
        $laatsteLogin  = if ($recordLaatsteLogin) { $recordLaatsteLogin.Tijdstip } else { $null }

        $gebruikerSamenvatting.Add([pscustomobject]@{
            UPN                     = $groep.Name
            Weergavenaam            = $alle[0].Weergavenaam
            Classificatie           = $classificatie
            TotaalLogins            = $alle.Count
            SuccesvolleLogins       = $succes.Count
            LoginsMetMFA            = $metMfa.Count
            LoginsZonderMFA         = $zonderMfa.Count
            PercentageMFA           = $percentage
            MFAOoitUitgevoerd       = [bool](@($metMfa | Where-Object { $_.MFAUitgevoerdDezeSessie }).Count)
            GebruikteMFAMethoden    = ($methoden -join ', ')
            LaatsteMFALogin         = $laatsteMfa
            EersteLogin             = $eersteLogin
            LaatsteLogin            = $laatsteLogin
            TopAppsZonderMFA        = ($appsZonderMfa -join ' | ')
            LegacyClientGebruikt    = [bool](@($alle | Where-Object { $_.ClientApp -match 'IMAP|POP|SMTP|MAPI|Exchange ActiveSync|Other clients|Authenticated SMTP' }).Count)
        })
    }

    $gebruikerSamenvatting = @($gebruikerSamenvatting | Sort-Object -Property PercentageMFA, UPN)
    Write-Log -Niveau OK -Bericht ("{0} unieke gebruikers geaggregeerd." -f $gebruikerSamenvatting.Count)
}

#endregion -------------------------------------------------------------------------------

#region Inventarisatie geregistreerde methoden -------------------------------------------

$methodeInventaris = [System.Collections.Generic.List[object]]::new()

if ($MethodInventory) {

    Write-Log -Niveau STAP -Bericht 'Stap 5 - Geregistreerde authenticatiemethoden inventariseren.'

    $registratieViaRapport = $false
    try {
        $reg = Get-GraphCollectie -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails' -Omschrijving 'registratiedetails'

        foreach ($r in $reg) {
            $methodeInventaris.Add([pscustomobject]@{
                UPN                       = Get-Eigenschap -Object $r -Naam 'userPrincipalName' -Standaard ''
                Weergavenaam              = Get-Eigenschap -Object $r -Naam 'userDisplayName' -Standaard ''
                IsAdmin                   = Get-Eigenschap -Object $r -Naam 'isAdmin' -Standaard $false
                MFAGeregistreerd          = Get-Eigenschap -Object $r -Naam 'isMfaRegistered' -Standaard $false
                MFACapabel                = Get-Eigenschap -Object $r -Naam 'isMfaCapable' -Standaard $false
                SSPRGeregistreerd         = Get-Eigenschap -Object $r -Naam 'isSsprRegistered' -Standaard $false
                PasswordlessCapabel       = Get-Eigenschap -Object $r -Naam 'isPasswordlessCapable' -Standaard $false
                StandaardMethode          = Get-Eigenschap -Object $r -Naam 'defaultMfaMethod' -Standaard ''
                GeregistreerdeMethoden    = ((@(Get-Eigenschap -Object $r -Naam 'methodsRegistered' -Standaard @())) -join ', ')
                Bron                      = 'userRegistrationDetails'
            })
        }

        $registratieViaRapport = $true
        Write-Log -Niveau OK -Bericht ("{0} registratiedetails opgehaald via het rapportage-endpoint." -f $methodeInventaris.Count)
    }
    catch {
        Write-Log -Niveau WAARSCHUWING -Bericht 'Rapportage-endpoint niet beschikbaar; per gebruiker uitlezen van authenticatiemethoden.'
    }

    if (-not $registratieViaRapport) {
        try {
            $gebruikers = Get-GraphCollectie -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,accountEnabled&$top=999' -Omschrijving 'gebruikers'
            $teller = 0

            foreach ($u in $gebruikers) {
                $teller++
                $upn = Get-Eigenschap -Object $u -Naam 'userPrincipalName' -Standaard ''
                $id  = Get-Eigenschap -Object $u -Naam 'id'

                if ($teller % 50 -eq 0) {
                    Write-Log -Bericht ("Voortgang: {0}/{1} gebruikers verwerkt." -f $teller, $gebruikers.Count)
                }

                $methoden = @()
                try {
                    $mResp = Invoke-GraphMetRetry -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $id)
                    $methoden = @(Get-Eigenschap -Object $mResp -Naam 'value' -Standaard @())
                }
                catch {
                    Write-Log -Niveau WAARSCHUWING -Bericht ("Methoden van {0} konden niet worden opgehaald: {1}" -f $upn, $_.Exception.Message)
                    continue
                }

                $types = @(
                    $methoden |
                        ForEach-Object { (Get-Eigenschap -Object $_ -Naam '@odata.type' -Standaard '') -replace '#microsoft\.graph\.', '' -replace 'AuthenticationMethod$', '' } |
                        Where-Object { $_ } | Sort-Object -Unique
                )

                $sterkeMethoden = @($types | Where-Object { $_ -notin @('password') })

                $methodeInventaris.Add([pscustomobject]@{
                    UPN                    = $upn
                    Weergavenaam           = Get-Eigenschap -Object $u -Naam 'displayName' -Standaard ''
                    IsAdmin                = ''
                    MFAGeregistreerd       = ($sterkeMethoden.Count -gt 0)
                    MFACapabel             = ($sterkeMethoden.Count -gt 0)
                    SSPRGeregistreerd      = ''
                    PasswordlessCapabel    = [bool](@($types | Where-Object { $_ -match 'fido2|passwordlessMicrosoftAuthenticator|windowsHelloForBusiness' }).Count)
                    StandaardMethode       = ''
                    GeregistreerdeMethoden = ($types -join ', ')
                    Bron                   = 'authentication/methods'
                })
            }

            Write-Log -Niveau OK -Bericht ("{0} gebruikers geïnventariseerd." -f $methodeInventaris.Count)
        }
        catch {
            Write-Log -Niveau FOUT -Bericht ("Inventarisatie mislukt: {0}" -f $_.Exception.Message)
        }
    }
}

#endregion -------------------------------------------------------------------------------

#region Export ---------------------------------------------------------------------------

Write-Log -Niveau STAP -Bericht 'Stap 6 - Rapportages wegschrijven.'

$csvParams = @{
    Delimiter         = ';'
    Encoding          = 'utf8BOM'
    NoTypeInformation = $true
    Force             = $true
}

$bestanden = @()

if ($signInRecords.Count -gt 0) {
    $padDetail = Join-Path $OutputPad ("01-SignIns-detail-{0}.csv" -f $stempel)
    $signInRecords | Sort-Object -Property Tijdstip -Descending | Export-Csv -Path $padDetail @csvParams
    $bestanden += $padDetail
    Write-Log -Niveau OK -Bericht ("Detailrapport weggeschreven: {0}" -f $padDetail)
}

if ($gebruikerSamenvatting.Count -gt 0) {
    $padSamen = Join-Path $OutputPad ("02-Gebruikers-MFA-samenvatting-{0}.csv" -f $stempel)
    $gebruikerSamenvatting | Export-Csv -Path $padSamen @csvParams
    $bestanden += $padSamen
    Write-Log -Niveau OK -Bericht ("Gebruikerssamenvatting weggeschreven: {0}" -f $padSamen)

    $zonderMfaGebruikers = @($gebruikerSamenvatting | Where-Object { $_.Classificatie -eq 'NOOIT MFA' })
    if ($zonderMfaGebruikers.Count -gt 0) {
        $padRisico = Join-Path $OutputPad ("03-Gebruikers-zonder-MFA-{0}.csv" -f $stempel)
        $zonderMfaGebruikers | Export-Csv -Path $padRisico @csvParams
        $bestanden += $padRisico
        Write-Log -Niveau OK -Bericht ("Risicolijst weggeschreven: {0}" -f $padRisico)
    }
}

if ($caConfiguratie.Count -gt 0) {
    $padCa = Join-Path $OutputPad ("04-ConditionalAccess-configuratie-{0}.csv" -f $stempel)
    $caConfiguratie | Sort-Object -Property VereistMFA, PolicyNaam -Descending | Export-Csv -Path $padCa @csvParams
    $bestanden += $padCa
    Write-Log -Niveau OK -Bericht ("CA-configuratie weggeschreven: {0}" -f $padCa)
}

$gapAnalyse = @()
if ($signInRecords.Count -gt 0) {
    $gapAnalyse = @(
        $signInRecords |
            Where-Object { $_.Succesvol -and -not $_.MFAGebruikt } |
            Group-Object -Property RedenGeenMFA, Applicatie, ClientApp |
            ForEach-Object {
                $eerste = $_.Group[0]
                [pscustomobject]@{
                    RedenGeenMFA     = $eerste.RedenGeenMFA
                    Applicatie       = $eerste.Applicatie
                    ClientApp        = $eerste.ClientApp
                    AantalLogins     = $_.Count
                    AantalGebruikers = @($_.Group | Select-Object -ExpandProperty UPN -Unique).Count
                    VoorbeeldUPN     = $eerste.UPN
                    CAResultaten     = $eerste.CAResultaten
                }
            } | Sort-Object -Property AantalLogins -Descending
    )

    if ($gapAnalyse.Count -gt 0) {
        $padGap = Join-Path $OutputPad ("05-MFA-gapanalyse-{0}.csv" -f $stempel)
        $gapAnalyse | Export-Csv -Path $padGap @csvParams
        $bestanden += $padGap
        Write-Log -Niveau OK -Bericht ("Gapanalyse weggeschreven: {0}" -f $padGap)
    }
}

if ($methodeInventaris.Count -gt 0) {
    $padMethoden = Join-Path $OutputPad ("06-Geregistreerde-MFA-methoden-{0}.csv" -f $stempel)
    $methodeInventaris | Sort-Object -Property MFAGeregistreerd, UPN | Export-Csv -Path $padMethoden @csvParams
    $bestanden += $padMethoden
    Write-Log -Niveau OK -Bericht ("Methode-inventarisatie weggeschreven: {0}" -f $padMethoden)
}

#endregion -------------------------------------------------------------------------------

#region Samenvatting op scherm -----------------------------------------------------------

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ' SAMENVATTING' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan

Write-Host ("Tenant                : {0}" -f $context.TenantId)
Write-Host ("Periode               : {0} t/m {1}" -f $vanaf.ToLocalTime().ToString('dd-MM-yyyy'), (Get-Date).ToString('dd-MM-yyyy'))
Write-Host ("Security Defaults     : {0}" -f $(if ($null -eq $securityDefaultsAan) { 'onbekend' } elseif ($securityDefaultsAan) { 'INGESCHAKELD' } else { 'uitgeschakeld' }))
Write-Host ("Conditional Access    : {0} policies" -f $caPolicies.Count)

if ($signInRecords.Count -gt 0) {
    $succesTotaal = @($signInRecords | Where-Object { $_.Succesvol })
    $mfaTotaal    = @($succesTotaal | Where-Object { $_.MFAGebruikt })
    $verseTotaal  = @($succesTotaal | Where-Object { $_.MFAUitgevoerdDezeSessie })
    $tokenTotaal  = @($succesTotaal | Where-Object { $_.MFAUitTokenOfClaim -and -not $_.MFAUitgevoerdDezeSessie })
    $pct          = if ($succesTotaal.Count -gt 0) { [math]::Round(($mfaTotaal.Count / $succesTotaal.Count) * 100, 1) } else { 0 }

    Write-Host ''
    Write-Host ("Sign-ins totaal       : {0}" -f $signInRecords.Count)
    Write-Host ("Succesvolle sign-ins  : {0}" -f $succesTotaal.Count)
    Write-Host ("  met MFA             : {0} ({1}%)" -f $mfaTotaal.Count, $pct) -ForegroundColor Green
    Write-Host ("  waarvan live MFA    : {0}" -f $verseTotaal.Count)
    Write-Host ("  waarvan uit token   : {0}" -f $tokenTotaal.Count)
    Write-Host ("  single-factor       : {0}" -f ($succesTotaal.Count - $mfaTotaal.Count)) -ForegroundColor Yellow

    Write-Host ''
    Write-Host 'Gebruikers per classificatie:'
    $gebruikerSamenvatting | Group-Object -Property Classificatie | Sort-Object -Property Count -Descending | ForEach-Object {
        $kleur = if ($_.Name -eq 'NOOIT MFA') { 'Red' } elseif ($_.Name -eq 'Altijd MFA') { 'Green' } else { 'Yellow' }
        Write-Host ("  {0,-24} : {1}" -f $_.Name, $_.Count) -ForegroundColor $kleur
    }

    $legacy = @($gebruikerSamenvatting | Where-Object { $_.LegacyClientGebruikt })
    if ($legacy.Count -gt 0) {
        Write-Host ''
        Write-Host ("LET OP: {0} gebruikers gebruiken legacy authenticatie-clients." -f $legacy.Count) -ForegroundColor Red
    }

    if ($gapAnalyse.Count -gt 0) {
        Write-Host ''
        Write-Host 'Waarom waren logins single-factor (top 8):' -ForegroundColor Cyan
        $gapAnalyse | Select-Object -First 8 | ForEach-Object {
            Write-Host ("  {0,6} logins | {1,-28} | {2} / {3}" -f $_.AantalLogins, $_.RedenGeenMFA, $_.Applicatie, $_.ClientApp) -ForegroundColor Yellow
        }
    }

    $mfaPolicies = @($caConfiguratie | Where-Object { $_.VereistMFA })
    if ($mfaPolicies.Count -gt 0) {
        Write-Host ''
        Write-Host 'MFA-policies en hun uitzonderingen:' -ForegroundColor Cyan
        foreach ($mp in $mfaPolicies) {
            $kleur = if ($mp.Status -ne 'enabled') { 'Red' } elseif ($mp.AantalUitzonderingen -gt 0) { 'Yellow' } else { 'Green' }
            Write-Host ("  {0,-45} status={1,-34} uitzonderingen={2}" -f $mp.PolicyNaam, $mp.Status, $mp.AantalUitzonderingen) -ForegroundColor $kleur
            if ($mp.UitgeslotenGebruikers) { Write-Host ("      gebruikers: {0}" -f $mp.UitgeslotenGebruikers) -ForegroundColor DarkYellow }
            if ($mp.UitgeslotenGroepen)    { Write-Host ("      groepen   : {0}" -f $mp.UitgeslotenGroepen) -ForegroundColor DarkYellow }
            if ($mp.UitgeslotenRollen)     { Write-Host ("      rollen    : {0}" -f $mp.UitgeslotenRollen) -ForegroundColor DarkYellow }
        }
    }
}

if ($methodeInventaris.Count -gt 0) {
    $geregistreerd = @($methodeInventaris | Where-Object { $_.MFAGeregistreerd -eq $true })
    Write-Host ''
    Write-Host ("Gebruikers met MFA geregistreerd : {0} van {1}" -f $geregistreerd.Count, $methodeInventaris.Count) -ForegroundColor Green
    Write-Host ("Gebruikers zonder MFA-registratie: {0}" -f ($methodeInventaris.Count - $geregistreerd.Count)) -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Bestanden:' -ForegroundColor Cyan
foreach ($b in $bestanden) { Write-Host ("  {0}" -f $b) }

Write-Host ''
Write-Log -Niveau OK -Bericht ("Analyse afgerond in {0:n1} seconden." -f ((Get-Date) - $startTijd).TotalSeconds)

#endregion -------------------------------------------------------------------------------
