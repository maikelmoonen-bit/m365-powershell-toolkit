#Requires -Version 7.2
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Importeert Conditional Access-policies uit een export van Export-TenantPolicies.ps1
    in een (andere) tenant.

.DESCRIPTION
    Per policy worden alle tenant-specifieke verwijzingen omgezet naar de doeltenant:
      - Groepen                  : op displayName (optioneel aanmaken met -CreateMissingGroups)
      - Gebruikers               : op UPN, daarna UPN met het standaarddomein van de doeltenant
      - Applicaties              : op appId (service principal moet in de doeltenant bestaan)
      - Named locations          : op displayName (moeten al bestaan in de doeltenant)
      - Authentication strengths : built-in op ID, custom op displayName
      - Authentication contexts  : moeten met hetzelfde ID (c1..c99) bestaan
      - Workload identities en Terms of Use: alleen via -MappingCsv
    Elke verwijzing kan handmatig overschreven worden via -MappingCsv.

    Niet-oplosbare verwijzingen:
      Een verwijzing die niet in de doeltenant gevonden wordt, wordt uit de policy
      verwijderd en de policy wordt toch aangemaakt. Elke aanpassing wordt gelogd en in
      het rapport (kolom Aanpassingen) vastgelegd. Verwijderde UITSLUITINGEN (bijv. een
      break-glass-account) worden extra gemarkeerd: controleer die vóór het inschakelen.
      Blijft er door verwijdering een verplichte conditie leeg achter, dan wordt die
      veilig gerepareerd:
        - geen gebruikers/groepen/rollen meer in include  -> includeUsers = 'None'
        - geen applicaties meer in include                -> includeApplications = 'None'
        - geen named locations meer in include            -> locatieconditie verwijderd
        - geen grant control meer over                    -> builtInControls = 'mfa'

    Productieveilig:
      - Standaard DRY-RUN: er wordt niets aangemaakt. Gebruik -Execute om te importeren.
      - ALLE policies worden aangemaakt als Report-only, ongeacht de status in de bron.
      - Het account uit -ExcludeUserUpn (bijv. een break-glass- of beheeraccount) wordt in
        elke geïmporteerde policy uitgesloten. Bestaat dit account niet in de doeltenant,
        dan wordt een -Execute-run afgebroken voordat er iets wordt aangemaakt.
      - Bij -Execute wordt na elke aangemaakte policy gevraagd of de volgende mag worden
        aangemaakt: [J]a, [A]lle resterende, of [S]toppen. Uitzetten met -SkipConfirmation.
      - Idempotent: een policy met dezelfde displayName in de doeltenant wordt overgeslagen.

.PARAMETER Path
    Pad naar het JSON-exportbestand.

.PARAMETER MappingCsv
    Optionele puntkomma-gescheiden CSV met kolommen Type;BronId;DoelId.
    Type is één van: group, user, application, namedLocation, authenticationStrength,
    authenticationContext, servicePrincipal, termsOfUse.

.PARAMETER CreateMissingGroups
    Maak groepen die in de doeltenant niet bestaan aan op basis van de exportgegevens
    (naam, beschrijving, type, dynamische regel). Groepen worden leeg aangemaakt.

.PARAMETER ExcludeUserUpn
    UPN van het account dat in de doeltenant in elke policy wordt uitgesloten,
    bijvoorbeeld een break-glass-account. Sterk aanbevolen om buitensluiten te voorkomen.
    Leeg laten (standaard) om geen account uit te sluiten.

.PARAMETER SkipConfirmation
    Niet na elke policy om bevestiging vragen (bijv. voor onbeheerde runs).

.PARAMETER Execute
    Voer de import daadwerkelijk uit. Zonder deze switch: dry-run.

.EXAMPLE
    .\Import-TenantPolicies.ps1 -Path .\PolicyExport\CAPolicies_Contoso_20261001_101500.json -TenantId 'fabrikam.onmicrosoft.com'
    Dry-run: toont per policy wat er zou gebeuren en welke verwijzingen verwijderd zouden worden.

.EXAMPLE
    .\Import-TenantPolicies.ps1 -Path .\export.json -TenantId 'fabrikam.onmicrosoft.com' -ExcludeUserUpn 'breakglass@fabrikam.onmicrosoft.com' -CreateMissingGroups -Execute

.EXAMPLE
    .\Import-TenantPolicies.ps1 -Path .\export.json -MappingCsv .\mapping.csv -Execute
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })][string]$Path,
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })][string]$MappingCsv,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$CreateMissingGroups,
    [AllowEmptyString()][string]$ExcludeUserUpn = '',
    [switch]$SkipConfirmation,
    [switch]$Execute
)

$ErrorActionPreference = 'Stop'
$GraphBeta   = 'https://graph.microsoft.com/beta'
$GraphV1     = 'https://graph.microsoft.com/v1.0'
$GuidPatroon = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$NulGuid     = '00000000-0000-0000-0000-000000000000'
$ReportOnly  = 'enabledForReportingButNotEnforced'

$script:LogRegels = [System.Collections.Generic.List[string]]::new()
$script:Cache     = @{}

#region Helpers ------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Bericht,
        [ValidateSet('INFO', 'OK', 'WAARSCHUWING', 'FOUT')][string]$Niveau = 'INFO'
    )
    $regel = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Niveau, $Bericht
    $kleur = switch ($Niveau) { 'OK' { 'Green' } 'WAARSCHUWING' { 'Yellow' } 'FOUT' { 'Red' } default { 'Gray' } }
    Write-Host $regel -ForegroundColor $kleur
    $script:LogRegels.Add($regel)
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST')][string]$Method = 'GET',
        $Body
    )
    $maxPogingen = 5
    for ($poging = 1; $poging -le $maxPogingen; $poging++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; OutputType = 'Json'; ErrorAction = 'Stop' }
            if ($null -ne $Body) {
                $params.Body        = ($Body | ConvertTo-Json -Depth 50 -Compress)
                $params.ContentType = 'application/json'
            }
            $raw = Invoke-MgGraphRequest @params
            if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
            return ($raw | ConvertFrom-Json -Depth 100)
        }
        catch {
            $status = $null
            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) {
                $status = [int]$_.Exception.Response.StatusCode
            }
            if ($status -in 429, 500, 502, 503, 504 -and $poging -lt $maxPogingen) {
                $wacht = [math]::Pow(2, $poging)
                Write-Log "Graph gaf HTTP $status, nieuwe poging $($poging + 1)/$maxPogingen over $wacht s." -Niveau WAARSCHUWING
                Start-Sleep -Seconds $wacht
                continue
            }
            throw
        }
    }
}

function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $volgende = $Uri
    while ($volgende) {
        $resp = Invoke-Graph -Uri $volgende
        foreach ($v in $resp.value) { $items.Add($v) }
        $volgende = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    }
    $items   # items op de pipeline zetten, zodat @() een platte array oplevert
}

function Get-FoutTekst {
    param($Fout)
    if ($Fout.ErrorDetails -and $Fout.ErrorDetails.Message) { return $Fout.ErrorDetails.Message }
    return $Fout.Exception.Message
}

function ConvertTo-FilterWaarde {
    param([string]$Waarde)
    # Eerst escapen in een aparte stap: binnen de method-call haakjes wordt de komma
    # van -replace als argumentscheiding gezien
    $geescaped = $Waarde -replace "'", "''"
    [uri]::EscapeDataString($geescaped)
}

# Verwijdert read-only/null-eigenschappen recursief (behoudt @odata.type)
function Remove-ReadOnly {
    param($Node)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($k in @($Node.Keys)) {
            if ($k -like '*@odata.context' -or $null -eq $Node[$k]) { $Node.Remove($k); continue }
            Remove-ReadOnly $Node[$k]
        }
    }
    elseif ($Node -is [System.Collections.IList]) {
        foreach ($i in $Node) { Remove-ReadOnly $i }
    }
}

#endregion

#region Doelobjecten opzoeken ----------------------------------------------------------

function New-DoelGroep {
    param($Bron)
    $groupTypes = @($Bron.groupTypes | Where-Object { $_ })
    $nickname   = if ($Bron.mailNickname) { $Bron.mailNickname } else { ($Bron.displayName -replace '[^A-Za-z0-9]', '') }
    $body = [ordered]@{
        displayName     = $Bron.displayName
        mailNickname    = $nickname
        securityEnabled = if ($null -ne $Bron.securityEnabled) { [bool]$Bron.securityEnabled } else { $true }
        mailEnabled     = $groupTypes -contains 'Unified'
        groupTypes      = $groupTypes
    }
    if ($Bron.description) { $body.description = $Bron.description }
    if ($groupTypes -contains 'DynamicMembership' -and $Bron.membershipRule) {
        $body.membershipRule                = $Bron.membershipRule
        $body.membershipRuleProcessingState = 'On'
    }
    $nieuw = Invoke-Graph -Method POST -Uri "$GraphV1/groups" -Body $body
    Write-Log "Groep '$($Bron.displayName)' aangemaakt ($($nieuw.id))." -Niveau OK
    return $nieuw.id
}

function Resolve-DoelId {
    param([string]$Type, [string]$BronId)

    $sleutel = "$Type|$BronId"
    if ($script:Cache.ContainsKey($sleutel)) { return $script:Cache[$sleutel] }

    # Handmatige mapping heeft altijd voorrang
    if ($Mapping.ContainsKey($sleutel)) {
        $script:Cache[$sleutel] = $Mapping[$sleutel]
        return $Mapping[$sleutel]
    }

    $doel = $null
    switch ($Type) {
        'group' {
            $bron = $BronGroepen[$BronId]
            if (-not $bron -or -not $bron.displayName) { break }
            $r = Invoke-Graph -Uri "$GraphV1/groups?`$filter=displayName eq '$(ConvertTo-FilterWaarde $bron.displayName)'&`$select=id,displayName"
            $treffers = @($r.value)
            if ($treffers.Count -eq 1) { $doel = $treffers[0].id }
            elseif ($treffers.Count -gt 1) {
                Write-Log "Groep '$($bron.displayName)' komt $($treffers.Count)x voor in de doeltenant; gebruik -MappingCsv." -Niveau WAARSCHUWING
            }
            elseif ($CreateMissingGroups) {
                if ($Execute) { $doel = New-DoelGroep -Bron $bron }
                else {
                    Write-Log "Groep '$($bron.displayName)' bestaat niet en zou worden aangemaakt." -Niveau WAARSCHUWING
                    $doel = "NIEUW:$($bron.displayName)"
                }
            }
        }
        'user' {
            $bron = $BronGebruikers[$BronId]
            if (-not $bron -or -not $bron.userPrincipalName) { break }
            $kandidaten = [System.Collections.Generic.List[string]]::new()
            $kandidaten.Add($bron.userPrincipalName)
            if ($bron.userPrincipalName -notlike '*#EXT#*') {
                $kandidaten.Add(($bron.userPrincipalName -split '@')[0] + "@$DoelDomein")
            }
            foreach ($upn in $kandidaten | Select-Object -Unique) {
                $r = Invoke-Graph -Uri "$GraphV1/users?`$filter=userPrincipalName eq '$(ConvertTo-FilterWaarde $upn)'&`$select=id"
                if (@($r.value).Count -eq 1) { $doel = $r.value[0].id; break }
            }
        }
        'application' {
            $r = Invoke-Graph -Uri "$GraphV1/servicePrincipals?`$filter=appId eq '$BronId'&`$select=appId"
            if (@($r.value).Count -ge 1) { $doel = $BronId }
        }
        'namedLocation' {
            $bron = $BronLocaties[$BronId]
            if (-not $bron -or -not $bron.displayName) { break }
            $treffers = @($DoelLocaties | Where-Object { $_.displayName -eq $bron.displayName })
            if ($treffers.Count -eq 1) { $doel = $treffers[0].id }
            elseif ($treffers.Count -gt 1) {
                Write-Log "Named location '$($bron.displayName)' komt meerdere keren voor; gebruik -MappingCsv." -Niveau WAARSCHUWING
            }
        }
        'authenticationStrength' {
            $bron = $BronStrengths[$BronId]
            if ($bron -and $bron.policyType -eq 'builtIn') {
                if ($DoelStrengths | Where-Object { $_.id -eq $BronId }) { $doel = $BronId }
            }
            elseif ($bron -and $bron.displayName) {
                $treffers = @($DoelStrengths | Where-Object { $_.displayName -eq $bron.displayName -and $_.policyType -eq 'custom' })
                if ($treffers.Count -eq 1) { $doel = $treffers[0].id }
            }
        }
        'authenticationContext' {
            if ($DoelContexten | Where-Object { $_.id -eq $BronId }) { $doel = $BronId }
        }
        default {
            # servicePrincipal, termsOfUse: alleen via -MappingCsv
        }
    }

    $script:Cache[$sleutel] = $doel
    return $doel
}

function Convert-IdLijst {
    param(
        [System.Collections.IDictionary]$Ouder,
        [string]$Sleutel,
        [string]$Type,
        [System.Collections.Generic.List[string]]$Aanpassingen,
        [switch]$AlleIds   # ook niet-GUID waarden mappen (auth contexts)
    )
    if (-not $Ouder -or -not $Ouder.Contains($Sleutel) -or $null -eq $Ouder[$Sleutel]) { return }
    $nieuw = foreach ($id in @($Ouder[$Sleutel])) {
        $mappen = if ($AlleIds) { $true } else { $id -match $GuidPatroon -and $id -ne $NulGuid }
        if (-not $mappen) { $id; continue }
        $doel = Resolve-DoelId -Type $Type -BronId $id
        if ($doel) { $doel }
        else {
            $label = if ($Sleutel -like 'exclude*') { 'LET OP UITSLUITING VERWIJDERD' } else { 'Verwijderd' }
            $Aanpassingen.Add("${label}: $Sleutel $Type '$(Get-BronNaam -Type $Type -Id $id)' ($id) niet gevonden")
            # niets teruggeven: de verwijzing verdwijnt uit de lijst
        }
    }
    $Ouder[$Sleutel] = @($nieuw)
}

function Get-BronNaam {
    param([string]$Type, [string]$Id)
    $bron = switch ($Type) {
        'group'                  { $BronGroepen[$Id] }
        'user'                   { $BronGebruikers[$Id] }
        'namedLocation'          { $BronLocaties[$Id] }
        'authenticationStrength' { $BronStrengths[$Id] }
        default                  { $null }
    }
    if ($bron -and $bron.userPrincipalName) { return $bron.userPrincipalName }
    if ($bron -and $bron.displayName)       { return $bron.displayName }
    return 'onbekend'
}

# Repareert condities die door verwijderde verwijzingen leeg (en dus ongeldig) zijn geworden
function Repair-Policy {
    param(
        [System.Collections.IDictionary]$Body,
        [System.Collections.Generic.List[string]]$Aanpassingen
    )
    $c = $Body.conditions

    $u = $c.users
    if ($u) {
        $includes = @($u.includeUsers) + @($u.includeGroups) + @($u.includeRoles) | Where-Object { $_ }
        $gasten   = $u.Contains('includeGuestsOrExternalUsers') -and $u.includeGuestsOrExternalUsers
        $workload = $c.clientApplications -and (@($c.clientApplications.includeServicePrincipals) | Where-Object { $_ })
        if (-not $includes -and -not $gasten -and -not $workload) {
            $u.includeUsers = @('None')
            $Aanpassingen.Add("Geen geldige gebruikers/groepen/rollen meer in include: includeUsers op 'None' gezet (policy geldt voor niemand)")
        }
    }

    $a = $c.applications
    if ($a) {
        $includes = @($a.includeApplications) + @($a.includeAuthenticationContextClassReferences) + @($a.includeUserActions) | Where-Object { $_ }
        if (-not $includes -and -not $a.applicationFilter) {
            $a.includeApplications = @('None')
            $Aanpassingen.Add("Geen geldige applicaties meer in include: includeApplications op 'None' gezet")
        }
    }

    $l = $c.locations
    if ($l -and -not (@($l.includeLocations) | Where-Object { $_ })) {
        [void]$c.Remove('locations')
        $Aanpassingen.Add('Geen geldige named locations meer in include: locatieconditie verwijderd (policy is breder dan in de bron)')
    }

    $g = $Body.grantControls
    if ($g) {
        $controles = @($g.builtInControls) + @($g.termsOfUse) + @($g.customAuthenticationFactors) | Where-Object { $_ }
        if (-not $controles -and -not $g.authenticationStrength) {
            $g.builtInControls = @('mfa')
            if (-not $g.operator) { $g.operator = 'OR' }
            $Aanpassingen.Add("Geen grant control meer over: builtInControls op 'mfa' gezet")
        }
    }
}

function Read-Doorgaan {
    param([int]$Resterend)
    while ($true) {
        $keuze = Read-Host "Doorgaan met de volgende policy? ($Resterend resterend) [J]a / [A]lle resterende zonder vragen / [S]toppen"
        switch (($keuze ?? '').Trim().ToUpper()) {
            'J' { return 'J' }
            'A' { return 'A' }
            'S' { return 'S' }
        }
        Write-Host 'Ongeldige keuze, kies J, A of S.' -ForegroundColor Yellow
    }
}

#endregion

#region Export inlezen -----------------------------------------------------------------

if (-not $Execute) {
    Write-Log 'DRY-RUN modus: er wordt niets aangemaakt. Gebruik -Execute om te importeren.' -Niveau WAARSCHUWING
}

Write-Log "Exportbestand inlezen: $Path"
$Data = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable -Depth 100
if (-not $Data.Contains('policies') -or -not $Data.Contains('references')) {
    throw 'Het bestand lijkt geen export van Export-TenantPolicies.ps1 te zijn (policies/references ontbreken).'
}

$BronGroepen = @{}; foreach ($o in @($Data.references.groups))                  { if ($o) { $BronGroepen[$o.id] = $o } }
$BronGebruikers = @{}; foreach ($o in @($Data.references.users))                { if ($o) { $BronGebruikers[$o.id] = $o } }
$BronLocaties = @{}; foreach ($o in @($Data.references.namedLocations))         { if ($o) { $BronLocaties[$o.id] = $o } }
$BronStrengths = @{}; foreach ($o in @($Data.references.authenticationStrengths)) { if ($o) { $BronStrengths[$o.id] = $o } }

# Platslaan: vangt ook exports op die met de vorige versie (geneste array) zijn gemaakt
$Policies = @($Data.policies | ForEach-Object { $_ } | Where-Object { $_ -is [System.Collections.IDictionary] })
Write-Log "$($Policies.Count) policy/policies in export van '$($Data.metadata.sourceTenantName)' ($($Data.metadata.exportDate))." -Niveau OK

$Mapping = @{}
if ($MappingCsv) {
    foreach ($r in Import-Csv -LiteralPath $MappingCsv -Delimiter ';' -Encoding utf8) {
        if ($r.Type -and $r.BronId -and $r.DoelId) { $Mapping["$($r.Type)|$($r.BronId)"] = $r.DoelId }
    }
    Write-Log "$($Mapping.Count) handmatige mapping(s) ingelezen uit $MappingCsv."
}

#endregion

#region Verbinden ----------------------------------------------------------------------

$Scopes = @('Policy.Read.All', 'Application.Read.All', 'Directory.Read.All')
if ($Execute) { $Scopes += 'Policy.ReadWrite.ConditionalAccess' }
if ($Execute -and $CreateMissingGroups) { $Scopes += 'Group.ReadWrite.All' }

$ctx = Get-MgContext
if ($ClientId -and $CertificateThumbprint) {
    if (-not $TenantId) { throw 'Bij app-only authenticatie is -TenantId verplicht.' }
    Write-Log "Verbinden met tenant $TenantId (app-only, client $ClientId)..."
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    $missendeScopes = if ($ctx) { $Scopes | Where-Object { $_ -notin $ctx.Scopes } } else { $Scopes }
    $andereTenant   = $ctx -and $TenantId -and ($ctx.TenantId -ne $TenantId)
    if (-not $ctx -or $missendeScopes -or $andereTenant) {
        Write-Log 'Interactief verbinden met Microsoft Graph...'
        $connectParams = @{ Scopes = $Scopes; NoWelcome = $true }
        if ($TenantId) { $connectParams.TenantId = $TenantId }
        Connect-MgGraph @connectParams
    }
    else {
        Write-Log "Bestaande Graph-sessie hergebruikt ($($ctx.Account))."
    }
}

$org = (Invoke-Graph -Uri "$GraphV1/organization?`$select=id,displayName,verifiedDomains").value | Select-Object -First 1
$DoelDomein = ($org.verifiedDomains | Where-Object { $_.isDefault }).name
Write-Log "Doeltenant: '$($org.displayName)' ($($org.id), $DoelDomein)." -Niveau OK

if ($org.id -eq $Data.metadata.sourceTenantId) {
    Write-Log 'Let op: de doeltenant is dezelfde als de brontenant van de export.' -Niveau WAARSCHUWING
}

#endregion

#region Doeltenant inventariseren ------------------------------------------------------

Write-Log 'Bestaande objecten in de doeltenant inlezen...'
$DoelPolicies  = @(Get-GraphCollection -Uri "$GraphBeta/identity/conditionalAccess/policies?`$select=id,displayName")
$DoelLocaties  = @(Get-GraphCollection -Uri "$GraphBeta/identity/conditionalAccess/namedLocations")
$DoelStrengths = @(Get-GraphCollection -Uri "$GraphBeta/identity/conditionalAccess/authenticationStrength/policies")
$DoelContexten = @(Get-GraphCollection -Uri "$GraphBeta/identity/conditionalAccess/authenticationContextClassReferences")

$BestaandeNamen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($p in $DoelPolicies) { [void]$BestaandeNamen.Add($p.displayName) }
Write-Log ("Doeltenant: {0} policy/policies, {1} named location(s), {2} authentication strength(s), {3} authentication context(s)." -f `
    $DoelPolicies.Count, $DoelLocaties.Count, $DoelStrengths.Count, $DoelContexten.Count)

$UitsluitingId = $null
if ([string]::IsNullOrWhiteSpace($ExcludeUserUpn)) {
    Write-Log 'Geen -ExcludeUserUpn opgegeven: er wordt geen vast account uitgesloten.' -Niveau WAARSCHUWING
}
else {
    $r = Invoke-Graph -Uri "$GraphV1/users?`$filter=userPrincipalName eq '$(ConvertTo-FilterWaarde $ExcludeUserUpn)'&`$select=id,userPrincipalName"
    if (@($r.value).Count -eq 1) {
        $UitsluitingId = $r.value[0].id
        Write-Log "Uitsluitingsaccount gevonden: $ExcludeUserUpn ($UitsluitingId). Wordt in elke policy uitgesloten." -Niveau OK
    }
    elseif ($Execute) {
        throw "Uitsluitingsaccount '$ExcludeUserUpn' niet gevonden in de doeltenant. Import afgebroken; er is niets aangemaakt."
    }
    else {
        Write-Log "Uitsluitingsaccount '$ExcludeUserUpn' niet gevonden in de doeltenant. Een -Execute-run wordt hierop afgebroken." -Niveau FOUT
    }
}

#endregion

#region Import -------------------------------------------------------------------------

$Resultaten = [System.Collections.Generic.List[object]]::new()

$AllesBevestigd = [bool]$SkipConfirmation
$GestoptBij     = $null

for ($i = 0; $i -lt $Policies.Count; $i++) {
    $policy = $Policies[$i]
    $naam = $policy.displayName
    $res  = [ordered]@{ Naam = $naam; BronId = $policy.id; Actie = ''; BronStatus = $policy.state; Status = ''; NieuwId = ''; Aanpassingen = ''; Opmerking = '' }

    if ($BestaandeNamen.Contains($naam)) {
        Write-Log "'$naam' bestaat al in de doeltenant, overgeslagen."
        $res.Actie = 'Overgeslagen'; $res.Opmerking = 'Bestaat al (zelfde displayName)'
        $Resultaten.Add([pscustomobject]$res); continue
    }

    # Body opschonen
    $body = $policy
    foreach ($k in 'id', 'createdDateTime', 'modifiedDateTime', 'deletedDateTime', 'templateId') { [void]$body.Remove($k) }
    Remove-ReadOnly $body

    # Verwijzingen omzetten (niet-oplosbare worden verwijderd en vastgelegd)
    $aanpassingen = [System.Collections.Generic.List[string]]::new()
    $c = $body.conditions
    Convert-IdLijst $c.users        'includeGroups'       group       $aanpassingen
    Convert-IdLijst $c.users        'excludeGroups'       group       $aanpassingen
    Convert-IdLijst $c.users        'includeUsers'        user        $aanpassingen
    Convert-IdLijst $c.users        'excludeUsers'        user        $aanpassingen
    Convert-IdLijst $c.applications 'includeApplications' application $aanpassingen
    Convert-IdLijst $c.applications 'excludeApplications' application $aanpassingen
    Convert-IdLijst $c.applications 'includeAuthenticationContextClassReferences' authenticationContext $aanpassingen -AlleIds
    Convert-IdLijst $c.locations    'includeLocations'    namedLocation $aanpassingen
    Convert-IdLijst $c.locations    'excludeLocations'    namedLocation $aanpassingen
    Convert-IdLijst $c.clientApplications 'includeServicePrincipals' servicePrincipal $aanpassingen
    Convert-IdLijst $c.clientApplications 'excludeServicePrincipals' servicePrincipal $aanpassingen

    $g = $body.grantControls
    if ($g -and $g.Contains('authenticationStrength') -and $g.authenticationStrength) {
        $bronStrength = $g.authenticationStrength.id
        $doelStrength = Resolve-DoelId -Type authenticationStrength -BronId $bronStrength
        if ($doelStrength) { $g.authenticationStrength = @{ id = $doelStrength } }
        else {
            [void]$g.Remove('authenticationStrength')
            $aanpassingen.Add("Verwijderd: authentication strength '$(Get-BronNaam -Type authenticationStrength -Id $bronStrength)' ($bronStrength) niet gevonden")
        }
    }
    Convert-IdLijst $g 'termsOfUse' termsOfUse $aanpassingen

    Repair-Policy -Body $body -Aanpassingen $aanpassingen

    # Vast uitsluitingsaccount toevoegen
    if ($UitsluitingId -or (-not $Execute -and $ExcludeUserUpn)) {
        $idOfPlaceholder = $UitsluitingId ?? "NIET-GEVONDEN:$ExcludeUserUpn"
        if (-not $body.conditions.users) { $body.conditions.users = [ordered]@{ includeUsers = @('None') } }
        $huidig = @($body.conditions.users.excludeUsers | Where-Object { $_ })
        if ($huidig -notcontains $idOfPlaceholder) {
            $body.conditions.users.excludeUsers = @($huidig + $idOfPlaceholder)
        }
    }

    # Altijd Report-only, ongeacht de bronstatus
    $res.BronStatus = $body.state
    $body.state     = $ReportOnly
    $res.Status     = $ReportOnly
    $res.Aanpassingen = $aanpassingen -join ' | '

    foreach ($a in $aanpassingen) {
        $niveau = if ($a -like 'LET OP*') { 'FOUT' } else { 'WAARSCHUWING' }
        Write-Log "'$naam': $a" -Niveau $niveau
    }

    $metAanpassingen = if ($aanpassingen.Count) { " met $($aanpassingen.Count) aanpassing(en)" } else { '' }
    if (-not $Execute) {
        Write-Log "'$naam' zou worden aangemaakt als Report-only$metAanpassingen." -Niveau OK
        $res.Actie = 'Zou aanmaken'
        $Resultaten.Add([pscustomobject]$res); continue
    }

    try {
        $nieuw = Invoke-Graph -Method POST -Uri "$GraphBeta/identity/conditionalAccess/policies" -Body $body
        [void]$BestaandeNamen.Add($naam)
        Write-Log "'$naam' aangemaakt als Report-only$metAanpassingen (id: $($nieuw.id))." -Niveau OK
        $res.Actie = 'Aangemaakt'; $res.NieuwId = $nieuw.id
    }
    catch {
        $fout = Get-FoutTekst $_
        Write-Log "'$naam' aanmaken mislukt: $fout" -Niveau FOUT
        $res.Actie = 'Mislukt'; $res.Opmerking = $fout
    }
    $Resultaten.Add([pscustomobject]$res)

    $resterend = $Policies.Count - $i - 1
    if (-not $AllesBevestigd -and $resterend -gt 0) {
        switch (Read-Doorgaan -Resterend $resterend) {
            'A' { $AllesBevestigd = $true; Write-Log 'Uitvoerder koos: alle resterende policies zonder bevestiging.' }
            'S' { $GestoptBij = $i + 1;   Write-Log 'Uitvoerder koos: stoppen.' -Niveau WAARSCHUWING }
            default { Write-Log 'Uitvoerder koos: doorgaan met de volgende policy.' }
        }
        if ($null -ne $GestoptBij) { break }
    }
}

if ($null -ne $GestoptBij) {
    for ($j = $GestoptBij; $j -lt $Policies.Count; $j++) {
        $Resultaten.Add([pscustomobject][ordered]@{
            Naam = $Policies[$j].displayName; BronId = $Policies[$j].id; Actie = 'Niet uitgevoerd'
            BronStatus = $Policies[$j].state; Status = ''; NieuwId = ''; Aanpassingen = ''
            Opmerking = 'Import gestopt door uitvoerder'
        })
    }
    Write-Log "Import gestopt door uitvoerder; $($Policies.Count - $GestoptBij) policy/policies niet verwerkt." -Niveau WAARSCHUWING
}

#endregion

#region Samenvatting -------------------------------------------------------------------

$telling = $Resultaten | Group-Object Actie | ForEach-Object { "$($_.Name): $($_.Count)" }
Write-Log "Resultaat: $($telling -join ', ')" -Niveau OK
$Resultaten | Format-Table Naam, Actie, BronStatus, Status, Aanpassingen, Opmerking -AutoSize -Wrap | Out-String | Write-Host

$uitsluitingen = @($Resultaten | Where-Object { $_.Aanpassingen -like '*LET OP UITSLUITING*' })
if ($uitsluitingen) {
    Write-Log "$($uitsluitingen.Count) policy/policies hebben verwijderde uitsluitingen. Herstel deze vóór het inschakelen: $($uitsluitingen.Naam -join ', ')" -Niveau FOUT
}

if (-not $Execute) {
    Write-Log 'DRY-RUN afgerond. Er is niets aangemaakt of weggeschreven.' -Niveau WAARSCHUWING
    return
}

$map       = Split-Path -LiteralPath (Resolve-Path -LiteralPath $Path) -Parent
$tenant    = ($org.displayName -replace '[^A-Za-z0-9\-]', '')
$basisNaam = "CAImport_${tenant}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$CsvPad    = Join-Path $map "$basisNaam.csv"
$LogPad    = Join-Path $map "${basisNaam}_Log.txt"

foreach ($pad in $CsvPad, $LogPad) {
    if (Test-Path -LiteralPath $pad) { throw "Uitvoerbestand bestaat al en wordt niet overschreven: $pad" }
}

$Resultaten | Export-Csv -LiteralPath $CsvPad -Delimiter ';' -Encoding utf8 -NoTypeInformation
Write-Log "Rapport: $CsvPad"
Write-Log "Log:     $LogPad"
$script:LogRegels | Set-Content -LiteralPath $LogPad -Encoding utf8

#endregion
