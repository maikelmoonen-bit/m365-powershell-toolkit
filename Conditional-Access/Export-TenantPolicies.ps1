#Requires -Version 7.2
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Exporteert alle Conditional Access-policies van een tenant naar één JSON-bestand,
    als bron voor een latere import in een andere tenant.

.DESCRIPTION
    Leest via Microsoft Graph alle Conditional Access-policies uit en schrijft ze naar
    één JSON-bestand, met een CSV-overzicht en een logbestand.

    De policies verwijzen naar tenant-specifieke ID's (groepen, gebruikers, applicaties,
    named locations, authentication strengths). Die objecten worden NIET geëxporteerd,
    maar hun ID's worden met displayName vastgelegd in de sectie 'references', zodat een
    importscript ze in de doeltenant op naam kan mappen.

    Standaard draait het script in DRY-RUN: alles wordt uitgelezen en gelogd, maar er
    worden geen bestanden weggeschreven. Gebruik -Execute om daadwerkelijk te exporteren.

.PARAMETER OutputFolder
    Map waarin de export, het CSV-overzicht en het logbestand worden geplaatst.

.PARAMETER TenantId
    Tenant om mee te verbinden. Verplicht bij app-only authenticatie.

.PARAMETER ClientId
    App registration (client) ID voor app-only authenticatie met certificaat.

.PARAMETER CertificateThumbprint
    Thumbprint van het certificaat voor app-only authenticatie.

.PARAMETER Execute
    Schrijf de export daadwerkelijk weg. Zonder deze switch: dry-run.

.EXAMPLE
    .\Export-TenantPolicies.ps1
    Dry-run: leest alles uit en toont wat geëxporteerd zou worden.

.EXAMPLE
    .\Export-TenantPolicies.ps1 -Execute

.EXAMPLE
    .\Export-TenantPolicies.ps1 -Execute -TenantId 'contoso.onmicrosoft.com' -ClientId '<appId>' -CertificateThumbprint '<thumbprint>'
#>
[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path $PWD -ChildPath 'PolicyExport'),
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$Execute
)

$ErrorActionPreference = 'Stop'
$ScriptVersie = '2.0.1'
$GraphBeta    = 'https://graph.microsoft.com/beta'
$GraphV1      = 'https://graph.microsoft.com/v1.0'
$GuidPatroon  = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$NulGuid      = '00000000-0000-0000-0000-000000000000'   # MFA trusted IPs in locations

$script:LogRegels   = [System.Collections.Generic.List[string]]::new()
$script:Referenties = @{}
foreach ($t in 'group', 'user', 'application', 'namedLocation', 'authenticationStrength') {
    $script:Referenties[$t] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
}

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

function Invoke-GraphGet {
    param([Parameter(Mandatory)][string]$Uri)
    $maxPogingen = 5
    for ($poging = 1; $poging -le $maxPogingen; $poging++) {
        try {
            $raw = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Json -ErrorAction Stop
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
        $resp = Invoke-GraphGet -Uri $volgende
        foreach ($v in $resp.value) { $items.Add($v) }
        $volgende = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    }
    $items   # items op de pipeline zetten, zodat @() een platte array oplevert
}

function Add-Reference {
    param([string]$Type, [string]$Id)
    if ($Id -and $Id -match $GuidPatroon -and $Id -ne $NulGuid) { [void]$script:Referenties[$Type].Add($Id) }
}

function Get-PolicyReferences {
    param($Policy)
    $c = $Policy.conditions
    foreach ($x in @($c.users.includeGroups) + @($c.users.excludeGroups) | Where-Object { $_ }) { Add-Reference -Type group -Id $x }
    foreach ($x in @($c.users.includeUsers) + @($c.users.excludeUsers) | Where-Object { $_ }) { Add-Reference -Type user -Id $x }
    foreach ($x in @($c.applications.includeApplications) + @($c.applications.excludeApplications) | Where-Object { $_ }) { Add-Reference -Type application -Id $x }
    foreach ($x in @($c.locations.includeLocations) + @($c.locations.excludeLocations) | Where-Object { $_ }) { Add-Reference -Type namedLocation -Id $x }
    if ($Policy.grantControls -and $Policy.grantControls.authenticationStrength) {
        Add-Reference -Type authenticationStrength -Id $Policy.grantControls.authenticationStrength.id
    }
}

function Resolve-References {
    $uitkomst = [ordered]@{}

    # Groepen en gebruikers in batches van max. 1000 via getByIds
    $dirIds   = @($script:Referenties.group) + @($script:Referenties.user)
    $gevonden = @{}
    for ($i = 0; $i -lt $dirIds.Count; $i += 1000) {
        $chunk = @($dirIds[$i..([math]::Min($i + 999, $dirIds.Count - 1))])
        $body  = @{ ids = $chunk; types = @('group', 'user') } | ConvertTo-Json -Depth 5
        $resp  = Invoke-MgGraphRequest -Method POST -Uri "$GraphV1/directoryObjects/getByIds" -Body $body `
                     -ContentType 'application/json' -OutputType Json | ConvertFrom-Json -Depth 20
        foreach ($o in $resp.value) { $gevonden[$o.id] = $o }
    }

    $uitkomst.groups = @(foreach ($id in $script:Referenties.group) {
        $o = $gevonden[$id]
        if (-not $o) { Write-Log "Groep $id niet gevonden (mogelijk verwijderd)." -Niveau WAARSCHUWING }
        [ordered]@{
            id              = $id
            displayName     = $o.displayName
            description     = $o.description
            mailNickname    = $o.mailNickname
            securityEnabled = $o.securityEnabled
            groupTypes      = $o.groupTypes
            membershipRule  = $o.membershipRule
            resolved        = [bool]$o
        }
    })

    $uitkomst.users = @(foreach ($id in $script:Referenties.user) {
        $o = $gevonden[$id]
        if (-not $o) { Write-Log "Gebruiker $id niet gevonden (mogelijk verwijderd)." -Niveau WAARSCHUWING }
        [ordered]@{ id = $id; displayName = $o.displayName; userPrincipalName = $o.userPrincipalName; resolved = [bool]$o }
    })

    $uitkomst.applications = @(foreach ($appId in $script:Referenties.application) {
        $sp = $null
        try {
            $r  = Invoke-GraphGet -Uri "$GraphV1/servicePrincipals?`$filter=appId eq '$appId'&`$select=appId,displayName,appOwnerOrganizationId"
            $sp = @($r.value) | Select-Object -First 1
        }
        catch { }
        if (-not $sp) { Write-Log "Applicatie $appId niet gevonden als service principal." -Niveau WAARSCHUWING }
        [ordered]@{ appId = $appId; displayName = $sp.displayName; appOwnerOrganizationId = $sp.appOwnerOrganizationId; resolved = [bool]$sp }
    })

    $uitkomst.namedLocations = @(foreach ($id in $script:Referenties.namedLocation) {
        $o = $null
        try { $o = Invoke-GraphGet -Uri "$GraphBeta/identity/conditionalAccess/namedLocations/$id" } catch { }
        if (-not $o) { Write-Log "Named location $id niet gevonden." -Niveau WAARSCHUWING }
        [ordered]@{ id = $id; displayName = $o.displayName; type = $o.'@odata.type'; resolved = [bool]$o }
    })

    $uitkomst.authenticationStrengths = @(foreach ($id in $script:Referenties.authenticationStrength) {
        $o = $null
        try { $o = Invoke-GraphGet -Uri "$GraphBeta/identity/conditionalAccess/authenticationStrength/policies/$id" } catch { }
        if (-not $o) { Write-Log "Authentication strength $id niet gevonden." -Niveau WAARSCHUWING }
        [ordered]@{ id = $id; displayName = $o.displayName; policyType = $o.policyType; resolved = [bool]$o }
    })

    return $uitkomst
}

#endregion

#region Verbinden ----------------------------------------------------------------------

if (-not $Execute) {
    Write-Log 'DRY-RUN modus: er wordt alles uitgelezen, maar niets weggeschreven. Gebruik -Execute om te exporteren.' -Niveau WAARSCHUWING
}

$Scopes = @('Policy.Read.All', 'Directory.Read.All')

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

$ctx = Get-MgContext
$org = (Invoke-GraphGet -Uri "$GraphV1/organization?`$select=id,displayName,verifiedDomains").value | Select-Object -First 1
$standaardDomein = ($org.verifiedDomains | Where-Object { $_.isDefault }).name
Write-Log "Verbonden met tenant '$($org.displayName)' ($($org.id), $standaardDomein)." -Niveau OK

#endregion

#region Uitvoerpaden (nooit overschrijven) ---------------------------------------------

$tenantSlug = ($org.displayName -replace '[^A-Za-z0-9\-]', '')
$basisNaam  = "CAPolicies_${tenantSlug}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$JsonPad    = Join-Path $OutputFolder "$basisNaam.json"
$CsvPad     = Join-Path $OutputFolder "${basisNaam}_Overzicht.csv"
$LogPad     = Join-Path $OutputFolder "${basisNaam}_Log.txt"

foreach ($pad in $JsonPad, $CsvPad, $LogPad) {
    if (Test-Path -LiteralPath $pad) { throw "Uitvoerbestand bestaat al en wordt niet overschreven: $pad" }
}

#endregion

#region Export -------------------------------------------------------------------------

Write-Log 'Conditional Access-policies uitlezen...'
$Policies = @(Get-GraphCollection -Uri "$GraphBeta/identity/conditionalAccess/policies")
Write-Log "$($Policies.Count) Conditional Access-policy/policies gevonden." -Niveau OK

$Overzicht = foreach ($p in $Policies) {
    Get-PolicyReferences -Policy $p
    [pscustomobject]@{
        Naam            = $p.displayName
        Id              = $p.id
        Status          = $p.state
        Aangemaakt      = $p.createdDateTime
        LaatstGewijzigd = $p.modifiedDateTime
    }
}

Write-Log 'Verwijzingen oplossen naar displayName...'
$References = Resolve-References
Write-Log ('Verwijzingen: {0} groep(en), {1} gebruiker(s), {2} applicatie(s), {3} named location(s), {4} authentication strength(s).' -f `
    $References.groups.Count, $References.users.Count, $References.applications.Count,
    $References.namedLocations.Count, $References.authenticationStrengths.Count) -Niveau OK

$Export = [ordered]@{
    metadata   = [ordered]@{
        exportDate          = (Get-Date).ToString('o')
        scriptVersion       = $ScriptVersie
        sourceTenantId      = $org.id
        sourceTenantName    = $org.displayName
        sourceDefaultDomain = $standaardDomein
        exportedBy          = $ctx.Account ?? $ctx.ClientId
        policyCount         = $Policies.Count
    }
    references = $References
    policies   = $Policies
}

#endregion

#region Wegschrijven -------------------------------------------------------------------

if (-not $Execute) {
    $Overzicht | Format-Table -AutoSize | Out-String | Write-Host
    Write-Log "DRY-RUN afgerond: $($Policies.Count) policy/policies gevonden. Er is niets weggeschreven." -Niveau WAARSCHUWING
    Write-Log "Met -Execute zou worden geschreven naar: $JsonPad"
    return
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
    Write-Log "Uitvoermap aangemaakt: $OutputFolder"
}

$Export | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $JsonPad -Encoding utf8
$Overzicht | Export-Csv -LiteralPath $CsvPad -Delimiter ';' -Encoding utf8 -NoTypeInformation

try {
    $null = Get-Content -LiteralPath $JsonPad -Raw | ConvertFrom-Json -Depth 100
    Write-Log 'JSON-validatie geslaagd.' -Niveau OK
}
catch {
    Write-Log "JSON-validatie mislukt: $($_.Exception.Message)" -Niveau FOUT
}

Write-Log "Export voltooid: $($Policies.Count) policy/policies." -Niveau OK
Write-Log "JSON:       $JsonPad"
Write-Log "Overzicht:  $CsvPad"
Write-Log "Log:        $LogPad"
$script:LogRegels | Set-Content -LiteralPath $LogPad -Encoding utf8

#endregion
