<#
.SYNOPSIS
    Rapporteert en verwijdert verweesde access assignments in GDAP admin
    relationships: koppelingen naar security groepen die niet meer bestaan
    in de partner tenant (zichtbaar in Partner Center als kale GUID).

.DESCRIPTION
    Het script logt in op de partner tenant via Microsoft Graph en doorloopt
    alle delegated admin relationships (GDAP) over alle klanten heen.

    Per relationship worden de access assignments opgehaald. Voor elke
    gekoppelde groep (accessContainerId) wordt gecontroleerd of deze nog
    bestaat in Entra ID. Bestaat de groep niet meer, dan wordt de assignment
    als "verweesd" aangemerkt.

    Standaard draait het script in dry-run modus en wordt alleen een overzicht
    plus CSV-rapport geproduceerd. Met -Execute worden de verweesde
    assignments na bevestiging daadwerkelijk verwijderd.

.PARAMETER Execute
    Verwijdert de verweesde assignments daadwerkelijk. Zonder deze switch
    is het een dry-run (alleen rapporteren).

.PARAMETER IncludeInactive
    Neemt ook relationships mee die niet de status "active" hebben
    (bijv. "terminated" of "expired"). Standaard alleen actieve relationships.
    NB: assignments in beëindigde relationships kunnen doorgaans niet meer
    verwijderd worden; deze worden alleen gerapporteerd.

.PARAMETER LogPath
    Pad voor het CSV-rapport. Standaard in de huidige map met timestamp.

.NOTES
    Benodigde Graph scopes (delegated, partner tenant):
      - DelegatedAdminRelationship.ReadWrite.All
      - Group.Read.All
    Vereist module: Microsoft.Graph.Authentication

.EXAMPLE
    .\Remove-GdapOrphanedAssignments.ps1
    Dry-run: rapporteert alle verweesde koppelingen, wijzigt niets.

.EXAMPLE
    .\Remove-GdapOrphanedAssignments.ps1 -Execute
    Verwijdert de verweesde koppelingen na bevestiging.
#>

[CmdletBinding()]
param(
    [switch]$Execute,

    [switch]$IncludeInactive,

    [string]$LogPath = ".\GdapOrphanedAssignments_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#region Helpers
$script:LogRecords = [System.Collections.Generic.List[pscustomobject]]::new()

function Write-Log {
    param(
        [ValidateSet('INFO','WAARSCHUWING','FOUT','ACTIE','DRYRUN')]
        [string]$Niveau = 'INFO',
        [string]$Bericht,
        [string]$Klant = '',
        [string]$Relationship = ''
    )
    $kleur = switch ($Niveau) {
        'INFO'         { 'Gray' }
        'WAARSCHUWING' { 'Yellow' }
        'FOUT'         { 'Red' }
        'ACTIE'        { 'Green' }
        'DRYRUN'       { 'Cyan' }
    }
    $tijd = Get-Date -Format 'HH:mm:ss'
    Write-Host "[$tijd] [$Niveau] $Bericht" -ForegroundColor $kleur
}

function Invoke-GraphGetAll {
    param([string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($resp.value) { $items.AddRange([object[]]$resp.value) }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

# Cache zodat elke groep maar één keer opgezocht wordt
$script:GroupCache = @{}

function Resolve-GroupName {
    param([string]$GroupId)

    if ($script:GroupCache.ContainsKey($GroupId)) {
        return $script:GroupCache[$GroupId]
    }

    $result = [pscustomobject]@{ Bestaat = $false; Naam = $null }
    try {
        $g = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=id,displayName" `
            -OutputType PSObject
        $result.Bestaat = $true
        $result.Naam    = $g.displayName
    }
    catch {
        # HTTP-statuscode uit de exception halen (betrouwbaarder dan tekstmatch).
        # Verschillende Graph-/PowerShell-versies leggen de code op andere plekken.
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        elseif ($_.Exception.PSObject.Properties.Name -contains 'StatusCode' -and $_.Exception.StatusCode) {
            $statusCode = [int]$_.Exception.StatusCode
        }

        $msg = $_.Exception.Message
        $isNotFound = ($statusCode -eq 404) -or
                      ($msg -match '(?i)\bnot\s*found\b') -or
                      ($msg -match '(?i)ResourceNotFound')

        if ($isNotFound) {
            # Groep bestaat niet meer -> verweesd
            $result.Bestaat = $false
        }
        elseif ($statusCode -eq 429 -or $msg -match '(?i)too many requests|throttl') {
            # Throttling: NIET als verweesd aanmerken. Korte pauze en één retry.
            Write-Log -Niveau WAARSCHUWING -Bericht "Throttling bij opzoeken groep $GroupId; 5s pauze en opnieuw proberen."
            Start-Sleep -Seconds 5
            try {
                $g = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=id,displayName" `
                    -OutputType PSObject
                $result.Bestaat = $true
                $result.Naam    = $g.displayName
            }
            catch {
                $retryMsg  = $_.Exception.Message
                $retryCode = $null
                if ($_.Exception.Response -and $_.Exception.Response.StatusCode) { $retryCode = [int]$_.Exception.Response.StatusCode }
                if ($retryCode -eq 404 -or $retryMsg -match '(?i)\bnot\s*found\b|ResourceNotFound') {
                    $result.Bestaat = $false
                }
                else {
                    # Nog steeds onduidelijk: veilig aannemen dat de groep bestaat.
                    $result.Bestaat = $true
                    $result.Naam    = "<niet gecontroleerd: $retryMsg>"
                    Write-Log -Niveau WAARSCHUWING -Bericht "Groep $GroupId kon na retry niet gecontroleerd worden: $retryMsg"
                }
            }
        }
        else {
            # Andere fout: NIET als verweesd aanmerken (veilig), maar wel melden.
            $result.Bestaat = $true
            $result.Naam    = "<niet gecontroleerd: $msg>"
            Write-Log -Niveau WAARSCHUWING -Bericht "Groep $GroupId kon niet gecontroleerd worden (status $statusCode): $msg"
        }
    }
    $script:GroupCache[$GroupId] = $result
    return $result
}
#endregion Helpers

#region Verbinden
Write-Log -Niveau INFO -Bericht 'Verbinden met Microsoft Graph (partner tenant)...'
try {
    Connect-MgGraph -Scopes 'DelegatedAdminRelationship.ReadWrite.All','Group.Read.All' -NoWelcome -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Log -Niveau INFO -Bericht "Verbonden als $($ctx.Account) op tenant $($ctx.TenantId)"
}
catch {
    Write-Log -Niveau FOUT -Bericht "Verbinden met Graph mislukt: $($_.Exception.Message)"
    return
}
#endregion Verbinden

#region Relationships en assignments doorlopen
Write-Log -Niveau INFO -Bericht 'Alle GDAP admin relationships ophalen...'
try {
    $allRelationships = Invoke-GraphGetAll -Uri 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships'
}
catch {
    Write-Log -Niveau FOUT -Bericht "Ophalen van relationships mislukt: $($_.Exception.Message)"
    return
}

if (-not $IncludeInactive) {
    $relationships = @($allRelationships | Where-Object { $_.status -eq 'active' })
}
else {
    $relationships = @($allRelationships)
}

Write-Log -Niveau INFO -Bericht "Totaal relationships: $($allRelationships.Count) | Te controleren: $($relationships.Count)"

$orphaned = [System.Collections.Generic.List[pscustomobject]]::new()
$teller = 0

foreach ($rel in $relationships) {
    $teller++
    $klantNaam = $rel.customer.displayName
    if (-not $klantNaam) { $klantNaam = $rel.customer.tenantId }

    Write-Progress -Activity 'Relationships controleren' -Status "$teller van $($relationships.Count): $klantNaam" `
        -PercentComplete (($teller / $relationships.Count) * 100)

    try {
        $assignments = Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$($rel.id)/accessAssignments"
    }
    catch {
        Write-Log -Niveau FOUT -Klant $klantNaam -Relationship $rel.displayName -Bericht "Access assignments ophalen mislukt: $($_.Exception.Message)"
        $script:LogRecords.Add([pscustomobject]@{
            Klant = $klantNaam; KlantTenantId = $rel.customer.tenantId
            Relationship = $rel.displayName; RelationshipStatus = $rel.status
            AssignmentId = ''; GroepId = ''; GroepNaam = ''
            Status = 'FOUT'; Detail = $_.Exception.Message
        })
        continue
    }

    foreach ($a in $assignments) {
        if ($a.accessContainer.accessContainerType -ne 'securityGroup') { continue }
        if ($a.status -eq 'deleted') { continue }

        $groupId = $a.accessContainer.accessContainerId
        $info = Resolve-GroupName -GroupId $groupId

        if (-not $info.Bestaat) {
            $orphaned.Add([pscustomobject]@{
                Relationship       = $rel
                Assignment         = $a
                KlantNaam          = $klantNaam
                KlantTenantId      = $rel.customer.tenantId
                RelationshipNaam   = $rel.displayName
                RelationshipStatus = $rel.status
                GroepId            = $groupId
            })
            $script:LogRecords.Add([pscustomobject]@{
                Klant = $klantNaam; KlantTenantId = $rel.customer.tenantId
                Relationship = $rel.displayName; RelationshipStatus = $rel.status
                AssignmentId = $a.id; GroepId = $groupId; GroepNaam = ''
                Status = 'VERWEESD'; Detail = "Groep bestaat niet meer in Entra ID (assignment status: $($a.status))"
            })
        }
        else {
            $script:LogRecords.Add([pscustomobject]@{
                Klant = $klantNaam; KlantTenantId = $rel.customer.tenantId
                Relationship = $rel.displayName; RelationshipStatus = $rel.status
                AssignmentId = $a.id; GroepId = $groupId; GroepNaam = $info.Naam
                Status = 'OK'; Detail = "Assignment status: $($a.status)"
            })
        }
    }
}
Write-Progress -Activity 'Relationships controleren' -Completed
#endregion Relationships en assignments doorlopen

#region Overzicht tonen
Write-Host ''
Write-Host '================ VERWEESDE KOPPELINGEN (GUID zonder groep) ================' -ForegroundColor White

if ($orphaned.Count -eq 0) {
    Write-Log -Niveau INFO -Bericht 'Geen verweesde koppelingen gevonden. Alles is opgeruimd.'
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    Write-Log -Niveau INFO -Bericht "Rapport weggeschreven naar: $LogPath"
    Disconnect-MgGraph | Out-Null
    return
}

$orphaned | Sort-Object KlantNaam | ForEach-Object {
    Write-Host ''
    Write-Host "Klant        : $($_.KlantNaam)" -ForegroundColor White
    Write-Host "Relationship : $($_.RelationshipNaam) (status: $($_.RelationshipStatus))"
    Write-Host "  Verweesde groep : $($_.GroepId)" -ForegroundColor Yellow
    Write-Host "  Assignment      : $($_.Assignment.id) (status: $($_.Assignment.status))"
}

Write-Host ''
Write-Host "Totaal verweesde koppelingen: $($orphaned.Count)" -ForegroundColor Yellow
Write-Host '===========================================================================' -ForegroundColor White
Write-Host ''

if (-not $Execute) {
    Write-Log -Niveau DRYRUN -Bericht 'Dry-run modus: er is NIETS verwijderd. Gebruik -Execute om de verweesde koppelingen te verwijderen.'
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    Write-Log -Niveau INFO -Bericht "Rapport weggeschreven naar: $LogPath"
    Disconnect-MgGraph | Out-Null
    return
}

$bevestiging = Read-Host "Weet je zeker dat je $($orphaned.Count) verweesde koppeling(en) wilt verwijderen? (JA om door te gaan)"
if ($bevestiging -cne 'JA') {
    Write-Log -Niveau WAARSCHUWING -Bericht 'Geannuleerd door gebruiker. Er is niets verwijderd.'
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    Disconnect-MgGraph | Out-Null
    return
}
#endregion Overzicht tonen

#region Verwijderen
foreach ($item in $orphaned) {
    $rel = $item.Relationship
    $a   = $item.Assignment
    $uri = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$($rel.id)/accessAssignments/$($a.id)"

    if ($rel.status -ne 'active') {
        Write-Log -Niveau WAARSCHUWING -Klant $item.KlantNaam -Relationship $item.RelationshipNaam `
            -Bericht "Relationship is niet actief ($($rel.status)); assignment $($a.id) wordt overgeslagen."
        continue
    }

    try {
        # Verse assignment ophalen voor actuele ETag (vereist voor DELETE)
        $fresh = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        $etag  = $fresh.'@odata.etag'

        $headers = @{}
        if ($etag) { $headers['If-Match'] = $etag }

        $null = Invoke-MgGraphRequest -Method DELETE -Uri $uri -Headers $headers
        Write-Log -Niveau ACTIE -Klant $item.KlantNaam -Relationship $item.RelationshipNaam `
            -Bericht "Verweesde koppeling verwijderd: groep $($item.GroepId) (assignment $($a.id))."

        ($script:LogRecords | Where-Object { $_.AssignmentId -eq $a.id -and $_.Status -eq 'VERWEESD' }) |
            ForEach-Object { $_.Status = 'VERWIJDERD' }
    }
    catch {
        Write-Log -Niveau FOUT -Klant $item.KlantNaam -Relationship $item.RelationshipNaam `
            -Bericht "Verwijderen van assignment $($a.id) mislukt: $($_.Exception.Message)"
        ($script:LogRecords | Where-Object { $_.AssignmentId -eq $a.id }) |
            ForEach-Object { $_.Detail = "Verwijderen mislukt: $($_.Exception.Message)" }
    }
}
#endregion Verwijderen

#region Afronden
$script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
Write-Log -Niveau INFO -Bericht "Klaar. Rapport weggeschreven naar: $LogPath"
Disconnect-MgGraph | Out-Null
#endregion Afronden
