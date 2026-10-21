<#
.SYNOPSIS
    Vervangt de gekoppelde security groep in alle GDAP admin relationships
    waarvan de naam "contributor" bevat, door de groep
    SG_MicrosoftPartnerCenter_ServiceSupportContributor.

.DESCRIPTION
    Het script logt in op de partner tenant via Microsoft Graph, haalt alle
    delegated admin relationships (GDAP) op over alle klanten heen, filtert op
    relationships met "contributor" in de displayName en status "active".

    Per gevonden relationship worden de bestaande access assignments
    (gekoppelde groepen) verwijderd en wordt een nieuwe assignment aangemaakt
    voor de doelgroep, met dezelfde rollen als de verwijderde assignment.

    Standaard draait het script in dry-run modus (alleen rapporteren).
    Gebruik -Execute om de wijzigingen daadwerkelijk door te voeren.

.PARAMETER Execute
    Voert de wijzigingen daadwerkelijk uit. Zonder deze switch is het een dry-run.

.PARAMETER TargetGroupName
    DisplayName van de Entra ID security groep die gekoppeld moet worden.
    Standaard: SG_MicrosoftPartnerCenter_ServiceSupportContributor

.PARAMETER NameFilter
    Tekst die in de relationship displayName moet voorkomen (case-insensitive).
    Standaard: contributor

.PARAMETER LogPath
    Pad voor het CSV-logbestand. Standaard in de huidige map met timestamp.

.NOTES
    Benodigde Graph scopes (delegated, partner tenant):
      - DelegatedAdminRelationship.ReadWrite.All
      - Group.Read.All
    Vereist module: Microsoft.Graph.Authentication (en Microsoft.Graph.Groups)

.EXAMPLE
    .\Set-GdapContributorGroup.ps1
    Dry-run: toont wat er zou gebeuren, wijzigt niets.

.EXAMPLE
    .\Set-GdapContributorGroup.ps1 -Execute
    Voert de wijzigingen door, met bevestiging vooraf.
#>

[CmdletBinding()]
param(
    [switch]$Execute,

    [string]$TargetGroupName = 'SG_MicrosoftPartnerCenter_ServiceSupportContributor',

    [string]$NameFilter = 'contributor',

    [string]$LogPath = ".\GdapContributorGroup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#region Helpers
$script:LogRecords = [System.Collections.Generic.List[pscustomobject]]::new()

function Write-Log {
    param(
        [ValidateSet('INFO','WAARSCHUWING','FOUT','ACTIE','DRYRUN')]
        [string]$Niveau = 'INFO',
        [string]$Bericht,
        [string]$Klant = '',
        [string]$Relationship = '',
        [string]$Detail = ''
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

    $script:LogRecords.Add([pscustomobject]@{
        Tijdstip     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Niveau       = $Niveau
        Klant        = $Klant
        Relationship = $Relationship
        Bericht      = $Bericht
        Detail       = $Detail
    })
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

#region Doelgroep opzoeken
Write-Log -Niveau INFO -Bericht "Doelgroep opzoeken: $TargetGroupName"
try {
    $escaped = $TargetGroupName.Replace("'", "''")
    $groupResp = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escaped'&`$select=id,displayName,securityEnabled" `
        -OutputType PSObject
    $targetGroup = $groupResp.value | Select-Object -First 1
    if (-not $targetGroup) {
        Write-Log -Niveau FOUT -Bericht "Groep '$TargetGroupName' niet gevonden in de partner tenant. Script gestopt."
        return
    }
    if (-not $targetGroup.securityEnabled) {
        Write-Log -Niveau FOUT -Bericht "Groep '$TargetGroupName' is geen security enabled groep. Script gestopt."
        return
    }
    Write-Log -Niveau INFO -Bericht "Doelgroep gevonden: $($targetGroup.displayName) ($($targetGroup.id))"
}
catch {
    Write-Log -Niveau FOUT -Bericht "Fout bij opzoeken van de doelgroep: $($_.Exception.Message)"
    return
}
#endregion Doelgroep opzoeken

#region Relationships ophalen
Write-Log -Niveau INFO -Bericht 'Alle GDAP admin relationships ophalen...'
try {
    $allRelationships = Invoke-GraphGetAll -Uri 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships'
}
catch {
    Write-Log -Niveau FOUT -Bericht "Ophalen van relationships mislukt: $($_.Exception.Message)"
    return
}

$matching = $allRelationships | Where-Object {
    $_.displayName -like "*$NameFilter*" -and $_.status -eq 'active'
}

Write-Log -Niveau INFO -Bericht "Totaal relationships: $($allRelationships.Count) | Actief en met '$NameFilter' in de naam: $($matching.Count)"

if ($matching.Count -eq 0) {
    Write-Log -Niveau WAARSCHUWING -Bericht 'Geen relationships gevonden die aan de criteria voldoen. Niets te doen.'
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    return
}
#endregion Relationships ophalen

#region Plan opbouwen
$plan = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($rel in $matching) {
    $klantNaam = $rel.customer.displayName
    if (-not $klantNaam) { $klantNaam = $rel.customer.tenantId }

    try {
        $assignments = Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$($rel.id)/accessAssignments"
    }
    catch {
        Write-Log -Niveau FOUT -Klant $klantNaam -Relationship $rel.displayName -Bericht "Access assignments ophalen mislukt: $($_.Exception.Message)"
        continue
    }

    $activeAssignments = @($assignments | Where-Object { $_.status -in @('active','pending') })

    if ($activeAssignments.Count -eq 0) {
        Write-Log -Niveau WAARSCHUWING -Klant $klantNaam -Relationship $rel.displayName -Bericht 'Geen bestaande access assignments gevonden; nieuwe groep wordt gekoppeld met alle rollen uit de relationship zelf.'
        # Fallback: rollen uit de relationship definitie gebruiken
        $roles = @($rel.accessDetails.unifiedRoles)
        $plan.Add([pscustomobject]@{
            Relationship   = $rel
            KlantNaam      = $klantNaam
            TeVerwijderen  = @()
            NieuweRollen   = $roles
            AlGekoppeld    = $false
        })
        continue
    }

    $alreadyLinked = $activeAssignments | Where-Object { $_.accessContainer.accessContainerId -eq $targetGroup.id }
    $toRemove      = @($activeAssignments | Where-Object { $_.accessContainer.accessContainerId -ne $targetGroup.id })

    # Rollen overnemen van de bestaande assignment(s)
    $roles = @($toRemove | ForEach-Object { $_.accessDetails.unifiedRoles } | ForEach-Object { $_ }) |
             Sort-Object -Property roleDefinitionId -Unique
    if (-not $roles -or $roles.Count -eq 0) {
        $roles = @($rel.accessDetails.unifiedRoles)
    }

    $plan.Add([pscustomobject]@{
        Relationship   = $rel
        KlantNaam      = $klantNaam
        TeVerwijderen  = $toRemove
        NieuweRollen   = $roles
        AlGekoppeld    = [bool]$alreadyLinked
    })
}
#endregion Plan opbouwen

#region Plan tonen
Write-Host ''
Write-Host '================ OVERZICHT GEPLANDE WIJZIGINGEN ================' -ForegroundColor White
foreach ($item in $plan) {
    $rel = $item.Relationship
    Write-Host ''
    Write-Host "Klant        : $($item.KlantNaam)" -ForegroundColor White
    Write-Host "Relationship : $($rel.displayName) (status: $($rel.status))"
    if ($item.AlGekoppeld) {
        Write-Host "  Doelgroep is al gekoppeld." -ForegroundColor DarkGreen
    }
    foreach ($a in $item.TeVerwijderen) {
        Write-Host "  - VERWIJDEREN: groep $($a.accessContainer.accessContainerId) (assignment $($a.id), status $($a.status))" -ForegroundColor Yellow
    }
    if (-not $item.AlGekoppeld) {
        Write-Host "  + TOEVOEGEN  : $($targetGroup.displayName) met $($item.NieuweRollen.Count) rol(len)" -ForegroundColor Green
    }
}
Write-Host ''
Write-Host '=================================================================' -ForegroundColor White
Write-Host ''

if (-not $Execute) {
    Write-Log -Niveau DRYRUN -Bericht "Dry-run modus: er is NIETS gewijzigd. Gebruik -Execute om de wijzigingen door te voeren."
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    Write-Log -Niveau INFO -Bericht "Logbestand weggeschreven naar: $LogPath"
    return
}

$bevestiging = Read-Host "Weet je zeker dat je bovenstaande wijzigingen wilt doorvoeren voor $($plan.Count) relationship(s)? (JA om door te gaan)"
if ($bevestiging -cne 'JA') {
    Write-Log -Niveau WAARSCHUWING -Bericht 'Geannuleerd door gebruiker. Er is niets gewijzigd.'
    $script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
    return
}
#endregion Plan tonen

#region Uitvoeren
foreach ($item in $plan) {
    $rel = $item.Relationship
    $relBase = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$($rel.id)"

    # 1) Nieuwe assignment aanmaken (eerst toevoegen, dan pas verwijderen: geen gat in toegang)
    if (-not $item.AlGekoppeld) {
        $body = @{
            accessContainer = @{
                accessContainerId   = $targetGroup.id
                accessContainerType = 'securityGroup'
            }
            accessDetails = @{
                unifiedRoles = @($item.NieuweRollen | ForEach-Object { @{ roleDefinitionId = $_.roleDefinitionId } })
            }
        } | ConvertTo-Json -Depth 6

        try {
            $null = Invoke-MgGraphRequest -Method POST -Uri "$relBase/accessAssignments" -Body $body -ContentType 'application/json'
            Write-Log -Niveau ACTIE -Klant $item.KlantNaam -Relationship $rel.displayName -Bericht "Groep '$($targetGroup.displayName)' gekoppeld met $($item.NieuweRollen.Count) rol(len)."
        }
        catch {
            Write-Log -Niveau FOUT -Klant $item.KlantNaam -Relationship $rel.displayName -Bericht "Aanmaken nieuwe assignment mislukt: $($_.Exception.Message)" -Detail 'Bestaande assignments worden NIET verwijderd voor deze relationship.'
            continue
        }
    }

    # 2) Oude assignments verwijderen (vereist ETag via If-Match)
    foreach ($a in $item.TeVerwijderen) {
        try {
            # Verse assignment ophalen voor actuele ETag
            $fresh = Invoke-MgGraphRequest -Method GET -Uri "$relBase/accessAssignments/$($a.id)" -OutputType PSObject
            $etag  = $fresh.'@odata.etag'

            if ($fresh.status -eq 'pending') {
                Write-Log -Niveau WAARSCHUWING -Klant $item.KlantNaam -Relationship $rel.displayName -Bericht "Assignment $($a.id) heeft status 'pending' en kan mogelijk nog niet verwijderd worden; poging wordt gedaan."
            }

            $headers = @{}
            if ($etag) { $headers['If-Match'] = $etag }

            $null = Invoke-MgGraphRequest -Method DELETE -Uri "$relBase/accessAssignments/$($a.id)" -Headers $headers
            Write-Log -Niveau ACTIE -Klant $item.KlantNaam -Relationship $rel.displayName -Bericht "Oude assignment verwijderd (groep $($a.accessContainer.accessContainerId))."
        }
        catch {
            Write-Log -Niveau FOUT -Klant $item.KlantNaam -Relationship $rel.displayName -Bericht "Verwijderen van assignment $($a.id) mislukt: $($_.Exception.Message)"
        }
    }
}
#endregion Uitvoeren

#region Afronden
$script:LogRecords | Export-Csv -Path $LogPath -Delimiter ';' -Encoding utf8 -NoTypeInformation
Write-Log -Niveau INFO -Bericht "Klaar. Logbestand weggeschreven naar: $LogPath"
Disconnect-MgGraph | Out-Null
#endregion Afronden
