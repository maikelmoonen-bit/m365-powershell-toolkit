<#
.SYNOPSIS
    Genereert een grafisch attestatierapport (.xlsx) van ALLE SharePoint Online-sitecollecties
    (ongeacht sjabloon) met hun werkelijke eigenaren, elke eigenaar in een eigen kolom.

.DESCRIPTION
      1. SITE-GEDREVEN
         Alle sitecollecties worden opgehaald via Get-PnPTenantSite en per site wordt bepaald
         waar de eigenaren vandaan komen. Zo vallen ook sites zonder M365-groep in het rapport.

         Ondersteunde sjablonen (standaard):
           GROUP#0                   Team site
           SITEPAGEPUBLISHING#0      Communication site
           STS#3                     Team site (no Microsoft 365 group)
           STS#0                     Team site (classic experience)
           PROJECTSITE#0             Project Site
           ENTERWIKI#0               Enterprise Wiki
           BDR#0                     Document Center
           APPCATALOG#0              App Catalog Site
           COMMUNITY#0               Community Site
           CMSPUBLISHING#0           Publishing Site
           BLANKINTERNETCONTAINER#0  Publishing Portal
           EDISC#0 / EDISC#1         eDiscovery Center / Case
           PWA#0 / PWS#0             Project Web App Site / Microsoft Project Site

      2. EIGENAARBRONNEN PER SJABLOON
         Groepssites (GROUP#0, TEAMCHANNEL#*) -> eigenaren van de gekoppelde M365-groep via Graph.
         Alle overige sjablonen              -> leden van de SharePoint Owners-groep (de
                                                associated owner group) via PnP.

         Sitecollectiebeheerders worden standaard NIET gerapporteerd.
         Sitecollectiebeheerders zijn een beheerconstruct - daar staan doorgaans SharePoint-
         beheerders, migratie-accounts en service-principals in. De functionele eigenaar van een
         site is het lid van de Owners-groep, en dat is wat bij attestatie geattesteerd moet
         worden. Met -IncludeSiteCollectionAdmins kunnen de beheerders desgewenst worden
         toegevoegd als aanvulling.

         LET OP: Microsoft Graph biedt GEEN API voor SharePoint-groepen. PnP.PowerShell is
         daarom een harde vereiste voor alles wat geen M365-groep heeft.

      3. EIGENAREN IN AFZONDERLIJKE KOLOMMEN
         In plaats van een enkele kolom 'Eigenaren (namen)' krijgt het rapport de kolommen
         'Eigenaar 1 (naam)' t/m 'Eigenaar N (naam)', gevolgd door 'Eigenaar 1 (e-mail)' t/m
         'Eigenaar N (e-mail)'. N is ONBEGRENSD en wordt bepaald door de site met de meeste
         eigenaren in de tenant; er wordt niets samengevoegd of afgekapt.

      4. ACTIVITEITSBEPALING MET FALLBACK
         Primair het Graph-gebruiksrapport (op Site Id); ontbreekt de site daarin, dan wordt
         LastContentModifiedDate uit Get-PnPTenantSite gebruikt. Classic sites ontbreken namelijk
         regelmatig in het gebruiksrapport, wat anders massaal 'Onbekend' zou opleveren.

    Attestatiecriteria:
      1. Te weinig eigenaren   (standaard: minder dan 2)
      2. Komma in de sitenaam  (naamgevingsconventie / CSV-risico)
      3. Lange inactiviteit    (standaard: geen activiteit in 180 dagen)

    Werkbladen: Dashboard, Sites & Eigenaren, Attestatie, Eigenaren (plat).

.PARAMETER OutputPath
    Pad van het te genereren xlsx-bestand. Standaard: .\SPO-Attestatie-<jjjjMMdd>.xlsx

.PARAMETER MinimumOwners
    Minimum aantal eigenaren voordat een site als risico wordt gemarkeerd. Standaard: 2

.PARAMETER InactiveDaysThreshold
    Aantal dagen zonder activiteit voordat een site als inactief wordt gemarkeerd. Standaard: 180

.PARAMETER IncludeTemplates
    Overschrijft de standaardlijst met sjablooncodes. Bijv. 'STS#0','STS#3'.

.PARAMETER IncludeAllTemplates
    Neem elk sjabloon mee behalve OneDrive, redirect- en systeemsites.

.PARAMETER IncludeSiteCollectionAdmins
    Vul de leden van de Owners-groep aan met de sitecollectiebeheerders. Standaard UIT: die
    groep bevat doorgaans beheer- en serviceaccounts die geen functionele eigenaar zijn.

.PARAMETER SkipOwnerLookup
    Sla het per-site ophalen van de Owners-groep over en gebruik alleen de primaire eigenaar uit
    Get-PnPTenantSite. Veel sneller, maar onvolledig; bedoeld om een run snel te valideren.

.PARAMETER PnPClientId
    App-registratie-id voor interactieve PnP-aanmelding. Sinds september 2024 heeft PnP.PowerShell
    geen ingebouwde multi-tenant app meer; in delegated modus is dit verplicht.

.PARAMETER UseManagedIdentity
    Verbind met de system-assigned managed identity van het Automation-account.

.PARAMETER ClientId / TenantId / CertificateThumbprint
    App-only (certificaat) auth voor zowel Graph als PnP.

.PARAMETER TenantDomain
    Bijv. contoso.onmicrosoft.com. Alleen nodig als automatische detectie faalt.

.PARAMETER MailTo / MailFrom
    Verzend het rapport als bijlage. Vereist de application-permissie Mail.Send.

.PARAMETER OpenAfterExport
    Open het rapport direct in Excel na afronding.

.EXAMPLE
    .\Export-SPOSiteAttestatieRapport.ps1 -PnPClientId '<app-id>'

.EXAMPLE
    .\Export-SPOSiteAttestatieRapport.ps1 -ClientId '<app-id>' -TenantId '<tenant-id>' -CertificateThumbprint '<thumbprint>'

.NOTES
    Vereiste modules : Microsoft.Graph.Authentication, PnP.PowerShell, ImportExcel

    Graph-permissies (app) : Group.Read.All, GroupMember.Read.All, User.Read.All, Sites.Read.All,
                             Reports.Read.All (+ Mail.Send bij -MailTo)
    SharePoint (PnP, app)  : Sites.FullControl.All  <-- Sites.Read.All is NIET voldoende;
                             het uitlezen van sitecollectiebeheerders vereist volledig beheer.

    LET OP: User.Read.All is onmisbaar. Zonder die permissie levert Graph bij /groups/{id}/owners
    alleen object-ID's; namen en e-mailadressen blijven dan leeg.
#>

[CmdletBinding(DefaultParameterSetName = 'Delegated')]
param(
    [string]$OutputPath = ".\SPO-Attestatie-$(Get-Date -Format 'yyyyMMdd').xlsx",
    [ValidateRange(1, 10)]
    [int]$MinimumOwners = 2,
    [ValidateRange(30, 180)]
    [int]$InactiveDaysThreshold = 180,

    [string[]]$IncludeTemplates,
    [switch]$IncludeAllTemplates,
    [switch]$IncludeSiteCollectionAdmins,
    [switch]$SkipOwnerLookup,

    [Parameter(ParameterSetName = 'Delegated')]
    [string]$PnPClientId,

    [Parameter(ParameterSetName = 'ManagedIdentity', Mandatory)]
    [switch]$UseManagedIdentity,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$ClientId,
    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$TenantId,
    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$CertificateThumbprint,

    [string]$TenantDomain,
    [string]$MailTo,
    [string]$MailFrom,

    [switch]$OpenAfterExport
)

$ProgressPreference = 'SilentlyContinue'

#region -- Sjabloondefinities ---------------------------------------------------
# Vriendelijke namen zoals ze in het SharePoint-admincenter worden getoond.
$SjabloonNamen = @{
    'GROUP#0'                  = 'Team site'
    'TEAMCHANNEL#0'            = 'Team site (kanaalsite)'
    'TEAMCHANNEL#1'            = 'Team site (privé kanaal)'
    'SITEPAGEPUBLISHING#0'     = 'Communication site'
    'STS#3'                    = 'Team site (no Microsoft 365 group)'
    'STS#0'                    = 'Team site (classic experience)'
    'PROJECTSITE#0'            = 'Project Site'
    'ENTERWIKI#0'              = 'Enterprise Wiki'
    'BDR#0'                    = 'Document Center'
    'APPCATALOG#0'             = 'App Catalog Site'
    'COMMUNITY#0'              = 'Community Site'
    'CMSPUBLISHING#0'          = 'Publishing Site'
    'BLANKINTERNETCONTAINER#0' = 'Publishing Portal'
    'EDISC#0'                  = 'eDiscovery Center'
    'EDISC#1'                  = 'eDiscovery Case'
    'PWA#0'                    = 'Project Web App Site'
    'PWS#0'                    = 'Microsoft Project Site'
    'SRCHCEN#0'                = 'Enterprise Search Center'
    'SRCHCENTERLITE#0'         = 'Basic Search Center'
    'BLANKINTERNET#0'          = 'Publishing Site (klassiek)'
    'POINTPUBLISHINGTOPIC#0'   = 'Viva Topics'
    'POINTPUBLISHINGHUB#0'     = 'Organisatienieuwsportaal'
    'STS#1'                    = 'Document Workspace (klassiek)'
    'BLANKINTERNETCONTAINER#1' = 'Publishing Portal'
}

# Standaardselectie van sjablonen; overschrijfbaar met -IncludeTemplates.
$StandaardSjablonen = @(
    'GROUP#0'
    'SITEPAGEPUBLISHING#0'
    'STS#3'
    'STS#0'
    'PROJECTSITE#0'
    'ENTERWIKI#0'
    'BDR#0'
    'APPCATALOG#0'
    'COMMUNITY#0'
    'CMSPUBLISHING#0'
    'BLANKINTERNETCONTAINER#0'
    'EDISC#0'
    'PWA#0'
)

# Nooit meenemen, ook niet bij -IncludeAllTemplates.
$UitgeslotenSjabloonPatronen = @(
    '^SPSPERS'          # OneDrive for Business
    '^REDIRECTSITE'     # redirect-stubs na hernoemen
    '^SPSMSITEHOST'     # My Site host
    '^POINTPUBLISHINGPERSONAL'
    '^TEAMCHANNEL'      # kanaalsites erven de eigenaren van het bovenliggende team
)
#endregion

#region -- Hulpfuncties ---------------------------------------------------------
function Write-Log {
    param(
        [string]$Bericht,
        [ValidateSet('INFO', 'WAARSCHUWING', 'FOUT', 'OK')]
        [string]$Niveau = 'INFO'
    )
    $regel = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Niveau, $Bericht
    if ($env:AUTOMATION_ASSET_ACCOUNTID) { Write-Output $regel }
    else {
        $kleur = switch ($Niveau) { 'INFO' { 'Cyan' } 'WAARSCHUWING' { 'Yellow' } 'FOUT' { 'Red' } 'OK' { 'Green' } }
        Write-Host $regel -ForegroundColor $kleur
    }
}

function Get-Veld {
    <#
        Leest een veld uit een Graph-antwoord, ongeacht of dat is gedeserialiseerd als Hashtable
        (standaard bij Invoke-MgGraphRequest) of als PSCustomObject. Zoekt hoofdletterongevoelig.
    #>
    param($Object, [string]$Naam)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Naam)) { return $Object[$Naam] }
        foreach ($sleutel in $Object.Keys) { if ($sleutel -ieq $Naam) { return $Object[$sleutel] } }
        return $null
    }
    $eigenschap = $Object.PSObject.Properties[$Naam]
    if ($eigenschap) { return $eigenschap.Value }
    foreach ($p in $Object.PSObject.Properties) { if ($p.Name -ieq $Naam) { return $p.Value } }
    return $null
}

function ConvertTo-KolomLetter {
    <# Zet een 1-gebaseerde kolomindex om naar de Excel-letternotatie (1 -> A, 27 -> AA). #>
    param([int]$Index)
    $letter = ''
    while ($Index -gt 0) {
        $rest = ($Index - 1) % 26
        $letter = [char](65 + $rest) + $letter
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $letter
}

function Get-KolomLetterVoorKop {
    <#
        Zoekt in rij 1 van een werkblad naar een kopnaam en geeft de kolomletter terug.
        Nodig omdat het aantal eigenaarkolommen dynamisch is; harde verwijzingen als "Q2:Q100"
        zouden na uitbreiding naar de verkeerde kolom wijzen.
    #>
    param($Werkblad, [string]$Kop)
    for ($k = 1; $k -le $Werkblad.Dimension.End.Column; $k++) {
        if ([string]$Werkblad.Cells[1, $k].Value -eq $Kop) { return (ConvertTo-KolomLetter $k) }
    }
    return $null
}

function Test-BestandVergrendeld {
    param([string]$Pad)
    if (-not (Test-Path $Pad)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Pad, 'Open', 'ReadWrite', 'None')
        $stream.Close(); $stream.Dispose()
        return $false
    }
    catch { return $true }
}

function ConvertTo-GenormaliseerdeUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    return $Url.TrimEnd('/').ToLowerInvariant()
}

function Test-SjabloonUitgesloten {
    param([string]$Sjabloon)
    foreach ($patroon in $script:UitgeslotenSjabloonPatronen) {
        if ($Sjabloon -match $patroon) { return $true }
    }
    return $false
}

function Invoke-GraphBatch {
    param([array]$Requests, [string]$Activiteit = 'Graph-batch')

    $resultaten = @{}
    if (-not $Requests -or $Requests.Count -eq 0) { return $resultaten }
    $totaal = [math]::Ceiling($Requests.Count / 20)
    $batchNr = 0
    for ($offset = 0; $offset -lt $Requests.Count; $offset += 20) {
        $batchNr++
        Write-Progress -Activity $Activiteit -Status "Batch $batchNr / $totaal" -PercentComplete ([math]::Min(100, ($batchNr / $totaal) * 100))
        $einde = [math]::Min($offset + 19, $Requests.Count - 1)
        $chunk = $Requests[$offset..$einde]
        $body = @{ requests = @($chunk | ForEach-Object { @{ id = $_.id; method = 'GET'; url = $_.url } }) } | ConvertTo-Json -Depth 5

        $pogingen = 0
        do {
            $herhaal = $false
            $pogingen++
            try {
                $antwoord = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/$batch' -Body $body -ContentType 'application/json' -ErrorAction Stop
                foreach ($response in $antwoord.responses) {
                    if ($response.status -eq 429 -and $pogingen -lt 5) {
                        $wachttijd = if ($response.headers.'Retry-After') { [int]$response.headers.'Retry-After' } else { 5 }
                        Start-Sleep -Seconds $wachttijd
                        $herhaal = $true
                        break
                    }
                    $resultaten[$response.id] = @{ Status = [int]$response.status; Body = $response.body }
                }
            }
            catch {
                if ($_.Exception.Message -match '429' -and $pogingen -lt 5) {
                    Start-Sleep -Seconds (5 * $pogingen); $herhaal = $true
                }
                else {
                    foreach ($request in $chunk) {
                        if (-not $resultaten.ContainsKey($request.id)) {
                            $resultaten[$request.id] = @{ Status = 0; Body = @{ error = @{ message = $_.Exception.Message } } }
                        }
                    }
                }
            }
        } while ($herhaal)
    }
    Write-Progress -Activity $Activiteit -Completed
    return $resultaten
}

function Connect-PnPDoelsite {
    <# Verbindt PnP met een specifieke site-URL in dezelfde modus als de hoofdverbinding. #>
    param([string]$Url)
    switch ($script:modus) {
        'ManagedIdentity' { Connect-PnPOnline -Url $Url -ManagedIdentity -ErrorAction Stop }
        'AppOnly'         { Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $script:pnpTenant -Thumbprint $CertificateThumbprint -ErrorAction Stop }
        default           { Connect-PnPOnline -Url $Url -Interactive -ClientId $PnPClientId -ErrorAction Stop }
    }
}

function Send-RapportMail {
    param([string]$Pad, [string]$Van, [string]$Naar, [string]$Samenvatting)
    $bytes = [System.IO.File]::ReadAllBytes($Pad)
    $bericht = @{
        message = @{
            subject = "SharePoint site-attestatierapport $(Get-Date -Format 'dd-MM-yyyy')"
            body    = @{ contentType = 'Text'; content = $Samenvatting }
            toRecipients = @(@{ emailAddress = @{ address = $Naar } })
            attachments  = @(@{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name          = [System.IO.Path]::GetFileName($Pad)
                contentType   = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                contentBytes  = [System.Convert]::ToBase64String($bytes)
            })
        }
        saveToSentItems = $true
    } | ConvertTo-Json -Depth 6
    Invoke-MgGraphRequest -Method POST -Uri "v1.0/users/$Van/sendMail" -Body $bericht -ContentType 'application/json' -ErrorAction Stop
}
#endregion

#region -- Voorcontrole ---------------------------------------------------------
if ($MailTo -and -not $MailFrom) {
    Write-Log '-MailTo opgegeven zonder -MailFrom; beide zijn nodig om te kunnen verzenden.' -Niveau FOUT
    return
}
if ($PSCmdlet.ParameterSetName -eq 'Delegated' -and -not $PnPClientId) {
    Write-Log 'In delegated modus is -PnPClientId verplicht: PnP.PowerShell heeft sinds september 2024 geen ingebouwde multi-tenant app meer.' -Niveau FOUT
    Write-Log "Registreer eenmalig een app met 'Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP SPO Attestatie' -Tenant <tenant>.onmicrosoft.com -Interactive'." -Niveau FOUT
    return
}
if (Test-BestandVergrendeld -Pad $OutputPath) {
    Write-Log "Het uitvoerbestand '$OutputPath' is in gebruik door een ander proces (waarschijnlijk Excel)." -Niveau FOUT
    return
}
foreach ($module in @('Microsoft.Graph.Authentication', 'PnP.PowerShell', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "Module '$module' ontbreekt. Installeren met: Install-Module $module -Scope CurrentUser" -Niveau FOUT
        return
    }
}
Import-Module ImportExcel -ErrorAction Stop
Import-Module PnP.PowerShell -ErrorAction Stop
try {
    $importExcelVersie = (Get-Module ImportExcel | Select-Object -First 1).Version
    $pnpVersie         = (Get-Module PnP.PowerShell | Select-Object -First 1).Version
    $epplusVersie      = [OfficeOpenXml.ExcelPackage].Assembly.GetName().Version
    Write-Log "ImportExcel $importExcelVersie | EPPlus $epplusVersie | PnP.PowerShell $pnpVersie"
}
catch {
    Write-Log "Versie-informatie kon niet worden uitgelezen: $($_.Exception.Message)" -Niveau WAARSCHUWING
}

# Bepaal de te rapporteren sjablonen.
if ($IncludeTemplates) {
    $doelSjablonen = @($IncludeTemplates | ForEach-Object { $_.ToUpperInvariant() })
    Write-Log "Sjabloonselectie via -IncludeTemplates: $($doelSjablonen -join ', ')"
}
elseif ($IncludeAllTemplates) {
    $doelSjablonen = $null   # null betekent: alles behalve de uitgesloten patronen
    Write-Log 'Sjabloonselectie: alle sjablonen behalve OneDrive-, redirect- en systeemsites.'
}
else {
    $doelSjablonen = $StandaardSjablonen
    Write-Log "Sjabloonselectie: standaardset van $($doelSjablonen.Count) sjablonen."
}
#endregion

#region -- Verbinding -----------------------------------------------------------
$modus = $PSCmdlet.ParameterSetName
$tenantBreed = ($modus -ne 'Delegated')
Write-Log "Verbinding maken met Microsoft Graph (modus: $modus)..."
try {
    switch ($modus) {
        'ManagedIdentity' { Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop }
        'AppOnly'         { Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop }
        default           { Connect-MgGraph -Scopes 'Group.Read.All', 'GroupMember.Read.All', 'User.Read.All', 'Sites.Read.All', 'Reports.Read.All' -NoWelcome -ErrorAction Stop }
    }
    $context = Get-MgContext
    Write-Log "Verbonden met tenant: $($context.TenantId) als $(if ($context.Account) { $context.Account } else { $context.AppName })" -Niveau OK
}
catch {
    Write-Log "Verbinden met Microsoft Graph mislukt: $($_.Exception.Message)" -Niveau FOUT
    return
}

# Admin-URL afleiden uit de roothostname; nodig voor Get-PnPTenantSite.
try {
    $rootSite = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/sites/root' -ErrorAction Stop
    $hostnaam = Get-Veld $rootSite 'siteCollectionHostname'
    if (-not $hostnaam) { $hostnaam = ([uri](Get-Veld $rootSite 'webUrl')).Host }
    $tenantPrefix = ($hostnaam -split '\.')[0]
    $adminUrl     = "https://$tenantPrefix-admin.sharepoint.com"
    $pnpTenant    = if ($TenantDomain) { $TenantDomain } else { "$tenantPrefix.onmicrosoft.com" }
    Write-Log "SharePoint-admincenter: $adminUrl | PnP-tenant: $pnpTenant"
}
catch {
    Write-Log "Roothostname kon niet worden bepaald: $($_.Exception.Message)" -Niveau FOUT
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Log 'Verbinding maken met SharePoint Online via PnP.PowerShell...'
try {
    Connect-PnPDoelsite -Url $adminUrl
    Write-Log 'PnP-verbinding met het admincenter tot stand gebracht.' -Niveau OK
}
catch {
    Write-Log "PnP-verbinding mislukt: $($_.Exception.Message)" -Niveau FOUT
    if ($tenantBreed) {
        Write-Log 'Controleer of de app/managed identity de SharePoint-application-permissie Sites.FullControl.All heeft met admin consent.' -Niveau FOUT
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
#endregion

#region -- Stap 1: activiteitsdata uit het gebruiksrapport (op Site Id) ---------
Write-Log 'SharePoint-gebruiksrapport ophalen (periode D180)...'
$activiteitPerSiteId = @{}
$tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "SPOUsage_$(Get-Random).csv"
try {
    Invoke-MgGraphRequest -Method GET -Uri "v1.0/reports/getSharePointSiteUsageDetail(period='D180')" -OutputFilePath $tempCsv -ErrorAction Stop
    foreach ($rij in (Import-Csv -Path $tempCsv)) {
        $sid = $rij.'Site Id'
        if ([string]::IsNullOrWhiteSpace($sid)) { continue }
        $activiteitPerSiteId[$sid.Trim().ToLowerInvariant()] = [PSCustomObject]@{
            LastActivityDate = if ($rij.'Last Activity Date') { [datetime]$rij.'Last Activity Date' } else { $null }
            FileCount        = if ($rij.'File Count')          { [int64]$rij.'File Count' }          else { 0 }
            StorageUsedBytes = if ($rij.'Storage Used (Byte)') { [int64]$rij.'Storage Used (Byte)' } else { 0 }
        }
    }
    Write-Log "$($activiteitPerSiteId.Count) sites met activiteitsdata geladen." -Niveau OK
}
catch {
    Write-Log "Gebruiksrapport kon niet worden opgehaald: $($_.Exception.Message). Er wordt teruggevallen op LastContentModifiedDate." -Niveau WAARSCHUWING
}
finally {
    if (Test-Path $tempCsv) { Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue }
}
#endregion

#region -- Stap 2: Graph-site-index (voor Site Id per URL) ----------------------
# Het gebruiksrapport is gesleuteld op de sitecollectie-GUID. Get-PnPTenantSite levert die niet
# in alle moduleversies, dus de koppeling loopt via de Graph-site-index op URL.
if ($tenantBreed) {
    Write-Log 'Graph-site-index opbouwen via /sites/getAllSites...'
    $uri = 'v1.0/sites/getAllSites'
}
else {
    Write-Log 'Graph-site-index opbouwen via /sites?search=* (delegated)...'
    $uri = 'v1.0/sites?search=*'
}
$siteIdPerUrl = @{}
try {
    while ($uri) {
        $pagina = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($site in $pagina.value) {
            $sleutel = ConvertTo-GenormaliseerdeUrl (Get-Veld $site 'webUrl')
            if ($sleutel -and -not $siteIdPerUrl.ContainsKey($sleutel)) {
                $siteIdPerUrl[$sleutel] = Get-Veld $site 'id'
            }
        }
        $uri = $pagina.'@odata.nextLink'
    }
    Write-Log "$($siteIdPerUrl.Count) sites in de Graph-index." -Niveau OK
}
catch {
    Write-Log "Graph-site-index kon niet volledig worden opgebouwd: $($_.Exception.Message). Activiteit valt terug op LastContentModifiedDate." -Niveau WAARSCHUWING
}
#endregion

#region -- Stap 3: sitecollecties ophalen via PnP -------------------------------
Write-Log 'Alle sitecollecties ophalen via Get-PnPTenantSite (dit kan enkele minuten duren)...'
try {
    $alleSites = @(Get-PnPTenantSite -IncludeOneDriveSites:$false -Detailed -ErrorAction Stop)
}
catch {
    Write-Log "Sitecollecties konden niet worden opgehaald: $($_.Exception.Message)" -Niveau FOUT
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
Write-Log "$($alleSites.Count) sitecollecties opgehaald." -Niveau OK

$geselecteerd = [System.Collections.Generic.List[object]]::new()
$sjabloonTellers = @{}
$overgeslagen = 0
foreach ($s in $alleSites) {
    $sjabloon = [string](Get-Veld $s 'Template')
    if ([string]::IsNullOrWhiteSpace($sjabloon)) { $sjabloon = 'ONBEKEND' }
    $sjabloon = $sjabloon.ToUpperInvariant()

    if (Test-SjabloonUitgesloten -Sjabloon $sjabloon) { $overgeslagen++; continue }
    if ($doelSjablonen -and $sjabloon -notin $doelSjablonen) { $overgeslagen++; continue }

    $status = [string](Get-Veld $s 'Status')
    if ($status -and $status -ne 'Active') { $overgeslagen++; continue }

    if (-not $sjabloonTellers.ContainsKey($sjabloon)) { $sjabloonTellers[$sjabloon] = 0 }
    $sjabloonTellers[$sjabloon]++
    $geselecteerd.Add($s)
}
Write-Log "$($geselecteerd.Count) sites geselecteerd | $overgeslagen overgeslagen (sjabloon/status)." -Niveau OK
foreach ($sleutel in ($sjabloonTellers.Keys | Sort-Object)) {
    $vriendelijk = if ($SjabloonNamen.ContainsKey($sleutel)) { $SjabloonNamen[$sleutel] } else { $sleutel }
    Write-Log ("  {0,-28} {1,5}  ({2})" -f $sleutel, $sjabloonTellers[$sleutel], $vriendelijk)
}
if ($geselecteerd.Count -eq 0) {
    Write-Log 'Geen enkele site voldoet aan de sjabloonselectie; rapport kan niet worden opgebouwd.' -Niveau FOUT
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
#endregion

#region -- Stap 4a: eigenaren van groepssites via Graph -------------------------
$groepsSites = @($geselecteerd | Where-Object {
    $gid = [string](Get-Veld $_ 'GroupId')
    $gid -and $gid -ne [guid]::Empty.ToString()
})
Write-Log "$($groepsSites.Count) groepsgekoppelde sites | $($geselecteerd.Count - $groepsSites.Count) sites zonder M365-groep."

$eigenarenPerGroep = @{}
$groepsDetails     = @{}

if ($groepsSites.Count -gt 0) {
    # Groepsmetadata (mail, zichtbaarheid, aanmaakdatum) in één keer ophalen.
    Write-Log 'Metadata van M365-groepen ophalen...'
    $groepIds = @($groepsSites | ForEach-Object { [string](Get-Veld $_ 'GroupId') } | Select-Object -Unique)
    $metaRequests = $groepIds | ForEach-Object {
        @{ id = $_; url = "/groups/$_" + '?$select=id,displayName,mail,visibility,createdDateTime' }
    }
    $metaResultaten = Invoke-GraphBatch -Requests $metaRequests -Activiteit 'Groepsmetadata ophalen'
    foreach ($gid in $groepIds) {
        $r = $metaResultaten[$gid]
        if ($r -and $r.Status -eq 200) {
            $groepsDetails[$gid] = [PSCustomObject]@{
                Mail            = Get-Veld $r.Body 'mail'
                Zichtbaarheid   = Get-Veld $r.Body 'visibility'
                CreatedDateTime = Get-Veld $r.Body 'createdDateTime'
            }
        }
    }

    Write-Log 'Eigenaren van M365-groepen ophalen via $batch (slag 1/2)...'
    $eigenaarRequests = $groepIds | ForEach-Object { @{ id = $_; url = "/groups/$_/owners" } }
    $eigenaarResultaten = Invoke-GraphBatch -Requests $eigenaarRequests -Activiteit 'Groepseigenaren ophalen (slag 1/2)'

    $ruweEigenaren = @{}
    $eigenaarFouten = 0; $eersteEigenaarFout = $null
    foreach ($gid in $groepIds) {
        $resultaat = $eigenaarResultaten[$gid]
        if ($resultaat -and $resultaat.Status -eq 200) {
            $lijst    = [System.Collections.Generic.List[object]]::new()
            $waarden  = Get-Veld $resultaat.Body 'value'
            $volgende = Get-Veld $resultaat.Body '@odata.nextLink'
            while ($true) {
                foreach ($e in $waarden) {
                    $lijst.Add([PSCustomObject]@{
                        Id    = Get-Veld $e 'id'
                        Naam  = Get-Veld $e 'displayName'
                        Email = if (Get-Veld $e 'mail') { Get-Veld $e 'mail' } else { Get-Veld $e 'userPrincipalName' }
                        Type  = ((Get-Veld $e '@odata.type') -replace '#microsoft.graph.', '')
                    })
                }
                if (-not $volgende) { break }
                $pagina   = Invoke-MgGraphRequest -Method GET -Uri $volgende -ErrorAction SilentlyContinue
                $waarden  = Get-Veld $pagina 'value'
                $volgende = Get-Veld $pagina '@odata.nextLink'
            }
            $ruweEigenaren[$gid] = $lijst
        }
        else {
            $eigenaarFouten++
            if (-not $eersteEigenaarFout -and $resultaat) {
                $eersteEigenaarFout = "HTTP $($resultaat.Status): $(Get-Veld (Get-Veld $resultaat.Body 'error') 'message')"
            }
            $ruweEigenaren[$gid] = @()
        }
    }
    if ($eigenaarFouten -gt 0) { Write-Log "$eigenaarFouten eigenaar-calls mislukt. Eerste fout: $eersteEigenaarFout" -Niveau WAARSCHUWING }

    # Slag 2: ontbrekende namen alsnog ophalen op object-ID.
    $openIds = @(
        $ruweEigenaren.Values | ForEach-Object { $_ } |
        Where-Object { $_.Id -and [string]::IsNullOrWhiteSpace($_.Naam) } |
        Select-Object -ExpandProperty Id -Unique
    )
    $profielen = @{}
    if ($openIds.Count -gt 0) {
        Write-Log "$($openIds.Count) eigenaren zonder naam; profielen ophalen op ID..."
        $profielRequests = $openIds | ForEach-Object { @{ id = $_; url = "/directoryObjects/$_" } }
        $profielResultaten = Invoke-GraphBatch -Requests $profielRequests -Activiteit 'Groepseigenaren ophalen (slag 2/2)'
        foreach ($oid in $openIds) {
            $r = $profielResultaten[$oid]
            if ($r -and $r.Status -eq 200) {
                $naam = Get-Veld $r.Body 'displayName'
                $upn  = Get-Veld $r.Body 'userPrincipalName'
                $mail = Get-Veld $r.Body 'mail'
                $profielen[$oid] = [PSCustomObject]@{
                    Naam  = if ($naam) { $naam } elseif ($upn) { $upn } else { $null }
                    Email = if ($mail) { $mail } elseif ($upn) { $upn } else { '' }
                }
            }
        }
    }

    foreach ($gid in $ruweEigenaren.Keys) {
        $lijst = [System.Collections.Generic.List[object]]::new()
        foreach ($e in $ruweEigenaren[$gid]) {
            $naam = $e.Naam; $mail = $e.Email
            if ([string]::IsNullOrWhiteSpace($naam) -and $e.Id -and $profielen.ContainsKey($e.Id)) {
                $naam = $profielen[$e.Id].Naam
                if ([string]::IsNullOrWhiteSpace($mail)) { $mail = $profielen[$e.Id].Email }
            }
            if ([string]::IsNullOrWhiteSpace($naam)) { $naam = '(naam onbekend)' }
            $lijst.Add([PSCustomObject]@{ Naam = $naam; Email = $mail; Type = $e.Type; Bron = 'M365-groep' })
        }
        $eigenarenPerGroep[$gid] = $lijst
    }

    $totaalGroepsEigenaren = ($eigenarenPerGroep.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    $naamloos = @($eigenarenPerGroep.Values | ForEach-Object { $_ } | Where-Object { $_.Naam -eq '(naam onbekend)' }).Count
    Write-Log "$totaalGroepsEigenaren groepseigenaren opgehaald ($naamloos zonder naam)." -Niveau $(if ($naamloos -gt 0) { 'WAARSCHUWING' } else { 'OK' })
    if ($naamloos -gt 0 -and $totaalGroepsEigenaren -gt 0 -and ($naamloos / $totaalGroepsEigenaren) -gt 0.5) {
        Write-Log "Vrijwel alle eigenaarnamen zijn leeg terwijl het AANTAL klopt. Dit duidt op een ontbrekende permissie: voeg 'User.Read.All' toe en verleen admin consent." -Niveau FOUT
    }
}
#endregion

#region -- Stap 4b: Site Owners van niet-groepssites via SharePoint -------------
# Microsoft Graph kent geen API voor SharePoint-groepen. Voor elke niet-groepssite wordt daarom
# een aparte PnP-verbinding opgezet. Dat is de traagste stap van het script.
#
# BRON = DE OWNERS-GROEP, NIET DE SITECOLLECTIEBEHEERDERS.
# De associated owner group ("<Sitenaam> Owners") bevat de functionele eigenaren van de site.
# Sitecollectiebeheerders zijn een beheerconstruct en bevatten doorgaans SharePoint-beheerders,
# migratie-accounts en service-principals; die horen niet in een attestatierapport thuis.
# Met -IncludeSiteCollectionAdmins kunnen ze alsnog als aanvulling worden meegenomen.
$eigenarenPerSiteUrl = @{}
$nietGroepsSites = @($geselecteerd | Where-Object {
    $gid = [string](Get-Veld $_ 'GroupId')
    -not $gid -or $gid -eq [guid]::Empty.ToString()
})

# Accounts die nooit als functionele eigenaar gelden.
$systeemPatroon = 'app@sharepoint|SHAREPOINT\\system|^c:0\(\.s\|true|^c:0!\.s\|windows|_SPOApp|NT AUTHORITY'

if ($nietGroepsSites.Count -gt 0 -and -not $SkipOwnerLookup) {
    Write-Log "Site Owners ophalen voor $($nietGroepsSites.Count) sites zonder M365-groep..."
    $teller = 0; $ownerFouten = 0; $eersteOwnerFout = $null; $zonderOwnersGroep = 0
    foreach ($s in $nietGroepsSites) {
        $teller++
        $siteUrl = [string](Get-Veld $s 'Url')
        Write-Progress -Activity 'Site Owners ophalen' -Status "$teller / $($nietGroepsSites.Count): $siteUrl" `
            -PercentComplete ([math]::Min(100, ($teller / $nietGroepsSites.Count) * 100))

        $lijst = [System.Collections.Generic.List[object]]::new()
        try {
            Connect-PnPDoelsite -Url $siteUrl

            # --- Primaire bron: de associated owner group van de site.
            try {
                $ownersGroep = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
                foreach ($lid in @(Get-PnPGroupMember -Identity $ownersGroep -ErrorAction Stop)) {
                    $login = [string]$lid.LoginName
                    if ($login -match $systeemPatroon) { continue }
                    $lijst.Add([PSCustomObject]@{
                        Naam  = if ($lid.Title) { $lid.Title } else { $login }
                        Email = [string]$lid.Email
                        Type  = if ($login -match '^c:0') { 'group' } else { 'user' }
                        Bron  = 'Site Owners-groep'
                    })
                }
            }
            catch {
                # Sommige sjablonen (o.a. APPCATALOG#0, EDISC#1) hebben geen associated owner group.
                $zonderOwnersGroep++
                Write-Verbose "Geen Owners-groep op $siteUrl : $($_.Exception.Message)"
            }

            # --- Optionele aanvulling: sitecollectiebeheerders.
            if ($IncludeSiteCollectionAdmins) {
                try {
                    foreach ($beheerder in @(Get-PnPSiteCollectionAdmin -ErrorAction Stop)) {
                        $login = [string]$beheerder.LoginName
                        if ($login -match $systeemPatroon) { continue }
                        $naam = if ($beheerder.Title) { $beheerder.Title } else { $login }
                        if (@($lijst | Where-Object { $_.Naam -eq $naam }).Count -gt 0) { continue }
                        $lijst.Add([PSCustomObject]@{
                            Naam  = $naam
                            Email = [string]$beheerder.Email
                            Type  = if ($login -match '^c:0') { 'group' } else { 'user' }
                            Bron  = 'Sitecollectiebeheerder'
                        })
                    }
                }
                catch { Write-Verbose "Sitecollectiebeheerders niet leesbaar op $siteUrl : $($_.Exception.Message)" }
            }
        }
        catch {
            $ownerFouten++
            if (-not $eersteOwnerFout) { $eersteOwnerFout = "$siteUrl -> $($_.Exception.Message)" }
        }
        $eigenarenPerSiteUrl[(ConvertTo-GenormaliseerdeUrl $siteUrl)] = $lijst
    }
    Write-Progress -Activity 'Site Owners ophalen' -Completed
    if ($ownerFouten -gt 0) {
        Write-Log "$ownerFouten sites konden niet worden benaderd. Eerste fout: $eersteOwnerFout" -Niveau WAARSCHUWING
        Write-Log 'Bij app-only vereist het uitlezen van SharePoint-groepen de permissie Sites.FullControl.All.' -Niveau WAARSCHUWING
    }
    if ($zonderOwnersGroep -gt 0) {
        Write-Log "$zonderOwnersGroep sites hebben geen Owners-groep; daar wordt teruggevallen op de primaire eigenaar." -Niveau WAARSCHUWING
    }
    if ($ownerFouten -eq 0 -and $zonderOwnersGroep -eq 0) {
        Write-Log 'Site Owners volledig opgehaald.' -Niveau OK
    }

    # Terug naar het admincenter voor de rest van het script.
    try { Connect-PnPDoelsite -Url $adminUrl } catch { }
}
elseif ($SkipOwnerLookup) {
    Write-Log '-SkipOwnerLookup actief: alleen de primaire eigenaar wordt gerapporteerd voor niet-groepssites.' -Niveau WAARSCHUWING
}
#endregion

#region -- Stap 5: rapportrijen samenstellen ------------------------------------
$vandaag       = Get-Date
$tussenRijen   = [System.Collections.Generic.List[object]]::new()
$eigenaarRijen = [System.Collections.Generic.List[object]]::new()

foreach ($s in $geselecteerd) {
    $siteUrl      = [string](Get-Veld $s 'Url')
    $siteNaam     = [string](Get-Veld $s 'Title')
    if ([string]::IsNullOrWhiteSpace($siteNaam)) { $siteNaam = ($siteUrl -split '/')[-1] }
    $sjabloonCode = ([string](Get-Veld $s 'Template')).ToUpperInvariant()
    $sjabloonNaam = if ($SjabloonNamen.ContainsKey($sjabloonCode)) { $SjabloonNamen[$sjabloonCode] } else { $sjabloonCode }
    $groepId      = [string](Get-Veld $s 'GroupId')
    $heeftGroep   = $groepId -and $groepId -ne [guid]::Empty.ToString()
    $genormUrl    = ConvertTo-GenormaliseerdeUrl $siteUrl

    # Eigenaren bepalen op basis van de bron.
    if ($heeftGroep -and $eigenarenPerGroep.ContainsKey($groepId)) {
        $eigenaren  = @($eigenarenPerGroep[$groepId])
        $eigenaarBron = 'M365-groep'
    }
    elseif ($eigenarenPerSiteUrl.ContainsKey($genormUrl)) {
        $eigenaren  = @($eigenarenPerSiteUrl[$genormUrl])
        $eigenaarBron = if ($IncludeSiteCollectionAdmins) { 'Site Owners + sitecollectiebeheer' } else { 'Site Owners-groep' }
    }
    else {
        $eigenaren  = @()
        $eigenaarBron = 'Primaire eigenaar'
    }

    # Fallback: als er niets is gevonden, alsnog de primaire eigenaar uit Get-PnPTenantSite.
    if ($eigenaren.Count -eq 0) {
        $primair = Get-Veld $s 'OwnerName'
        if (-not $primair) { $primair = Get-Veld $s 'Owner' }
        $primairMail = Get-Veld $s 'OwnerEmail'
        if ($primair) {
            $eigenaren = @([PSCustomObject]@{
                Naam  = [string]$primair
                Email = [string]$primairMail
                Type  = 'user'
                Bron  = 'Primaire eigenaar'
            })
            $eigenaarBron = 'Primaire eigenaar'
        }
    }

    # Activiteit: primair het gebruiksrapport, anders LastContentModifiedDate.
    $laatsteActiviteit = $null; $bestanden = $null; $opslagGb = $null; $activiteitBron = 'Onbekend'
    $graphSiteId = if ($genormUrl -and $siteIdPerUrl.ContainsKey($genormUrl)) { $siteIdPerUrl[$genormUrl] } else { $null }
    if ($graphSiteId) {
        $siteCollectionId = ($graphSiteId -split ',')[1]
        if ($siteCollectionId -and $activiteitPerSiteId.ContainsKey($siteCollectionId.ToLowerInvariant())) {
            $a = $activiteitPerSiteId[$siteCollectionId.ToLowerInvariant()]
            $laatsteActiviteit = $a.LastActivityDate
            $bestanden = $a.FileCount
            $opslagGb  = [math]::Round($a.StorageUsedBytes / 1GB, 2)
            $activiteitBron = 'Gebruiksrapport'
        }
    }
    if (-not $laatsteActiviteit) {
        $lcmd = Get-Veld $s 'LastContentModifiedDate'
        if ($lcmd) {
            try { $laatsteActiviteit = [datetime]$lcmd; $activiteitBron = 'LastContentModifiedDate' } catch { }
        }
    }
    if ($null -eq $opslagGb) {
        $storageMb = Get-Veld $s 'StorageUsageCurrent'
        if ($storageMb) { $opslagGb = [math]::Round(([double]$storageMb) / 1024, 2) }
    }

    $dagenInactief = if ($laatsteActiviteit) { [int]($vandaag - $laatsteActiviteit).TotalDays } else { $null }

    $teWeinigEigenaren = ($eigenaren.Count -lt $MinimumOwners)
    $kommaInNaam       = ($siteNaam -match ',')
    $langInactief      = ($null -eq $laatsteActiviteit) -or ($dagenInactief -ge $InactiveDaysThreshold)

    $redenen = @()
    if ($teWeinigEigenaren) { $redenen += "Te weinig eigenaren ($($eigenaren.Count) van min. $MinimumOwners)" }
    if ($kommaInNaam)       { $redenen += 'Komma in sitenaam' }
    if ($langInactief)      { $redenen += if ($null -eq $laatsteActiviteit) { 'Geen activiteitsdata (mogelijk nooit gebruikt)' } else { "Inactief sinds $($laatsteActiviteit.ToString('dd-MM-yyyy')) ($dagenInactief dagen)" } }

    $risicoScore = @($teWeinigEigenaren, $kommaInNaam, $langInactief).Where({ $_ }).Count
    $status = switch ($risicoScore) { 0 { 'OK' } 1 { 'Aandacht' } default { 'Attestatie vereist' } }

    $groepsMail = ''; $zichtbaarheid = ''; $aangemaakt = $null
    if ($heeftGroep -and $groepsDetails.ContainsKey($groepId)) {
        $d = $groepsDetails[$groepId]
        $groepsMail    = [string]$d.Mail
        $zichtbaarheid = [string]$d.Zichtbaarheid
        if ($d.CreatedDateTime) { $aangemaakt = ([datetime]$d.CreatedDateTime).ToLocalTime().ToString('dd-MM-yyyy') }
    }
    if (-not $zichtbaarheid) { $zichtbaarheid = [string](Get-Veld $s 'SharingCapability') }

    $tussenRijen.Add([PSCustomObject]@{
        Sitenaam          = $siteNaam
        SiteUrl           = $siteUrl
        SjabloonNaam      = $sjabloonNaam
        SjabloonCode      = $sjabloonCode
        EigenaarBron      = $eigenaarBron
        GroepsMail        = $groepsMail
        Zichtbaarheid     = $zichtbaarheid
        Aangemaakt        = $aangemaakt
        Eigenaren         = $eigenaren
        LaatsteActiviteit = $laatsteActiviteit
        ActiviteitBron    = $activiteitBron
        DagenInactief     = $dagenInactief
        Bestanden         = $bestanden
        OpslagGb          = $opslagGb
        TeWeinig          = $teWeinigEigenaren
        Komma             = $kommaInNaam
        Inactief          = $langInactief
        RisicoScore       = $risicoScore
        Status            = $status
        Redenen           = ($redenen -join ' | ')
    })

    foreach ($e in $eigenaren) {
        $eigenaarRijen.Add([PSCustomObject]@{
            'Sitenaam'        = $siteNaam
            'Site-URL'        = $siteUrl
            'Sjabloon'        = $sjabloonNaam
            'Eigenaar naam'   = $e.Naam
            'Eigenaar e-mail' = $e.Email
            'Objecttype'      = $e.Type
            'Bron'            = $e.Bron
        })
    }
}

# Bepaal het aantal eigenaarkolommen op basis van de werkelijke maximumbezetting.
# Er wordt NIET afgekapt: elke eigenaar krijgt gegarandeerd een eigen kolom.
$maxGevonden = 0
foreach ($r in $tussenRijen) { if ($r.Eigenaren.Count -gt $maxGevonden) { $maxGevonden = $r.Eigenaren.Count } }
$aantalEigenaarKolommen = [math]::Max(1, $maxGevonden)
Write-Log "Maximaal $maxGevonden eigenaren op één site; er worden $aantalEigenaarKolommen naam- en $aantalEigenaarKolommen e-mailkolommen aangemaakt." -Niveau OK
if ($aantalEigenaarKolommen -gt 100) {
    Write-Log "Let op: $aantalEigenaarKolommen eigenaarkolommen maakt het werkblad erg breed. Controleer of er sites zijn waar de Owners-groep als distributielijst wordt gebruikt." -Niveau WAARSCHUWING
}

# Definitieve rijen met platte eigenaarkolommen.
$rapportRijen = [System.Collections.Generic.List[object]]::new()
foreach ($r in $tussenRijen) {
    $rij = [ordered]@{
        'Sitenaam'          = $r.Sitenaam
        'Site-URL'          = $r.SiteUrl
        'Sjabloon'          = $r.SjabloonNaam
        'Eigenaarbron'      = $r.EigenaarBron
        'Groeps-e-mail'     = $r.GroepsMail
        'Zichtbaarheid'     = $r.Zichtbaarheid
        'Aangemaakt op'     = $r.Aangemaakt
        'Aantal eigenaren'  = $r.Eigenaren.Count
    }
    for ($i = 0; $i -lt $aantalEigenaarKolommen; $i++) {
        $rij["Eigenaar $($i + 1) (naam)"] = if ($i -lt $r.Eigenaren.Count) { $r.Eigenaren[$i].Naam } else { '' }
    }
    for ($i = 0; $i -lt $aantalEigenaarKolommen; $i++) {
        $rij["Eigenaar $($i + 1) (e-mail)"] = if ($i -lt $r.Eigenaren.Count) { $r.Eigenaren[$i].Email } else { '' }
    }
    $rij['Laatste activiteit']  = if ($r.LaatsteActiviteit) { $r.LaatsteActiviteit.ToString('dd-MM-yyyy') } else { 'Onbekend' }
    $rij['Activiteitsbron']     = $r.ActiviteitBron
    $rij['Dagen inactief']      = $r.DagenInactief
    $rij['Bestanden']           = $r.Bestanden
    $rij['Opslag (GB)']         = $r.OpslagGb
    $rij['Te weinig eigenaren'] = if ($r.TeWeinig) { 'Ja' } else { 'Nee' }
    $rij['Komma in naam']       = if ($r.Komma)    { 'Ja' } else { 'Nee' }
    $rij['Lang inactief']       = if ($r.Inactief) { 'Ja' } else { 'Nee' }
    $rij['Risicoscore']         = $r.RisicoScore
    $rij['Status']              = $r.Status
    $rij['Reden(en)']           = $r.Redenen
    # Sjablooncode bewust als laatste kolom: technisch nuttig om op te filteren, maar
    # visueel niet relevant genoeg om vooraan te staan.
    $rij['Sjablooncode']        = $r.SjabloonCode

    $rapportRijen.Add([PSCustomObject]$rij)
}
Write-Log "$($rapportRijen.Count) sites verwerkt voor het rapport." -Niveau OK
#endregion

#region -- Stap 6: Excel-rapport genereren --------------------------------------
Write-Log "Excel-rapport opbouwen: $OutputPath"
if (Test-BestandVergrendeld -Pad $OutputPath) {
    Write-Log 'Uitvoerbestand is vergrendeld (open in Excel). Sluit het bestand en start opnieuw.' -Niveau FOUT
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
if (Test-Path $OutputPath) {
    try { Remove-Item $OutputPath -Force -ErrorAction Stop }
    catch {
        Write-Log "Bestaand rapport kon niet worden verwijderd: $($_.Exception.Message)" -Niveau FOUT
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
}

$kleurPrimair  = '#1F3864'; $kleurAccent = '#2E75B6'; $kleurOK = '#70AD47'
$kleurAandacht = '#FFC000'; $kleurKritiek = '#C00000'; $kleurLicht = '#DEEBF7'

$attestatieRijen = @($rapportRijen | Where-Object { $_.Risicoscore -gt 0 } | Sort-Object -Property Risicoscore -Descending)
$excel = $null
$opgeslagen = $false
try {
    $excel = $rapportRijen | Sort-Object Sjabloon, Sitenaam | Export-Excel -Path $OutputPath -WorksheetName 'Sites & Eigenaren' `
        -TableName 'SitesEigenaren' -TableStyle Medium2 -AutoSize -FreezeTopRow -AutoFilter -PassThru
    $ws = $excel.Workbook.Worksheets['Sites & Eigenaren']
    $laatsteRij = $ws.Dimension.End.Row
    $ws.Cells.Style.Font.Name = 'Segoe UI'; $ws.Cells.Style.Font.Size = 9
    $ws.Row(1).Height = 22

    # Kolomposities dynamisch opzoeken: het aantal eigenaarkolommen varieert per run,
    # dus harde letterverwijzingen zouden naar de verkeerde kolom wijzen.
    $kolStatus   = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Status'
    $kolAantal   = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Aantal eigenaren'
    $kolDagen    = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Dagen inactief'
    $kolTeWeinig = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Te weinig eigenaren'
    $kolKomma    = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Komma in naam'
    $kolInactief = Get-KolomLetterVoorKop -Werkblad $ws -Kop 'Lang inactief'

    foreach ($kop in 'Sitenaam', 'Site-URL', 'Reden(en)') {
        $letter = Get-KolomLetterVoorKop -Werkblad $ws -Kop $kop
        if ($letter) { $ws.Column($ws.Cells["${letter}1"].Start.Column).Width = 45 }
    }

    if ($kolStatus) {
        Add-ConditionalFormatting -Worksheet $ws -Range "$kolStatus`2:$kolStatus$laatsteRij" -RuleType ContainsText -ConditionValue 'Attestatie vereist' -BackgroundColor $kleurKritiek -ForegroundColor White -Bold
        Add-ConditionalFormatting -Worksheet $ws -Range "$kolStatus`2:$kolStatus$laatsteRij" -RuleType ContainsText -ConditionValue 'Aandacht' -BackgroundColor $kleurAandacht
        Add-ConditionalFormatting -Worksheet $ws -Range "$kolStatus`2:$kolStatus$laatsteRij" -RuleType ContainsText -ConditionValue 'OK' -BackgroundColor $kleurOK -ForegroundColor White
    }
    foreach ($letter in @($kolTeWeinig, $kolKomma, $kolInactief)) {
        if ($letter) {
            Add-ConditionalFormatting -Worksheet $ws -Range "$letter`2:$letter$laatsteRij" -RuleType ContainsText -ConditionValue 'Ja' -BackgroundColor '#F8CBAD' -Bold
        }
    }
    if ($kolDagen) {
        Add-ConditionalFormatting -Worksheet $ws -Range "$kolDagen`2:$kolDagen$laatsteRij" -DataBarColor ([System.Drawing.ColorTranslator]::FromHtml($kleurAccent))
    }
    if ($kolAantal) {
        # Drempels expliciet als getal zetten; de standaard 0/33/67 zou elk realistisch
        # aantal eigenaren onder de laagste drempel laten vallen (alles rood).
        $iconEigenaren = Add-ConditionalFormatting -Worksheet $ws -Range "$kolAantal`2:$kolAantal$laatsteRij" -ThreeIconsSet Symbols -PassThru
        $iconEigenaren.Icon1.Type = 'Num'; $iconEigenaren.Icon1.Value = 0
        $iconEigenaren.Icon2.Type = 'Num'; $iconEigenaren.Icon2.Value = $MinimumOwners
        $iconEigenaren.Icon3.Type = 'Num'; $iconEigenaren.Icon3.Value = ($MinimumOwners + 1)
    }

    if ($attestatieRijen.Count -gt 0) {
        $attKolommen = @('Sitenaam', 'Site-URL', 'Sjabloon', 'Eigenaarbron', 'Aantal eigenaren')
        for ($i = 1; $i -le $aantalEigenaarKolommen; $i++) { $attKolommen += "Eigenaar $i (naam)" }
        $attKolommen += @('Laatste activiteit', 'Risicoscore', 'Status', 'Reden(en)')

        $excel = $attestatieRijen | Select-Object $attKolommen |
            Export-Excel -ExcelPackage $excel -WorksheetName 'Attestatie' -TableName 'Attestatie' `
                -TableStyle Medium3 -AutoSize -FreezeTopRow -AutoFilter -PassThru
        $wsAtt = $excel.Workbook.Worksheets['Attestatie']
        $wsAtt.Cells.Style.Font.Name = 'Segoe UI'; $wsAtt.Cells.Style.Font.Size = 9
        foreach ($kop in 'Sitenaam', 'Site-URL', 'Reden(en)') {
            $letter = Get-KolomLetterVoorKop -Werkblad $wsAtt -Kop $kop
            if ($letter) { $wsAtt.Column($wsAtt.Cells["${letter}1"].Start.Column).Width = 45 }
        }
        $kolRisico = Get-KolomLetterVoorKop -Werkblad $wsAtt -Kop 'Risicoscore'
        if ($kolRisico) {
            # Reverse maakt score 3 rood en score 1 groen.
            $iconRisico = Add-ConditionalFormatting -Worksheet $wsAtt -Range "$kolRisico`2:$kolRisico$($wsAtt.Dimension.End.Row)" -ThreeIconsSet TrafficLights1 -PassThru
            $iconRisico.Reverse = $true
            $iconRisico.Icon1.Type = 'Num'; $iconRisico.Icon1.Value = 0
            $iconRisico.Icon2.Type = 'Num'; $iconRisico.Icon2.Value = 2
            $iconRisico.Icon3.Type = 'Num'; $iconRisico.Icon3.Value = 3
        }
    }

    if ($eigenaarRijen.Count -gt 0) {
        $excel = $eigenaarRijen | Sort-Object Sitenaam, 'Eigenaar naam' |
            Export-Excel -ExcelPackage $excel -WorksheetName 'Eigenaren (plat)' -TableName 'EigenarenPlat' `
                -TableStyle Light9 -AutoSize -FreezeTopRow -AutoFilter -PassThru
        $excel.Workbook.Worksheets['Eigenaren (plat)'].Cells.Style.Font.Name = 'Segoe UI'
    }

    # ---- Dashboard
    if ($excel.Workbook.Worksheets['Dashboard']) { $excel.Workbook.Worksheets.Delete('Dashboard') }
    $wsDash = $excel.Workbook.Worksheets.Add('Dashboard')
    $excel.Workbook.Worksheets.MoveToStart('Dashboard')
    $wsDash.View.ShowGridLines = $false
    $wsDash.Cells.Style.Font.Name = 'Segoe UI'

    $wsDash.Cells['B2:M3'].Merge = $true
    $titel = $wsDash.Cells['B2']
    $titel.Value = 'SharePoint Online - Site-attestatierapport'
    $titel.Style.Font.Size = 20; $titel.Style.Font.Bold = $true; $titel.Style.Font.Color.SetColor('White')
    $wsDash.Cells['B2:M3'].Style.Fill.PatternType = 'Solid'
    $wsDash.Cells['B2:M3'].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurPrimair))
    $wsDash.Cells['B2:M3'].Style.VerticalAlignment = 'Center'; $titel.Style.Indent = 1

    $wsDash.Cells['B4:M4'].Merge = $true
    $sub = $wsDash.Cells['B4']
    $sub.Value = "Gegenereerd op $(Get-Date -Format 'dd-MM-yyyy HH:mm') | Tenant: $($context.TenantId) | Criteria: min. $MinimumOwners eigenaren, inactief >= $InactiveDaysThreshold dagen | $($rapportRijen.Count) sites over $($sjabloonTellers.Keys.Count) sjablonen"
    $sub.Style.Font.Size = 9; $sub.Style.Font.Italic = $true; $sub.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#595959'))

    $kpis = @(
        @{ Label = 'SITES TOTAAL';        Waarde = $rapportRijen.Count;                                                   Kleur = $kleurAccent }
        @{ Label = 'ATTESTATIE VEREIST';  Waarde = @($rapportRijen | Where-Object Status -eq 'Attestatie vereist').Count; Kleur = $kleurKritiek }
        @{ Label = 'TE WEINIG EIGENAREN'; Waarde = @($rapportRijen | Where-Object 'Te weinig eigenaren' -eq 'Ja').Count;  Kleur = $kleurAandacht }
        @{ Label = 'KOMMA IN NAAM';       Waarde = @($rapportRijen | Where-Object 'Komma in naam' -eq 'Ja').Count;        Kleur = $kleurAandacht }
        @{ Label = 'LANG INACTIEF';       Waarde = @($rapportRijen | Where-Object 'Lang inactief' -eq 'Ja').Count;        Kleur = $kleurAandacht }
        @{ Label = 'SITES OK';            Waarde = @($rapportRijen | Where-Object Status -eq 'OK').Count;                 Kleur = $kleurOK }
    )
    $kolomStart = 2
    foreach ($kpi in $kpis) {
        $kolEind = $kolomStart + 1
        $bereikWaarde = $wsDash.Cells[6, $kolomStart, 8, $kolEind]
        $bereikLabel  = $wsDash.Cells[9, $kolomStart, 9, $kolEind]
        $bereikWaarde.Merge = $true; $bereikLabel.Merge = $true
        $bereikWaarde.Value = $kpi.Waarde
        $bereikWaarde.Style.Font.Size = 28; $bereikWaarde.Style.Font.Bold = $true; $bereikWaarde.Style.Font.Color.SetColor('White')
        $bereikWaarde.Style.HorizontalAlignment = 'Center'; $bereikWaarde.Style.VerticalAlignment = 'Center'
        $bereikWaarde.Style.Fill.PatternType = 'Solid'
        $bereikWaarde.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kpi.Kleur))
        $bereikLabel.Value = $kpi.Label
        $bereikLabel.Style.Font.Size = 8; $bereikLabel.Style.Font.Bold = $true
        $bereikLabel.Style.HorizontalAlignment = 'Center'
        $bereikLabel.Style.Fill.PatternType = 'Solid'
        $bereikLabel.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurLicht))
        $kolomStart = $kolEind + 1
    }

    # Hulpkolommen voor de grafieken. Bewust NIET verbergen: Excel plot standaard geen cellen
    # uit verborgen kolommen, waardoor de grafieken als leeg kader zouden verschijnen.
    # Gebruik hier geen .PlotVisibleOnly; die property bestaat niet op het ExcelChart-object.
    $wsDash.Cells['P2'].Value = 'Categorie';           $wsDash.Cells['Q2'].Value = 'Aantal'
    $wsDash.Cells['P3'].Value = 'Te weinig eigenaren'; $wsDash.Cells['Q3'].Value = $kpis[2].Waarde
    $wsDash.Cells['P4'].Value = 'Komma in naam';       $wsDash.Cells['Q4'].Value = $kpis[3].Waarde
    $wsDash.Cells['P5'].Value = 'Lang inactief';       $wsDash.Cells['Q5'].Value = $kpis[4].Waarde
    $wsDash.Cells['P8'].Value = 'Status';              $wsDash.Cells['Q8'].Value = 'Aantal'
    $wsDash.Cells['P9'].Value = 'OK';                  $wsDash.Cells['Q9'].Value  = $kpis[5].Waarde
    $wsDash.Cells['P10'].Value = 'Aandacht';           $wsDash.Cells['Q10'].Value = @($rapportRijen | Where-Object Status -eq 'Aandacht').Count
    $wsDash.Cells['P11'].Value = 'Attestatie vereist'; $wsDash.Cells['Q11'].Value = $kpis[1].Waarde
    $wsDash.Cells['P2:Q11'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsDash.Column(16).Width = 2; $wsDash.Column(17).Width = 2

    $chart1 = $wsDash.Drawings.AddChart('RisicoCategorieen', [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
    $chart1.Title.Text = 'Bevindingen per risicocategorie'
    $chart1.Title.Font.Size = 11; $chart1.Title.Font.Bold = $true
    $chart1.SetPosition(10, 5, 1, 0); $chart1.SetSize(430, 250)
    $serie1 = $chart1.Series.Add($wsDash.Cells['Q3:Q5'], $wsDash.Cells['P3:P5'])
    $serie1.Header = 'Aantal sites'; $chart1.Legend.Remove()
    $chart1.DataLabel.ShowValue = $true

    $chart2 = $wsDash.Drawings.AddChart('StatusVerdeling', [OfficeOpenXml.Drawing.Chart.eChartType]::Doughnut)
    $chart2.Title.Text = 'Statusverdeling van alle sites'
    $chart2.Title.Font.Size = 11; $chart2.Title.Font.Bold = $true
    $chart2.SetPosition(10, 5, 8, 0); $chart2.SetSize(430, 250)
    $chart2.Series.Add($wsDash.Cells['Q9:Q11'], $wsDash.Cells['P9:P11']) | Out-Null
    $chart2.Legend.Position = [OfficeOpenXml.Drawing.Chart.eLegendPosition]::Bottom
    $chart2.DataLabel.ShowValue = $true

    # ---- Verdeling per sjabloon
    $sjabloonRij = 28
    $wsDash.Cells[$sjabloonRij, 2, $sjabloonRij, 13].Merge = $true
    $kopS = $wsDash.Cells[$sjabloonRij, 2]
    $kopS.Value = 'Sites per sjabloon'
    $kopS.Style.Font.Bold = $true; $kopS.Style.Font.Size = 12
    $kopS.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurPrimair))

    $sHeaders = 'Sjabloon', 'Sjablooncode', 'Aantal sites', 'Attestatie vereist'
    $sKolMap  = 2, 6, 9, 11
    for ($i = 0; $i -lt $sHeaders.Count; $i++) {
        $cel = $wsDash.Cells[($sjabloonRij + 1), $sKolMap[$i]]
        $cel.Value = $sHeaders[$i]; $cel.Style.Font.Bold = $true; $cel.Style.Font.Color.SetColor('White')
        $cel.Style.Fill.PatternType = 'Solid'
        $cel.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurPrimair))
    }
    $rijNr = $sjabloonRij + 2
    foreach ($groepje in ($rapportRijen | Group-Object Sjablooncode | Sort-Object Count -Descending)) {
        $code = $groepje.Name
        $wsDash.Cells[$rijNr, 2].Value  = if ($SjabloonNamen.ContainsKey($code)) { $SjabloonNamen[$code] } else { $code }
        $wsDash.Cells[$rijNr, 6].Value  = $code
        $wsDash.Cells[$rijNr, 9].Value  = $groepje.Count
        $wsDash.Cells[$rijNr, 11].Value = @($groepje.Group | Where-Object Status -eq 'Attestatie vereist').Count
        if ($rijNr % 2 -eq 0) {
            $wsDash.Cells[$rijNr, 2, $rijNr, 13].Style.Fill.PatternType = 'Solid'
            $wsDash.Cells[$rijNr, 2, $rijNr, 13].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurLicht))
        }
        $rijNr++
    }

    # ---- Top 10 hoogste risicoscore
    $topRij = $rijNr + 2
    $wsDash.Cells[$topRij, 2, $topRij, 13].Merge = $true
    $kop = $wsDash.Cells[$topRij, 2]
    $kop.Value = 'Top 10 sites met hoogste risicoscore'
    $kop.Style.Font.Bold = $true; $kop.Style.Font.Size = 12
    $kop.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurPrimair))
    $headers = 'Sitenaam', 'Sjabloon', 'Eigenaren', 'Laatste activiteit', 'Risicoscore', 'Reden(en)'
    $kolMap  = 2, 5, 7, 8, 10, 11
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cel = $wsDash.Cells[($topRij + 1), $kolMap[$i]]
        $cel.Value = $headers[$i]; $cel.Style.Font.Bold = $true; $cel.Style.Font.Color.SetColor('White')
        $cel.Style.Fill.PatternType = 'Solid'
        $cel.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurPrimair))
    }
    $rijNr = $topRij + 2
    foreach ($topItem in ($attestatieRijen | Select-Object -First 10)) {
        $wsDash.Cells[$rijNr, 2].Value  = $topItem.Sitenaam
        $wsDash.Cells[$rijNr, 5].Value  = $topItem.Sjabloon
        $wsDash.Cells[$rijNr, 7].Value  = $topItem.'Aantal eigenaren'
        $wsDash.Cells[$rijNr, 8].Value  = $topItem.'Laatste activiteit'
        $wsDash.Cells[$rijNr, 10].Value = $topItem.Risicoscore
        $wsDash.Cells[$rijNr, 11].Value = $topItem.'Reden(en)'
        if ($rijNr % 2 -eq 0) {
            $wsDash.Cells[$rijNr, 2, $rijNr, 13].Style.Fill.PatternType = 'Solid'
            $wsDash.Cells[$rijNr, 2, $rijNr, 13].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($kleurLicht))
        }
        $rijNr++
    }
    foreach ($kolom in 2..13) { if ($wsDash.Column($kolom).Width -lt 12) { $wsDash.Column($kolom).Width = 12 } }
    $wsDash.Column(2).Width = 40; $wsDash.Column(11).Width = 55

    Close-ExcelPackage $excel -ErrorAction Stop
    $excel = $null
    $opgeslagen = $true
}
catch {
    Write-Log "Fout bij het opbouwen/opslaan van het Excel-rapport: $($_.Exception.Message)" -Niveau FOUT
    if ($excel) { try { Close-ExcelPackage $excel -NoSave } catch { } }
}

$samenvatting = "$($rapportRijen.Count) sites | $($attestatieRijen.Count) met bevindingen | $(@($rapportRijen | Where-Object Status -eq 'Attestatie vereist').Count) attestatie vereist"
if ($opgeslagen) {
    Write-Log "Rapport opgeslagen: $((Resolve-Path $OutputPath).Path)" -Niveau OK
    Write-Log "Samenvatting: $samenvatting" -Niveau OK
    if ($MailTo) {
        try {
            Send-RapportMail -Pad $OutputPath -Van $MailFrom -Naar $MailTo -Samenvatting "Bijgevoegd het maandelijkse SharePoint site-attestatierapport.`r`n`r`n$samenvatting"
            Write-Log "Rapport verzonden naar $MailTo." -Niveau OK
        }
        catch { Write-Log "Verzenden mislukt: $($_.Exception.Message)" -Niveau FOUT }
    }
    if ($OpenAfterExport) { Invoke-Item $OutputPath }
}
else {
    Write-Log 'Er is GEEN rapport opgeslagen. Zie de fout hierboven.' -Niveau FOUT
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Log 'Klaar.' -Niveau OK
#endregion
