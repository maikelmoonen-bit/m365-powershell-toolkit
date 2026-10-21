<#
.SYNOPSIS
    Diagnostiek voor Outlook-fout 0x8004010F (OST/mailbox bereikbaarheidsfout).

.DESCRIPTION
    Verzamelt de volgende gegevens voor een opgegeven gebruiker:
      - Shared mailboxen met FullAccess + automapping-status
      - Mailboxgrootte en itemtelling (primair + archief)
      - Quotuminformatie en soft-deleted items
      - Algemene mailboxconfiguratie (taal, tijdzone, OWA-toegang)
    Genereert een HTML-rapport met remediatie-adviezen.

    AutoMapping-detectie:
      In Exchange Online heeft Get-MailboxPermission GEEN .AutoMapping-property.
      AutoMapping wordt bepaald via de delegate-lijst op de mailbox zelf:
      Get-Mailbox <shared> | Select GrantSendOnBehalfTo is niet relevant;
      de echte AutoMapping-staat zit in het interne msExchDelegateListLink-attribuut,
      dat via EXO niet direct leesbaar is.

      Betrouwbare detectiemethode: als een gebruiker via Add-MailboxPermission met
      -AutoMapping $false is toegevoegd, staat de user NIET in de interne delegate-lijst
      van de mailbox. We benaderen dit door voor elke gevonden shared mailbox opnieuw
      Get-MailboxPermission op te vragen en de hele permissie-string te inspecteren.
      Wanneer AutoMapping niet bepaald kan worden, rapporteren we 'Onbekend' en geven
      we de aanbeveling om het te controleren.

      Snellere aanpak voor grote omgevingen: Get-EXOMailboxPermission (REST-gebaseerd)
      is aanzienlijk sneller dan Get-MailboxPermission (RPS) bij bulk-queries.

.PARAMETER UserPrincipalName
    UPN van de gebruiker, bijv. jan.de.vries@contoso.com

.PARAMETER ReportPath
    Pad voor het HTML-rapport. Standaard: bureaublad van de huidige gebruiker.

.EXAMPLE
    .\Diagnose-0x8004010F.ps1 -UserPrincipalName jan.de.vries@contoso.com

.EXAMPLE
    .\Diagnose-0x8004010F.ps1 -UserPrincipalName jan.de.vries@contoso.com -ReportPath C:\Reports
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = [Environment]::GetFolderPath('Desktop')
)

#region Hulpfuncties

function Write-Step {
    param([string]$Message)
    Write-Host "`n[ $Message ]" -ForegroundColor Cyan
}

function Write-OK   { param([string]$Msg) Write-Host "  OK  $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg) Write-Host "  !!  $Msg" -ForegroundColor Yellow }
function Write-Fail { param([string]$Msg) Write-Host "  XX  $Msg" -ForegroundColor Red }

function Get-BytesFromString {
    param([string]$SizeString)
    if ($SizeString -match '\(([0-9,]+) bytes\)') {
        return [long]($Matches[1] -replace ',', '')
    }
    return 0
}

#endregion

#region Verbinding met Exchange Online

Write-Step "Verbinding controleren met Exchange Online"

try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-OK "Al verbonden met Exchange Online."
} catch {
    Write-Warn "Geen actieve Exchange Online-sessie. Verbinding wordt gemaakt..."
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        Write-OK "Verbonden met Exchange Online."
    } catch {
        Write-Fail "Kan niet verbinden met Exchange Online: $_"
        exit 1
    }
}

# Controleer of Get-EXOMailboxPermission beschikbaar is (moderne REST-cmdlet, veel sneller)
$UseEXO = $null -ne (Get-Command Get-EXOMailboxPermission -ErrorAction SilentlyContinue)
if ($UseEXO) {
    Write-OK "Get-EXOMailboxPermission beschikbaar — wordt gebruikt voor snellere queries."
} else {
    Write-Warn "Get-EXOMailboxPermission niet beschikbaar. Terugvallen op Get-MailboxPermission (langzamer)."
}

#endregion

#region Gebruikerscontrole

Write-Step "Gebruiker ophalen: $UserPrincipalName"

try {
    $Mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    Write-OK "Mailbox gevonden: $($Mailbox.DisplayName) [$($Mailbox.RecipientTypeDetails)]"
} catch {
    Write-Fail "Mailbox niet gevonden voor '$UserPrincipalName': $_"
    exit 1
}

# Haal ook de ExchangeGuid en SID op van de gebruiker — nodig voor delegate-check
$UserMailboxObj = Get-EXOMailbox -Identity $UserPrincipalName -Properties ExchangeGuid -ErrorAction SilentlyContinue
$UserSID        = $null
try {
    # In EXO is de SID niet altijd rechtstreeks beschikbaar; we gebruiken de UPN als fallback
    $UserSID = (Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop).Guid.ToString()
} catch { }

#endregion

#region STAP 1 – Shared mailboxen met FullAccess en automapping-status

Write-Step "STAP 1 – Shared mailboxen (FullAccess) en automapping"

$SharedMailboxen = [System.Collections.Generic.List[PSObject]]::new()

try {
    # -----------------------------------------------------------------------
    # Haal ALLE shared mailboxen op in bulk via Get-EXOMailbox (sneller dan
    # Get-Mailbox voor grote omgevingen). Filter op RecipientTypeDetails.
    # -----------------------------------------------------------------------
    Write-Host "  Shared mailboxen ophalen uit tenant..." -ForegroundColor Gray

    $SharedMBList = Get-EXOMailbox -ResultSize Unlimited `
        -Filter "RecipientTypeDetails -eq 'SharedMailbox'" `
        -Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, GrantSendOnBehalfTo `
        -ErrorAction SilentlyContinue

    $total = ($SharedMBList | Measure-Object).Count
    Write-Host "  $total shared mailboxen gevonden. Rechten controleren..." -ForegroundColor Gray

    # -----------------------------------------------------------------------
    # AutoMapping-detectie in Exchange Online:
    #
    # Get-MailboxPermission en Get-EXOMailboxPermission hebben GEEN .AutoMapping
    # property — die bestaat alleen in on-premises Exchange.
    #
    # In EXO wordt AutoMapping intern bepaald via de 'delegate list' die Exchange
    # bijhoudt op de mailbox. Er zijn twee methoden om dit te benaderen:
    #
    # Methode A (Get-EXOMailbox .GrantSendOnBehalfTo): dit is NIET AutoMapping —
    #   dit is SendOnBehalf, een ander recht.
    #
    # Methode B (betrouwbaar): Controleer de interne attribute via MS Graph:
    #   GET /users/{mailboxId}/mailboxSettings — niet beschikbaar in EXO-module.
    #
    # Methode C (meest praktisch voor EXO-module):
    #   Voer Get-MailboxPermission uit met -IncludeSoftDeletedPermissions.
    #   De output bevat geen AutoMapping-boolean, maar via de richting van de
    #   toewijzing kunnen we redeneren:
    #   - Als de gebruiker staat in de 'msExchDelegateListLink' van de mailbox
    #     => AutoMapping is AAN (standaard gedrag bij Add-MailboxPermission)
    #   - Als uitdrukkelijk -AutoMapping $false is gebruikt bij het toevoegen,
    #     staat de gebruiker NIET in die lijst.
    #
    #   We benaderen dit door voor elke shared mailbox te controleren of de user
    #   voorkomt in de output van Get-MailboxPermission zonder SoftDelete-filter.
    #   Zonder directe boolean rapporteren we 'Waarschijnlijk Aan' als aanname,
    #   want dat is de Exchange-standaard tenzij anders ingesteld.
    #
    #   Voor organisaties met Exchange Hybrid of AzureAD Connect kan de waarde
    #   wel beschikbaar zijn via Get-ADUser msExchDelegateListLink.
    # -----------------------------------------------------------------------

    $counter = 0

    foreach ($SMB in $SharedMBList) {
        $counter++
        if ($counter % 50 -eq 0) {
            Write-Host "  Voortgang: $counter / $total" -ForegroundColor Gray
        }

        # Gebruik REST-gebaseerde cmdlet als beschikbaar, anders klassiek
        if ($UseEXO) {
            $Perm = Get-EXOMailboxPermission -Identity $SMB.Identity -User $UserPrincipalName `
                -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'FullAccess' }
        } else {
            $Perm = Get-MailboxPermission -Identity $SMB.Identity -User $UserPrincipalName `
                -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'FullAccess' }
        }

        if ($Perm) {
            # ----------------------------------------------------------------
            # AutoMapping bepalen:
            # We proberen via Get-MailboxPermission (niet EXO) een uitgebreide
            # output te krijgen. In sommige EXO-tenants toont de klassieke cmdlet
            # wel een AutoMapping-veld (afhankelijk van module-versie).
            # We controleren dit dynamisch.
            # ----------------------------------------------------------------
            $AutoMappingStatus = 'Onbekend'
            $AutoMappingNote   = 'Niet bepaalbaar via EXO-module; vermoedelijk Aan (standaard)'

            # Probeer AutoMapping te lezen via klassieke cmdlet (heeft soms meer properties)
            $PermKlassiek = Get-MailboxPermission -Identity $SMB.Identity -User $UserPrincipalName `
                -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'FullAccess' } |
                Select-Object -First 1

            if ($PermKlassiek -and ($PermKlassiek.PSObject.Properties.Name -contains 'AutoMapping')) {
                # Property bestaat in deze module-versie
                $AutoMappingStatus = $PermKlassiek.AutoMapping
                $AutoMappingNote   = ''
            } else {
                # Fallback: controleer of de UPN voorkomt in de GrantSendOnBehalfTo
                # als proxy voor 'is er expliciete automapping-configuratie?'
                # Dit is geen directe meting maar een indicator.
                #
                # Betrouwbaardere fallback: probeer de tenant-brede delegate-lijst
                # via Get-Mailbox -Identity <shared> en kijk naar de interne
                # msExchDelegateListLink via het Guid-attribuut.
                # Dit vereist on-premises of Hybrid; in cloud-only EXO niet beschikbaar.
                #
                # Conclusie: we markeren als 'Vermoedelijk Aan' — standaard Exchange-gedrag.
                $AutoMappingStatus = 'Vermoedelijk Aan'
                $AutoMappingNote   = 'AutoMapping-property niet beschikbaar in deze EXO-moduleversie. Standaard staat AutoMapping AAN tenzij bij toewijzing -AutoMapping $false is gebruikt.'
            }

            $SharedMailboxen.Add([PSCustomObject]@{
                DisplayName       = $SMB.DisplayName
                EmailAddress      = $SMB.PrimarySmtpAddress
                Identity          = $SMB.Identity
                AutoMapping       = $AutoMappingStatus
                AutoMappingNote   = $AutoMappingNote
                RecipientType     = $SMB.RecipientTypeDetails
            })
        }
    }

    Write-OK "Gevonden: $($SharedMailboxen.Count) shared mailbox(en) met FullAccess."

    # Statistieken tonen in console
    $AMOnbekend = ($SharedMailboxen | Where-Object { $_.AutoMapping -notlike '*Nee*' -and $_.AutoMapping -ne $false } | Measure-Object).Count
    if ($SharedMailboxen.Count -ge 15) {
        Write-Warn "Gebruiker heeft toegang tot $($SharedMailboxen.Count) shared mailboxen."
        Write-Warn "Als AutoMapping voor de meeste AAN staat, is dit een waarschijnlijke oorzaak van 0x8004010F."
    }

    if (($SharedMailboxen | Where-Object { $_.AutoMapping -eq 'Onbekend' -or $_.AutoMapping -eq 'Vermoedelijk Aan' } | Measure-Object).Count -gt 0) {
        Write-Warn "AutoMapping kon niet exact worden bepaald via de EXO-module (zie rapport voor uitleg)."
        Write-Host "  Gebruik onderstaand commando om AutoMapping per mailbox handmatig te verifiëren:" -ForegroundColor Gray
        Write-Host "  Get-MailboxPermission -Identity <mailbox> -User $UserPrincipalName | Select AccessRights, AutoMapping" -ForegroundColor Gray
    }

} catch {
    Write-Warn "Fout bij ophalen shared mailbox-rechten: $_"
}

#endregion

#region STAP 2 – Mailboxgrootte en itemtelling

Write-Step "STAP 2 – Mailboxgrootte en itemtelling"

$MailboxStats = $null
$ArchiveStats = $null
$PrimaryBytes = 0
$ArchiveBytes = 0

try {
    $MailboxStats = Get-MailboxStatistics -Identity $UserPrincipalName -ErrorAction Stop
    $PrimaryBytes = Get-BytesFromString $MailboxStats.TotalItemSize.ToString()

    Write-OK "Primaire mailbox:"
    Write-Host "    Grootte    : $($MailboxStats.TotalItemSize)" -ForegroundColor White
    Write-Host "    Items      : $($MailboxStats.ItemCount)" -ForegroundColor White
    Write-Host "    Verwijderd : $($MailboxStats.DeletedItemCount) items ($($MailboxStats.TotalDeletedItemSize))" -ForegroundColor White
    Write-Host "    Laatste log: $($MailboxStats.LastLogonTime)" -ForegroundColor White

    if ($PrimaryBytes -ge 50GB) {
        Write-Warn "Mailbox is groter dan 50 GB. Grote mailboxen verhogen de kans op OST-corruptie."
    } elseif ($PrimaryBytes -ge 25GB) {
        Write-Warn "Mailbox is groter dan 25 GB. Houd dit in de gaten."
    }
} catch {
    Write-Warn "Kan mailboxstatistieken niet ophalen: $_"
}

# Archief
if ($Mailbox.ArchiveStatus -ne 'None') {
    try {
        $ArchiveStats = Get-MailboxStatistics -Identity $UserPrincipalName -Archive -ErrorAction Stop
        $ArchiveBytes = Get-BytesFromString $ArchiveStats.TotalItemSize.ToString()

        Write-OK "Archiefmailbox:"
        Write-Host "    Grootte    : $($ArchiveStats.TotalItemSize)" -ForegroundColor White
        Write-Host "    Items      : $($ArchiveStats.ItemCount)" -ForegroundColor White
    } catch {
        Write-Warn "Kan archiefstatistieken niet ophalen: $_"
    }
} else {
    Write-Host "  Geen archiefmailbox geconfigureerd." -ForegroundColor Gray
}

#endregion

#region STAP 3 – Quotum en soft-deleted items

Write-Step "STAP 3 – Quotuminformatie en soft-deleted items"

$QuotumInfo = $null

try {
    $QuotumInfo = Get-Mailbox -Identity $UserPrincipalName |
        Select-Object IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota,
                      UseDatabaseQuotaDefaults, RecoverableItemsQuota, RecoverableItemsWarningQuota

    Write-OK "Quotuminstellingen:"
    Write-Host "    Waarschuwing verzenden : $($QuotumInfo.IssueWarningQuota)" -ForegroundColor White
    Write-Host "    Blokkeer verzenden     : $($QuotumInfo.ProhibitSendQuota)" -ForegroundColor White
    Write-Host "    Blokkeer alles         : $($QuotumInfo.ProhibitSendReceiveQuota)" -ForegroundColor White
    Write-Host "    Gebruik DB-standaard   : $($QuotumInfo.UseDatabaseQuotaDefaults)" -ForegroundColor White

    if ($MailboxStats -and $QuotumInfo.ProhibitSendReceiveQuota -ne 'Unlimited') {
        $LimitBytes = Get-BytesFromString $QuotumInfo.ProhibitSendReceiveQuota.ToString()
        if ($LimitBytes -gt 0) {
            $PctUsed = [math]::Round(($PrimaryBytes / $LimitBytes) * 100, 1)
            if ($PctUsed -ge 90) {
                Write-Warn "Mailbox is $PctUsed% vol (boven 90%). Dit kan sync-fouten veroorzaken."
            } elseif ($PctUsed -ge 75) {
                Write-Warn "Mailbox is $PctUsed% vol."
            } else {
                Write-OK "Mailbox is $PctUsed% vol — geen quotumprobleem."
            }
        }
    }
} catch {
    Write-Warn "Kan quotuminformatie niet ophalen: $_"
}

# Recoverable Items (soft-deleted)
try {
    $RecStats = Get-MailboxFolderStatistics -Identity $UserPrincipalName `
        -FolderScope RecoverableItems -ErrorAction Stop

    $RecoverableTotal = ($RecStats | Measure-Object -Property ItemsInFolder -Sum).Sum

    Write-OK "Recoverable Items (soft-deleted): $RecoverableTotal items totaal"
    foreach ($Folder in $RecStats) {
        if ($Folder.ItemsInFolder -gt 0) {
            Write-Host "    $($Folder.Name.PadRight(30)) : $($Folder.ItemsInFolder) items" -ForegroundColor Gray
        }
    }
} catch {
    Write-Warn "Kan Recoverable Items niet ophalen: $_"
}

#endregion

#region STAP 4 – Aanvullende configuratiegegevens

Write-Step "STAP 4 – Aanvullende mailboxconfiguratie"

$RegionalConfig = $null
$CASConfig      = $null

try {
    $RegionalConfig = Get-MailboxRegionalConfiguration -Identity $UserPrincipalName -ErrorAction Stop
    Write-OK "Regionale instellingen:"
    Write-Host "    Tijdzone    : $($RegionalConfig.TimeZone)" -ForegroundColor White
    Write-Host "    Taal        : $($RegionalConfig.Language)" -ForegroundColor White
    Write-Host "    Datumnotatie: $($RegionalConfig.DateFormat)" -ForegroundColor White
} catch {
    Write-Warn "Kan regionale configuratie niet ophalen: $_"
}

try {
    $CASConfig = Get-CASMailbox -Identity $UserPrincipalName -ErrorAction Stop
    Write-OK "Client-toegang (CAS):"
    Write-Host "    MAPI (Outlook) : $($CASConfig.MAPIEnabled)" -ForegroundColor White
    Write-Host "    OWA            : $($CASConfig.OWAEnabled)" -ForegroundColor White
    Write-Host "    ActiveSync     : $($CASConfig.ActiveSyncEnabled)" -ForegroundColor White

    if (-not $CASConfig.MAPIEnabled) {
        Write-Fail "MAPI is UITGESCHAKELD — directe oorzaak van 0x8004010F!"
    }
} catch {
    Write-Warn "Kan CAS-mailboxconfiguratie niet ophalen: $_"
}

try {
    $AutoReply = Get-MailboxAutoReplyConfiguration -Identity $UserPrincipalName -ErrorAction Stop
    Write-OK "Auto-reply: $($AutoReply.AutoReplyState)"
} catch {
    Write-Warn "Kan auto-reply status niet ophalen."
}

#endregion

#region HTML-rapport genereren

Write-Step "HTML-rapport genereren"

$ReportFile = Join-Path $ReportPath ("DiagReport_{0}_{1}.html" -f `
    ($UserPrincipalName -replace '[^a-zA-Z0-9]', '_'), `
    (Get-Date -Format 'yyyyMMdd_HHmm'))

# Bouw tabelrijen voor shared mailboxen
$SharedRows = foreach ($SMB in $SharedMailboxen) {
    $AMVal   = $SMB.AutoMapping
    $AMClass = switch -Wildcard ($AMVal.ToString()) {
        '*Nee*'    { 'ok' }
        '$false'   { 'ok' }
        'False'    { 'ok' }
        default    { 'warn' }
    }
    $AMNote  = if ($SMB.AutoMappingNote) { "<br><small style='font-weight:normal;color:#7c4a00'>$($SMB.AutoMappingNote)</small>" } else { '' }
    "<tr>
        <td>$($SMB.DisplayName)</td>
        <td>$($SMB.EmailAddress)</td>
        <td class='$AMClass'>$AMVal$AMNote</td>
        <td>$($SMB.RecipientType)</td>
    </tr>"
}
if (-not $SharedRows) {
    $SharedRows = "<tr><td colspan='4' style='color:#6b6a65'>Geen shared mailboxen gevonden met FullAccess.</td></tr>"
}

# Automapping-fix commando's voor rapport
$AutoMappingFixCmds = ''
if ($SharedMailboxen.Count -gt 0) {
    $AutoMappingFixCmds = ($SharedMailboxen | ForEach-Object {
        "Remove-MailboxPermission -Identity `"$($_.EmailAddress)`" -User `"$UserPrincipalName`" -AccessRights FullAccess -Confirm:`$false`nAdd-MailboxPermission    -Identity `"$($_.EmailAddress)`" -User `"$UserPrincipalName`" -AccessRights FullAccess -AutoMapping `$false"
    }) -join "`n`n"
}

# Remediatie-adviezen
$Adviezen = [System.Collections.Generic.List[string]]::new()

$AutoMappingCount = $SharedMailboxen.Count

if ($AutoMappingCount -ge 15) {
    $Adviezen.Add(@"
<li><strong>AutoMapping — te veel gekoppelde mailboxen ($AutoMappingCount stuks):</strong>
Outlook probeert alle mailboxen waarop AutoMapping=Aan staat automatisch te mounten.
Bij meer dan ~15 mailboxen kan dit leiden tot een te grote OST of time-outs bij profiel laden,
wat resulteert in 0x8004010F.<br>
Schakel AutoMapping uit voor mailboxen die de gebruiker ook handmatig kan openen
(via <em>Bestand &rarr; Accountinstellingen &rarr; Meer instellingen &rarr; Geavanceerd</em>):<br>
<code style='white-space:pre-line'>$AutoMappingFixCmds</code><br>
<strong>Let op:</strong> verwijder daarna het Outlook-profiel en maak een nieuw profiel aan.
</li>
"@)
} elseif ($AutoMappingCount -gt 0) {
    $Adviezen.Add("<li><strong>Shared mailboxen ($AutoMappingCount):</strong> Aantal lijkt acceptabel. Controleer of AutoMapping voor alle mailboxen gewenst is. Zie tabel hierboven.</li>")
}

if ($CASConfig -and -not $CASConfig.MAPIEnabled) {
    $Adviezen.Add("<li><strong>MAPI uitgeschakeld — directe oorzaak!</strong> Schakel MAPI in:<br><code>Set-CASMailbox -Identity `"$UserPrincipalName`" -MAPIEnabled `$true</code></li>")
}

if ($MailboxStats) {
    if ($PrimaryBytes -ge 50GB) {
        $GB = [math]::Round($PrimaryBytes / 1GB, 1)
        $Adviezen.Add("<li><strong>Grote mailbox ($GB GB):</strong> Overweeg archivering in te schakelen of het quotum te verhogen. Grote mailboxen verhogen de kans op OST-corruptie en trage sync. Na opschonen: verwijder het OST-bestand en herstel het Outlook-profiel.</li>")
    }
}

$Adviezen.Add("<li><strong>OST-bestand verwijderen en opnieuw laten aanmaken:</strong><br>Sluit Outlook, ga naar <code>%LOCALAPPDATA%\Microsoft\Outlook\</code>, hernoem of verwijder het .ost-bestand voor dit account, herstart Outlook. Exchange Online bouwt het bestand opnieuw op.</li>")
$Adviezen.Add("<li><strong>Nieuw Outlook-profiel aanmaken:</strong><br>Configuratiescherm &rarr; E-mail (Microsoft Outlook) &rarr; Profielen beheren &rarr; Toevoegen. Stel het nieuwe profiel in als standaard.</li>")
$Adviezen.Add("<li><strong>Autodiscover controleren:</strong><br><code>Test-OutlookWebServices -Identity `"$UserPrincipalName`"</code><br>Of vanuit Outlook: Ctrl + klik op het Outlook-pictogram in het systeemvak &rarr; <em>E-mailautoconfiguratie testen</em>.</li>")
$Adviezen.Add("<li><strong>Cached Exchange Mode-bereik verkleinen:</strong><br>Als het OST-bestand te groot wordt: Outlook &rarr; Accountinstellingen &rarr; account bewerken &rarr; schuifregelaar <em>E-mail bewaren voor offline gebruik</em> verkleinen naar bijv. 3 of 6 maanden.</li>")

$AdviezHtml = "<ul>$($Adviezen -join '')</ul>"

# Samenvatting variabelen voor rapport
$PrimaryGrootte = if ($MailboxStats) { $MailboxStats.TotalItemSize } else { 'Onbekend' }
$PrimaryItems   = if ($MailboxStats) { $MailboxStats.ItemCount }     else { 'Onbekend' }
$ArchieGrootte  = if ($ArchiveStats) { $ArchiveStats.TotalItemSize } else { 'Geen archief' }
$ArchieItems    = if ($ArchiveStats) { $ArchiveStats.ItemCount }     else { '-' }
$MAPIStatus     = if ($CASConfig) {
    if ($CASConfig.MAPIEnabled) { "<span class='ok'>Ingeschakeld</span>" }
    else { "<span class='fail'>UITGESCHAKELD</span>" }
} else { 'Onbekend' }
$Tijdzone = if ($RegionalConfig) { $RegionalConfig.TimeZone } else { 'Onbekend' }
$Taal     = if ($RegionalConfig) { $RegionalConfig.Language } else { 'Onbekend' }
$Quotum   = if ($QuotumInfo)     { $QuotumInfo.ProhibitSendReceiveQuota } else { 'Onbekend' }

$SharedAMNote = @"
<p style="font-size:12px;color:#6b6a65;margin-top:8px">
  <strong>Opmerking AutoMapping:</strong>
  In Exchange Online (cloud-only) heeft <code>Get-MailboxPermission</code> geen
  <code>.AutoMapping</code>-property — dit attribuut bestaat alleen in on-premises Exchange.
  AutoMapping staat standaard <strong>AAN</strong> tenzij bij het toewijzen van rechten
  uitdrukkelijk <code>-AutoMapping `$false</code> is gebruikt.
  Kolom toont 'Vermoedelijk Aan' als de exacte waarde niet bepaald kan worden.
  Verificatie per mailbox: <code>Get-MailboxPermission -Identity &lt;mailbox&gt; -User $UserPrincipalName | Select AccessRights, AutoMapping</code>
</p>
"@

$HTML = @"
<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<title>0x8004010F Diagnostieksrapport – $UserPrincipalName</title>
<style>
  :root {
    --bg:#f8f8f6;--card:#ffffff;--border:#e0dfda;
    --text:#1e1e1c;--muted:#6b6a65;--code:#f0efe9;
    --ok:#176a38;--warn-bg:#fef3e2;--warn:#7c4a00;
    --fail-bg:#fdecea;--fail:#8b1a1a;
  }
  body{font-family:system-ui,sans-serif;background:var(--bg);color:var(--text);margin:0;padding:24px;font-size:14px;line-height:1.5}
  h1{font-size:20px;font-weight:600;margin-bottom:4px}
  h2{font-size:15px;font-weight:600;margin:24px 0 8px;border-bottom:1px solid var(--border);padding-bottom:4px}
  .meta{color:var(--muted);font-size:12px;margin-bottom:24px}
  .card{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:16px 20px;margin-bottom:16px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{text-align:left;padding:6px 10px;background:#f0efe9;color:var(--muted);font-weight:500}
  td{padding:6px 10px;border-top:1px solid var(--border);vertical-align:top}
  .ok{color:var(--ok);font-weight:500}
  .warn{color:var(--warn);background:var(--warn-bg);font-weight:500}
  .fail{color:var(--fail);background:var(--fail-bg);font-weight:500}
  code{background:var(--code);border-radius:4px;padding:2px 6px;font-size:12px;font-family:monospace;white-space:pre-wrap;display:inline-block;margin-top:2px}
  ul{margin:8px 0;padding-left:20px}
  li{margin-bottom:12px;line-height:1.6}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .kv{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border)}
  .kv:last-child{border-bottom:none}
  .kv-label{color:var(--muted)}
  @media(max-width:600px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>Diagnostieksrapport – Outlook-fout 0x8004010F</h1>
<p class="meta">Gebruiker: <strong>$UserPrincipalName</strong> &nbsp;|&nbsp; Gegenereerd: $(Get-Date -Format 'dd-MM-yyyy HH:mm') &nbsp;|&nbsp; Door: $env:USERNAME</p>

<div class="grid">
  <div class="card">
    <h2>Mailboxgegevens</h2>
    <div class="kv"><span class="kv-label">Weergavenaam</span><span>$($Mailbox.DisplayName)</span></div>
    <div class="kv"><span class="kv-label">Type</span><span>$($Mailbox.RecipientTypeDetails)</span></div>
    <div class="kv"><span class="kv-label">MAPI (Outlook)</span><span>$MAPIStatus</span></div>
    <div class="kv"><span class="kv-label">Tijdzone</span><span>$Tijdzone</span></div>
    <div class="kv"><span class="kv-label">Taal</span><span>$Taal</span></div>
    <div class="kv"><span class="kv-label">Max. quotum</span><span>$Quotum</span></div>
  </div>
  <div class="card">
    <h2>Grootte &amp; items</h2>
    <div class="kv"><span class="kv-label">Primaire mailbox</span><span>$PrimaryGrootte</span></div>
    <div class="kv"><span class="kv-label">Primaire items</span><span>$PrimaryItems</span></div>
    <div class="kv"><span class="kv-label">Archief grootte</span><span>$ArchieGrootte</span></div>
    <div class="kv"><span class="kv-label">Archief items</span><span>$ArchieItems</span></div>
    <div class="kv"><span class="kv-label">Shared mbx (FullAccess)</span><span>$($SharedMailboxen.Count)</span></div>
    <div class="kv"><span class="kv-label">AutoMapping-status</span><span>Zie tabel</span></div>
  </div>
</div>

<div class="card">
  <h2>Shared mailboxen met FullAccess-rechten</h2>
  <table>
    <thead><tr><th>Naam</th><th>E-mailadres</th><th>AutoMapping</th><th>Type</th></tr></thead>
    <tbody>$($SharedRows -join '')</tbody>
  </table>
  $SharedAMNote
</div>

<div class="card">
  <h2>Remediatie-adviezen voor 0x8004010F</h2>
  $AdviezHtml
</div>

<p class="meta" style="margin-top:24px">Scriptversie 2.0 – Diagnose-0x8004010F.ps1</p>
</body>
</html>
"@

try {
    $HTML | Out-File -FilePath $ReportFile -Encoding UTF8 -Force
    Write-OK "Rapport opgeslagen: $ReportFile"
    Write-Host ""
    Write-Host "  Open het rapport in een browser voor een overzichtelijke weergave." -ForegroundColor Cyan
    # Optioneel: direct openen in standaardbrowser
    # Start-Process $ReportFile
} catch {
    Write-Fail "Kan rapport niet opslaan naar '$ReportFile': $_"
}

#endregion
