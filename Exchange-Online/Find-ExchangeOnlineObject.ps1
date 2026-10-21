#requires -Version 5.1

<#
.SYNOPSIS
    Zoekt een e-mailadres breed op binnen Exchange Online.

.DESCRIPTION
    Handig wanneer een gebruiker stelt een shared mailbox te gebruiken, maar het
    opgegeven SMTP-adres niet direct als shared mailbox terug te vinden is.

    Het script controleert:
      - directe Exchange-recipient resolutie;
      - alle mail-enabled recipients op PrimarySmtpAddress en proxy/alias-adressen;
      - actieve mailboxes (incl. controle op RecipientTypeDetails);
      - Microsoft 365 Groups;
      - soft-deleted mailboxes.

    Er worden geen wijzigingen aangebracht in Exchange Online.

.EXAMPLE
    .\Find-ExchangeOnlineObject.ps1 -EmailAddress "verkoop@contoso.com"

.EXAMPLE
    .\Find-ExchangeOnlineObject.ps1 -EmailAddress "verkoop@contoso.com" -Connect
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$EmailAddress,

    [Parameter(Mandatory = $false)]
    [switch]$Connect
)

$ErrorActionPreference = 'Stop'
$Target = $EmailAddress.Trim().ToLowerInvariant()
$Findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding {
    param (
        [string]$SearchMethod,
        [object]$Object
    )

    if ($null -eq $Object) { return }

    $emailAddresses = @($Object.EmailAddresses | ForEach-Object { $_.ToString() })
    $primary = if ($Object.PrimarySmtpAddress) { $Object.PrimarySmtpAddress.ToString() } else { $null }

    $Findings.Add([pscustomobject]@{
        SearchMethod        = $SearchMethod
        DisplayName         = $Object.DisplayName
        Name                = $Object.Name
        Alias               = $Object.Alias
        PrimarySmtpAddress  = $primary
        RecipientType       = $Object.RecipientType
        RecipientTypeDetails= $Object.RecipientTypeDetails
        ExternalDirectoryObjectId = $Object.ExternalDirectoryObjectId
        Identity            = $Object.Identity
        EmailAddresses      = ($emailAddresses -join '; ')
    })
}

function Test-SmtpMatch {
    param ([object]$Object)

    if ($Object.PrimarySmtpAddress -and
        $Object.PrimarySmtpAddress.ToString().Equals($Target, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    foreach ($proxy in @($Object.EmailAddresses)) {
        $address = $proxy.ToString()
        if ($address -match '^(?i)smtp:(.+)$' -and
            $Matches[1].Equals($Target, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

Write-Host "Exchange Online object search" -ForegroundColor Cyan
Write-Host "Target: $Target" -ForegroundColor Cyan
Write-Host ""

if ($Connect) {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'ExchangeOnlineManagement module is niet geinstalleerd. Installeer deze met: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    }
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false
}

if (-not (Get-Command Get-Recipient -ErrorAction SilentlyContinue)) {
    throw 'Geen actieve Exchange Online PowerShell-sessie gevonden. Maak eerst verbinding met Connect-ExchangeOnline, of start het script met -Connect.'
}

# 1. Directe recipient lookup. Dit is snel en vindt veel aliassen direct.
Write-Host '[1/5] Directe recipient lookup...' -ForegroundColor Yellow
try {
    $direct = Get-Recipient -Identity $Target -ErrorAction Stop
    Add-Finding -SearchMethod 'Get-Recipient -Identity' -Object $direct
}
catch {
    Write-Verbose "Geen directe recipient-match: $($_.Exception.Message)"
}

# 2. Tenant-brede controle op primaire SMTP- en proxyadressen.
Write-Host '[2/5] Alle mail-enabled recipients en proxyadressen doorzoeken...' -ForegroundColor Yellow
try {
    Get-Recipient -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
        if (Test-SmtpMatch -Object $_) {
            Add-Finding -SearchMethod 'Recipient SMTP/proxy scan' -Object $_
        }
    }
}
catch {
    Write-Warning "Recipient-scan kon niet volledig worden uitgevoerd: $($_.Exception.Message)"
}

# 3. Actieve mailboxen expliciet controleren, zodat zichtbaar wordt of het echt
#    om RecipientTypeDetails = SharedMailbox gaat.
Write-Host '[3/5] Actieve mailboxen controleren...' -ForegroundColor Yellow
try {
    Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
        if (Test-SmtpMatch -Object $_) {
            Add-Finding -SearchMethod 'Active mailbox scan' -Object $_
        }
    }
}
catch {
    Write-Warning "Mailbox-scan kon niet volledig worden uitgevoerd: $($_.Exception.Message)"
}

# 4. Microsoft 365 Groups expliciet meenemen.
Write-Host '[4/5] Microsoft 365 Groups controleren...' -ForegroundColor Yellow
if (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue) {
    try {
        Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
            if (Test-SmtpMatch -Object $_) {
                Add-Finding -SearchMethod 'Microsoft 365 Group scan' -Object $_
            }
        }
    }
    catch {
        Write-Warning "Microsoft 365 Group-scan kon niet volledig worden uitgevoerd: $($_.Exception.Message)"
    }
}

# 5. Soft-deleted mailboxen. Dit kan verklaren waarom de mailbox niet in de
#    normale Exchange Admin Center-weergave staat.
Write-Host '[5/5] Soft-deleted mailboxes controleren...' -ForegroundColor Yellow
try {
    Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
        if (Test-SmtpMatch -Object $_) {
            Add-Finding -SearchMethod 'Soft-deleted mailbox scan' -Object $_
        }
    }
}
catch {
    Write-Verbose "Soft-deleted mailbox-scan niet beschikbaar of niet toegestaan: $($_.Exception.Message)"
}

Write-Host ""

if ($Findings.Count -eq 0) {
    Write-Host "GEEN MATCH GEVONDEN voor $Target" -ForegroundColor Red
    Write-Host 'Het adres is niet aangetroffen als primaire SMTP/proxy op de onderzochte Exchange Online-objecten.' -ForegroundColor Red
    Write-Host 'Controleer in dat geval ook of het adres correct gespeld is en of het mogelijk in een andere tenant/domein wordt beheerd.' -ForegroundColor DarkYellow
}
else {
    # Dezelfde recipient kan via meerdere zoekmethoden voorkomen. Dit is bewust:
    # SearchMethod laat zien WAAROM en HOE het object is gevonden.
    Write-Host "MATCH(ES) GEVONDEN: $($Findings.Count)" -ForegroundColor Green
    $Findings |
        Sort-Object PrimarySmtpAddress, SearchMethod |
        Format-List SearchMethod, DisplayName, Name, Alias, PrimarySmtpAddress,
                    RecipientType, RecipientTypeDetails, ExternalDirectoryObjectId,
                    Identity, EmailAddresses

    $shared = @($Findings | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' })
    if ($shared.Count -gt 0) {
        Write-Host 'Conclusie: het adres is gekoppeld aan een SharedMailbox.' -ForegroundColor Green
    }
    else {
        Write-Host 'Let op: er is wel een object gevonden, maar geen match met RecipientTypeDetails = SharedMailbox.' -ForegroundColor Yellow
    }
}
