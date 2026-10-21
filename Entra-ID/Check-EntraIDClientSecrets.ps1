<#
.SYNOPSIS
    Controleert de status van client secrets voor alle App Registrations in Entra ID.

.DESCRIPTION
    Dit script maakt verbinding met Microsoft Graph en controleert alle app registrations
    om te zien welke client secrets verlopen zijn, bijna verlopen of nog geldig zijn.

.PARAMETER DaysWarning
    Aantal dagen vooruit om te waarschuwen voor verlopende secrets (standaard: 30)

.EXAMPLE
    .\Check-EntraIDClientSecrets.ps1
    
.EXAMPLE
    .\Check-EntraIDClientSecrets.ps1 -DaysWarning 60
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$DaysWarning = 30
)

# Vereiste modules controleren en installeren
$requiredModule = "Microsoft.Graph.Applications"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module $requiredModule niet gevonden. Installeren..." -ForegroundColor Yellow
    Install-Module -Name $requiredModule -Scope CurrentUser -Force
}

# Importeer de module
Import-Module Microsoft.Graph.Applications

# Verbinding maken met Microsoft Graph
Write-Host "Verbinding maken met Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    Write-Host "Verbinding succesvol!" -ForegroundColor Green
} catch {
    Write-Error "Kan geen verbinding maken met Microsoft Graph: $_"
    exit 1
}

# Huidige datum
$today = Get-Date

# Alle app registrations ophalen
Write-Host "`nApp Registrations ophalen..." -ForegroundColor Cyan
$apps = Get-MgApplication -All

# Tellers initialiseren
$totalApps = $apps.Count
$appsWithSecrets = 0
$expiredSecrets = 0
$expiringSecrets = 0
$validSecrets = 0
$noExpirySecrets = 0

# Resultaten verzamelen
$results = @()

Write-Host "Analyseren van $totalApps app registrations...`n" -ForegroundColor Cyan

foreach ($app in $apps) {
    $secrets = $app.PasswordCredentials
    
    if ($secrets.Count -gt 0) {
        $appsWithSecrets++
        
        foreach ($secret in $secrets) {
            $secretInfo = [PSCustomObject]@{
                AppName = $app.DisplayName
                AppId = $app.AppId
                ObjectId = $app.Id
                SecretKeyId = $secret.KeyId
                SecretDisplayName = $secret.DisplayName
                StartDate = $secret.StartDateTime
                EndDate = $secret.EndDateTime
                DaysRemaining = $null
                Status = ""
            }
            
            if ($secret.EndDateTime) {
                $daysRemaining = ($secret.EndDateTime - $today).Days
                $secretInfo.DaysRemaining = $daysRemaining
                
                if ($daysRemaining -lt 0) {
                    $secretInfo.Status = "VERLOPEN"
                    $expiredSecrets++
                } elseif ($daysRemaining -le $DaysWarning) {
                    $secretInfo.Status = "VERLOOPT BINNENKORT"
                    $expiringSecrets++
                } else {
                    $secretInfo.Status = "GELDIG"
                    $validSecrets++
                }
            } else {
                $secretInfo.Status = "GEEN VERVALDATUM"
                $noExpirySecrets++
            }
            
            $results += $secretInfo
        }
    }
}

# Samenvatting weergeven
Write-Host "=== SAMENVATTING ===" -ForegroundColor Cyan
Write-Host "Totaal aantal app registrations: $totalApps" -ForegroundColor White
Write-Host "App registrations met secrets: $appsWithSecrets" -ForegroundColor White
Write-Host ""
Write-Host "Status van client secrets:" -ForegroundColor White
Write-Host "  - Verlopen: $expiredSecrets" -ForegroundColor Red
Write-Host "  - Verloopt binnen $DaysWarning dagen: $expiringSecrets" -ForegroundColor Yellow
Write-Host "  - Geldig: $validSecrets" -ForegroundColor Green
Write-Host "  - Geen vervaldatum: $noExpirySecrets" -ForegroundColor Gray
Write-Host ""

# Verlopen secrets tonen
if ($expiredSecrets -gt 0) {
    Write-Host "=== VERLOPEN SECRETS ===" -ForegroundColor Red
    $results | Where-Object { $_.Status -eq "VERLOPEN" } | 
        Sort-Object DaysRemaining | 
        Format-Table AppName, SecretDisplayName, EndDate, DaysRemaining -AutoSize
}

# Bijna verlopende secrets tonen
if ($expiringSecrets -gt 0) {
    Write-Host "=== SECRETS DIE BINNENKORT VERLOPEN ===" -ForegroundColor Yellow
    $results | Where-Object { $_.Status -eq "VERLOOPT BINNENKORT" } | 
        Sort-Object DaysRemaining | 
        Format-Table AppName, SecretDisplayName, EndDate, DaysRemaining -AutoSize
}

# Optie om alle resultaten te exporteren naar CSV
$export = Read-Host "`nWil je de volledige rapportage exporteren naar CSV? (J/N)"
if ($export -eq "J" -or $export -eq "j") {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $filename = "EntraID-ClientSecrets-$timestamp.csv"
    $results | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
    Write-Host "Rapportage geëxporteerd naar: $filename" -ForegroundColor Green
}

# Verbinding verbreken
Disconnect-MgGraph | Out-Null
Write-Host "`nKlaar!" -ForegroundColor Green