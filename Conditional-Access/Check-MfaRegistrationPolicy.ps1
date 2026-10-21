#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Controleert of er een Conditional Access policy bestaat die MFA-registratie afdwingt.

.DESCRIPTION
    Dit script verbindt met Microsoft Graph en analyseert alle Conditional Access policies
    op de aanwezigheid van een policy die MFA-registratie afdwingt voor reguliere gebruikers.
#>

param(
    [switch]$DetailedOutput,
    [switch]$ExportCsv,
    [string]$CsvPath = ".\CA_MFA_Registration_Check_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# ─── Connectie met Microsoft Graph ───────────────────────────────────────────
Write-Host "`nVerbinden met Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All" -ErrorAction Stop
    Write-Host "[OK] Verbonden met Microsoft Graph.`n" -ForegroundColor Green
} catch {
    Write-Error "[FOUT] Verbinden mislukt: $_"
    exit 1
}

# ─── Ophalen van alle Conditional Access Policies ────────────────────────────
Write-Host "Ophalen van alle Conditional Access policies..." -ForegroundColor Cyan

try {
    $policies = Get-MgIdentityConditionalAccessPolicy -All
    Write-Host "[OK] $($policies.Count) policies gevonden.`n" -ForegroundColor Green
} catch {
    Write-Error "[FOUT] Ophalen policies mislukt: $_"
    exit 1
}

# ─── Analysefuncties ─────────────────────────────────────────────────────────

function Test-IsMfaRegistrationPolicy {
    param($policy)

    $conditions   = $policy.Conditions
    $grantControls = $policy.GrantControls
    $sessionControls = $policy.SessionControls

    # Controleer of de policy gericht is op MFA-registratie actie
    $targetsMfaRegistration = $false

    # Methode 1: UserActions bevat "urn:user:registerSecurityInfo"
    if ($conditions.Applications.IncludeUserActions -contains "urn:user:registerSecurityInfo") {
        $targetsMfaRegistration = $true
    }

    # Methode 2: AuthenticationStrength of MFA grant control aanwezig
    $requiresMfa = $false
    if ($grantControls) {
        if ($grantControls.BuiltInControls -contains "mfa") {
            $requiresMfa = $true
        }
        if ($null -ne $grantControls.AuthenticationStrength) {
            $requiresMfa = $true
        }
    }

    return [PSCustomObject]@{
        TargetsMfaRegistration = $targetsMfaRegistration
        RequiresMfa            = $requiresMfa
        IsMfaRegistrationPolicy = ($targetsMfaRegistration -and $requiresMfa)
    }
}

function Get-PolicySummary {
    param($policy)

    $userCount      = $policy.Conditions.Users.IncludeUsers.Count
    $groupCount     = $policy.Conditions.Users.IncludeGroups.Count
    $excludeCount   = $policy.Conditions.Users.ExcludeUsers.Count + 
                      $policy.Conditions.Users.ExcludeGroups.Count

    $includesAllUsers = $policy.Conditions.Users.IncludeUsers -contains "All"
    $excludesGuests   = $policy.Conditions.Users.ExcludeGroups.Count -gt 0 -or
                        ($policy.Conditions.Users.ExcludeUsers | Where-Object { $_ -ne $null }).Count -gt 0

    return [PSCustomObject]@{
        PolicyName              = $policy.DisplayName
        Status                  = $policy.State
        IncludesAllUsers        = $includesAllUsers
        IncludedUserCount       = $userCount
        IncludedGroupCount      = $groupCount
        ExcludedCount           = $excludeCount
        UserActions             = ($policy.Conditions.Applications.IncludeUserActions -join ", ")
        GrantControls           = ($policy.GrantControls.BuiltInControls -join ", ")
        AuthStrength            = $policy.GrantControls.AuthenticationStrength.DisplayName
        CreatedDateTime         = $policy.CreatedDateTime
        ModifiedDateTime        = $policy.ModifiedDateTime
    }
}

# ─── Analyse uitvoeren ────────────────────────────────────────────────────────
Write-Host "Analyseren van policies op MFA-registratieafdwinging...`n" -ForegroundColor Cyan

$results = @()
$mfaRegistrationPolicies  = @()
$partialPolicies          = @()

foreach ($policy in $policies) {
    $check   = Test-IsMfaRegistrationPolicy -policy $policy
    $summary = Get-PolicySummary -policy $policy

    $result = [PSCustomObject]@{
        PolicyName              = $policy.DisplayName
        Status                  = $policy.State
        TargetsMfaRegistration  = $check.TargetsMfaRegistration
        RequiresMfaOrStrength   = $check.RequiresMfa
        IsCompleteMfaRegPolicy  = $check.IsMfaRegistrationPolicy
        IncludesAllUsers        = $summary.IncludesAllUsers
        IncludedUserCount       = $summary.IncludedUserCount
        IncludedGroupCount      = $summary.IncludedGroupCount
        ExcludedCount           = $summary.ExcludedCount
        UserActions             = $summary.UserActions
        GrantControls           = $summary.GrantControls
        AuthStrength            = $summary.AuthStrength
        CreatedDateTime         = $summary.CreatedDateTime
        ModifiedDateTime        = $summary.ModifiedDateTime
    }

    $results += $result

    if ($check.IsMfaRegistrationPolicy) {
        $mfaRegistrationPolicies += $result
    } elseif ($check.TargetsMfaRegistration -or $check.RequiresMfa) {
        $partialPolicies += $result
    }
}

# ─── Resultaten weergeven ─────────────────────────────────────────────────────
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor White
Write-Host "  RESULTAAT: MFA-REGISTRATIE CONDITIONAL ACCESS ANALYSE" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor White

if ($mfaRegistrationPolicies.Count -gt 0) {
    Write-Host "[OK] GEVONDEN: $($mfaRegistrationPolicies.Count) policy/policies die MFA-registratie afdwingen:`n" -ForegroundColor Green

    foreach ($p in $mfaRegistrationPolicies) {
        Write-Host "  - Naam    : $($p.PolicyName)" -ForegroundColor Yellow
        Write-Host "  - Status  : $($p.Status)" -ForegroundColor $(if ($p.Status -eq "enabled") { "Green" } elseif ($p.Status -eq "enabledForReportingButNotEnforced") { "Yellow" } else { "Red" })
        Write-Host "  - Alle gebruikers : $($p.IncludesAllUsers)"
        Write-Host "  - User Actions    : $($p.UserActions)"
        Write-Host "  - Grant Controls  : $($p.GrantControls)"
        if ($p.AuthStrength) {
        Write-Host "  - Auth Strength   : $($p.AuthStrength)" }
        Write-Host "  - Aangemaakt      : $($p.CreatedDateTime)"
        Write-Host ""
    }

    # Waarschuwing als policy alleen in report-only staat
    $reportOnly = $mfaRegistrationPolicies | Where-Object { $_.Status -eq "enabledForReportingButNotEnforced" }
    if ($reportOnly.Count -gt 0) {
        Write-Host "[!] WAARSCHUWING: $($reportOnly.Count) policy/policies staan in 'Report-only' modus en dwingen NIET af!`n" -ForegroundColor Yellow
    }

} else {
    Write-Host "[FOUT] GEEN volledige MFA-registratie policy gevonden!`n" -ForegroundColor Red
    Write-Host "   Er is geen Conditional Access policy die zowel:" -ForegroundColor Red
    Write-Host "   1. Gericht is op de actie 'registerSecurityInfo'" -ForegroundColor Red
    Write-Host "   2. MFA of Authentication Strength vereist`n" -ForegroundColor Red
}

# Toon gedeeltelijke matches
if ($partialPolicies.Count -gt 0) {
    Write-Host "[!] GEDEELTELIJKE MATCHES ($($partialPolicies.Count) policies - mogelijk relevant):`n" -ForegroundColor Yellow
    foreach ($p in $partialPolicies) {
        Write-Host "  - $($p.PolicyName) [Status: $($p.Status)]" -ForegroundColor Yellow
        Write-Host "     Targets MFA Registratie : $($p.TargetsMfaRegistration)"
        Write-Host "     Vereist MFA             : $($p.RequiresMfaOrStrength)"
        Write-Host "     User Actions            : $($p.UserActions)"
        Write-Host ""
    }
}

# ─── Gedetailleerde output (optioneel) ───────────────────────────────────────
if ($DetailedOutput) {
    Write-Host "`nALLE POLICIES (overzicht):`n" -ForegroundColor Cyan
    $results | Format-Table -AutoSize -Property PolicyName, Status, TargetsMfaRegistration, RequiresMfaOrStrength, IsCompleteMfaRegPolicy, IncludesAllUsers
}

# ─── CSV export (optioneel) ───────────────────────────────────────────────────
if ($ExportCsv) {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Resultaten geëxporteerd naar: $CsvPath`n" -ForegroundColor Green
}

# ─── Afsluiten ────────────────────────────────────────────────────────────────
Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor White
Disconnect-MgGraph | Out-Null
Write-Host "Verbinding met Microsoft Graph verbroken.`n" -ForegroundColor Gray