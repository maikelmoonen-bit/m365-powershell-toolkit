#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Retroactively deletes all mailbox items older than a retention period (default 7 years)
    from a specified mailbox.

.DESCRIPTION
    Connects to Exchange Online and uses Compliance Search to find and purge all mailbox
    items received before the cutoff date (today minus -RetentionYears) from the mailbox
    given in -TargetMailbox.

.NOTES
    Requirements:
    - Exchange Online Management module: Install-Module ExchangeOnlineManagement
    - Security and Compliance PowerShell access
    - The account running this script must have the Compliance Search and
      New-ComplianceSearchAction roles (eDiscovery Manager or Compliance Administrator)
    - This action is IRREVERSIBLE. Purged items cannot be recovered.

.EXAMPLE
    .\Delete-OldMailboxItems.ps1 -TargetMailbox administratie@contoso.com -WhatIf
    .\Delete-OldMailboxItems.ps1 -TargetMailbox administratie@contoso.com -RetentionYears 10
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$TargetMailbox,

    [ValidateRange(1, 50)]
    [int]$RetentionYears = 7,

    [switch]$WhatIf
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$CutoffDate      = (Get-Date).Date.AddYears(-$RetentionYears)
$SearchName      = "Delete_OldItems_$($TargetMailbox -replace '[^a-zA-Z0-9_-]', '_')"
$CutoffFormatted = $CutoffDate.ToString("yyyy-MM-dd")

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Mailbox Cleanup Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Target mailbox : $TargetMailbox"
Write-Host " Cutoff date    : $CutoffFormatted (items BEFORE this date will be deleted)"
Write-Host " WhatIf mode    : $($WhatIf.IsPresent)"
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Safety confirmation
# ---------------------------------------------------------------------------
if (-not $WhatIf) {
    Write-Warning "You are about to PERMANENTLY DELETE all mailbox items older than $CutoffFormatted from '$TargetMailbox'."
    $confirm = Read-Host "Type 'YES' to continue, or anything else to abort"
    if ($confirm -ne "YES") {
        Write-Host "Aborted by user." -ForegroundColor Yellow
        exit 0
    }
}

# ---------------------------------------------------------------------------
# Connect to Security and Compliance Center
# ---------------------------------------------------------------------------
try {
    Write-Host "[1/5] Connecting to Security and Compliance PowerShell..." -ForegroundColor Green
    # Connect-IPPSSession establishes the remote session but does not always
    # auto-import compliance cmdlets. We explicitly import the resulting
    # temporary module so Get-ComplianceSearch etc. are available.
    Connect-IPPSSession -ErrorAction Stop
    # Give the implicit remoting module a moment to register, then import it
    Start-Sleep -Seconds 3
    Get-Module -Name "tmp_*" | Import-Module -ErrorAction SilentlyContinue
    # Verify the compliance cmdlets loaded
    if (-not (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue)) {
        throw "Compliance cmdlets not available after connecting. Ensure your account has the Compliance Search role and try again."
    }
    Write-Host "      Connected successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Security and Compliance PowerShell: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Clean up any existing search with the same name
# ---------------------------------------------------------------------------
Write-Host "[2/5] Checking for existing compliance search '$SearchName'..." -ForegroundColor Green
$existingSearch = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
if ($existingSearch) {
    Write-Host "      Removing existing search..." -ForegroundColor Yellow
    Remove-ComplianceSearch -Identity $SearchName -Confirm:$false
}

# ---------------------------------------------------------------------------
# Create the Compliance Search
# ---------------------------------------------------------------------------
$KQLQuery = "received<$CutoffFormatted"

Write-Host "[3/5] Creating compliance search..." -ForegroundColor Green
Write-Host "      Query : $KQLQuery"
Write-Host "      Scope : $TargetMailbox"

try {
    New-ComplianceSearch `
        -Name              $SearchName `
        -ExchangeLocation  $TargetMailbox `
        -ContentMatchQuery $KQLQuery `
        -ErrorAction       Stop | Out-Null

    Write-Host "      Compliance search created. Starting search..." -ForegroundColor Green
    Start-ComplianceSearch -Identity $SearchName -ErrorAction Stop

    Write-Host "      Waiting for search to complete (this may take a while)..." -ForegroundColor Green
    do {
        Start-Sleep -Seconds 15
        $search = Get-ComplianceSearch -Identity $SearchName
        Write-Host "      Status: $($search.Status) | Items found: $($search.Items)"
    } while ($search.Status -notin @("Completed", "Failed"))

    if ($search.Status -eq "Failed") {
        Write-Error "Compliance search failed. Please check the Security and Compliance Center for details."
        exit 1
    }

    Write-Host ""
    Write-Host "      Search completed." -ForegroundColor Green
    Write-Host "      Total items found : $($search.Items)"
    Write-Host "      Total size        : $($search.Size) bytes"
    Write-Host ""
}
catch {
    Write-Error "Failed to create or run compliance search: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Purge the items
# ---------------------------------------------------------------------------
if ($search.Items -eq 0) {
    Write-Host "[4/5] No items found matching the criteria. Nothing to delete." -ForegroundColor Yellow
}
elseif ($WhatIf) {
    Write-Host "[4/5] WhatIf mode: Would delete $($search.Items) items from '$TargetMailbox'." -ForegroundColor Yellow
    Write-Host "      Run without -WhatIf to perform the actual deletion." -ForegroundColor Yellow
}
else {
    Write-Host "[4/5] Purging $($search.Items) items from '$TargetMailbox'..." -ForegroundColor Green
    Write-Host "      Using HardDelete - items will be permanently removed." -ForegroundColor Red

    try {
        New-ComplianceSearchAction `
            -SearchName  $SearchName `
            -Purge `
            -PurgeType   HardDelete `
            -Confirm:$false `
            -ErrorAction Stop | Out-Null

        $actionName = $SearchName + "_Purge"
        Write-Host "      Waiting for purge action to complete..." -ForegroundColor Green
        do {
            Start-Sleep -Seconds 15
            $action = Get-ComplianceSearchAction -Identity $actionName
            Write-Host "      Purge status: $($action.Status)"
        } while ($action.Status -notin @("Completed", "Failed"))

        if ($action.Status -eq "Failed") {
            Write-Error "Purge action failed. Please check the Security and Compliance Center for details."
            exit 1
        }

        Write-Host "      Purge completed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to purge items: $_"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------
Write-Host "[5/5] Cleaning up compliance search object..." -ForegroundColor Green
Remove-ComplianceSearch -Identity $SearchName -Confirm:$false -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Done." -ForegroundColor Green
Write-Host " All mailbox items older than $CutoffFormatted have been" -ForegroundColor Green
Write-Host " permanently deleted from '$TargetMailbox'." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
