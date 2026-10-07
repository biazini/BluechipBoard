# Launcher for Bluechip-Board.ps1 (used by the desktop shortcut).
# The window closes by itself when everything goes well; it stays open if there are errors or warnings.
# The SEC contact e-mail is not in this file (it is in git), nor in any command line: Bluechip-Board.ps1 reads it itself from
# the BLUECHIP_SEC_EMAIL environment variable or, if that is empty, from bluechip-board.config.json next to it (not in git;
# copy bluechip-board.config.example.json). Without it, the board still runs, with the SEC sources skipped.
try { [Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; [Globalization.CultureInfo]::DefaultThreadCurrentUICulture = 'en-US' } catch { }
$script = Join-Path $PSScriptRoot 'Bluechip-Board.ps1'
$warnings = @()
$failed = $false

try {
    # The board runs with $ErrorActionPreference = 'Stop', so any unhandled error reaches this catch.
    # (-ErrorVariable is not used: it would also catch errors the script already handles, such as download retries.)
    & $script -WarningVariable +warnings
} catch {
    $failed = $true
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkRed
}

if ($failed -or $warnings.Count -gt 0) {
    Write-Host ''
    Write-Host "Finished with problems ($([int]$failed) error(s), $($warnings.Count) warning(s))." -ForegroundColor Yellow
    Write-Host 'Press Enter to close...'
    [void][Console]::ReadLine()
}
