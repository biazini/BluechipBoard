# Launcher for Bluechip-Board.ps1 (used by the desktop shortcut).
# The window closes by itself when everything goes well; it stays open if there are errors or warnings.
# The SEC contact e-mail is not in this file (it is in git): it comes from the BLUECHIP_SEC_EMAIL environment variable or,
# if that is empty, from bluechip-board.config.json next to this file (not in git; copy bluechip-board.config.example.json).
# Without it, the board still runs, with the SEC sources skipped.
try { [Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; [Globalization.CultureInfo]::DefaultThreadCurrentUICulture = 'en-US' } catch { }
$script = Join-Path $PSScriptRoot 'Bluechip-Board.ps1'
$warnings = @()
$failed = $false

$secEmail = "$env:BLUECHIP_SEC_EMAIL".Trim()
$configPath = Join-Path $PSScriptRoot 'bluechip-board.config.json'
if (-not $secEmail -and (Test-Path -LiteralPath $configPath)) {
    try { $secEmail = "$((Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json).secEmail)".Trim() }
    catch { Write-Warning "bluechip-board.config.json could not be read ($($_.Exception.Message)): running without the SEC sources." }
}
if ($secEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or $secEmail -eq 'your@email.com') {
    if ($secEmail) { Write-Warning 'The SEC e-mail in bluechip-board.config.json is not valid: running without the SEC sources.' }
    else { Write-Host 'No SEC e-mail (bluechip-board.config.json or BLUECHIP_SEC_EMAIL): the SEC sources are skipped.' -ForegroundColor DarkYellow }
    $secEmail = ''
}

try {
    # The board runs with $ErrorActionPreference = 'Stop', so any unhandled error reaches this catch.
    # (-ErrorVariable is not used: it would also catch errors the script already handles, such as download retries.)
    if ($secEmail) { & $script -SecEmail $secEmail -WarningVariable +warnings }
    else { & $script -WarningVariable +warnings }
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