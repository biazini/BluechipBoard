<#
.SYNOPSIS
    Creates (or updates) the "Bluechip Board" shortcut on the desktop, pointing at Start-BluechipBoard.ps1 in this folder.
.DESCRIPTION
    Run it once after cloning the project, or after moving the folder. The shortcut runs PowerShell 7 (pwsh.exe) when it is
    installed, otherwise Windows PowerShell 5.1, with the Bluechip-Board.ico icon. Nothing else is changed: no scheduled
    task, no settings, no data.

        pwsh -ExecutionPolicy Bypass -File .\Install-Shortcut.ps1
        pwsh -ExecutionPolicy Bypass -File .\Install-Shortcut.ps1 -Shell powershell    # use Windows PowerShell 5.1
#>
param(
    # pasta onde o atalho é criado (por omissão, o ambiente de trabalho do utilizador)
    [string]$Destino = [Environment]::GetFolderPath('Desktop'),
    [ValidateSet('auto', 'pwsh', 'powershell')][string]$Shell = 'auto'
)
$ErrorActionPreference = 'Stop'
$pasta = $PSScriptRoot
$lancador = Join-Path $pasta 'Start-BluechipBoard.ps1'
if (-not (Test-Path -LiteralPath $lancador)) { throw "Start-BluechipBoard.ps1 not found next to this script ($pasta)." }

# PowerShell 7 se existir (é o que o atalho usa desde 5 out 2026); senão o Windows PowerShell 5.1, que vem com o Windows
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($Shell -eq 'pwsh' -and -not $pwsh) { throw 'PowerShell 7 (pwsh.exe) is not installed. Install it, or use -Shell powershell.' }
$exe = if ($Shell -ne 'powershell' -and $pwsh) { $pwsh.Source } else { Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }

if (-not (Test-Path -LiteralPath $Destino)) { New-Item -ItemType Directory -Path $Destino | Out-Null }
$caminho = Join-Path $Destino 'Bluechip Board.lnk'
$atalho = (New-Object -ComObject WScript.Shell).CreateShortcut($caminho)
$atalho.TargetPath = $exe
$atalho.Arguments = "-ExecutionPolicy Bypass -File `"$lancador`""
$atalho.WorkingDirectory = $pasta
$icone = Join-Path $pasta 'Bluechip-Board.ico'
if (Test-Path -LiteralPath $icone) { $atalho.IconLocation = "$icone,0" }
$atalho.Description = 'Bluechip Board: stocks, ETF and Bitcoin'
$atalho.WindowStyle = 1
$atalho.Save()
Write-Host "Shortcut saved: $caminho" -ForegroundColor Green
Write-Host "  runs $exe" -ForegroundColor DarkGray
