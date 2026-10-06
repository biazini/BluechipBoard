<#
.SYNOPSIS
    Browser tests for the Bluechip Board website (headless Chrome or Edge; no files changed in the project).
.DESCRIPTION
    Builds a test page from the template in Bluechip-Board.ps1 and a data file (by default the latest
    bluechip-board-data.json), adds a hostile headline that goes through the same "<" escaping as the script, and opens
    it once per scenario of site-scenarios.js, each in a fresh browser profile. Scenarios cover a new user, old and
    new backups, merging between browsers, deletions, a future stock split, currency failures, stale prices,
    FIFO sales, the exchange calendar, the statistics, HTML injection, the dividend projection (div*) and the
    Anexo J tax export (tax*).

        powershell -ExecutionPolicy Bypass -File .\Tests\Test-Site.ps1
        .\Tests\Test-Site.ps1 -Only split,fifo

    taxBaseline compares the Anexo J CSV files byte by byte with Tests\fixtures\taxBaseline. Those reference files are
    written once with -WriteFixtures, and only when missing: an existing reference is never overwritten.
#>
param([string]$Data = '', [string[]]$Only = @(), [switch]$WriteFixtures)
$ErrorActionPreference = 'Stop'
$raiz = Split-Path $PSScriptRoot
if (-not $Data) { $Data = Join-Path $raiz 'bluechip-board-data.json' }
# num clone do git (sem dados locais, que não vão para o git) usa a amostra guardada em Tests\fixtures
if (-not (Test-Path -LiteralPath $Data)) { $Data = Join-Path $PSScriptRoot 'fixtures\bluechip-board-data.sample.json' }
$browser = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $browser) { throw 'Chrome or Edge not found' }

# ---- página de teste: template do script + dados + cenários ----
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $raiz 'Bluechip-Board.ps1'), [ref]$null, [ref]$null)
$fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-Plantilla' }, $true) | Select-Object -First 1
. ([scriptblock]::Create($fn.Extent.Text))
$json = [IO.File]::ReadAllText($Data).Trim()
if (-not $json.StartsWith('{"geradoEm"')) { throw "Unexpected data file (expected the format of version 1.1 or later): $Data" }
# manchete hostil, escapada pelo mesmo método do script
$hostil = '{"chave":"xss","titulo":"<img src=x onerror=\"window.__xss=1\"></script><script>window.__xss=2</script>","link":"javascript:window.__xss=3","fonte":"<b>evil</b>","dominio":"x","data":null,"feed":"x","empresas":["AAPL"],"temas":["Earnings"],"nivel":"orange","score":99,"severo":false,"sentimento":"neutro","tier":"outra","novo":false,"soFeed":false,"outras":[{"titulo":"<svg onload=window.__xss=4>","link":"javascript:1","fonte":"<i>x</i>","tier":"outra","data":null,"chave":"y"}]}'
$i = $json.IndexOf('"noticias":['); if ($i -lt 0) { throw 'no news array in the data' }
$json = $json.Insert($i + 12, $hostil + ',').Replace('<', '\' + 'u003c')
$site = '{"backup":null,' + $json.Substring(1)
$cenarios = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'site-scenarios.js'))
# ficheiros de referência (bytes em base64) para as comparações byte a byte, por exemplo os CSV do anexo J
$fixDir = Join-Path $PSScriptRoot 'fixtures\taxBaseline'
$fix = [ordered]@{}
if (Test-Path -LiteralPath $fixDir) { foreach ($f in Get-ChildItem -LiteralPath $fixDir -Filter '*.csv' -File) { $fix[$f.Name] = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f.FullName)) } }
$cenarios = 'window.__FIX=' + (ConvertTo-Json $fix -Compress) + ';' + $cenarios
$tpl = Get-Plantilla
$tpl = $tpl.Replace('<meta charset="utf-8">', '<meta charset="utf-8"><script>window.__errs=[];addEventListener("error",e=>window.__errs.push(String(e.message)+" @"+e.lineno));</script>')
$tpl = $tpl.Replace('__DADOS_JSON__</script>', '__DADOS_JSON__</script><script>' + $cenarios + '</script>')
$pasta = Join-Path $env:TEMP ('bb-site-test-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
New-Item -ItemType Directory $pasta | Out-Null
$pagina = Join-Path $pasta 'site-test.html'
[IO.File]::WriteAllText($pagina, $tpl.Replace('__DADOS_JSON__', $site), (New-Object Text.UTF8Encoding($false)))

$lista = @('newuser', 'v1backup', 'reopen1', 'reopen2', 'legacy', 'merge', 'tomb', 'wrongfolder', 'restore', 'migrate', 'split', 'fxlong', 'fxnone', 'fxecb', 'stale', 'prev', 'fifo', 'cal', 'stats', 'xss', 'corrupt',
    'div', 'divnone', 'divfail', 'divbad', 'divstale', 'divfx', 'div8a', 'div8afx', 'div8anone', 'newsrel', 'newsrelempty', 'evts', 'evtsempty', 'fund', 'fundempty', 'macro', 'macroempty', 'tax', 'taxsplit', 'taxcrypto', 'taxold', 'taxbad', 'taxfx', 'taxfx0', 'taxBaseline', 'ret', 'retsales', 'retshort', 'retmissing', 'retempty', 'tgtsplit', 'tgtsplit2', 'tgtnoprice', 'tgtsum', 'tgtmerge', 'tgtold', 'polctx', 'polnone', 'polzero', 'polform', 'polnotes', 'polmerge', 'polold', 'stress', 'stressnone', 'strat', 'roll', 'rollshort', 'rollbtc', 'expo', 'expostocks', 'expoold', 'conc', 'sellsim', 'sellbtc', 'fees', 'feesold',
    'etfui', 'etfpf', 'etfbackup', 'etfmissing', 'etfstale', 'etfhold', 'etfbad')
$Only = @($Only | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })   # "-Only a,b" chega como um só texto com -File
if ($Only.Count) { $lista = @($lista | Where-Object { $_ -in $Only }) }
$falhas = 0; $total = 0
try {
    foreach ($s in $lista) {
        $perfil = Join-Path $pasta $(if ($s -eq 'reopen2') { 'p-reopen1' } else { "p-$s" })   # reopen2 reabre o perfil de reopen1
        $url = 'file:///' + $pagina.Replace('\', '/') + "?s=$s"
        # (o browser escreve avisos em stderr; no Windows PowerShell 5.1 isso seria um erro fatal com 'Stop')
        $ErrorActionPreference = 'Continue'
        $dom = & $browser --headless=new --disable-gpu --no-first-run --no-default-browser-check --user-data-dir="$perfil" --virtual-time-budget=8000 --dump-dom $url 2>$null | Out-String
        $ErrorActionPreference = 'Stop'
        $m = [regex]::Match($dom, '<pre id="__res">(.*?)</pre>', 'Singleline')
        if (-not $m.Success) { $falhas++; $total++; Write-Host "  FAIL [$s] the page did not finish" -ForegroundColor Red; continue }
        $r = [Net.WebUtility]::HtmlDecode($m.Groups[1].Value) | ConvertFrom-Json
        # -WriteFixtures: grava os ficheiros de referência que ainda não existem (nunca substitui um que já exista)
        if ($WriteFixtures -and $r.out) {
            New-Item -ItemType Directory -Force $fixDir | Out-Null
            foreach ($o in $r.out.PSObject.Properties) {
                $alvo = Join-Path $fixDir $o.Name
                if (Test-Path -LiteralPath $alvo) { Write-Host "  (fixture kept: $($o.Name) already exists)" -ForegroundColor DarkGray; continue }
                [IO.File]::WriteAllBytes($alvo, [Convert]::FromBase64String($o.Value)); Write-Host "  fixture written: $($o.Name)" -ForegroundColor Yellow
            }
        }
        foreach ($e in @($r.errs)) { if ($e) { $falhas++; $total++; Write-Host "  FAIL [$s] JavaScript error: $e" -ForegroundColor Red } }
        foreach ($x in @($r.r)) {
            $total++
            if ($x[1]) { Write-Host "  ok   [$s] $($x[0])" -ForegroundColor DarkGreen }
            else { $falhas++; Write-Host "  FAIL [$s] $($x[0])  $($x[2])" -ForegroundColor Red }
        }
    }
} finally { Remove-Item $pasta -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host ''
if ($falhas) { Write-Host "$falhas of $total checks FAILED" -ForegroundColor Red; exit 1 }
Write-Host "All $total checks passed" -ForegroundColor Green
