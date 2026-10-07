<#
.SYNOPSIS
    Failure-mode tests for Bluechip-Board.ps1, run on a patched copy in a temporary folder (the project folder, your
    Downloads folder and Windows Task Scheduler are never touched). No internet needed: every download is blocked.
.DESCRIPTION
    Scenarios: fresh install with every source down; previous run's prices (and dividends, up to 180 days old) used when the sources fail; backup brought
    from a (fake) Downloads folder only after validation; a second instance blocked by the mutex; a run that fails
    halfway leaves the previous site and vistos.json untouched; scheduled-task registration (dry run).

        powershell -ExecutionPolicy Bypass -File .\Tests\Test-Resilience.ps1
#>
param([string]$Shell = 'powershell.exe')
$ErrorActionPreference = 'Stop'
# as cópias do script também leem o e-mail da SEC de BLUECHIP_SEC_EMAIL: os testes decidem quando há e-mail (só neste processo)
Remove-Item env:BLUECHIP_SEC_EMAIL -ErrorAction SilentlyContinue
$raiz = Split-Path $PSScriptRoot
# dados da execução anterior para os testes: os locais, ou a amostra em Tests\fixtures num clone do git (os dados não vão para o git)
$dadosRef = Join-Path $raiz 'bluechip-board-data.json'; if (-not (Test-Path -LiteralPath $dadosRef)) { $dadosRef = Join-Path $PSScriptRoot 'fixtures\bluechip-board-data.sample.json' }
$base = Join-Path $env:TEMP ('bb-resil-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
New-Item -ItemType Directory $base | Out-Null
$Script:Falhas = 0; $Script:Total = 0
function Check([string]$Nome, [bool]$Cond, $Detalhe = '') { $Script:Total++; if ($Cond) { Write-Host "  ok   $Nome" -ForegroundColor DarkGreen } else { $Script:Falhas++; Write-Host "  FAIL $Nome  $Detalhe" -ForegroundColor Red } }

# ---- cópia do script com ganchos de teste ----
$t = [IO.File]::ReadAllText((Join-Path $raiz 'Bluechip-Board.ps1'))
$t = $t.Replace('function Get-Url {', "function Get-Url {`n    param([string]`$Url, [string]`$UserAgent = `$Script:UA, [int]`$Timeout = 25)`n    if (`$env:BB_TEST_BLOCK -and `$Url -match `$env:BB_TEST_BLOCK) { throw `"blocked by test: `$Url`" }`n    Get-UrlReal -Url `$Url -UserAgent `$UserAgent -Timeout `$Timeout`n}`nfunction Get-UrlReal {")
$t = $t.Replace("(New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path", '$env:BB_TEST_DOWNLOADS')
$t = $t.Replace("Write-Passo 'Building the website'", "if (`$env:BB_TEST_FAIL) { throw 'simulated failure' }`nWrite-Passo 'Building the website'")
# evento manual com a data mal escrita (só quando o teste o pede)
$t = $t.Replace('# Horário das bolsas', 'if ($env:BB_TEST_CAL_BAD) { $Calendario += @{ d = ''2026-13-45''; e = ''MKT''; ev = ''bad''; imp = ''High''; st = ''E'' } }' + "`n# Horário das bolsas")
# calendário e datas da Nasdaq de teste (por variáveis de ambiente, em JSON)
$t = $t.Replace("foreach (`$r in @(Get-DatasResultados)) {", "if (`$env:BB_TEST_CAL) { `$eventos = @(`$env:BB_TEST_CAL | ConvertFrom-Json | ForEach-Object { `$_ }) }`nforeach (`$r in @(`$(if (`$env:BB_TEST_NASDAQ) { `$env:BB_TEST_NASDAQ | ConvertFrom-Json | ForEach-Object { `$_ } } else { Get-DatasResultados }))) {")
# nome de tarefa só deste teste: mesmo que o registo verdadeiro chegasse a correr, nunca tocaria na tarefa real
$tarefaTeste = 'BluechipBoard-Test-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$t = $t.Replace("-TaskName 'BluechipBoard'", "-TaskName '$tarefaTeste'")
# limite de tamanho do Archive: com BB_TEST_ARQ_MAX (bytes) o teste usa um limite pequeno em vez de 300 MB
$t = $t.Replace('$total -gt 300MB', '$total -gt $(if ($env:BB_TEST_ARQ_MAX) { [int64]$env:BB_TEST_ARQ_MAX } else { 300MB })')
$copia = Join-Path $base 'Bluechip-Board.ps1'
[IO.File]::WriteAllText($copia, $t, (New-Object Text.UTF8Encoding($true)))

function Invoke-Copia([string]$Pasta, [hashtable]$Env, [string[]]$Args2 = @()) {
    foreach ($k in $Env.Keys) { Set-Item "env:$k" $Env[$k] }
    try {
        $ErrorActionPreference = 'Continue'
        $o = & $Shell -NoProfile -ExecutionPolicy Bypass -File $copia -NoOpen -Folder $Pasta @Args2 2>&1 | Out-String
        $ErrorActionPreference = 'Stop'
        return $o
    } finally { foreach ($k in $Env.Keys) { Remove-Item "env:$k" -ErrorAction SilentlyContinue } }
}
$dl = Join-Path $base 'Downloads'; New-Item -ItemType Directory $dl | Out-Null
Write-Host "Using $Shell" -ForegroundColor Cyan

try {
    Write-Host 'Fresh install, every source down'
    $p1 = Join-Path $base 'fresh'; New-Item -ItemType Directory $p1 | Out-Null
    $o = Invoke-Copia $p1 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'run finishes' ($o -match 'Done in') $o
    Check 'site, data, vistos.json and an Archive copy are written' ((Test-Path "$p1\bluechip-board.html") -and (Test-Path "$p1\bluechip-board-data.json") -and (Test-Path "$p1\vistos.json") -and @(Get-ChildItem "$p1\Archive").Count -eq 1)
    $d = [IO.File]::ReadAllText("$p1\bluechip-board-data.json") | ConvertFrom-Json
    Check 'data file is valid JSON and marks the first run' ($d.primeiraExecucao -eq $true)
    Check 'no source reported as ok' (@($d.fontes | Where-Object { $_.estado -like 'ok*' }).Count -eq 0)
    Check 'halving date estimated without the network' (@($d.calendario | Where-Object { $_.ev -match 'approximate date' }).Count -eq 1)
    Check 'no temporary files left' (@(Get-ChildItem $p1 -Recurse -Filter *.tmp).Count -eq 0)
    $novos = 'EUNK', 'IS3N', 'EUNN'
    Check 'new ETFs, every source down: no prices invented (empty series), price errors listed' (@($novos | Where-Object { $i = $_; $a = $d.ativos | Where-Object { $_.id -eq $i }; $a -and @(@($a.pontos) | Where-Object { @($_).Count -ge 2 }).Count -eq 0 -and $a.tipo -eq 'ETF' }).Count -eq 3 -and @($d.fontes | Where-Object { $_.nome -match '^Prices: .*\((EUNK|IS3N|EUNN)\)$' -and $_.estado -eq 'error' }).Count -eq 3)
    Check 'new ETFs, holdings down: each uses its own reference weights, metadata kept' (@($novos | Where-Object { $e = $d.etfs.$_; $e -and -not $e.aoVivo -and $e.fonte -like 'Reference weights*' -and $e.isin -like 'IE00*' -and @($e.top10).Count -eq 0 }).Count -eq 3 -and @($d.fontes | Where-Object { $_.nome -like 'iShares: ETF holdings (*' -and $_.estado -eq 'error (using reference weights)' }).Count -eq 4)
    Check 'offline: every fund''s country, sector and currency aggregates are Unavailable (null); data version 1.4' ($d.versao -eq '1.4' -and @('SXR8', 'EUNK', 'IS3N', 'EUNN' | Where-Object { $null -eq $d.etfs.$_.agregados }).Count -eq 4)
    Check 'SXR8 holdings fallback unchanged (its reference weights, also in "etf")' ($d.etf.AAPL -eq 7.81 -and -not $d.etf.aoVivo -and $d.etfs.SXR8.AAPL -eq 7.81)
    Check 'dividend source down: site built, one error entry per asset, no numbers invented' (@('AAPL', 'NVDA', 'GOOGL' | Where-Object { $d.dividendos.$_.estado -eq 'error' -and $null -eq $d.dividendos.$_.anualPorAcao }).Count -eq 3 -and @($d.fontes | Where-Object { $_.tipo -eq 'dividends' -and $_.estado -eq 'error' }).Count -eq 3 -and ([IO.File]::ReadAllText("$p1\bluechip-board.html")).Contains('id="tbl-div"'))
    $hist1 = [IO.File]::ReadAllText("$p1\noticias-historico.json") | ConvertFrom-Json
    Check 'news history: no file yet, a new empty one is written at the end (with a notice); its start is in the data file' ($o -match 'News history file not found' -and $hist1.versao -eq 1 -and @($hist1.noticias).Count -eq 0 -and $hist1.inicio -and $d.historicoNoticias.inicio -and $d.historicoNoticias.dias -eq 400) $o
    Check 'without -SecEmail: no SEC request, past earnings dates Unavailable (skipped, with the reason)' (@('AAPL', 'NVDA', 'GOOGL' | Where-Object { $d.resultadosSec.$_.estado -eq 'skipped' -and @($d.resultadosSec.$_.resultados).Count -eq 0 -and $d.resultadosSec.$_.erro -match 'EmailSEC' }).Count -eq 3 -and @($d.fontes | Where-Object { $_.nome -like 'SEC: past earnings*' }).Count -eq 0)
    Check 'ECB indicators offline: one error per series, no values invented' (@('BCE_DFR', 'HICP_EA' | Where-Object { $d.macro.$_.estado -eq 'error' -and @($d.macro.$_.pontos).Count -eq 0 }).Count -eq 2 -and @($d.fontes | Where-Object { $_.nome -like 'ECB: *' -and $_.tipo -eq 'macro' -and $_.estado -eq 'error' }).Count -eq 2)
    Check 'without -SecEmail: no fundamentals request, every company skipped with the reason' (@('AAPL', 'NVDA', 'GOOGL' | Where-Object { $d.fundamentais.$_.estado -eq 'skipped' -and @($d.fundamentais.$_.trimestres).Count -eq 0 -and $d.fundamentais.$_.erro -match 'EmailSEC' }).Count -eq 3 -and @($d.fontes | Where-Object { $_.nome -like 'SEC: fundamentals*' }).Count -eq 0)

    Write-Host "Sources down, previous run's data available"
    $p2 = Join-Path $base 'cache'; New-Item -ItemType Directory $p2 | Out-Null
    # dividendos da execução "anterior": AAPL obtidos há 10 dias (reutilizáveis), NVDA há 200 dias (velhos demais)
    $iso = { param($n) (Get-Date).ToUniversalTime().AddDays(-$n).ToString('yyyy-MM-ddTHH:mm:ss.fffffff+00:00') }
    $divAnt = '{"AAPL":{"id":"AAPL","simbolo":"AAPL","moeda":"USD","estado":"ok","fonte":"Yahoo Finance","obtidoEm":"' + (& $iso 10) + '","anualPorAcao":1.08,"ttmPorAcao":1.06,"frequencia":4,"rendimentoPct":0.32,"ultimo":["2026-08-10",0.27],"pagamentos":[["2026-05-11",0.27],["2026-08-10",0.27]],"preco":333.69,"precoData":"' + (& $iso 12) + '","nota":"","erro":""},' +
        '"NVDA":{"id":"NVDA","simbolo":"NVDA","moeda":"USD","estado":"ok","fonte":"Yahoo Finance","obtidoEm":"' + (& $iso 200) + '","anualPorAcao":1,"rendimentoPct":0.4,"ultimo":["2026-03-10",0.25],"pagamentos":[["2026-03-10",0.25]],"preco":233,"nota":"","erro":""}}'
    $ant = [IO.File]::ReadAllText($dadosRef)
    $i = $ant.IndexOf(',"dividendos":'); $j = $ant.IndexOf(',"noticias":'); if ($i -ge 0 -and $j -gt $i) { $ant = $ant.Remove($i, $j - $i) }
    $ant = $ant.Insert($ant.IndexOf(',"noticias":'), ',"dividendos":' + $divAnt)
    [IO.File]::WriteAllText("$p2\bluechip-board-data.json", $ant, (New-Object Text.UTF8Encoding($false)))
    $o = Invoke-Copia $p2 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    $d = [IO.File]::ReadAllText("$p2\bluechip-board-data.json") | ConvertFrom-Json
    $aapl = $d.ativos | Where-Object { $_.id -eq 'AAPL' }
    Check "prices come from the previous run, labelled as such" ($aapl.fonte -like 'previous run*' -and @($aapl.pontos).Count -gt 200) $aapl.fonte
    Check "source table says the previous run's data is shown" (@($d.fontes | Where-Object { $_.estado -like "error (showing previous run's data*" }).Count -ge 5)
    Check 'long history and splits also kept from the previous run' (@($d.historico.NVDA).Count -gt 4000 -and @($d.splits.NVDA).Count -eq 2)
    Check 'each new ETF falls back to its own previous-run prices and history, labelled' (@('EUNK', 'IS3N', 'EUNN' | Where-Object { $i = $_; $a = $d.ativos | Where-Object { $_.id -eq $i }; $a.fonte -like 'previous run*' -and @($a.pontos).Count -gt 200 -and @($d.historico.$i).Count -gt 3000 }).Count -eq 3 -and @($d.fontes | Where-Object { $_.nome -match '\((EUNK|IS3N|EUNN)\)$' -and $_.estado -like "error (showing previous run's data*" }).Count -eq 6)
    Check 'previous-run data is not counted as a working source' (@($d.fontes | Where-Object { $_.estado -like 'ok*' }).Count -eq 0)
    $da = $d.dividendos.AAPL; $quando = if ($da.obtidoEm -is [datetime]) { [DateTimeOffset]$da.obtidoEm } else { [DateTimeOffset]::Parse("$($da.obtidoEm)", [Globalization.CultureInfo]::InvariantCulture) }
    Check "dividends: previous run's data reused, labelled, with its original retrieval date" ($da.estado -eq 'previous run' -and $da.fonte -like 'previous run (retrieved *' -and [math]::Abs(((Get-Date).ToUniversalTime() - $quando.UtcDateTime).TotalDays - 10) -lt 0.1 -and $da.anualPorAcao -eq 1.08 -and @($da.pagamentos).Count -eq 2 -and "$(@($da.pagamentos)[1][0])" -eq '2026-08-10') ($da | ConvertTo-Json -Compress -Depth 5)
    Check 'dividends: previous data older than 180 days is not reused' ($d.dividendos.NVDA.estado -eq 'error' -and $null -eq $d.dividendos.NVDA.anualPorAcao -and $d.dividendos.GOOGL.estado -eq 'error')
    Check "dividends: source table says the previous run's data is shown" (@($d.fontes | Where-Object { $_.nome -eq 'Dividends: Apple' -and $_.estado -like "error (showing previous run's data, retrieved *" }).Count -eq 1)

    Write-Host 'Past earnings dates (SEC) and the news history'
    $p7 = Join-Path $base 'sec'; New-Item -ItemType Directory $p7 | Out-Null
    # dados SEC da execução "anterior": AAPL obtidos há 10 dias (reutilizáveis), NVDA há 200 dias (velhos demais)
    $secAnt = '{"AAPL":{"id":"AAPL","estado":"ok","fonte":"SEC EDGAR","obtidoEm":"' + (& $iso 10) + '","resultados":[{"acc":"0000320193-26-000018","entrega":"2026-07-30","aceite":"2026-07-30T20:30:28Z","horaNY":"2026-07-30 16:30","quando":"after","sessao":"2026-07-31","nota":""}],"ultimoRelatorio":{"form":"10-Q","data":"2026-07-31","periodo":"2026-06-27","acc":"0000320193-26-000020"},"erro":"","nota":""},' +
        '"NVDA":{"id":"NVDA","estado":"ok","fonte":"SEC EDGAR","obtidoEm":"' + (& $iso 200) + '","resultados":[{"acc":"0001045810-26-000019","entrega":"2026-02-25","aceite":"2026-02-25T21:31:25Z","horaNY":"2026-02-25 16:31","quando":"after","sessao":"2026-02-26","nota":""}],"ultimoRelatorio":null,"erro":"","nota":""}}'
    $ant7 = [IO.File]::ReadAllText($dadosRef)
    $i = $ant7.IndexOf(',"historicoNoticias":'); $j = $ant7.IndexOf(',"bolsas":'); if ($i -ge 0 -and $j -gt $i) { $ant7 = $ant7.Remove($i, $j - $i) }
    $ant7 = $ant7.Insert($ant7.IndexOf(',"bolsas":'), ',"resultadosSec":' + $secAnt)
    # fundamentais da execução "anterior": AAPL recolhidos há 10 dias (reutilizáveis), NVDA há 130 dias (mais de 120)
    $fq = '"tags":{"receita":["Revenues"]},"faltam":{},"trimestres":[{"fim":"2026-06-27","inicio":"2026-03-29","ano":"FY2026","q":3,"m":{"receita":{"v":109420000000,"f":"2026-07-31","p":"2026-07-31"}}}],"ttm":null,"erro":"","nota":""'
    $fundAnt7 = '{"AAPL":{"id":"AAPL","estado":"ok","fonte":"SEC XBRL (companyfacts)","obtidoEm":"' + (& $iso 10) + '","relatorio":{"form":"10-Q","data":"2026-07-31","periodo":"2026-06-27","acc":"x"},' + $fq + '},' +
        '"NVDA":{"id":"NVDA","estado":"ok","fonte":"SEC XBRL (companyfacts)","obtidoEm":"' + (& $iso 130) + '","relatorio":null,' + $fq + '}}'
    $ant7 = $ant7.Insert($ant7.IndexOf(',"bolsas":'), ',"fundamentais":' + $fundAnt7)
    # indicadores do BCE da execução "anterior": a taxa de depósito obtida há 10 dias (reutilizável), a inflação há 40 (mais de 30)
    $ant7 = $ant7.Insert($ant7.IndexOf(',"bolsas":'), ',"macro":{"BCE_DFR":{"id":"BCE_DFR","estado":"ok","fonte":"ECB Data Portal","obtidoEm":"' + (& $iso 10) + '","freq":"B","pontos":[["2019-09-18",-0.5],["2026-09-16",2.5]],"erro":""},' +
        '"HICP_EA":{"id":"HICP_EA","estado":"ok","fonte":"ECB Data Portal","obtidoEm":"' + (& $iso 40) + '","freq":"M","pontos":[["2025-12-01",1.9]],"erro":""}}')
    [IO.File]::WriteAllText("$p7\bluechip-board-data.json", $ant7, (New-Object Text.UTF8Encoding($false)))
    # histórico de notícias já existente: uma notícia de há 10 dias e outra de há 401 dias (fora da retenção)
    $hn = { param($t, $n) '{"chave":"' + $t.ToLower() + '","titulo":"' + $t + '","link":"https://example.com/' + $n + '","fonte":"Reuters","data":"' + (& $iso $n) + '","empresas":["AAPL"],"nivel":"red","score":8,"temas":["Earnings"]}' }
    [IO.File]::WriteAllText("$p7\noticias-historico.json", '{"versao":1,"inicio":"' + (& $iso 420) + '","dias":400,"noticias":[' + (& $hn 'Recent story' 10) + ',' + (& $hn 'Old story' 401) + ']}')
    $o = Invoke-Copia $p7 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl } @('-SecEmail', 'test@example.com')
    $raw7 = [IO.File]::ReadAllText("$p7\bluechip-board-data.json"); $d = $raw7 | ConvertFrom-Json
    Check 'run with -SecEmail, SEC blocked: one error per company, the run finishes' ($o -match 'Done in' -and @($d.fontes | Where-Object { $_.nome -like 'SEC: past earnings dates (*' -and $_.estado -eq 'error' -and $_.erro -match 'blocked' }).Count -eq 3) $o
    Check "SEC down: the previous run's past earnings (10 days old) are reused, labelled, dates unchanged" ($d.resultadosSec.AAPL.estado -eq 'previous run' -and $d.resultadosSec.AAPL.fonte -like 'previous run (retrieved *' -and @($d.resultadosSec.AAPL.resultados).Count -eq 1 -and $raw7.Contains('"aceite":"2026-07-30T20:30:28Z","horaNY":"2026-07-30 16:30","quando":"after","sessao":"2026-07-31"') -and $d.resultadosSec.AAPL.ultimoRelatorio.data -eq '2026-07-31') ($d.resultadosSec.AAPL | ConvertTo-Json -Depth 5 -Compress)
    Check 'SEC data over 180 days old is not reused; a company without data is Unavailable (error), nothing invented' ($d.resultadosSec.NVDA.estado -eq 'error' -and @($d.resultadosSec.NVDA.resultados).Count -eq 0 -and $d.resultadosSec.GOOGL.estado -eq 'error' -and @($d.resultadosSec.GOOGL.resultados).Count -eq 0)
    Check "fundamentals, SEC down: the previous run's facts (10 days old) are reused, labelled, values unchanged" ($d.fundamentais.AAPL.estado -eq 'previous run' -and $d.fundamentais.AAPL.fonte -like 'previous run (retrieved *' -and @($d.fundamentais.AAPL.trimestres).Count -eq 1 -and $d.fundamentais.AAPL.trimestres[0].m.receita.v -eq 109420000000 -and @($d.fontes | Where-Object { $_.nome -like 'SEC: fundamentals (*' -and $_.estado -eq 'error' }).Count -eq 3) ($d.fundamentais.AAPL | ConvertTo-Json -Depth 6 -Compress)
    Check 'fundamentals over 120 days old are not reused; a company without them is Unavailable (error), nothing invented' ($d.fundamentais.NVDA.estado -eq 'error' -and @($d.fundamentais.NVDA.trimestres).Count -eq 0 -and $d.fundamentais.GOOGL.estado -eq 'error' -and @($d.fundamentais.GOOGL.trimestres).Count -eq 0)
    Check "ECB down: the previous run's deposit rate (10 days old) is reused, labelled, negative values kept; inflation 40 days old is not" ($d.macro.BCE_DFR.estado -eq 'previous run' -and $d.macro.BCE_DFR.fonte -like 'previous run (retrieved *' -and @($d.macro.BCE_DFR.pontos).Count -eq 2 -and $raw7.Contains('"pontos":[["2019-09-18",-0.5],["2026-09-16",2.5]]') -and $d.macro.HICP_EA.estado -eq 'error' -and @($d.macro.HICP_EA.pontos).Count -eq 0) ($d.macro | ConvertTo-Json -Depth 5 -Compress)
    $hist7 = [IO.File]::ReadAllText("$p7\noticias-historico.json") | ConvertFrom-Json
    $ini7 = if ($hist7.inicio -is [datetime]) { [DateTimeOffset]$hist7.inicio } else { [DateTimeOffset]::Parse("$($hist7.inicio)", [Globalization.CultureInfo]::InvariantCulture) }
    Check 'news history kept: retention 400 days (the 401-day-old story dropped), start moved to the cut' (@($hist7.noticias).Count -eq 1 -and $hist7.noticias[0].titulo -eq 'Recent story' -and [math]::Abs(((Get-Date).ToUniversalTime() - $ini7.UtcDateTime).TotalDays - 400) -lt 0.1 -and @($d.historicoNoticias.noticias).Count -eq 1) ($hist7 | ConvertTo-Json -Depth 5 -Compress)
    Check 'the news history has no personal data and no backup' (-not ([IO.File]::ReadAllText("$p7\noticias-historico.json")).Contains('"backup"'))
    $p8 = Join-Path $base 'histbad'; New-Item -ItemType Directory $p8 | Out-Null
    [IO.File]::WriteAllText("$p8\noticias-historico.json", '{"noticias":[{"titulo"')
    $o = Invoke-Copia $p8 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'news history corrupted: a warning, the old file kept as .bad, a new valid file written' ($o -match 'could not be read' -and (Test-Path "$p8\noticias-historico.json.bad") -and [IO.File]::ReadAllText("$p8\noticias-historico.json.bad") -eq '{"noticias":[{"titulo"' -and ([IO.File]::ReadAllText("$p8\noticias-historico.json") | ConvertFrom-Json).versao -eq 1) $o

    Write-Host 'Backup from Downloads: validated before it replaces anything'
    $p3 = Join-Path $base 'backup'; New-Item -ItemType Directory $p3 | Out-Null
    $velho = '{"app":"Bluechip Board","version":4,"saved":"2026-09-01T10:00:00.000Z","buys":[{"id":"a","a":"AAPL","d":"2026-08-01","q":1,"p":200}],"lots":[],"sales":[],"deleted":{}}'
    [IO.File]::WriteAllText("$p3\bluechip-board-backup.json", $velho)
    [IO.File]::WriteAllText("$dl\bluechip-board-backup (1).json", '{"app":"Bluechip Board","version":4,"saved":"2026-10-01T10:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{}}')
    [IO.File]::WriteAllText("$dl\bluechip-board-backup (2).json", '{ this is not json')
    (Get-Item "$dl\bluechip-board-backup (2).json").LastWriteTime = (Get-Date).AddMinutes(5)   # o inválido é o mais recente
    $o = Invoke-Copia $p3 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'invalid file in Downloads: warning, left where it was' ($o -match 'not a valid Bluechip Board backup: it was left there' -and (Test-Path "$dl\bluechip-board-backup (2).json"))
    Check 'newer valid file moved into the project folder' ((-not (Test-Path "$dl\bluechip-board-backup (1).json")) -and ([IO.File]::ReadAllText("$p3\bluechip-board-backup.json") -match '2026-10-01'))
    Check 'previous backup kept as .previous.json' ((Test-Path "$p3\bluechip-board-backup.previous.json") -and ([IO.File]::ReadAllText("$p3\bluechip-board-backup.previous.json") -match '2026-09-01'))
    Remove-Item "$dl\*" -Force
    [IO.File]::WriteAllText("$dl\bluechip-board-backup.json", '{"app":"Bluechip Board","version":4,"saved":"2026-08-01T10:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{}}')
    (Get-Item "$dl\bluechip-board-backup.json").LastWriteTime = (Get-Date).AddMinutes(10)   # ficheiro recente, mas com dados antigos
    $o = Invoke-Copia $p3 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'older backup (by its saved date) does not replace a newer one' ((Test-Path "$dl\bluechip-board-backup.json") -and ([IO.File]::ReadAllText("$p3\bluechip-board-backup.json") -match '2026-10-01'))
    # um backup com a data de gravação no futuro (relógio errado) nunca decide: fica nas Transferências, com um aviso
    Remove-Item "$dl\*" -Force
    [IO.File]::WriteAllText("$dl\bluechip-board-backup (5).json", '{"app":"Bluechip Board","version":4,"saved":"2099-01-01T10:00:00.000Z","buys":[{"id":"fut","a":"AAPL","d":"2026-08-01","q":9,"p":1}],"lots":[],"sales":[],"deleted":{}}')
    $o = Invoke-Copia $p3 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'backup in Downloads saved "in the future": left there with a warning, the project backup unchanged' (($o -replace '\s+', ' ') -match 'in the future: it was left there' -and (Test-Path "$dl\bluechip-board-backup (5).json") -and ([IO.File]::ReadAllText("$p3\bluechip-board-backup.json") -match '2026-10-01')) $o
    Remove-Item "$dl\*" -Force
    $p3f = Join-Path $base 'backupfuture'; New-Item -ItemType Directory $p3f | Out-Null
    [IO.File]::WriteAllText("$p3f\bluechip-board-backup.json", '{"app":"Bluechip Board","version":4,"saved":"2099-01-01T10:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{}}')
    $o = Invoke-Copia $p3f @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    Check 'project backup saved "in the future": a warning says no Downloads backup can be brought in' (($o -replace '\s+', ' ') -match 'in the future: no backup from Downloads can be newer') $o
    $h = [IO.File]::ReadAllText("$p3\bluechip-board.html"); $ar = [IO.File]::ReadAllText((Get-ChildItem "$p3\Archive" | Select-Object -First 1).FullName)
    Check 'backup embedded in the main site only (not in Archive or the data file)' ($h.Contains('{"backup":{"app"') -and $ar.Contains('{"backup":null,') -and -not ([IO.File]::ReadAllText("$p3\bluechip-board-data.json")).Contains('"backup"'))
    # backup v5 com alvos de alocação (dados pessoais): só no site principal, nunca no Archive nem no ficheiro de dados
    $p3b = Join-Path $base 'backup5'; New-Item -ItemType Directory $p3b | Out-Null
    [IO.File]::WriteAllText("$p3b\bluechip-board-backup.json", '{"app":"Bluechip Board","version":5,"saved":"2026-10-06T09:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{},"targets":{"weights":{"SXR8":60,"EUNK":40},"band":5,"monthly":300,"at":"2026-10-06T09:00:00.000Z"}}')
    $o = Invoke-Copia $p3b @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    $h5 = [IO.File]::ReadAllText("$p3b\bluechip-board.html"); $ar5 = [IO.File]::ReadAllText((Get-ChildItem "$p3b\Archive" | Select-Object -First 1).FullName); $d5 = [IO.File]::ReadAllText("$p3b\bluechip-board-data.json")
    Check 'backup v5: targets embedded in the main site only (not in Archive or the data file)' ($h5.Contains('"targets":{"weights":{"SXR8":60,"EUNK":40}') -and -not $ar5.Contains('"targets":{') -and -not $d5.Contains('"targets"')) $o
    # política de investimento e notas por entrada (textos pessoais): idem
    $p3c = Join-Path $base 'backup5pol'; New-Item -ItemType Directory $p3c | Out-Null
    [IO.File]::WriteAllText("$p3c\bluechip-board-backup.json", '{"app":"Bluechip Board","version":5,"saved":"2026-10-06T09:00:00.000Z","buys":[{"id":"n1","a":"SXR8","d":"2026-01-05","q":1,"p":600}],"lots":[],"sales":[],"deleted":{},"policy":{"horizon":"PRIVATE-HORIZON-TEXT","allocation":"","monthly":"","drop20":"PRIVATE-RULE-20","drop30":"","sell":"","at":"2026-10-06T09:00:00.000Z"},"notes":{"n1":{"t":"PRIVATE-NOTE-TEXT","at":"2026-10-06T09:00:00.000Z"}},"fees":{"n1":{"v":1234.56,"at":"2026-10-06T09:00:00.000Z"}}}')
    $o = Invoke-Copia $p3c @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
    $h6 = [IO.File]::ReadAllText("$p3c\bluechip-board.html"); $ar6 = [IO.File]::ReadAllText((Get-ChildItem "$p3c\Archive" | Select-Object -First 1).FullName); $d6 = [IO.File]::ReadAllText("$p3c\bluechip-board-data.json")
    Check 'backup v5: fees embedded in the main site only (not in Archive or the data file)' ($h6.Contains('"fees":{"n1":{"v":1234.56') -and -not ($ar6 + $d6).Contains('1234.56') -and -not $d6.Contains('"fees"'))
    Check 'backup v5: policy and notes embedded in the main site only (not in Archive or the data file)' ($h6.Contains('PRIVATE-HORIZON-TEXT') -and $h6.Contains('PRIVATE-NOTE-TEXT') -and -not ($ar6 + $d6).Contains('PRIVATE-') -and -not $d6.Contains('"policy"') -and -not $d6.Contains('"notes"')) $o

    Write-Host 'Archive: at most 30 copies and about 300 MB in total'
    $p12 = Join-Path $base 'archive'; New-Item -ItemType Directory "$p12\Archive" -Force | Out-Null
    for ($k = 1; $k -le 8; $k++) { $fa = "$p12\Archive\bluechip-board-2026-09-0$($k)_0800.html"; [IO.File]::WriteAllText($fa, ('x' * 2048)); (Get-Item $fa).LastWriteTime = (Get-Date).AddDays(-20 + $k) }
    $dl12 = Join-Path $base 'Downloads-arq'; New-Item -ItemType Directory $dl12 | Out-Null
    $o = Invoke-Copia $p12 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl12; BB_TEST_ARQ_MAX = '1' }
    $restam = @(Get-ChildItem "$p12\Archive" -Filter 'bluechip-board-2*.html' | Sort-Object LastWriteTime -Descending)
    Check 'Archive over its size limit: the oldest copies go, the 5 most recent stay (the new one included)' ($restam.Count -eq 5 -and $restam[0].Length -gt 2048 -and -not (Test-Path "$p12\Archive\bluechip-board-2026-09-01_0800.html") -and (Test-Path "$p12\Archive\bluechip-board-2026-09-08_0800.html")) (@($restam | ForEach-Object { $_.Name }) -join ', ')

    Write-Host 'A run that fails halfway'
    $antesHtml = [IO.File]::ReadAllText("$p3\bluechip-board.html"); $antesV = '{"abc":"2026-10-01T00:00:00.0000000+00:00"}'
    [IO.File]::WriteAllText("$p3\vistos.json", $antesV)
    $antesH = [IO.File]::ReadAllText("$p3\noticias-historico.json")
    $o = Invoke-Copia $p3 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl; BB_TEST_FAIL = '1' }
    Check 'the failure is reported' ($o -match 'simulated failure')
    Check 'previous site left untouched' ([IO.File]::ReadAllText("$p3\bluechip-board.html") -eq $antesHtml)
    Check 'vistos.json not updated by the failed run' ([IO.File]::ReadAllText("$p3\vistos.json") -eq $antesV)
    Check 'news history not updated by the failed run' ([IO.File]::ReadAllText("$p3\noticias-historico.json") -eq $antesH)
    $p9 = Join-Path $base 'failfresh'; New-Item -ItemType Directory $p9 | Out-Null
    $dl9 = Join-Path $base 'Downloads-empty'; New-Item -ItemType Directory $dl9 | Out-Null
    $o = Invoke-Copia $p9 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl9; BB_TEST_FAIL = '1' } @('-SecEmail', 'test@example.com')
    Check 'a first run that fails halfway writes no file (no news history, no data, no site)' ($o -match 'simulated failure' -and @(Get-ChildItem $p9 -File).Count -eq 0) (@(Get-ChildItem $p9 -File | ForEach-Object Name) -join ', ')

    Write-Host 'Calendar: automatic earnings dates (Nasdaq)'
    $o0 = Invoke-Copia (Join-Path $base 'calbad') @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl; BB_TEST_CAL_BAD = '1' }
    Check 'a manual event with a malformed date is skipped with a warning, the run continues' ($o0 -match 'Calendar event ignored' -and $o0 -match 'Done in') $o0
    $dia = { param($n) (Get-Date).ToUniversalTime().Date.AddDays($n).ToString('yyyy-MM-dd') }
    $cal = @(@{ d = (& $dia 20); e = 'AAPL'; ev = 'Q4 earnings'; imp = 'High'; st = 'C' }, @{ d = (& $dia 25); e = 'GOOGL'; ev = 'Q3 earnings'; imp = 'High'; st = 'E' }, @{ d = (& $dia 30); e = 'NVDA'; ev = 'Q3 earnings'; imp = 'High'; st = 'E' }) | ConvertTo-Json -Compress
    $nas = @(@{ e = 'AAPL'; d = (& $dia 22); st = 'E' }, @{ e = 'GOOGL'; d = (& $dia 27); st = 'E' }, @{ e = 'NVDA'; d = (& $dia 31); st = 'C' }) | ConvertTo-Json -Compress
    $p6 = Join-Path $base 'cal'; New-Item -ItemType Directory $p6 | Out-Null
    $o = Invoke-Copia $p6 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl; BB_TEST_CAL = $cal; BB_TEST_NASDAQ = $nas }
    if (-not (Test-Path "$p6\bluechip-board-data.json")) { Write-Host $o }
    $ev = @(([IO.File]::ReadAllText("$p6\bluechip-board-data.json") | ConvertFrom-Json).calendario)
    $ap = $ev | Where-Object { $_.e -eq 'AAPL' }; $go = $ev | Where-Object { $_.e -eq 'GOOGL' }; $nv = $ev | Where-Object { $_.e -eq 'NVDA' }
    Check 'confirmed manual date is not replaced by a Nasdaq estimate' ("$($ap.d)" -eq (& $dia 20) -and $ap.st -eq 'C' -and -not $ap.auto) ($ap | ConvertTo-Json -Compress)
    Check 'estimated manual date follows the Nasdaq date' ("$($go.d)" -eq (& $dia 27) -and $go.auto -eq 'Nasdaq') ($go | ConvertTo-Json -Compress)
    Check 'a confirmed Nasdaq date replaces an estimated one' ("$($nv.d)" -eq (& $dia 31) -and $nv.st -eq 'C') ($nv | ConvertTo-Json -Compress)

    Write-Host 'Second instance while one is running'
    $m = New-Object System.Threading.Mutex($false, 'Local\BluechipBoard'); [void]$m.WaitOne(0)
    try {
        $p4 = Join-Path $base 'mutex'; New-Item -ItemType Directory $p4 | Out-Null
        $o = Invoke-Copia $p4 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl }
        Check 'second instance refuses to run and writes nothing' ($o -match 'already running' -and -not (Test-Path "$p4\bluechip-board.html"))
    } finally { $m.ReleaseMutex(); $m.Dispose() }

    Write-Host 'Scheduled task (dry run: Register-ScheduledTask is replaced, nothing is registered)'
    $pasta5 = Join-Path $base 'folder with spaces\'
    # o módulo ScheduledTasks tem de ser carregado antes do substituto: ao carregar sozinho, voltaria a pôr o verdadeiro
    $cmd = "Import-Module ScheduledTasks; function Register-ScheduledTask { param(`$TaskName, `$Action, `$Trigger, `$Settings, `$Description, [switch]`$Force) [IO.File]::WriteAllText('$base\task.json', (@{ args = `$Action.Arguments; exe = `$Action.Execute; bat = `$Settings.DisallowStartIfOnBatteries; stop = `$Settings.StopIfGoingOnBatteries; avail = `$Settings.StartWhenAvailable; net = `$Settings.RunOnlyIfNetworkAvailable; desc = `$Description } | ConvertTo-Json)) }; & '$copia' -ScheduleDaily -Time 07:45 -Folder '$pasta5' -SecEmail 'test@example.com'"
    $ErrorActionPreference = 'Continue'; $o = & $Shell -NoProfile -ExecutionPolicy Bypass -Command $cmd 2>&1 | Out-String; $ErrorActionPreference = 'Stop'
    $tk = Get-Content "$base\task.json" -Raw | ConvertFrom-Json
    Check 'task runs on battery and is not stopped when unplugged' (-not $tk.bat -and -not $tk.stop) ($tk | ConvertTo-Json -Compress)
    Check 'task runs late if the PC was off, only with network' ($tk.avail -and $tk.net)
    Check 'folder ending in a backslash is quoted safely' ($tk.args -match [regex]::Escape("-Pasta `"$($pasta5)\`"")) $tk.args
    Check 'hidden window, no browser opened, English description' ($tk.args -match '-WindowStyle Hidden' -and $tk.args -match '-NaoAbrir' -and $tk.desc -match '^Bluechip Board: news and prices')
    Check 'reported with the chosen time' ($o -match 'scheduled every day at 07:45')
    Check 'the SEC e-mail is never stored in the task''s arguments (visible to anyone who lists the tasks); a warning says where the runs read it' ($tk.args -notmatch 'example\.com' -and $tk.args -notmatch 'EmailSEC' -and ($o -replace '\s+', ' ') -match 'does not store the e-mail given with -SecEmail') "$($tk.args) | $o"

    Write-Host 'SEC e-mail from the local configuration (as the scheduled task and the launcher use it)'
    # bluechip-board.config.json ao lado da cópia do script: a execução sem -SecEmail usa as fontes da SEC (aqui bloqueadas)
    [IO.File]::WriteAllText((Join-Path $base 'bluechip-board.config.json'), '{ "secEmail": "cfg@example.com" }')
    try {
        $p10 = Join-Path $base 'cfgmail'; New-Item -ItemType Directory $p10 | Out-Null
        $o = Invoke-Copia $p10 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl9 }
        $d10 = [IO.File]::ReadAllText("$p10\bluechip-board-data.json") | ConvertFrom-Json
        Check 'no -SecEmail, e-mail in bluechip-board.config.json: the SEC sources are requested (blocked here: errors, nothing invented)' (@($d10.fontes | Where-Object { $_.nome -like 'SEC: past earnings dates*' -and $_.estado -eq 'error' -and $_.erro -match 'blocked by test' }).Count -eq 3 -and -not ([IO.File]::ReadAllText("$p10\bluechip-board.html")).Contains('cfg@example.com') -and -not ([IO.File]::ReadAllText("$p10\bluechip-board-data.json")).Contains('cfg@example.com')) $o
        [IO.File]::WriteAllText((Join-Path $base 'bluechip-board.config.json'), '{ "secEmail": "your@email.com" }')
        $p11 = Join-Path $base 'cfgplaceholder'; New-Item -ItemType Directory $p11 | Out-Null
        $o = Invoke-Copia $p11 @{ BB_TEST_BLOCK = '.'; BB_TEST_DOWNLOADS = $dl9 }
        $d11 = [IO.File]::ReadAllText("$p11\bluechip-board-data.json") | ConvertFrom-Json
        Check 'the example placeholder e-mail: a warning, and the SEC sources are skipped' (($o -replace '\s+', ' ') -match 'is not valid: running without the SEC sources' -and $d11.resultadosSec.AAPL.estado -eq 'skipped') $o
    } finally { Remove-Item (Join-Path $base 'bluechip-board.config.json') -Force -ErrorAction SilentlyContinue }
    Check 'nothing was registered in Task Scheduler' ($null -eq (Get-ScheduledTask -TaskName $tarefaTeste -ErrorAction SilentlyContinue))
} finally {
    # proteção: se, apesar de tudo, a tarefa de teste tiver sido registada, é removida
    if (Get-ScheduledTask -TaskName $tarefaTeste -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $tarefaTeste -Confirm:$false }
    Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ''
if ($Script:Falhas) { Write-Host "$($Script:Falhas) of $($Script:Total) checks FAILED" -ForegroundColor Red; exit 1 }
Write-Host "All $($Script:Total) checks passed" -ForegroundColor Green
