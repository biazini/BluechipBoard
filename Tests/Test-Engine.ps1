<#
.SYNOPSIS
    Offline tests for the PowerShell side of Bluechip-Board.ps1 (no internet, no files changed in the project).
.DESCRIPTION
    Loads the configuration and the functions of the script without running it, replaces Get-Url with canned responses
    and checks parsing, classification, duplicate grouping, price validation, stock splits, the dividend collector, the
    halving estimate and the backup helpers. Works in Windows PowerShell 5.1 and PowerShell 7.

        powershell -ExecutionPolicy Bypass -File .\Tests\Test-Engine.ps1
        pwsh -File .\Tests\Test-Engine.ps1
#>
$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path $PSScriptRoot) 'Bluechip-Board.ps1'
$texto = [IO.File]::ReadAllText($scriptPath)

# ---- carrega a configuração e as funções, sem executar o script ----
$Dias = 7; $Pasta = $env:TEMP
$Script:UA = 'test'; $Script:Inv = [Globalization.CultureInfo]::InvariantCulture
$Script:Agora = [DateTimeOffset]::new(2026, 10, 4, 7, 0, 0, [TimeSpan]::Zero)
$Script:Fontes = New-Object System.Collections.Generic.List[object]
$ini = $texto.IndexOf('# 1. CONFIGURA'); $fim = $texto.IndexOf('# 3. MODELO DO SITE')
. ([scriptblock]::Create($texto.Substring($ini, $fim - $ini)))
$ast = [System.Management.Automation.Language.Parser]::ParseInput($texto, [ref]$null, [ref]$null)
foreach ($nome in 'Get-PontosGuardados', 'Get-InstanteBackup', 'Read-Backup', 'Write-Atomico') {
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $nome }, $true) | Select-Object -First 1
    . ([scriptblock]::Create($fn.Extent.Text))
}
$utf8 = New-Object System.Text.UTF8Encoding($false)

$Script:Falhas = 0; $Script:Total = 0
function Check([string]$Nome, [bool]$Cond, $Detalhe = '') {
    $Script:Total++
    if ($Cond) { Write-Host "  ok   $Nome" -ForegroundColor DarkGreen }
    else { $Script:Falhas++; Write-Host "  FAIL $Nome  $Detalhe" -ForegroundColor Red }
}
function Noticia([string]$Titulo, [string]$Feed = 'Google News: test', [string]$Dica = '', [string]$Fonte = 'Reuters', [string]$Exigir = '') {
    Measure-Noticia ([pscustomobject]@{ titulo = $Titulo; link = 'https://example.com'; fonte = $Fonte; dominio = 'example.com'; data = $Script:Agora; feed = $Feed; dica = $Dica; exigir = $Exigir; extraTema = ''; extraPeso = 0 })
}
Write-Host "PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Cyan

Write-Host 'Dates and decoding'
Check 'RFC 822 date with GMT' ((ConvertTo-Data 'Sun, 04 Oct 2026 07:00:00 GMT').UtcDateTime -eq [datetime]'2026-10-04T07:00:00')
Check 'date with CEST is converted to UTC' ((ConvertTo-Data 'Sat, 03 Oct 2026 10:00:00 CEST').UtcDateTime -eq [datetime]'2026-10-03T08:00:00')
Check 'date with EDT is converted to UTC' ((ConvertTo-Data 'Fri, 02 Oct 2026 16:05:00 EDT').UtcDateTime -eq [datetime]'2026-10-02T20:05:00')
Check 'invalid date gives $null' ($null -eq (ConvertTo-Data 'not a date'))
$latin = [Text.Encoding]::GetEncoding('iso-8859-1').GetBytes('<?xml version="1.0" encoding="ISO-8859-1"?><a>ação</a>')
Check 'charset from the XML prolog (ISO-8859-1)' ((ConvertFrom-Bytes $latin '') -match 'ação')
Check 'charset from the HTTP header' ((ConvertFrom-Bytes $latin 'text/xml; charset=ISO-8859-1') -match 'ação')
$bom = [byte[]](0xEF, 0xBB, 0xBF) + [Text.Encoding]::UTF8.GetBytes('<a>é</a>')
Check 'UTF-8 with BOM, BOM removed' ((ConvertFrom-Bytes $bom 'text/xml; charset=iso-8859-1') -eq '<a>é</a>')

Write-Host 'Safe XML'
$x = ConvertTo-XmlSeguro '<rss><channel><item><title>T</title></item></channel></rss>'
Check 'RSS parses into an XmlDocument' ($x -is [System.Xml.XmlDocument] -and $x.rss.channel.item.title -eq 'T')
$bomba = '<?xml version="1.0"?><!DOCTYPE l [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;"><!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">]><rss><channel><item><title>&c;</title></item></channel></rss>'
$expandiu = $true; try { $y = ConvertTo-XmlSeguro $bomba; $expandiu = ("$($y.rss.channel.item.title)".Length -ge 1000) } catch { $expandiu = $false }
Check 'DTD entities are not expanded' (-not $expandiu)

Write-Host 'Classification'
$n = Noticia 'BlackRock iShares Bitcoin Trust sees record $1B outflows' 'CoinDesk' '' 'CoinDesk' 'BTC'
Check 'iShares Bitcoin Trust is Bitcoin, not SXR8' ($n -and $n.empresas -contains 'BTC' -and $n.empresas -notcontains 'SXR8')
$n = Noticia 'iShares Core S&P 500 UCITS ETF tops $100bn' 'Google News: ETF SXR8 / CSPX' 'SXR8'
Check 'iShares Core S&P 500 is still SXR8' ($n.empresas -contains 'SXR8')
$n = Noticia 'Big Apple real estate slump deepens' 'Yahoo Finance: S&P 500' 'MKT'
Check '"Big Apple" is not Apple' ($n.empresas -notcontains 'AAPL')
$n = Noticia "China's AI agents can lie and scheme - just like their US rivals" 'Google News: Apple (products and regulation)' 'AAPL'
Check 'search-feed story without keyword: weak match, at most moderate' ($n.empresas -contains 'AAPL' -and $n.soFeed -and $n.nivel -in 'yellow', 'white') "nivel=$($n.nivel)"
$n = Noticia 'Introducing the next generation of Siri' 'Apple Newsroom' 'AAPL' 'Apple Newsroom'
Check 'primary feed keeps the full assignment' ($n.empresas -contains 'AAPL' -and -not $n.soFeed)
$n = Noticia "Nvidia's new AI platform, Nor'easter flight delays, NFL's drone focus and more in Morning Squawk" 'Google News: NVIDIA (stock)' 'NVDA' 'CNBC'
Check '"flight delays" is not a severe event' (-not $n.severo -and $n.nivel -ne 'red') "nivel=$($n.nivel)"
$n = Noticia 'Nvidia delays Rubin launch to 2027'
Check '"delays launch" is severe' ($n.severo)
$n = Noticia "Fed's Barr says more rate hikes likely to be needed to curb inflation" 'Google News: Fed, inflation and rates' 'MKT'
Check '"curb inflation" is not China and exports' ($n.temas -notcontains 'China and exports' -and -not $n.severo)
$n = Noticia 'California tech CEO arrested for smuggling $300 million in Nvidia AI chips to China'
Check 'smuggling arrest is material' ($n.nivel -eq 'red') "nivel=$($n.nivel) score=$($n.score)"
$n = Noticia "Watch Nvidia's AI Chips Reach China Despite US Curbs" 'Google News: NVIDIA (chips, China, AI)' 'NVDA' 'Bloomberg.com'
Check '"US curbs" is severe' ($n.severo)
$n = Noticia 'Mica mining firm reports quarterly profit' 'Google News: Bitcoin (market)' '' 'Reuters' 'BTC'
Check 'mineral "mica" is not Bitcoin' ($null -eq $n)
$n = Noticia 'EU grants MiCA licence to Coinbase' 'Google News: Bitcoin (regulation and adoption)' '' 'Reuters' 'BTC'
Check 'MiCA (EU rules) is Bitcoin/crypto' ($n -and $n.empresas -contains 'BTC')
$n = Noticia 'Goldman Sachs pushes Fed rate hike forecast to December' 'Google News: Fed, inflation and rates' 'MKT'
Check '"rate forecast" is not Earnings' ($n.temas -notcontains 'Earnings')
$n = Noticia 'Apple Q4 revenue beats estimates on record iPhone demand'
Check 'real earnings story keeps the Earnings theme' ($n.temas -contains 'Earnings')
$n = Noticia 'Bitcoin treasury company Strategy buys more BTC' 'Google News: Bitcoin (market)' '' 'Reuters' 'BTC'
Check '"Bitcoin treasury company" is not Macro and rates' ($n.temas -notcontains 'Macro and rates' -and $n.temas -contains 'Institutional adoption')
$n = Noticia 'Nvidia stock rises as AI chip demand from China grows'
Check 'three ordinary themes no longer reach material' ($n.nivel -ne 'red') "score=$($n.score)"

Write-Host 'Top holdings of the ETFs (news aliases)'
$n = Noticia 'Microsoft shares plunge after EU antitrust fine'
Check 'alias only: the story is linked to the fund (SXR8), marked via top holding' ($n -and $n.empresas -contains 'SXR8' -and @($n.viaPosicao) -contains 'Microsoft') "empresas=$(if ($n) { $n.empresas -join ',' })"
Check 'alias only: at most moderate, even with a severe event from leading press' ($n.nivel -in 'yellow', 'white' -and $n.severo -and $n.score -ge 7) "nivel=$($n.nivel) score=$($n.score)"
$a = Noticia 'Apple shares plunge after EU antitrust fine'
Check 'alias only: no company bonus (one point less than the same story about Apple)' ($n.score -eq $a.score - 1) "$($n.score) vs $($a.score)"
$n = Noticia 'TSMC and ASML report record orders'
Check 'two funds through their holdings (IS3N and EUNK)' ($n -and $n.empresas -contains 'IS3N' -and $n.empresas -contains 'EUNK' -and @($n.viaPosicao).Count -eq 2) "empresas=$(if ($n) { $n.empresas -join ',' })"
$n = Noticia 'Apple and Microsoft sign AI deal'
Check 'a direct keyword wins: Apple, not via a holding' ($n.empresas -contains 'AAPL' -and $n.empresas -notcontains 'SXR8' -and -not @($n.viaPosicao).Count)
$n = Noticia 'Our new partnership with Samsung on XR headsets' 'Google Blog' 'GOOGL' 'Google Blog'
Check 'a primary feed with its own asset keeps the assignment' ($n.empresas -contains 'GOOGL' -and $n.empresas -notcontains 'IS3N' -and -not @($n.viaPosicao).Count) "empresas=$($n.empresas -join ',')"
$n = Noticia 'TSMC raises prices for advanced chips' 'Google News: NVIDIA (chips, China, AI)' 'NVDA'
Check 'a search feed with its own asset keeps it, as before (weak feed match, no alias)' ($n.empresas -contains 'NVDA' -and $n.empresas -notcontains 'IS3N' -and $n.soFeed -and -not @($n.viaPosicao).Count) "empresas=$($n.empresas -join ',')"
$n = Noticia 'Toyota recalls 1 million cars' 'Google News: Bitcoin (market)' '' 'Reuters' 'BTC'
Check 'a Bitcoin-only feed still requires Bitcoin' ($null -eq $n)
# cada alias: uma notícia verdadeira e um falso positivo conhecido (sem outra palavra-chave, por isso o falso positivo não entra)
$fp = @(
    @('Microsoft', 'Microsoft raises its cloud spending plans', 'Microsoft Word tips: ten shortcuts to save time'),
    @('Amazon', 'Amazon to cut 14,000 corporate jobs', 'Fires in the Amazon rainforest reach a record'),
    @('Broadcom', 'Broadcom wins a custom chip order', 'Student wins Broadcom MASTERS science award'),
    @('Meta Platforms', 'Meta raises its capex forecast for AI data centres', 'Governo falha a meta do défice'),
    @('Micron Technology', 'Micron forecasts record memory sales', 'New 5 micron filter removes microplastics from tap water'),
    @('Tesla', 'Tesla deliveries beat expectations', 'Nikola Tesla museum reopens in Belgrade'),
    @('AMD', 'AMD unveils new MI400 accelerators', 'New drug for wet AMD approved in Europe'),
    @('ASML', 'ASML orders jump on AI chip demand', 'ASM International shares rise on strong orders'),
    @('HSBC', 'HSBC to cut costs in Asia', 'HSBC SVNS rugby series returns to Hong Kong'),
    @('Roche', 'Roche obesity drug succeeds in late-stage trial', 'La Roche-Posay launches a new sunscreen'),
    @('Banco Santander', 'Santander profit rises on Brazil lending', 'Racing Santander win promotion to La Liga'),
    @('Shell', 'Shell profits fall on weaker gas trading', 'Shell company used to hide funds, prosecutors say'),
    @('Novartis', 'Novartis to buy a US biotech for $12bn', 'Sandoz shares rise after generic drug approvals'),
    @('AstraZeneca', 'AstraZeneca lists its shares in New York', 'Opel Astra recall announced in Germany'),
    @('Nestlé', 'Nestlé names a new chief executive', 'Ferrero recalls Nestle Toll House cookie dough'),
    @('Siemens', 'Siemens raises its outlook on factory automation', 'Siemens Energy raises its outlook'),
    @('SAP', 'SAP cloud backlog beats estimates', 'Maple Sap Season Starts Early in Vermont'),
    @('TSMC', 'TSMC quarterly profit jumps 40%', 'Taiwan semiconductor exports hit a record in September'),
    @('Samsung Electronics', 'Samsung memory profit soars', 'Samsung Heavy Industries wins an LNG carrier order'),
    @('SK Hynix', 'SK Hynix starts mass production of HBM4', 'SK Telecom fined over a data breach'),
    @('Tencent', 'Tencent games revenue beats estimates', 'Tencent Music Entertainment reports subscriber growth'),
    @('MediaTek', 'MediaTek launches a new smartphone chip', 'Media tech start-ups raise new funding'),
    @('Alibaba', 'Alibaba cloud sales accelerate', 'Ali Baba and the Forty Thieves returns to the stage'),
    @('Delta Electronics', 'Delta Electronics expands its power supply plant', 'Delta Air Lines cuts its forecast'),
    @('China Construction Bank', 'China Construction Bank profit edges higher', 'China State Construction wins a $2bn contract'),
    @('Hon Hai (Foxconn)', 'Foxconn sales rise on AI server demand', 'Fox Corporation earnings beat estimates'),
    @('Mitsubishi UFJ', 'Mitsubishi UFJ raises its profit target', 'Mitsubishi Motors recalls 100,000 cars'),
    @('Advantest', 'Advantest lifts its forecast on AI chip testing', 'Advantech posts higher quarterly sales'),
    @('Tokyo Electron', 'Tokyo Electron raises its annual forecast', 'Tokyo Electron Device raises its forecast'),
    @('Toyota', 'Toyota global sales hit a record', 'Toyota Tsusho expands its battery business'),
    @('Hitachi', 'Hitachi wins a rail signalling contract', 'Hitachi Zosen wins a waste plant order'),
    @('Sumitomo Mitsui Financial', 'Sumitomo Mitsui Financial raises its dividend', 'Sumitomo Mitsui Trust Holdings lifts its dividend'),
    @('SoftBank Group', 'SoftBank Group sells its T-Mobile stake', 'SoftBank Corp mobile network outage hits Japan'),
    @('Kioxia', 'Kioxia shares surge on memory prices', 'Toshiba unveils new hard drives'),
    @('Recruit Holdings', 'Recruit Holdings buys an HR software firm', 'Army struggles to recruit new soldiers'),
    @('Sony', 'Sony raises its PlayStation sales target', 'Sony Financial Group shares debut in Tokyo')
)
$todos = @($AliasesPosicoes.Keys | ForEach-Object { @($AliasesPosicoes[$_]) } | ForEach-Object { $_.Nome })
Check 'every alias has a true story and a known false positive below' ($todos.Count -eq $fp.Count -and @($todos | Where-Object { $_ -notin @($fp | ForEach-Object { $_[0] }) }).Count -eq 0) "aliases=$($todos.Count) cases=$($fp.Count)"
foreach ($c in $fp) {
    $sim = Noticia $c[1]; $nao = Noticia $c[2]
    Check "alias $($c[0]): '$($c[1])' matches; '$($c[2])' does not" ($sim -and @($sim.viaPosicao) -contains $c[0] -and $null -eq $nao) "sim=$(if ($sim) { @($sim.viaPosicao) -join ',' }) nao=$(if ($nao) { $nao.empresas -join ',' })"
}
# os níveis das notícias que já eram classificadas não mudam: a mesma lista com e sem os aliases
$casos = @(
    @('BlackRock iShares Bitcoin Trust sees record $1B outflows', 'CoinDesk', '', 'CoinDesk', 'BTC'),
    @('Big Apple real estate slump deepens', 'Yahoo Finance: S&P 500', 'MKT', 'Reuters', ''),
    @("China's AI agents can lie and scheme - just like their US rivals", 'Google News: Apple (products and regulation)', 'AAPL', 'Reuters', ''),
    @('Introducing the next generation of Siri', 'Apple Newsroom', 'AAPL', 'Apple Newsroom', ''),
    @('California tech CEO arrested for smuggling $300 million in Nvidia AI chips to China', 'Google News: test', '', 'Reuters', ''),
    @("Fed's Barr says more rate hikes likely to be needed to curb inflation", 'Google News: Fed, inflation and rates', 'MKT', 'Reuters', ''),
    @('Nvidia stock rises as AI chip demand from China grows', 'Google News: test', '', 'Reuters', ''),
    @('Apple and Microsoft sign AI deal', 'Google News: test', '', 'Reuters', ''),
    @('Nvidia supplier TSMC raises capex', 'Google News: NVIDIA (stock)', 'NVDA', 'CNBC', ''),
    @('Microsoft and Amazon lead Wall Street higher', 'Yahoo Finance: S&P 500', 'MKT', 'Reuters', ''),
    @('European shares slip as ASML falls', 'Google News: Europe (MSCI Europe, STOXX 600)', 'EUNK', 'Reuters', ''),
    @('Federal Reserve issues FOMC statement', 'Federal Reserve (monetary policy)', 'MKT', 'Federal Reserve', '')
)
$com = @(foreach ($c in $casos) { , (Noticia $c[0] $c[1] $c[2] $c[3] $c[4]) })
$guarda = $AliasesPosicoes; $AliasesPosicoes = [ordered]@{}
$sem = @(foreach ($c in $casos) { , (Noticia $c[0] $c[1] $c[2] $c[3] $c[4]) })
$AliasesPosicoes = $guarda
$dif = @(for ($i = 0; $i -lt $casos.Count; $i++) { if ($sem[$i] -and (-not $com[$i] -or $com[$i].nivel -ne $sem[$i].nivel -or $com[$i].score -ne $sem[$i].score -or ($com[$i].empresas -join ',') -ne ($sem[$i].empresas -join ','))) { $casos[$i][0] } })
Check 'stories that already matched keep their assets, level and score' ($dif.Count -eq 0 -and @($sem | Where-Object { $_ }).Count -eq $casos.Count) ($dif -join ' | ')
# lembrete de manutenção: posições do top 10 atual sem alias
$h = { param($s) @($s -split '\|' | ForEach-Object { $x = $_ -split '=', 2; [pscustomobject]@{ t = $x[0]; n = $x[1] } }) }
$p = [ordered]@{
    SXR8 = [pscustomobject]@{ aoVivo = $true; top10 = (& $h 'NVDA=NVIDIA|MSFT=Microsoft|BRKB=Berkshire Hathaway') }
    EUNK = [pscustomobject]@{ aoVivo = $true; top10 = (& $h 'SHEL=Shell|ZZZ1=Roche Holding') }
    EUNN = [pscustomobject]@{ aoVivo = $false; top10 = (& $h '9999=Unknown Corp') }
}
$s = @(Get-PosicoesSemAlias $p)
Check 'reminder: a top 10 holding without an alias is named (SXR8 Berkshire Hathaway)' ($s.Count -eq 1 -and $s[0] -eq 'SXR8 Berkshire Hathaway (BRKB)') ($s -join '; ')
Check 'reminder: no current holdings file (reference weights), no reminder' (@(Get-PosicoesSemAlias ([ordered]@{ EUNN = $p.EUNN; IS3N = [pscustomobject]@{ aoVivo = $false; top10 = @() } })).Count -eq 0)
$real = [ordered]@{
    SXR8 = [pscustomobject]@{ aoVivo = $true; top10 = (& $h 'NVDA=NVIDIA|AAPL=Apple|MSFT=Microsoft|GOOGL=Alphabet (A + C)|AMZN=Amazon|AVGO=Broadcom|META=Meta Platforms|MU=Micron Technology|TSLA=Tesla|AMD=AMD') }
    EUNK = [pscustomobject]@{ aoVivo = $true; top10 = (& $h 'ASML=Asml Holding|HSBA=Hsbc Holdings|ROP=Roche Ps Par Ag|SAN=Banco Santander|SHEL=Shell|NOVN=Novartis Ag|AZN=Astrazeneca|NESN=Nestle Sa|SIE=Siemens N Ag|SAP=Sap') }
    IS3N = [pscustomobject]@{ aoVivo = $true; top10 = (& $h '2330=Taiwan Semiconductor Manufacturing|005930=Samsung Electronics|000660=Sk Hynix|700=Tencent Holdings|2454=Mediatek|9988=Alibaba Group Holding|2308=Delta Electronics|005935=Samsung Electronics Non Voting Pre|939=China Construction Bank Corp H|2317=Hon Hai Precision Industry') }
    EUNN = [pscustomobject]@{ aoVivo = $true; top10 = (& $h '8306=Mitsubishi Ufj Financial Group|6857=Advantest|8035=Tokyo Electron|7203=Toyota Motor|6501=Hitachi|8316=Sumitomo Mitsui Financial Group|9984=Softbank Group|285A=Kioxia Holdings|6098=Recruit Holdings|6758=Sony Group') }
}
Check 'the aliases cover the top 10 of each fund on 5 Oct 2026 (no reminder today)' (@(Get-PosicoesSemAlias $real).Count -eq 0) (@(Get-PosicoesSemAlias $real) -join '; ')

Write-Host 'Duplicate grouping'
# Rare words weigh more than common ones, so each test case is mixed with unrelated stories, as in a real run
$Fundo = @('Apple unveils new MacBook Air with M5 chip', 'Nvidia shares rise after strong data center sales', 'Alphabet Waymo expands robotaxi service to Miami',
    'Google Gemini app adds video editing features', 'Bitcoin price climbs above 90,000 dollars', 'Fed minutes show officials split on December cut',
    'S&P 500 futures edge higher ahead of jobs report', 'Apple iPhone sales in India hit quarterly record', 'Nvidia CEO Jensen Huang to visit Taiwan suppliers',
    'YouTube launches cheaper premium plan in Europe', 'Treasury yields fall as inflation cools', 'Bitcoin ETF inflows reach three week high',
    'Google Cloud signs AI deal with Deutsche Bank', 'Apple Vision Pro gets new enterprise apps', 'Nvidia Blackwell servers ship to Microsoft',
    'ECB keeps deposit rate unchanged at 2 percent', 'Dollar weakens against euro after soft data', 'Alphabet DeepMind releases new weather model',
    'Microsoft and Nvidia expand partnership on AI factories', 'Apple TV price increase announced for subscribers')
function Grupo($Itens) {
    $k = 0
    $L = @(foreach ($i in $Itens) { $m = Noticia $i.t; $m.data = $i.d; $m.score = $i.s; $m.chave = "case$k"; $k++; $m })
    $F = @(foreach ($t in $Fundo) { $m = Noticia $t; if ($m) { $m.data = '2026-09-25T09:00:00.0000000+00:00'; $m.score = 1; $m.chave = "bg$k"; $k++; $m } })
    , @(Join-NoticiasDuplicadas ($L + $F) | Where-Object { $_.chave -like 'case*' })
}
$d0 = '2026-10-01T10:00:00.0000000+00:00'
$r = Grupo @(
    @{ t = 'California tech executive arrested for smuggling $300M in Nvidia AI chips to China'; d = $d0; s = 9 },
    @{ t = 'California man charged with smuggling $300 million Nvidia AI chips to China'; d = '2026-10-01T15:00:00.0000000+00:00'; s = 8 })
Check 'same event, different wording: grouped' ($r.Count -eq 1 -and @($r[0].outras).Count -eq 1) "groups=$($r.Count)"
$r = Grupo @(
    @{ t = 'California man charged with smuggling $300 million Nvidia AI chips to China'; d = $d0; s = 9 },
    @{ t = 'California man charged with smuggling $500 million Nvidia AI chips to China'; d = $d0; s = 8 })
Check 'conflicting amounts: not grouped' ($r.Count -eq 2)
$r = Grupo @(
    @{ t = 'EU regulators fine Apple over App Store rules in antitrust case'; d = '2026-09-27T10:00:00.0000000+00:00'; s = 9 },
    @{ t = 'Apple fined by EU regulators over App Store rules antitrust case'; d = '2026-09-29T10:00:00.0000000+00:00'; s = 8 },
    @{ t = 'Apple appeals EU regulators App Store rules antitrust fine case'; d = '2026-10-01T12:00:00.0000000+00:00'; s = 7 })
$span = 0; foreach ($g in $r) { $ds = @(@($g.data) + @($g.outras | ForEach-Object { $_.data }) | ForEach-Object { [DateTimeOffset]::Parse($_, $Script:Inv) }); if ($ds.Count -gt 1) { $s = (($ds | Measure-Object -Maximum UtcTicks).Maximum - ($ds | Measure-Object -Minimum UtcTicks).Minimum) / 864e9; if ($s -gt $span) { $span = $s } } }
Check 'a group never spans more than 3 days' ($span -le 3) "span=$span days, groups=$($r.Count)"

Write-Host 'Prices from Yahoo (canned responses)'
$Script:Resposta = ''
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) if ($Url -match 'stooq') { return '<!DOCTYPE html><html></html>' }; if ($Url -match 'kraken') { throw 'offline' }; return $Script:Resposta }
function Yahoo([string]$Simbolo, [string]$Moeda, [double[]]$Precos, [string]$Extra = '') {
    $t0 = 1759224600; $ts = @(); for ($i = 0; $i -lt $Precos.Count; $i++) { $ts += $t0 + 86400 * $i }
    $cl = ($Precos | ForEach-Object { if ([double]::IsNaN($_)) { 'null' } else { $_.ToString($Script:Inv) } }) -join ','
    '{"chart":{"result":[{"meta":{"symbol":"' + $Simbolo + '","currency":"' + $Moeda + '","exchangeTimezoneName":"America/New_York","gmtoffset":-14400,"regularMarketPrice":5,"regularMarketTime":' + ($t0 + 86400 * ($Precos.Count - 1) + 23400) + ',"currentTradingPeriod":{"regular":{"start":1,"end":2}}},"timestamp":[' + ($ts -join ',') + '],' + $Extra + '"indicators":{"quote":[{"close":[' + $cl + ']}]}}],"error":null}}'
}
$Script:Resposta = Yahoo 'AAPL' 'USD' @(1, 2, 3, 4, 5, 6)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = 'aapl.us'; Moeda = 'USD' }
Check 'valid series is accepted' ($s -and $s.pontos.Count -eq 6 -and $s.pontos[0][0] -eq '2025-09-30')
$Script:Resposta = Yahoo 'MSFT' 'USD' @(1, 2, 3, 4, 5, 6)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = 'aapl.us'; Moeda = 'USD' }
Check 'another symbol in the response is rejected' ($null -eq $s -and $Script:Fontes[-1].erro -match 'MSFT')
Check 'Stooq anti-bot page is reported clearly' ($Script:Fontes[-1].erro -match 'web page instead of CSV')
$Script:Resposta = Yahoo 'AAPL' 'EUR' @(1, 2, 3, 4, 5, 6)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
Check 'unexpected currency is rejected' ($null -eq $s)
$Script:Resposta = Yahoo 'AAPL' 'USD' @(1, -2, 0, 4, 5, 6, 7, 8)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
Check 'zero and negative prices are dropped' ($s -and $s.pontos.Count -eq 6 -and $Script:Fontes[-1].erro -match '2 invalid')
$Script:Resposta = Yahoo 'NVDA' 'USD' @(1, 2, 3, 4, 5, 6) '"events":{"splits":{"1718026200":{"date":1718026200,"numerator":10,"denominator":1},"1626787800":{"date":1626787800,"numerator":4,"denominator":1}}},'
$s = Get-Serie @{ Id = 'NVDA'; Nome = 'NVIDIA'; Yahoo = 'NVDA'; Stooq = ''; Moeda = 'USD' } -Desde 1262304000
Check 'split events are read and sorted' ($s.splits.Count -eq 2 -and $s.splits[0][0] -eq '2021-07-20' -and $s.splits[1][1] -eq 10) ($s.splits | ConvertTo-Json -Compress)
Check 'split pairs serialise as plain JSON arrays' ((@{ x = $s.splits } | ConvertTo-Json -Compress -Depth 5) -match '^\{"x":\[\["2021-07-20",4(\.0)?\],\["2024-06-10",10(\.0)?\]\]\}$') (@{ x = $s.splits } | ConvertTo-Json -Compress -Depth 5)

Write-Host 'Dividends from Yahoo (canned responses)'
# resposta do gráfico do Yahoo com eventos de dividendo; cada pagamento é @(dias atrás, valor como texto JSON)
function YahooDiv([string]$Simbolo, [string]$Moeda, [string]$Preco, $Pagamentos) {
    $agora = $Script:Agora.ToUnixTimeSeconds()
    $ev = @(foreach ($p in $Pagamentos) { $t = $agora - [int64]([double]$p[0] * 86400); '"' + $t + '":{"amount":' + $p[1] + ',"date":' + $t + '}' }) -join ','
    '{"chart":{"result":[{"meta":{"symbol":"' + $Simbolo + '","currency":"' + $Moeda + '","regularMarketPrice":' + $Preco + ',"regularMarketTime":' + $agora + ',"gmtoffset":-14400},"timestamp":[1],"events":{"dividends":{' + $ev + '}},"indicators":{"quote":[{"close":[1]}]}}],"error":null}}'
}
$aapl = @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Moeda = 'USD' }
$Script:Resposta = YahooDiv 'AAPL' 'USD' '333.69' @(@(420, '0.26'), @(330, '0.26'), @(239, '0.26'), @(147, '0.27'), @(55, '0.27'))
$d = Get-Dividendo $aapl
Check 'quarterly dividends: annual per share = latest × 4' ($d.estado -eq 'ok' -and $d.frequencia -eq 4 -and [math]::Abs($d.anualPorAcao - 1.08) -lt 1e-9) ($d | ConvertTo-Json -Compress)
Check 'trailing 12 months = payments of the last 365 days' ([math]::Abs($d.ttmPorAcao - 1.06) -lt 1e-9) $d.ttmPorAcao
Check 'yield in percentage points from the price in the response (0.3237 = 0.32%)' ([math]::Abs($d.rendimentoPct - [math]::Round(1.08 / 333.69 * 100, 4)) -lt 1e-9) $d.rendimentoPct
Check 'payment pairs serialise as plain JSON arrays' ((@{ x = $d.pagamentos; u = $d.ultimo } | ConvertTo-Json -Compress -Depth 5) -match '"x":\[\["\d{4}-\d{2}-\d{2}",0\.26\]' -and (@{ u = $d.ultimo } | ConvertTo-Json -Compress) -match '^\{"u":\["\d{4}-\d{2}-\d{2}",0\.27\]\}$') (@{ x = $d.pagamentos; u = $d.ultimo } | ConvertTo-Json -Compress -Depth 5)
Check 'source recorded as ok, with retrieval time' ($Script:Fontes[-1].nome -eq 'Dividends: Apple' -and $Script:Fontes[-1].estado -eq 'ok' -and $d.obtidoEm -like '2026-10-04T07:00:00*')
$Script:Resposta = YahooDiv 'AAPL' 'USD' '50' @(@(300, '0.5'), @(210, '0.5'), @(120, '0.5'), @(30, '0.5'))
$d = Get-Dividendo $aapl
Check 'a $0.50 dividend is not read as 50%: $2 a year on $50 = 4%' ($d.rendimentoPct -eq 4 -and $d.anualPorAcao -eq 2) ($d | ConvertTo-Json -Compress)
$Script:Resposta = YahooDiv 'MSFT' 'USD' '300' @(, @(30, '0.5'))
$d = Get-Dividendo $aapl
Check 'another symbol in the response is rejected' ($d.estado -eq 'error' -and $null -eq $d.anualPorAcao -and $Script:Fontes[-1].erro -match 'MSFT')
$Script:Resposta = YahooDiv 'AAPL' 'EUR' '300' @(, @(30, '0.5'))
Check 'unexpected currency is rejected' ((Get-Dividendo $aapl).estado -eq 'error')
$Script:Resposta = (YahooDiv 'AAPL' 'USD' '300' @()).Replace('"dividends":{}', '"dividends":null')
$d = Get-Dividendo $aapl
Check 'no dividends in the response: error, nothing invented' ($d.estado -eq 'error' -and $d.erro -match 'No dividend payments')
$Script:Resposta = YahooDiv 'AAPL' 'USD' '300' @(@(300, '0'), @(250, '-1'), @(200, 'null'), @(190, '"abc"'), @(150, '100'), @(120, '0.25'), @(30, '0.25'))
$d = Get-Dividendo $aapl
Check 'zero, negative, missing, non-numeric and absurd amounts are ignored and counted' ($d.estado -eq 'ok' -and @($d.pagamentos).Count -eq 2 -and $Script:Fontes[-1].erro -match '^5 invalid') "$(@($d.pagamentos).Count) / $($Script:Fontes[-1].erro)"
$Script:Resposta = YahooDiv 'AAPL' 'USD' '300' @(, @(30, '0.25'))
$d = Get-Dividendo $aapl
Check 'a single payment: no annual figure (frequency unknown)' ($d.estado -eq 'ok' -and $null -eq $d.anualPorAcao -and $d.nota -match 'frequency is unknown')
$Script:Resposta = YahooDiv 'AAPL' 'USD' '300' @(@(700, '0.25'), @(610, '0.25'), @(520, '0.25'), @(430, '0.25'))
$d = Get-Dividendo $aapl
Check 'payments stopped over a year ago: no annual figure (maybe suspended)' ($null -eq $d.anualPorAcao -and $d.nota -match 'suspended')
$Script:Resposta = YahooDiv 'AAPL' 'USD' '100' @(@(90, '20'), @(60, '20'), @(30, '20'))
Check 'implausible yield (240%) is rejected' ((Get-Dividendo $aapl).erro -match 'Implausible dividend yield')
Check 'dividend assets: AAPL, NVDA and GOOGL only (not SXR8 or BTC)' (($AtivosDividendos -join ',') -eq 'AAPL,NVDA,GOOGL')
$guardaUrl = ${function:Get-Url}; function Get-Url { throw 'offline' }
$todos = Get-Dividendos
${function:Get-Url} = $guardaUrl
Check 'source down: one error entry per asset, no exception' (@($todos.Keys) -join ',' -eq 'AAPL,NVDA,GOOGL' -and @($todos.Values | Where-Object { $_.estado -eq 'error' }).Count -eq 3)

Write-Host 'Reaction session in New York time'
$fer = @(@($Bolsas | Where-Object { $_.id -eq 'US' })[0].feriados)
# sessões de teste: dias úteis de 2026 sem os feriados de $Bolsas, menos 1 jun (um fecho que só o histórico conhece)
$sess = [string[]]@(for ($d = [datetime]'2026-01-02'; $d -le [datetime]'2026-10-02'; $d = $d.AddDays(1)) { $iso = $d.ToString('yyyy-MM-dd'); if ($d.DayOfWeek -notin 'Saturday', 'Sunday' -and $iso -notin $fer -and $iso -ne '2026-06-01') { $iso } })
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-30T16:30:28-04:00', $Script:Inv)) $sess $fer
Check 'accepted after 16:00 New York: the next session' ($r.horaNY -eq '2026-07-30 16:30' -and $r.quando -eq 'after' -and $r.sessao -eq '2026-07-31' -and -not $r.nota) ($r | ConvertTo-Json -Compress)
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-31T16:05:00-04:00', $Script:Inv)) $sess $fer
Check 'Friday after the close: Monday' ($r.sessao -eq '2026-08-03') $r.sessao
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-30T12:00:00Z', $Script:Inv)) $sess $fer
Check 'accepted before 09:30 New York (08:00 EDT): that same session' ($r.quando -eq 'before' -and $r.sessao -eq '2026-07-30' -and $r.horaNY -eq '2026-07-30 08:00') ($r | ConvertTo-Json -Compress)
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-30T12:00:00-04:00', $Script:Inv)) $sess $fer
Check 'accepted during the session: that same session, with a note' ($r.quando -eq 'during' -and $r.sessao -eq '2026-07-30' -and $r.nota -match 'during the session') ($r | ConvertTo-Json -Compress)
$r1 = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-03-06T20:30:00Z', $Script:Inv)) $sess $fer
$r2 = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-03-09T20:30:00Z', $Script:Inv)) $sess $fer
Check 'US clock change: 20:30 UTC is 15:30 EST (during) on 6 Mar and 16:30 EDT (after) on 9 Mar' ($r1.horaNY -eq '2026-03-06 15:30' -and $r1.quando -eq 'during' -and $r1.sessao -eq '2026-03-06' -and $r2.horaNY -eq '2026-03-09 16:30' -and $r2.quando -eq 'after' -and $r2.sessao -eq '2026-03-10') "$($r1.horaNY) $($r1.quando) / $($r2.horaNY) $($r2.quando)"
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-02T16:30:00-04:00', $Script:Inv)) $sess $fer
Check 'a holiday in the history is skipped (3 Jul → Monday 6 Jul)' ($r.sessao -eq '2026-07-06') $r.sessao
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-05-29T16:30:00-04:00', $Script:Inv)) $sess $fer
Check 'the price history decides the sessions (a day missing from it is skipped)' ($r.sessao -eq '2026-06-02') $r.sessao
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-11-25T16:30:00-05:00', $Script:Inv)) $sess $fer
Check 'after the end of the history: weekdays without the exchange holidays (Thanksgiving → Friday 27 Nov)' ($r.sessao -eq '2026-11-27') $r.sessao
$r1 = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-30T16:00:00-04:00', $Script:Inv)) $sess $fer
$r2 = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-07-30T09:30:00-04:00', $Script:Inv)) $sess $fer
Check 'boundaries: 16:00 counts as after the close, 09:30 as during the session' ($r1.quando -eq 'after' -and $r2.quando -eq 'during') "$($r1.quando) $($r2.quando)"

Write-Host 'Past earnings dates from the SEC (canned responses)'
# resposta de data.sec.gov/submissions: listas paralelas, como a real; cada linha é @(form, acesso, entrega, período, itens, acceptanceDateTime)
function SecJson([string]$Cik, $L) {
    $c = { param($k) , @(foreach ($x in $L) { $x[$k] }) }
    @{ cik = $Cik; name = 'Test Inc.'; filings = @{ recent = [ordered]@{ accessionNumber = (& $c 1); filingDate = (& $c 2); reportDate = (& $c 3); acceptanceDateTime = (& $c 5); form = (& $c 0); items = (& $c 4) }; files = @() } } | ConvertTo-Json -Depth 5 -Compress
}
# feed Atom de 8-K do EDGAR; cada entrada é @(acesso, updated com fuso, itens)
function SecAtom($E) {
    '<?xml version="1.0" encoding="ISO-8859-1" ?><feed xmlns="http://www.w3.org/2005/Atom"><title>8-K</title>' + (@(foreach ($x in $E) { "<entry><category term=""8-K"" /><content type=""text/xml""><accession-number>$($x[0])</accession-number><filing-date>$($x[1].Substring(0, 10))</filing-date><filing-type>8-K</filing-type><items-desc>$($x[2])</items-desc></content><updated>$($x[1])</updated></entry>" }) -join '') + '</feed>'
}
$Script:SecJ = SecJson '0000320193' @(
    @('10-Q', '0000320193-26-000020', '2026-07-31', '2026-06-27', '', '2026-07-31T14:01:02.000Z'),
    @('8-K', '0000320193-26-000018', '2026-07-30', '2026-07-30', '2.02,9.01', '2026-07-31T00:30:28.000Z'),
    @('8-K', '0001140361-26-015711', '2026-04-20', '2026-04-17', '5.02', '2026-04-21T01:29:51.000Z'),
    @('8-K', '0000320193-26-000011', '2026-04-30', '2026-04-30', '2.02,9.01', '2026-05-01T00:30:41.000Z'),
    @('10-K', '0000320193-25-000079', '2025-10-31', '2025-09-27', '', '2025-10-31T14:01:26.000Z'),
    @('8-K', '0000320193-22-000070', '2022-07-28', '2022-07-28', '2.02,9.01', '2022-07-28T20:30:00.000Z'),
    @('8-K', '0000320193-19-000066', '2019-07-30', '2019-07-30', '2.02,9.01', '2019-07-30T20:30:00.000Z'))
$Script:SecA = SecAtom @(@('0000320193-26-000018', '2026-07-30T16:30:28-04:00', 'items 2.02 and 9.01'), @('0000320193-26-000011', '2026-04-30T16:30:41-04:00', 'items 2.02 and 9.01'), @('0001140361-26-015711', '2026-04-20T17:29:51-04:00', 'item 5.02'))
$Script:SecFalha = ''; $Script:SecPedidos = New-Object System.Collections.Generic.List[string]
$guardaUrl2 = ${function:Get-Url}
# cabeçalhos das entregas no arquivo do EDGAR (…-index-headers.html): número de acesso → ACCEPTANCE-DATETIME (hora de Nova
# Iorque, aaaammddhhmmss) ou o texto inteiro da resposta; sem entrada, a SEC responde 404
$Script:SecCab = @{}
function SecCabecalho([string]$Acc, [string]$Dt) { "<HTML><HEAD><TITLE>SEC EDGAR Submission $Acc</TITLE></HEAD><BODY><PRE>`n&lt;SEC-HEADER&gt;$Acc.hdr.sgml : 20260730`n<ACCEPTANCE-DATETIME>$Dt`n<ACCESSION-NUMBER>$Acc`n<TYPE>8-K`n<ITEMS>2.02</PRE></BODY></HTML>" }
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) $Script:SecPedidos.Add("$UserAgent|$Url"); if ($Url -match '^https://data\.sec\.gov/submissions/CIK\d{10}\.json$') { if ($Script:SecFalha -eq 'json') { throw '503 Service Unavailable' }; return $Script:SecJ }; if ($Url -match 'browse-edgar.*type=8-K.*output=atom') { if ($Script:SecFalha -eq 'atom') { throw 'timeout' }; return $Script:SecA }
    if ($Url -match '^https://www\.sec\.gov/Archives/edgar/data/\d+/\d{18}/(\d{10}-\d{2}-\d{6})-index-headers\.html$') { $acc = $Matches[1]; $c = $Script:SecCab[$acc]; if (-not $c) { throw '404 Not Found' }; if ($c -match '^\d{14}$') { return (SecCabecalho $acc $c) }; return $c }
    throw "unexpected $Url" }
$apple = @{ Id = 'AAPL'; Cik = '0000320193' }
$x = Get-ResultadosSEC $apple 'BluechipBoard/1.0 test@example.com' $sess $fer
$ra = { param($a) @($x.resultados | Where-Object { $_.acc -eq $a })[0] }
Check 'valid: the 8-K with item 2.02 of the last 5 years (not item 5.02, not 2019)' ($x.estado -eq 'ok' -and (@($x.resultados | ForEach-Object { $_.acc }) -join ',') -eq '0000320193-26-000018,0000320193-26-000011,0000320193-22-000070') (@($x.resultados | ForEach-Object { $_.acc }) -join ',')
$a = & $ra '0000320193-26-000018'
Check 'acceptance time from the EDGAR 8-K feed (16:30 New York), not the JSON field (00:30Z)' ($a.aceite -eq '2026-07-30T20:30:28Z' -and $a.horaNY -eq '2026-07-30 16:30' -and $a.quando -eq 'after' -and $a.sessao -eq '2026-07-31' -and $a.entrega -eq '2026-07-30') ($a | ConvertTo-Json -Compress)
$a = & $ra '0000320193-22-000070'
Check 'an 8-K missing from the feed and without a filing header: no time and no session (Unavailable), nothing invented' ($null -eq $a.aceite -and $null -eq $a.sessao -and $a.nota -match 'unavailable' -and $x.nota -match '1 8-K without an acceptance time' -and $x.nota -match 'filing headers: 404') ($a | ConvertTo-Json -Compress)
Check 'latest 10-Q/10-K kept (form, filing date, period)' ($x.ultimoRelatorio.form -eq '10-Q' -and $x.ultimoRelatorio.data -eq '2026-07-31' -and $x.ultimoRelatorio.periodo -eq '2026-06-27') ($x.ultimoRelatorio | ConvertTo-Json -Compress)
Check 'three requests (filings list, 8-K feed, the filing header of the one 8-K missing from the feed), all with the SEC User-Agent' ($Script:SecPedidos.Count -eq 3 -and $Script:SecPedidos[2] -like '*/Archives/edgar/data/320193/000032019322000070/0000320193-22-000070-index-headers.html' -and @($Script:SecPedidos | Where-Object { $_ -notlike 'BluechipBoard/1.0 test@example.com|https://*' }).Count -eq 0) ($Script:SecPedidos -join ' ; ')
Check 'source recorded as ok' ($Script:Fontes[-1].nome -eq 'SEC: past earnings dates (AAPL)' -and $Script:Fontes[-1].estado -eq 'ok' -and $Script:Fontes[-1].itens -eq 3)
Check 'results serialise as plain JSON (dates as text)' (($x | ConvertTo-Json -Depth 5 -Compress) -match '"aceite":"2026-07-30T20:30:28Z","horaNY":"2026-07-30 16:30","quando":"after","sessao":"2026-07-31"') ($x | ConvertTo-Json -Depth 5 -Compress)
$Script:SecJ = SecJson '0001045810' @(, @('8-K', '0001045810-26-000073', '2026-08-26', '2026-08-26', '2.02,9.01', '2026-08-26T20:21:19.000Z'))
$x = Get-ResultadosSEC $apple 'ua' $sess $fer
Check 'invalid: another company (CIK) in the response is rejected' ($x.estado -eq 'error' -and $x.erro -match 'CIK' -and @($x.resultados).Count -eq 0)
$Script:SecJ = (SecJson '0000320193' @(@('8-K', 'a1', '2026-07-30', '2026-07-30', '2.02', 'x'), @('10-Q', 'a2', '2026-07-31', '2026-06-27', '', 'x'))).Replace('"items":[', '"items":["extra",')
$x = Get-ResultadosSEC $apple 'ua' $sess $fer
Check 'invalid: filing lists of different lengths are rejected' ($x.estado -eq 'error' -and $x.erro -match 'different lengths')
$Script:SecJ = SecJson '0000320193' @(@('8-K', '0000320193-26-000018', '2026-07-30', '2026-07-30', '2.02,9.01', 'x'), @('8-K', '0000320193-26-000011', '2026-04-30', '2026-04-30', '2.02,9.01', 'x'))
$Script:SecFalha = 'json'; $x = Get-ResultadosSEC $apple 'ua' $sess $fer
Check 'SEC down: error recorded, nothing invented, no exception' ($x.estado -eq 'error' -and @($x.resultados).Count -eq 0 -and $Script:Fontes[-1].estado -eq 'error' -and $Script:Fontes[-1].erro -match '503')
$Script:SecFalha = 'atom'; $x = Get-ResultadosSEC $apple 'ua' $sess $fer @{ '0000320193-26-000011' = '2026-04-30T20:30:41Z' }
Check 'acceptance feed down: dates kept, times from the previous run reused, the others Unavailable' ($x.estado -eq 'ok' -and $null -eq (& $ra '0000320193-26-000018').sessao -and (& $ra '0000320193-26-000011').sessao -eq '2026-05-01' -and $x.nota -match 'timeout') ($x | ConvertTo-Json -Depth 5 -Compress)
$Script:SecFalha = ''
# hora de aceitação pelo cabeçalho da entrega no arquivo do EDGAR (remediação de 7 out 2026: o feed Atom só tem as 40 mais recentes)
$Script:SecJ = SecJson '0000320193' @(
    @('8-K', '0000320193-26-000018', '2026-07-30', '2026-07-30', '2.02,9.01', 'x'),
    @('8-K', '0000320193-22-000070', '2022-07-28', '2022-07-28', '2.02,9.01', 'x'),
    @('8-K', '0000320193-23-000005', '2023-02-02', '2023-02-02', '2.02,9.01', 'x'),
    @('8-K', '0000320193-24-000007', '2024-01-30', '2024-01-30', '2.02,9.01', 'x'),
    @('8-K', '0000320193-21-000105', '2021-10-28', '2021-10-28', '2.02,9.01', 'x'))
$Script:SecCab = @{ '0000320193-22-000070' = '20220728163012'; '0000320193-23-000005' = '20230202163107' }
$Script:SecCab['0000320193-24-000007'] = (SecCabecalho '0000320193-24-000999' '20240130163000')   # cabeçalho de outra entrega
$Script:SecPedidos.Clear()
$x = Get-ResultadosSEC $apple 'ua' $sess $fer @{ '0000320193-21-000105' = '2021-10-28T20:30:00Z' }
$a = & $ra '0000320193-22-000070'
Check 'filing header: summer time (EDT) 16:30:12 New York = 20:30:12 UTC, reaction the next session' ($a.aceite -eq '2022-07-28T20:30:12Z' -and $a.horaNY -eq '2022-07-28 16:30' -and $a.quando -eq 'after' -and $a.sessao -eq '2022-07-29') ($a | ConvertTo-Json -Compress)
$a = & $ra '0000320193-23-000005'
Check 'filing header: winter time (EST) 16:31:07 New York = 21:31:07 UTC' ($a.aceite -eq '2023-02-02T21:31:07Z' -and $a.quando -eq 'after' -and $a.sessao -eq '2023-02-03') ($a | ConvertTo-Json -Compress)
$a = & $ra '0000320193-24-000007'
Check 'a filing header of another accession number is rejected: no time, nothing invented' ($null -eq $a.aceite -and $null -eq $a.sessao -and $x.nota -match 'not for 0000320193-24-000007') ($x.nota)
Check 'times already known (previous run) and from the 8-K feed are not requested again; the note counts the headers used' (@($Script:SecPedidos | Where-Object { $_ -match '0000320193-21-000105|0000320193-26-000018' -and $_ -match 'index-headers' }).Count -eq 0 -and (& $ra '0000320193-21-000105').aceite -eq '2021-10-28T20:30:00Z' -and $x.nota -match '2 acceptance time\(s\) from the EDGAR filing headers') ($Script:SecPedidos -join ' ; ')
$guardaMax = $MaxCabecalhos; $MaxCabecalhos = 1; $Script:SecPedidos.Clear(); $Script:SecCab = @{}
$x = Get-ResultadosSEC $apple 'ua' $sess $fer
$MaxCabecalhos = $guardaMax
Check 'filing headers: at most $MaxCabecalhos requests per company and run (the rest stay Unavailable until the next run)' (@($Script:SecPedidos | Where-Object { $_ -match 'index-headers' }).Count -eq 1 -and @($x.resultados | Where-Object { -not $_.aceite }).Count -eq 4) ($Script:SecPedidos -join ' ; ')
Check 'filing header with the acceptance time in a winter-time gap or malformed: no time' ($null -eq (& { $Script:SecCab = @{ '0000320193-26-000018' = (SecCabecalho '0000320193-26-000018' '2026073016') }; Get-AceiteCabecalho '0000320193' '0000320193-26-000018' 'ua' }) -and $null -eq (& { $Script:SecCab = @{ '0000320193-26-000018' = '20260308023000' }; Get-AceiteCabecalho '0000320193' '0000320193-26-000018' 'ua' }))
$Script:SecCab = @{}
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-04-03T11:00:00-04:00', $Script:Inv)) $sess $fer
Check 'released on a day the exchange is closed (Good Friday, 11:00): "closed", reaction the next session, with a note' ($r.quando -eq 'closed' -and $r.sessao -eq '2026-04-06' -and $r.nota -match 'closed') ($r | ConvertTo-Json -Compress)
$r = Get-SessaoReacao ([DateTimeOffset]::Parse('2026-08-01T10:00:00-04:00', $Script:Inv)) $sess $fer
Check 'released on a Saturday: "closed", reaction on Monday' ($r.quando -eq 'closed' -and $r.sessao -eq '2026-08-03') ($r | ConvertTo-Json -Compress)
${function:Get-Url} = $guardaUrl2

Write-Host 'Fundamentals from SEC XBRL companyfacts (canned responses)'
# companyfacts reduzido ao essencial, no formato real: facts.us-gaap.<tag>.units.<unidade> = [{start, end, val, form, filed, ...}]
function Fato([string]$Ini, [string]$Fim, $Val, [string]$Filed, [string]$Form = '10-Q') { [ordered]@{ start = $Ini; end = $Fim; val = $Val; accn = '0000000000-00-000000'; fy = 2025; fp = 'Q1'; form = $Form; filed = $Filed; frame = '' } }
function Factos([string]$Cik, [hashtable]$Tags) {
    $g = [ordered]@{}; foreach ($t in $Tags.Keys) { $u = [ordered]@{}; foreach ($k in $Tags[$t].Keys) { $u[$k] = @($Tags[$t][$k]) }; $g[$t] = [ordered]@{ label = $t; units = $u } }
    @{ cik = [int64]$Cik; entityName = 'Test Inc.'; facts = @{ 'us-gaap' = $g; dei = @{} } } | ConvertTo-Json -Depth 8 -Compress
}
# ano fiscal como o da Apple: começa a 29 set 2024; trimestres de 90 dias, 9 meses e ano no 10-K
$S = '2024-09-29'
$Script:FactosA = Factos '0000320193' @{
    Revenues = @{ USD = @((Fato $S '2024-12-28' 120 '2025-01-31'), (Fato $S '2024-12-28' 118 '2026-01-30'), (Fato '2024-12-29' '2025-03-29' 95 '2025-05-02'), (Fato $S '2025-03-29' 215 '2025-05-02'),
        (Fato '2025-03-30' '2025-06-28' 90 '2025-08-01'), (Fato $S '2025-06-28' 305 '2025-08-01'), (Fato $S '2025-09-27' 400 '2025-10-31' '10-K'), (Fato $S '2025-01-27' 160 '2025-02-28'), (Fato '2025-06-29' '2025-09-27' 999 '2025-10-30' '8-K')) }
    NetCashProvidedByUsedInOperatingActivities = @{ USD = @((Fato $S '2024-12-28' 30 '2025-01-31'), (Fato $S '2025-03-29' 70 '2025-05-02'), (Fato $S '2025-06-28' 100 '2025-08-01'), (Fato $S '2025-09-27' 150 '2025-10-31' '10-K')) }
    PaymentsToAcquirePropertyPlantAndEquipment = @{ USD = @((Fato $S '2024-12-28' 5 '2025-01-31'), (Fato $S '2025-03-29' 12 '2025-05-02'), (Fato $S '2025-06-28' 18 '2025-08-01'), (Fato $S '2025-09-27' 30 '2025-10-31' '10-K')) }
    EarningsPerShareDiluted = @{ 'USD/shares' = @((Fato $S '2024-12-28' 1.5 '2025-01-31'), (Fato '2024-12-29' '2025-03-29' 1.6 '2025-05-02'), (Fato '2025-03-30' '2025-06-28' 1.5 '2025-08-01'), (Fato $S '2025-06-28' 4.6 '2025-08-01'), (Fato $S '2025-09-27' 6.0 '2025-10-31' '10-K')) }
    WeightedAverageNumberOfDilutedSharesOutstanding = @{ shares = @((Fato $S '2024-12-28' 15e9 '2025-01-31'), (Fato $S '2025-06-28' 14.9e9 '2025-08-01'), (Fato $S '2025-09-27' 14.8e9 '2025-10-31' '10-K')) }
    OperatingIncomeLoss = @{ EUR = @(, (Fato $S '2024-12-28' 40 '2025-01-31')) }
}
# ano fiscal como o da NVIDIA (fim de janeiro) e o split de 10:1 de 10 jun 2024: EPS publicado antes (5,98) e reexpresso depois (0,6)
$Script:FactosN = Factos '0001045810' @{
    Revenues = @{ USD = @((Fato '2025-07-28' '2025-10-26' 57 '2025-11-19'), (Fato '2025-01-27' '2025-10-26' 147 '2025-11-19')) }
    EarningsPerShareDiluted = @{ 'USD/shares' = @((Fato '2024-01-29' '2024-04-28' 5.98 '2024-05-29'), (Fato '2024-01-29' '2024-04-28' 0.6 '2025-05-28')) }
    WeightedAverageNumberOfDilutedSharesOutstanding = @{ shares = @((Fato '2024-01-29' '2024-04-28' 2.5e9 '2024-05-29'), (Fato '2024-01-29' '2024-04-28' 25e9 '2025-05-28')) }
}
$Script:FactosResp = $Script:FactosA; $Script:FactosFalha = $false; $Script:FactosPedidos = 0
$guardaUrl3 = ${function:Get-Url}
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) if ($Url -notmatch '^https://data\.sec\.gov/api/xbrl/companyfacts/CIK\d{10}\.json$') { throw "unexpected $Url" }; $Script:FactosPedidos++; if ($Script:FactosFalha) { throw '503 Service Unavailable' }; return $Script:FactosResp }
$rel = [pscustomobject]@{ form = '10-K'; data = '2025-10-31'; periodo = '2025-09-27'; acc = 'x' }
$fa = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'BluechipBoard/1.0 test@example.com' @() $null $rel
$tq = { param($o, $fim) @($o.trimestres | Where-Object { $_.fim -eq $fim })[0] }
$q4 = & $tq $fa '2025-09-27'; $q1 = & $tq $fa '2024-12-28'; $q2 = & $tq $fa '2025-03-29'
Check 'collected: state ok, one request, the tags used are recorded' ($fa.estado -eq 'ok' -and $Script:FactosPedidos -eq 1 -and (@($fa.tags.receita) -join ',') -eq 'Revenues' -and $fa.relatorio.data -eq '2025-10-31') ($fa | ConvertTo-Json -Depth 3 -Compress)
Check '4th quarter = year − 9 months (400 − 305 = 95), derived, published with the 10-K' ($q4.m.receita.v -eq 95 -and $q4.m.receita.d -and $q4.m.receita.f -eq '2025-10-31' -and $q4.m.receita.p -eq '2025-10-31' -and $q4.ano -eq 'FY2025' -and $q4.q -eq 4) ($q4 | ConvertTo-Json -Depth 4 -Compress)
Check 'restatement: the latest filing wins (118), the first published value and date are kept (120, 2025-01-31)' ($q1.m.receita.v -eq 118 -and $q1.m.receita.f -eq '2026-01-30' -and $q1.m.receita.v0 -eq 120 -and $q1.m.receita.p -eq '2025-01-31' -and $q1.q -eq 1) ($q1.m.receita | ConvertTo-Json -Compress)
Check 'cash flow is year to date: quarters derived from the cumulative figures (Q2 = 70 − 30); FCF = CFO − capex (Q4 = 50 − 12)' ($q2.m.cfo.v -eq 40 -and $q2.m.cfo.d -and $q1.m.cfo.v -eq 30 -and -not $q1.m.cfo.d -and $q4.m.fcf.v -eq 38) ($q4.m | ConvertTo-Json -Depth 3 -Compress)
Check 'EPS of the 4th quarter = year − 9 months (6.0 − 4.6); diluted shares are never derived by subtraction' ([math]::Abs($q4.m.eps.v - 1.4) -lt 1e-9 -and $q4.m.eps.d -and $null -eq $q4.m.acoes -and $q1.m.acoes.v -eq 15e9) ($q4.m | ConvertTo-Json -Depth 3 -Compress)
Check 'TTM of the latest quarter: 4 consecutive quarters (118 + 95 + 90 + 95 = 398)' ($fa.ttm.receita.v -eq 398 -and $fa.ttm.receita.f -eq '2026-01-30' -and $fa.ttm.fim -eq '2025-09-27') ($fa.ttm | ConvertTo-Json -Depth 3 -Compress)
Check 'durations other than ~90/180/270/365 days and forms other than 10-Q/10-K are ignored' (-not @($fa.trimestres | Where-Object { $_.fim -eq '2025-01-27' }).Count -and $q4.m.receita.v -ne 999)
Check 'missing tag: Unavailable for every quarter, with the reason' ($null -eq $q1.m.lucroBruto -and $fa.faltam.lucroBruto -match 'GrossProfit') ($fa.faltam | ConvertTo-Json -Compress)
Check 'wrong units (EUR instead of USD): not used, with the reason' ($null -eq $q1.m.lucroOperacional -and $fa.faltam.lucroOperacional -match 'USD') ($fa.faltam | ConvertTo-Json -Compress)
Check 'a quarter without data is not interpolated' (@($fa.trimestres | Where-Object { $_.m.lucroBruto }).Count -eq 0 -and $null -eq (& $tq $fa '2025-06-28').m.acoes)
$Script:FactosResp = $Script:FactosN
$fn = Get-FundamentaisEmpresa @{ Id = 'NVDA'; Cik = '0001045810' } 'ua' @(, @('2024-06-10', 10)) $null $null
$n1 = & $tq $fn '2024-04-28'; $n3 = & $tq $fn '2025-10-26'
Check 'split: per-share figures adjusted for splits after each filing (0.6 filed after the split stays; 5.98 filed before → 0.598)' ([math]::Abs($n1.m.eps.v - 0.6) -lt 1e-9 -and [math]::Abs($n1.m.eps.v0 - 0.598) -lt 1e-9 -and $n1.m.eps.p -eq '2024-05-29' -and $n1.m.acoes.v -eq 25e9 -and $null -eq $n1.m.acoes.v0) ($n1.m | ConvertTo-Json -Depth 3 -Compress)
Check 'NVIDIA fiscal year (ends late January): quarter ended 26 Oct 2025 = Q3 FY2026' ($n3.ano -eq 'FY2026' -and $n3.q -eq 3 -and $n3.m.receita.v -eq 57) ($n3 | ConvertTo-Json -Depth 4 -Compress)
# cache e fallback
$Script:FactosResp = $Script:FactosA; $Script:FactosPedidos = 0
$prev = [pscustomobject]@{ id = 'AAPL'; estado = 'ok'; fonte = 'SEC XBRL (companyfacts)'; obtidoEm = $Script:Agora.AddDays(-10).ToString('o'); relatorio = [pscustomobject]@{ form = '10-Q'; data = '2026-07-31'; periodo = '2026-06-27'; acc = 'y' }; tags = $fa.tags; faltam = $fa.faltam; trimestres = $fa.trimestres; ttm = $fa.ttm; erro = ''; nota = '' }
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'ua' @() $prev ([pscustomobject]@{ form = '10-Q'; data = '2026-07-31'; periodo = '2026-06-27'; acc = 'y' })
Check 'cache: no 10-Q/10-K newer than the last collection → nothing downloaded, the stored facts are used' ($Script:FactosPedidos -eq 0 -and $x.estado -eq 'ok' -and @($x.trimestres).Count -eq @($fa.trimestres).Count -and $Script:Fontes[-1].estado -eq 'ok (cached)') "$($Script:FactosPedidos) $($x.estado) $($Script:Fontes[-1].estado)"
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'ua' @() $prev ([pscustomobject]@{ form = '10-K'; data = '2026-10-30'; periodo = '2026-09-26'; acc = 'z' })
Check 'a newer 10-K at the SEC → downloaded again' ($Script:FactosPedidos -eq 1 -and $x.estado -eq 'ok' -and $x.relatorio.data -eq '2026-10-30')
$Script:FactosFalha = $true
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'ua' @() $prev ([pscustomobject]@{ form = '10-K'; data = '2026-10-30'; periodo = '2026-09-26'; acc = 'z' })
Check "source down: the previous run's facts (10 days old) are used, labelled" ($x.estado -eq 'previous run' -and $x.fonte -like 'previous run (retrieved *' -and @($x.trimestres).Count -eq @($fa.trimestres).Count -and $x.erro -match '503' -and $Script:Fontes[-1].estado -eq 'error') ($x | ConvertTo-Json -Depth 2 -Compress)
$prev.obtidoEm = $Script:Agora.AddDays(-130).ToString('o')
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'ua' @() $prev ([pscustomobject]@{ form = '10-K'; data = '2026-10-30'; periodo = '2026-09-26'; acc = 'z' })
Check 'fallback limit: previous facts over 120 days old are not used (Unavailable, nothing invented)' ($x.estado -eq 'error' -and @($x.trimestres).Count -eq 0)
$Script:FactosFalha = $false; $Script:FactosPedidos = 0
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } '' @() $null $null
Check 'without -SecEmail: no request, skipped with the reason' ($Script:FactosPedidos -eq 0 -and $x.estado -eq 'skipped' -and $x.erro -match 'EmailSEC')
$Script:FactosResp = $Script:FactosN
$x = Get-FundamentaisEmpresa @{ Id = 'AAPL'; Cik = '0000320193' } 'ua' @() $null $null
Check 'another company (CIK) in the response is rejected' ($x.estado -eq 'error' -and $x.erro -match 'CIK')
${function:Get-Url} = $guardaUrl3

Write-Host 'Euro area indicators from the ECB Data Portal (canned responses)'
# CSV "dataonly" como o real: KEY, as dimensões da série, TIME_PERIOD, OBS_VALUE
$csvH = { param($Chave, $Linhas) "KEY,FREQ,REF_AREA,ADJUSTMENT,ICP_ITEM,STS_INSTITUTION,ICP_SUFFIX,TIME_PERIOD,OBS_VALUE`n" + (@(foreach ($l in $Linhas) { "$Chave,M,U2,N,000000,4,ANR,$($l[0]),$($l[1])" }) -join "`n") }
$kH = 'ICP.M.U2.N.000000.4.ANR'
$r = ConvertFrom-CsvSerie (& $csvH $kH @(@('2025-12', '1.9'), @('2025-10', '2.1'), @('2025-11', '2.1'))) $kH -5 25
Check 'valid CSV: monthly values read and sorted (Dec 2025 = 1.9)' (@($r.pontos).Count -eq 3 -and $r.pontos[-1][0] -eq '2025-12-01' -and $r.pontos[-1][1] -eq 1.9 -and $r.pontos[0][0] -eq '2025-10-01') ($r | ConvertTo-Json -Compress -Depth 4)
$e = ''; try { ConvertFrom-CsvSerie '<!DOCTYPE html><html><body>Service unavailable</body></html>' $kH -5 25 | Out-Null } catch { $e = $_.Exception.Message }
Check 'an HTML error page is rejected, never read as data' ($e -match 'HTML') $e
$r = ConvertFrom-CsvSerie (& $csvH $kH @(@('2025-10', '.'), @('2025-11', 'NaN'), @('2025-12', ''), @('2026-01', '2.0'))) $kH -5 25
Check '".", "NaN" and empty values are missing, never 0' (@($r.pontos).Count -eq 1 -and $r.faltas -eq 3 -and $r.pontos[0][1] -eq 2.0) ($r | ConvertTo-Json -Compress -Depth 4)
$e = ''; try { ConvertFrom-CsvSerie (& $csvH $kH @(@('2025-10', '.'))) $kH -5 25 | Out-Null } catch { $e = $_.Exception.Message }
Check 'empty series (no value): error' ($e -match 'no value') $e
$r = ConvertFrom-CsvSerie (& $csvH $kH @(@('2025-11', '2.1'), @('2025-12', '150'))) $kH -5 25
Check 'implausible values are ignored and counted' (@($r.pontos).Count -eq 1 -and $r.invalidos -eq 1)
$e = ''; try { ConvertFrom-CsvSerie (& $csvH 'ICP.M.DE.N.000000.4.ANR' @(, @('2025-12', '2.0'))) $kH -5 25 | Out-Null } catch { $e = $_.Exception.Message }
Check 'another series in the response (its key) is rejected' ($e -match 'key') $e
$e = ''; try { ConvertFrom-CsvSerie "DATE,VALUE`n2025-12-01,2.0" $kH -5 25 | Out-Null } catch { $e = $_.Exception.Message }
Check 'an unexpected header is rejected' ($e -match 'header') $e
$kD = 'FM.B.U2.EUR.4F.KR.DFR.LEV'
$csvD = "KEY,FREQ,REF_AREA,CURRENCY,PROVIDER_FM,INSTRUMENT_FM,PROVIDER_FM_ID,DATA_TYPE_FM,TIME_PERIOD,OBS_VALUE`n$kD,B,U2,EUR,4F,KR,DFR,LEV,2019-09-18,-0.5`n$kD,B,U2,EUR,4F,KR,DFR,LEV,2026-09-16,2.5"
$r = ConvertFrom-CsvSerie $csvD $kD -2 20
Check 'daily dates and negative rates are kept (deposit rate −0.50 from 18 Sep 2019)' (@($r.pontos).Count -eq 2 -and $r.pontos[0][0] -eq '2019-09-18' -and $r.pontos[0][1] -eq -0.5) ($r | ConvertTo-Json -Compress -Depth 4)
$guardaUrl4 = ${function:Get-Url}
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) throw 'offline' }
$m = Get-SerieMacro $SeriesMacro[1]
Check 'source down: error recorded, no values, no exception' ($m.estado -eq 'error' -and @($m.pontos).Count -eq 0 -and $Script:Fontes[-1].estado -eq 'error' -and $Script:Fontes[-1].nome -like 'ECB: *')
$Script:CsvResp = & $csvH $kH @(@('2025-11', '2.1'), @('2025-12', '1.9'))
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) if ($Url -notmatch '^https://data-api\.ecb\.europa\.eu/service/data/ICP/M\.U2\.N\.000000\.4\.ANR\?format=csvdata&detail=dataonly') { throw "unexpected $Url" }; return $Script:CsvResp }
$m = Get-SerieMacro $SeriesMacro[1]
Check 'valid response: ok, monthly, recorded as a source' ($m.estado -eq 'ok' -and $m.freq -eq 'M' -and @($m.pontos).Count -eq 2 -and $Script:Fontes[-1].estado -eq 'ok') ($m | ConvertTo-Json -Compress -Depth 4)
${function:Get-Url} = $guardaUrl4

Write-Host 'Central-bank meetings in the calendar'
$fed27 = @($Calendario | Where-Object { "$($_.d)" -like '2027-*' -and $_.ev -match '^Fed decision' -and $_.st -eq 'C' -and $_.e -eq 'MKT' })
$bce26 = @($Calendario | Where-Object { "$($_.d)" -like '2026-*' -and $_.ev -match '^ECB monetary policy decision' -and $_.st -eq 'C' -and $_.e -eq 'MKT' })
$bce27 = @($Calendario | Where-Object { "$($_.d)" -like '2027-*' -and $_.ev -match '^ECB monetary policy decision' -and $_.st -eq 'C' -and $_.e -eq 'MKT' })
Check 'calendar: the 8 Fed decisions of 2027 and the ECB decisions of 2026 (2 left) and 2027 (8), confirmed, Market' ($fed27.Count -eq 8 -and $bce26.Count -eq 2 -and $bce27.Count -eq 8) "$($fed27.Count) $($bce26.Count) $($bce27.Count)"
Check 'calendar dates are valid and in order' (@($Calendario | Where-Object { "$($_.d)" -notmatch '^\d{4}-\d{2}-\d{2}$' }).Count -eq 0 -and (@($Calendario | ForEach-Object { $_.d }) -join ',') -eq (@($Calendario | ForEach-Object { $_.d } | Sort-Object) -join ','))
Check 'no reminder while the Fed and the ECB both have a decision more than 60 days ahead' ($null -eq (Get-LembreteReunioes $Calendario $Script:Agora))
$l = Get-LembreteReunioes @($Calendario | Where-Object { $_.ev -notmatch '^ECB' }) $Script:Agora
Check 'reminder when no ECB decision is more than 60 days ahead' ($l -match 'ECB' -and $l -notmatch 'Fed') $l
$l = Get-LembreteReunioes $Calendario ([DateTimeOffset]::new(2027, 11, 1, 0, 0, 0, [TimeSpan]::Zero))
Check 'reminder for both when the list ends within 60 days (1 Nov 2027)' ($l -match 'Fed' -and $l -match 'ECB') $l

Write-Host 'News history (noticias-historico.json)'
$tmpH = Join-Path $env:TEMP "bb-hist-test-$PID.json"
$h = Read-HistoricoNoticias "$tmpH.missing"
Check 'missing file: starts empty, with a notice' ($h.estado -eq 'new' -and @($h.itens).Count -eq 0 -and $h.aviso -match 'not found')
[IO.File]::WriteAllText($tmpH, '{"inicio":"2026-01-01T00:00:00Z","noticias":[{"titulo":"x"', $utf8)
$h = Read-HistoricoNoticias $tmpH
Check 'corrupted file: starts empty, with a warning (the old file is kept as .bad)' ($h.estado -eq 'corrupted' -and @($h.itens).Count -eq 0 -and $h.aviso -match '\.bad') $h.aviso
$ag = $Script:Agora
$hi = { param($t, $dias, $nivel, $s) [pscustomobject]@{ chave = $t.ToLowerInvariant(); titulo = $t; link = 'https://example.com/' + $dias; fonte = 'Reuters'; dominio = 'reuters.com'; data = $ag.AddDays(-$dias).ToString('o'); empresas = @('AAPL'); temas = @('Earnings'); nivel = $nivel; score = $s; novo = $true; outras = @(); soFeed = $false } }
$base = [pscustomobject]@{ inicio = (ConvertTo-IsoUtc $ag.AddDays(-450)); itens = @((ConvertTo-NoticiaHistorico (& $hi 'Old story' 401 'red' 8)), (ConvertTo-NoticiaHistorico (& $hi 'Kept story' 399 'orange' 5)), (ConvertTo-NoticiaHistorico (& $hi 'Same title' 3 'orange' 4.5))); estado = 'ok'; aviso = '' }
$m = Merge-HistoricoNoticias $base @((& $hi 'Same title' 0 'red' 7.5), (& $hi 'Moderate story' 0 'yellow' 3), (& $hi 'New material story' 0 'red' 8)) $ag
$tit = @($m.itens | ForEach-Object { $_.titulo })
Check 'retention: 400 days (a story 401 days old is dropped, one 399 days old is kept)' ($tit -notcontains 'Old story' -and $tit -contains 'Kept story') ($tit -join ' | ')
Check 'only material and important stories are kept' ($tit -notcontains 'Moderate story' -and $tit -contains 'New material story')
Check 'a story already in the history is kept once, with its highest score and level' (@($m.itens | Where-Object { $_.titulo -eq 'Same title' }).Count -eq 1 -and @($m.itens | Where-Object { $_.titulo -eq 'Same title' })[0].score -eq 7.5 -and @($m.itens | Where-Object { $_.titulo -eq 'Same title' })[0].nivel -eq 'red')
Check 'only public fields are kept (title, link, source, date, assets, level, score, themes)' ((@($m.itens[0].PSObject.Properties.Name) -join ',') -eq 'chave,titulo,link,fonte,data,empresas,nivel,score,temas')
Check 'history start never before the 400-day cut' ((ConvertTo-Data $m.inicio) -ge $ag.AddDays(-400).AddSeconds(-1) -and (ConvertTo-Data $m.inicio) -le $ag.AddDays(-399)) $m.inicio
$n0 = Merge-HistoricoNoticias (Read-HistoricoNoticias "$tmpH.missing") @() $ag
Check 'a new history starts now, empty' ($n0.inicio -eq (ConvertTo-IsoUtc $ag) -and @($n0.itens).Count -eq 0) $n0.inicio
[IO.File]::WriteAllText($tmpH, ([ordered]@{ versao = 1; inicio = $m.inicio; dias = 400; noticias = @($m.itens) } | ConvertTo-Json -Depth 5 -Compress), $utf8)
$back = Read-HistoricoNoticias $tmpH
Check 'written and read back: the same stories, dates still ISO text (PowerShell 7 reads them as dates)' ($back.estado -eq 'ok' -and @($back.itens).Count -eq @($m.itens).Count -and $back.inicio -eq $m.inicio -and @($back.itens | Where-Object { $_.data -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$' }).Count -eq 0) ($back | ConvertTo-Json -Depth 5 -Compress)
try { [IO.File]::Delete($tmpH) } catch { }

Write-Host 'New ETFs (EUNK, IS3N, EUNN): configuration, prices and holdings'
$cfgEtf = @{ EUNK = @('EUNK.DE', 'IE00B4K48X80', 251861, 'Europe'); IS3N = @('IS3N.DE', 'IE00BKM4GZ66', 264659, 'Emerging markets'); EUNN = @('EUNN.DE', 'IE00B4L5YX21', 251867, 'Japan') }
foreach ($id in 'EUNK', 'IS3N', 'EUNN') {
    $a = $Ativos | Where-Object { $_.Id -eq $id }; $e = $ETFs | Where-Object { $_.Id -eq $id }; $c = $cfgEtf[$id]; $x = [datetime]::MinValue
    Check "$id registered: Yahoo $($c[0]) in EUR, ISIN $($c[1]), accumulating, Irish, on Xetra" ($a -and $a.Yahoo -eq $c[0] -and $a.Moeda -eq 'EUR' -and $e -and $e.Isin -eq $c[1] -and $e.Acumulacao -and $e.Domicilio -eq 'IE' -and $e.Bolsa -eq 'DE' -and @($Bolsas | Where-Object { $_.id -eq $e.Bolsa }).Count -eq 1 -and $e.Chip -eq $c[3])
    Check "$id has its own holdings file and its own reference weights" ($e.Url -match "portfolioId=$($c[2])&" -and $e.Url -ne $UrlPesosETF -and $e.Referencia -and $e.Referencia.AAPL -eq 0 -and [datetime]::TryParseExact("$($e.Referencia.Data)", 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$x))
}
$sx = $ETFs | Where-Object { $_.Id -eq 'SXR8' }
Check 'SXR8 unchanged: same holdings file, reference weights and expected companies' ($sx.Url -eq $UrlPesosETF -and $sx.Referencia.AAPL -eq $PesosReferencia.AAPL -and ($sx.Empresas -join ',') -eq 'AAPL,NVDA,GOOGL' -and ($Ativos | Where-Object { $_.Id -eq 'SXR8' }).Yahoo -eq 'SXR8.DE')
Check 'every ETF is also a price asset, ids unique' (@($ETFs | Where-Object { $i = $_.Id; -not @($Ativos | Where-Object { $_.Id -eq $i }).Count }).Count -eq 0 -and @($Ativos.Id | Select-Object -Unique).Count -eq $Ativos.Count)
$Script:Resposta = Yahoo 'EUNK.DE' 'EUR' @(100, 101, 102, 101.5, 102.2, 103)
Check 'EUNK prices: valid euro series accepted' ($null -ne (Get-Serie ($Ativos | Where-Object { $_.Id -eq 'EUNK' })))
$Script:Resposta = Yahoo 'IMAE.AS' 'EUR' @(100, 101, 102, 101.5, 102.2, 103)
Check 'EUNK prices: another listing (IMAE.AS) in the response is rejected' ($null -eq (Get-Serie ($Ativos | Where-Object { $_.Id -eq 'EUNK' })) -and $Script:Fontes[-1].erro -match 'IMAE')
$Script:Resposta = Yahoo 'IS3N.DE' 'USD' @(50, 51, 52, 51, 52, 53)
Check 'IS3N prices: a dollar response for the euro listing is rejected' ($null -eq (Get-Serie ($Ativos | Where-Object { $_.Id -eq 'IS3N' })) -and $Script:Fontes[-1].erro -match 'currency USD')
$Script:Resposta = Yahoo 'EUNN.DE' 'EUR' @(75, 0, -1, 76, 77, 78, 79)
$s = Get-Serie ($Ativos | Where-Object { $_.Id -eq 'EUNN' })
Check 'EUNN prices: zero and negative prices dropped' ($s -and $s.pontos.Count -eq 5)

# ficheiro "Detailed Holdings and Analytics" da iShares, reduzido ao essencial (folhas Key Facts e Holdings)
function iShares([string]$Isin, $Linhas, [switch]$SemHoldings, $Geo = $null) {
    $cel = { param($v) '<ss:Cell><ss:Data ss:Type="String">' + [Net.WebUtility]::HtmlEncode("$v") + '</ss:Data></ss:Cell>' }
    $lin = { param([object[]]$v) '<ss:Row>' + (($v | ForEach-Object { & $cel $_ }) -join '') + '</ss:Row>' }
    $kf = '<ss:Worksheet ss:Name="Key Facts"><ss:Table>' + (& $lin @('Fund Name', 'Test fund')) + (& $lin @('ISIN', $Isin)) + '</ss:Table></ss:Worksheet>'
    $h = '<ss:Worksheet ss:Name="Holdings"><ss:Table>' + (& $lin @('All')) + (& $lin @('as of', '02/Oct/2026')) + (& $lin @('Issuer Ticker', 'Name', 'Sector', 'Asset Class', 'Market Value', 'Weight (%)', 'Notional Value', 'Nominal', 'Market Currency')) +
        (($Linhas | ForEach-Object { & $lin $_ }) -join '') + '</ss:Table></ss:Worksheet>'
    # folha "Exposure Breakdowns" como a real: um bloco de setores e, nos fundos com vários países, "Geography/Locations"
    $eb = '<ss:Worksheet ss:Name="Exposure Breakdowns"><ss:Table>' + (& $lin @('Sector')) + (& $lin @('as of', '02/Oct/2026')) + (& $lin @('Type', 'Fund', '')) + (& $lin @('Financials', '50.00')) +
        $(if ($Geo) { (& $lin @('Geography/Locations')) + (& $lin @('as of', '02/Oct/2026')) + (& $lin @('Type', 'Fund', '', '')) + (($Geo | ForEach-Object { & $lin $_ }) -join '') + (& $lin @('Geographic exposure relates principally to the domicile of the issuers.')) } else { '' }) +
        (& $lin @('Allocations are subject to change.')) + '</ss:Table></ss:Worksheet>'
    '<?xml version="1.0"?><ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' + $kf + $eb + $(if ($SemHoldings) { '' } else { $h }) + '</ss:Workbook>'
}
$setores = 'Financials', 'Health Care', 'Industrials', 'Consumer Staples'
$linhasEU = @(, @('ASML', 'ASML HOLDING', 'Information Technology', 'Equity', 'EUR 1', '5.33', '1', '1', 'EUR')) + @(, @('V', 'VOLVO TEST', 'Industrials', 'Equity', 'EUR 1', '2.00', '1', '1', 'SEK')) +
    @(for ($i = 1; $i -le 8; $i++) { , @("C$i", "COMPANY $i PLC", $setores[$i % 4], 'Equity', 'EUR 1', '10.50', '1', '1', 'EUR') }) + @(, @('EUR', 'EUR CASH', 'Cash and/or Derivatives', 'Cash', 'EUR 1', '1.20', '1', '1', 'EUR'))
$linhasUS = @(, @('NVDA', 'NVIDIA', 'Information Technology', 'Equity', 'USD 1', '8.49', '1', '1', 'USD')) + @(, @('AAPL', 'APPLE', 'Information Technology', 'Equity', 'USD 1', '7.30', '1', '1', 'USD')) +
    @(, @('GOOGL', 'ALPHABET INC CLASS A', 'Communication', 'Equity', 'USD 1', '3.00', '1', '1', 'USD')) + @(, @('GOOG', 'ALPHABET INC CLASS C', 'Communication', 'Equity', 'USD 1', '2.45', '1', '1', 'USD')) +
    @(for ($i = 1; $i -le 6; $i++) { , @("U$i", "US COMPANY $i INC", $setores[$i % 4], 'Equity', 'USD 1', '13.00', '1', '1', 'USD') })
$Script:Ficheiros = @{}
function Get-Url { param([string]$Url, [string]$UserAgent, [int]$Timeout) $m = [regex]::Match($Url, 'portfolioId=(\d+)'); if ($m.Success -and $Script:Ficheiros.ContainsKey($m.Groups[1].Value)) { return $Script:Ficheiros[$m.Groups[1].Value] }; throw 'offline' }
$eunk = $ETFs | Where-Object { $_.Id -eq 'EUNK' }
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' $linhasEU
$r = Get-PesosETF $eunk
Check 'EUNK holdings: its own file read (10 companies, top 10, sectors), none of the three companies' ($r.aoVivo -and $r.posicoes -eq 10 -and @($r.top10).Count -eq 10 -and $r.top10[0].w -eq 10.5 -and @($r.setores).Count -eq 5 -and $r.AAPL -eq 0 -and $r.NVDA -eq 0 -and $r.GOOGL -eq 0 -and $r.dataIso -eq '2026-10-02') ($r | ConvertTo-Json -Compress -Depth 4)
Check 'EUNK holdings: metadata from the configuration (ISIN, chip, accumulating, index)' ($r.id -eq 'EUNK' -and $r.isin -eq 'IE00B4K48X80' -and $r.chip -eq 'Europe' -and $r.acumulacao -and $r.indice -eq 'MSCI Europe' -and $r.domicilio -eq 'IE')
Check 'EUNK holdings: its own source row' ($Script:Fontes[-1].nome -eq 'iShares: ETF holdings (EUNK)' -and $Script:Fontes[-1].estado -eq 'ok' -and $Script:Fontes[-1].url -match 'portfolioId=251861')
Check 'a short non-US ticker (V, Swedish) keeps its own name (not renamed Visa)' (@($r.top10 | Where-Object { $_.t -eq 'V' })[0].n -eq 'Volvo Test')
Check 'ETF data serialises without {value, Count} pairs' (-not ($r | ConvertTo-Json -Compress -Depth 6).Contains('"Count":'))
$Script:Ficheiros['251861'] = iShares 'IE00BKM4GZ66' $linhasEU
$r = Get-PesosETF $eunk
Check 'wrong fund in the file (IS3N ISIN for EUNK): rejected, EUNK reference weights used' (-not $r.aoVivo -and $r.fonte -match 'Reference weights' -and @($r.top10).Count -eq 0 -and $r.AAPL -eq 0 -and $Script:Fontes[-1].estado -eq 'error (using reference weights)' -and $Script:Fontes[-1].erro -match 'not IE00B4K48X80') $Script:Fontes[-1].erro
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' $linhasEU -SemHoldings
$r = Get-PesosETF $eunk
Check 'malformed file (no Holdings sheet): reference weights, nothing invented' (-not $r.aoVivo -and @($r.top10).Count -eq 0 -and $Script:Fontes[-1].erro -match 'Holdings')
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' @(, @('C1', 'COMPANY 1', 'Financials', 'Equity', 'EUR 1', '10.00', '1', '1', 'EUR'))
$r = Get-PesosETF $eunk
Check 'weights that do not add up to about 100%: rejected' (-not $r.aoVivo -and $Script:Fontes[-1].erro -match 'add up to')
$Script:Ficheiros.Remove('251861')
$r = Get-PesosETF $eunk
Check 'fund on its reference weights: no aggregates (Unavailable)' ($null -eq $r.agregados)
Check 'holdings source down: reference weights with their date, error shown' (-not $r.aoVivo -and $r.dataIso -eq $eunk.Referencia.Data -and $r.isin -eq 'IE00B4K48X80' -and $Script:Fontes[-1].estado -eq 'error (using reference weights)')
$Script:Ficheiros['253743'] = iShares 'IE00B5BMR087' $linhasEU
$r = Get-PesosETF
Check 'SXR8 rules unchanged: a file without the three companies is rejected (SXR8 reference weights)' (-not $r.aoVivo -and $r.AAPL -eq $PesosReferencia.AAPL -and $Script:Fontes[-1].nome -eq 'iShares: ETF holdings (SXR8)' -and $Script:Fontes[-1].erro -match 'zero|Implausible')
$Script:Ficheiros['253743'] = iShares 'IE00B5BMR087' $linhasUS
$r = Get-PesosETF
Check 'SXR8 without arguments: the three weights, Alphabet A + C combined' ($r.aoVivo -and $r.id -eq 'SXR8' -and $r.AAPL -eq 7.3 -and $r.NVDA -eq 8.49 -and $r.GOOGL -eq 5.45 -and @($r.top10 | Where-Object { $_.t -eq 'NVDA' })[0].n -eq 'NVIDIA')
Check 'single-country index without a Geography block (SXR8): its index country, marked as such' ($r.agregados.fontePaises -eq 'index' -and @($r.agregados.paises).Count -eq 1 -and $r.agregados.paises[0].n -eq 'United States' -and $r.agregados.paises[0].w -eq 99.24 -and $r.paisIndice -eq 'United States' -and ($r | ConvertTo-Json -Compress -Depth 6).Contains('"paises":[{'))
# agregados lidos de uma cópia do formato real: país (bloco Geography), setor e moeda; caixa e derivados como Cash/Other
$linhasAg = @(, @('ASML', 'ASML HOLDING', 'Information Technology', 'Equity', 'EUR 1', '40.00', '1', '1', 'EUR')) + @(, @('NESN', 'NESTLE', 'Consumer Staples', 'Equity', 'CHF 1', '30.00', '1', '1', 'CHF')) +
    @(, @('SHEL', 'SHELL', 'Energy', 'Equity', 'GBP 1', '28.00', '1', '1', 'GBP')) + @(, @('EUR', 'EUR CASH', 'Cash and/or Derivatives', 'Cash', 'EUR 1', '1.50', '1', '1', 'EUR')) + @(, @('FUT', 'STOXX 50 FUTURE', 'Cash and/or Derivatives', 'Futures', 'EUR 1', '0.50', '1', '1', 'EUR'))
$geoAg = @(@('Netherlands', '40.00', 'NL'), @('Switzerland', '30.00', 'CH'), @('United Kingdom', '28.00', 'GB'), @('Cash and/or Derivatives', '2.00', 'CASHD_C'))
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' $linhasAg -Geo $geoAg
$r = Get-PesosETF $eunk; $a = $r.agregados; $w = { param($l, $n) (@($l) | Where-Object { $_.n -eq $n } | Select-Object -First 1).w }
Check 'aggregates from the file: countries (Geography block), sectors and underlying currencies; cash and futures as Cash/Other' ($r.aoVivo -and $a.fontePaises -eq 'file' -and (& $w $a.paises 'Netherlands') -eq 40 -and (& $w $a.paises 'Cash/Other') -eq 2 -and (& $w $a.setores 'Energy') -eq 28 -and (& $w $a.setores 'Cash/Other') -eq 2 -and (& $w $a.moedas 'CHF') -eq 30 -and (& $w $a.moedas 'Cash/Other') -eq 2) ($a | ConvertTo-Json -Compress -Depth 4)
Check 'only totals are kept, never the holdings themselves' (@($a.paises).Count -eq 4 -and @($a.setores).Count -eq 4 -and @($a.moedas).Count -eq 4 -and -not ($r | ConvertTo-Json -Compress -Depth 6).Contains('"Count":'))
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' $linhasEU
$r = Get-PesosETF $eunk
Check 'aggregates that do not add up to about 100% are rejected (Unavailable); the holdings are still read' ($r.aoVivo -and $null -eq $r.agregados.setores -and $null -eq $r.agregados.moedas -and $null -eq $r.agregados.paises)
# vários ETF ao mesmo tempo: um falha (IS3N), os outros não são afetados
$Script:Ficheiros['251861'] = iShares 'IE00B4K48X80' $linhasEU; $Script:Ficheiros['251867'] = iShares 'IE00B4L5YX21' $linhasEU
$todos = [ordered]@{}; foreach ($e in $ETFs) { $todos[$e.Id] = Get-PesosETF $e }
Check 'four ETFs at once: one source down (IS3N) does not affect the others' ((@($todos.Keys) -join ',') -eq 'SXR8,EUNK,IS3N,EUNN' -and $todos.SXR8.aoVivo -and $todos.EUNK.aoVivo -and -not $todos.IS3N.aoVivo -and $todos.EUNN.aoVivo -and $todos.EUNN.isin -eq 'IE00B4L5YX21' -and @($Script:Fontes | Select-Object -Last 4 | Where-Object { $_.estado -eq 'ok' }).Count -eq 3)
$guardaDiv = $AtivosDividendos; $AtivosDividendos = @('AAPL', 'NVDA', 'GOOGL', 'EUNK', 'IS3N')
$dv = Get-Dividendos; $AtivosDividendos = $guardaDiv
Check 'accumulating ETFs never get dividends, even if added to the dividend list' ((@($dv.Keys) -join ',') -eq 'AAPL,NVDA,GOOGL')
$n = Noticia 'European shares rise as STOXX 600 hits record high' 'Google News: Europe (MSCI Europe, STOXX 600)' 'EUNK'
Check 'news: "European shares … STOXX 600" is EUNK (strong match)' ($n.empresas -contains 'EUNK' -and -not $n.soFeed)
$n = Noticia 'iShares Core MSCI Europe UCITS ETF passes EUR 10bn' 'Google News: Europe (MSCI Europe, STOXX 600)' 'EUNK'
Check 'news: "iShares Core MSCI Europe" is EUNK, not SXR8' ($n.empresas -contains 'EUNK' -and $n.empresas -notcontains 'SXR8')
$n = Noticia 'Nikkei 225 falls 2% as yen strengthens'
Check 'news: "Nikkei 225" is EUNN' ($n.empresas -contains 'EUNN')
$n = Noticia 'Apple to cut iPhone production, Nikkei reports'
Check 'news: "Nikkei reports" (the newspaper) is not EUNN' ($n.empresas -notcontains 'EUNN' -and $n.empresas -contains 'AAPL')
$n = Noticia 'Emerging-market stocks rally as dollar weakens'
Check 'news: "emerging-market stocks" is IS3N' ($n.empresas -contains 'IS3N')
$n = Noticia 'Samsung Electronics shares jump on chip demand' 'Google News: Emerging markets (MSCI EM)' 'IS3N'
Check 'news: one EM company without keyword is only a weak feed match' ($n.soFeed -and $n.nivel -in 'yellow', 'white')
$n = Noticia 'iShares Core S&P 500 UCITS ETF tops $100bn' 'Google News: ETF SXR8 / CSPX' 'SXR8'
Check 'news: SXR8 classification unchanged' ($n.empresas -contains 'SXR8' -and $n.empresas -contains 'MKT' -and $n.empresas -notcontains 'EUNK')

Write-Host 'Bitcoin halving estimate (offline)'
$h = Get-HalvingAproximado
Check 'next halving after 2026: block 1,050,000, reward 3.125 to 1.5625' ($h.ev -match '1,050,000' -and $h.ev -match '3\.125 to 1\.5625' -and $h.d -like '2028-*') $h.ev
$guarda = $Script:Agora; $Script:Agora = [DateTimeOffset]::new(2029, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
$h = Get-HalvingAproximado; $Script:Agora = $guarda
Check 'after the 2028 halving the estimate moves to the next one' ($h.ev -match '1,260,000' -and $h.d -like '2032-*') $h.ev

Write-Host 'Stored data and backups'
$pts = @(Get-PontosGuardados @(@('2026-01-02', 1.5), [pscustomobject]@{ value = @('2026-01-03', 2.5); Count = 2 }, @('bad', 3), @('2026-01-05', -1)))
Check 'previous-run points: both pair forms read, invalid ones dropped' ($pts.Count -eq 2 -and $pts[1][0] -eq '2026-01-03' -and $pts[1][1] -eq 2.5)
$tmp = Join-Path $env:TEMP ('bb-test-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
New-Item -ItemType Directory $tmp | Out-Null
try {
    $f = Join-Path $tmp 'a.json'
    Write-Atomico $f 'one'; Write-Atomico $f 'two'
    Check 'atomic write replaces the file and leaves no .tmp' ([IO.File]::ReadAllText($f) -eq 'two' -and -not (Test-Path "$f.tmp"))
    [IO.File]::WriteAllText($f, '{"app":"Bluechip Board","version":1,"exported":"2026-10-03T15:32:10.177Z","etfLots":[]}')
    $b = Read-Backup $f
    Check 'old backup (version 1) is read' ($b -and (Get-InstanteBackup $b).UtcDateTime -eq [datetime]'2026-10-03T15:32:10.177')
    [IO.File]::WriteAllText($f, '{"app":"Bluechip Board","version":3,"saved":"2026-01-01T00:00:00.000Z","buys":[{"id":"o","a":"NVDA","d":"2025-12-01","q":2,"p":150}],"lots":[]}')
    $b3 = Read-Backup $f
    [IO.File]::WriteAllText($f, '{"app":"Bluechip Board","version":4,"saved":"2026-10-03T12:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{}}')
    $b4 = Read-Backup $f
    Check 'backups of versions 3 and 4 (no targets) are read' ($b3 -and @($b3.buys).Count -eq 1 -and $b4 -and (Get-InstanteBackup $b4).UtcDateTime -eq [datetime]'2026-10-03T12:00:00' -and $null -eq $b4.targets)
    [IO.File]::WriteAllText($f, '{"app":"Bluechip Board","version":5,"saved":"2026-10-06T09:00:00.000Z","buys":[],"lots":[],"sales":[],"deleted":{},"targets":{"weights":{"SXR8":60,"EUNK":30,"BTC":10},"band":5,"monthly":300,"at":"2026-10-06T09:00:00.000Z"}}')
    $b5 = Read-Backup $f
    $j5 = ($b5 | ConvertTo-Json -Depth 10 -Compress)
    # (o PowerShell 7 lê as datas ISO como DateTime e reescreve-as sem milissegundos, como já fazia com "saved": o instante é o mesmo)
    $at5 = [regex]::Match($j5, '"targets":\{"weights":\{"SXR8":60,"EUNK":30,"BTC":10\},"band":5,"monthly":300,"at":"([^"]+)"\}').Groups[1].Value
    Check 'version-5 backup with targets is read, and the targets survive the copy into the site' ($b5 -and $b5.version -eq 5 -and $b5.targets.weights.SXR8 -eq 60 -and $at5 -and ([DateTimeOffset]::Parse($at5, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime -eq [datetime]'2026-10-06T09:00:00') $j5
    [IO.File]::WriteAllText($f, '{"app":"Bluechip Board","version":5,"saved":"2026-10-06T09:00:00.000Z","buys":[{"id":"n1","a":"SXR8","d":"2026-01-05","q":1,"p":600}],"lots":[],"sales":[],"deleted":{},"policy":{"horizon":"15 years","allocation":"","monthly":"","drop20":"Keep going","drop30":"","sell":"","at":"2026-10-06T09:00:00.000Z"},"notes":{"n1":{"t":"Core position","at":"2026-10-06T09:00:00.000Z"}}}')
    $b6 = Read-Backup $f; $j6 = ($b6 | ConvertTo-Json -Depth 10 -Compress)
    Check 'version-5 backup with a policy and notes is read, and both survive the copy into the site' ($b6 -and $b6.policy.drop20 -eq 'Keep going' -and $b6.notes.n1.t -eq 'Core position' -and $j6.Contains('"drop20":"Keep going"') -and $j6.Contains('"notes":{"n1":{"t":"Core position"')) $j6
    [IO.File]::WriteAllText($f, '{ not json')
    Check 'corrupt backup is rejected' ($null -eq (Read-Backup $f))
    [IO.File]::WriteAllText($f, '{"app":"Other"}')
    Check 'file from another app is rejected' ($null -eq (Read-Backup $f))
} finally { Remove-Item $tmp -Recurse -Force }

Write-Host 'Remediation of 7 Oct 2026'
# --- datas independentes da cultura do Windows (um calendário não gregoriano, como o tailandês, dava "2569-…") ---
$culturaAntes = [Threading.Thread]::CurrentThread.CurrentCulture
try {
    [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('th-TH')
    ${function:Get-Url} = { param([string]$Url, [string]$UserAgent, [int]$Timeout) return $Script:Resposta }
    $Script:Resposta = Yahoo 'AAPL' 'USD' @(1, 2, 3, 4, 5, 6)
    $s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
    $Script:Resposta = YahooDiv 'AAPL' 'USD' '333.69' @(@(147, '0.27'), @(55, '0.27'))
    $d = Get-Dividendo $aapl
    Check 'Thai (Buddhist) calendar in Windows: price and dividend dates stay Gregorian ISO (2025-…, not 2568-…)' ($s.pontos[0][0] -eq '2025-09-30' -and "$($d.pagamentos[0][0])" -like '2026-*') "$($s.pontos[0][0]) / $($d.pagamentos[0][0])"
} finally { [Threading.Thread]::CurrentThread.CurrentCulture = $culturaAntes }
$partesPs = $texto.Substring(0, $texto.IndexOf('function Get-Plantilla {')) + $texto.Substring($texto.IndexOf('# 4. EXECU'))
Check 'static: every date formatted by the script names its culture (no ToString(''yyyy-MM-dd'') on its own)' (-not ([regex]::IsMatch($partesPs, "ToString\('yyyy-MM-dd[^']*'\)")))

# --- pontos de preço isolados e impossíveis (um "tick" errado do fornecedor) ---
$Script:Resposta = Yahoo 'AAPL' 'USD' @(100, 101, 1010, 102, 103, 104)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
Check 'an isolated price 10x its two neighbours is dropped and counted' ($s -and $s.pontos.Count -eq 5 -and @($s.pontos | Where-Object { $_[1] -eq 1010 }).Count -eq 0 -and $Script:Fontes[-1].erro -match '1 invalid') "$($s.pontos.Count) / $($Script:Fontes[-1].erro)"
$Script:Resposta = Yahoo 'AAPL' 'USD' @(100, 101, 10.1, 10.2, 10.3, 10.4)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
Check 'a lasting change of level (like an unadjusted split) is kept: never "corrected" silently' ($s -and $s.pontos.Count -eq 6) $s.pontos.Count
$Script:Resposta = Yahoo 'AAPL' 'USD' @(100, 101, 102, 103, 104, 300)
$s = Get-Serie @{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = ''; Moeda = 'USD' }
Check 'a huge jump on the latest price is kept but flagged in the source row' ($s -and $s.pontos.Count -eq 6 -and $Script:Fontes[-1].erro -match 'latest price \+188% from the previous close') $Script:Fontes[-1].erro
$r = Remove-PicoIsolado @(@('2026-01-01', 50), @('2026-01-02', 20), @('2026-01-03', 51))
Check 'a deep isolated dip (to 40% of both neighbours) is also dropped' ($r.removidos -eq 1 -and @($r.pontos).Count -eq 2)

# --- Kraken: os mesmos controlos de preço do Yahoo ---
$kr = '{"error":[],"result":{"XXBTZEUR":[' + ((0..9 | ForEach-Object { $t = $Script:Agora.AddDays(-9 + $_).ToUnixTimeSeconds(); $c = @('60000', '0', '-5', 'abc', '61000', '62000', '63000', '64000', '65000', '66000')[$_]; "[$t,""1"",""1"",""1"",""$c"",""1"",""1"",1]" }) -join ',') + '],"last":1}}'
${function:Get-Url} = { param([string]$Url, [string]$UserAgent, [int]$Timeout) if ($Url -match 'kraken') { return $kr }; throw 'Yahoo offline' }
$s = Get-Serie ($Ativos | Where-Object { $_.Id -eq 'BTC' })
Check 'Kraken fallback: zero, negative and non-numeric closes are dropped and counted' ($s -and $s.fonte -eq 'Kraken' -and $s.pontos.Count -eq 7 -and @($s.pontos | Where-Object { $_[1] -le 0 }).Count -eq 0 -and $Script:Fontes[-1].erro -match '3 invalid') "$($s.pontos.Count) / $($Script:Fontes[-1].erro)"

# --- CoinGecko: um campo em falta fica vazio (nunca 0,00 % nem 0 %) ---
${function:Get-Url} = { param([string]$Url, [string]$UserAgent, [int]$Timeout)
    if ($Url -match 'simple/price') { return '{"bitcoin":{"eur":60000,"usd":70000,"eur_market_cap":1200000000000}}' }
    if ($Url -match 'coingecko.com/api/v3/global') { return '{"data":{"market_cap_percentage":{"eth":12},"total_market_cap":{"eur":2000000000000}}}' }
    throw 'offline' }
$b = Get-DadosBitcoin
Check 'CoinGecko without the 24 h change: null (the site falls back to the change since 00:00 UTC), not 0.00%' ($b.mercado -and $b.mercado.eur -eq 60000 -and $null -eq $b.mercado.var24 -and $null -eq $b.mercado.volEur -and $b.mercado.capEur -eq 1200000000000) ($b.mercado | ConvertTo-Json -Compress)
Check 'CoinGecko without the Bitcoin dominance: no 0% shown, the error is listed' ($null -eq $b.mercado.dominio -and @($Script:Fontes | Where-Object { $_.nome -eq 'CoinGecko: Bitcoin dominance' })[-1].estado -eq 'error')

# --- e-mail da SEC: da configuração local, nunca dos argumentos ---
$cfgDir = Join-Path $env:TEMP ('bb-cfg-' + [guid]::NewGuid().ToString('N').Substring(0, 6)); New-Item -ItemType Directory $cfgDir | Out-Null
$envAntes = $env:BLUECHIP_SEC_EMAIL
try {
    $env:BLUECHIP_SEC_EMAIL = ''
    $e0 = Get-EmailSecLocal $cfgDir
    [IO.File]::WriteAllText((Join-Path $cfgDir 'bluechip-board.config.json'), '{ "secEmail": "  me@example.org " }')
    $e1 = Get-EmailSecLocal $cfgDir
    [IO.File]::WriteAllText((Join-Path $cfgDir 'bluechip-board.config.json'), '{ "secEmail": "your@email.com" }')
    $e2 = Get-EmailSecLocal $cfgDir
    [IO.File]::WriteAllText((Join-Path $cfgDir 'bluechip-board.config.json'), '{ broken')
    $e3 = Get-EmailSecLocal $cfgDir
    $env:BLUECHIP_SEC_EMAIL = 'env@example.net'
    $e4 = Get-EmailSecLocal $cfgDir
    Check 'SEC e-mail: none without configuration (no warning); trimmed from bluechip-board.config.json' ($e0.email -eq '' -and $e0.aviso -eq '' -and $e1.email -eq 'me@example.org' -and $e1.aviso -eq '')
    Check 'SEC e-mail: the example placeholder and an unreadable file are refused with a warning' ($e2.email -eq '' -and $e2.aviso -match 'not valid' -and $e3.email -eq '' -and $e3.aviso -match 'could not be read')
    Check 'SEC e-mail: BLUECHIP_SEC_EMAIL wins over the file' ($e4.email -eq 'env@example.net')
} finally { $env:BLUECHIP_SEC_EMAIL = $envAntes; Remove-Item $cfgDir -Recurse -Force }
Check 'static: the scheduled task never gets the SEC e-mail in its arguments' ($texto -notmatch '\$argumentos \+= " -EmailSEC' -and $texto -match 'The SEC e-mail is NOT put in the task')

# --- Nasdaq: segunda fonte para as ações dos EUA, só se coincidir com a execução anterior ---
function NasdaqHist([string]$Simbolo, $Pts) { '{"data":{"symbol":"' + $Simbolo + '","totalRecords":' + @($Pts).Count + ',"tradesTable":{"rows":[' + ((@($Pts) | Sort-Object { $_[0] } -Descending | ForEach-Object { $q = $_[0].Split('-'); '{"date":"' + "$($q[1])/$($q[2])/$($q[0])" + '","close":"$' + ([double]$_[1]).ToString('#,##0.00', $Script:Inv) + '","volume":"1"}' }) -join ',') + ']}},"status":{"rCode":200}}' }
$refN = @(for ($i = 0; $i -lt 40; $i++) { , @($Script:Agora.UtcDateTime.AddDays(-60 + $i).ToString('yyyy-MM-dd', $Script:Inv), (1000 + $i)) })
$hojeNy = [TimeZoneInfo]::ConvertTime($Script:Agora, (Get-FusoNY)).ToString('yyyy-MM-dd', $Script:Inv)
$nasPts = @($refN) + @(, @($hojeNy, 1100))
$Script:NasResp = NasdaqHist 'AAPL' $nasPts
${function:Get-Url} = { param([string]$Url, [string]$UserAgent, [int]$Timeout) if ($Url -match 'api\.nasdaq\.com/api/quote/AAPL/historical') { return $Script:NasResp }; throw 'Yahoo offline' }
$aaplN = $Ativos | Where-Object { $_.Id -eq 'AAPL' }
$s = Get-Serie $aaplN -Referencia $refN
Check 'Yahoo down: US stock prices from Nasdaq, checked against the previous run, labelled "ok (Nasdaq fallback)"' ($s -and $s.fonte -eq 'Nasdaq' -and $s.moeda -eq 'USD' -and $Script:Fontes[-1].estado -eq 'ok (Nasdaq fallback)' -and $s.pontos[0][1] -eq 1000) ($Script:Fontes[-1] | ConvertTo-Json -Compress)
Check 'Nasdaq: today''s session still open in New York (before 16:15) is not taken as a close' (@($s.pontos | Where-Object { $_[0] -eq $hojeNy }).Count -eq 0 -and $s.pontos.Count -eq 40) "$($s.pontos.Count) $hojeNy"
$Script:NasResp = NasdaqHist 'AAPL' @($refN | ForEach-Object { , @($_[0], ($_[1] * 10)) })
$s = Get-Serie $aaplN -Referencia $refN
Check 'Nasdaq prices on another split basis (10x the previous run): refused, nothing invented' ($null -eq $s -and $Script:Fontes[-1].estado -eq 'error' -and $Script:Fontes[-1].erro -match 'Nasdaq prices do not match') $Script:Fontes[-1].erro
$Script:NasResp = NasdaqHist 'AAPL' $refN
$s = Get-Serie $aaplN -Referencia @($refN | Select-Object -First 10)
Check 'Nasdaq without enough dates to compare with the previous run: not used' ($null -eq $s -and $Script:Fontes[-1].erro -match 'could not be checked')
$s = Get-Serie $aaplN
Check 'Nasdaq on a fresh install (no previous run): not used' ($null -eq $s -and $Script:Fontes[-1].erro -match 'could not be checked')
$Script:NasResp = NasdaqHist 'MSFT' $refN
$s = Get-Serie $aaplN -Referencia $refN
Check 'Nasdaq answer for another symbol: refused' ($null -eq $s -and $Script:Fontes[-1].erro -match "returned 'MSFT'")
$Script:NasResp = '{"data":null,"message":null,"status":{"rCode":400,"bCodeMessage":[{"code":1001,"errorMessage":"Symbol not exists."}]}}'
$s = Get-Serie $aaplN -Referencia $refN
Check 'Nasdaq error answer: refused with its message' ($null -eq $s -and $Script:Fontes[-1].erro -match 'Symbol not exists')

# --- Get-Url: uma resposta definitiva (404) não é repetida; um erro temporário é ---
Add-Type -TypeDefinition 'public class BbTesteResp { public int StatusCode; public BbTesteResp(int c) { StatusCode = c; } } public class BbTesteHttp : System.Exception { public object Response; public BbTesteHttp(string m, int c) : base(m) { Response = new BbTesteResp(c); } }' -ErrorAction SilentlyContinue
$fnUrl = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-Url' }, $true) | Select-Object -First 1
. ([scriptblock]::Create($fnUrl.Extent.Text.Replace('function Get-Url {', 'function Get-UrlOriginal {')))
$Script:Chamadas = 0
function Invoke-WebRequest { $Script:Chamadas++; throw (New-Object BbTesteHttp ('HTTP ' + $Script:Codigo), $Script:Codigo) }
function Start-Sleep { }
$Script:Codigo = 404; try { Get-UrlOriginal 'https://example.com/x' | Out-Null } catch { }
$c404 = $Script:Chamadas; $Script:Chamadas = 0
$Script:Codigo = 500; try { Get-UrlOriginal 'https://example.com/x' | Out-Null } catch { }
$c500 = $Script:Chamadas
Remove-Item function:Invoke-WebRequest, function:Start-Sleep
Check 'download: a permanent answer (404) is not retried; a server error (500) is retried once' ($c404 -eq 1 -and $c500 -eq 2) "404: $c404 call(s), 500: $c500"

# --- agrupamento de duplicadas: palavras (o mesmo resultado do ciclo de antes) e resultado independente da ordem de entrada ---
$pal = { param($t) (@((Get-PalavrasTitulo $t) | ForEach-Object { $_ })) -join ' ' }
Check 'title words: accents, curly apostrophes, amounts, suffixes (stem of 4+ letters), synonyms and stop words, in order' ((& $pal ("California man$([char]0x2019)s smuggling of `$300 million in Nvidia chips: shares tumble – Reuters")) -eq 'california man smuggle 300m nvidia chip fall' -and (& $pal 'Ações da Apple sobem; things rallies kings -ASML- u.s.') -eq 'apple sobem ralli king asml u.s') "$(& $pal ("California man$([char]0x2019)s smuggling of `$300 million in Nvidia chips: shares tumble – Reuters")) / $(& $pal 'Ações da Apple sobem; things rallies kings -ASML- u.s.')"
$k = 0; $base = @(foreach ($tt in @($Fundo + @('Nvidia shares rise after strong data center sales', 'Nvidia shares rise after strong data center sales growth', 'Nvidia shares rise on strong data center sales'))) { $m = Noticia $tt; if ($m) { $m.data = '2026-09-25T09:00:00.0000000+00:00'; $m.score = 3; $m.chave = "t$k"; $k++; $m } })
$sig = { param($g) (@($g | ForEach-Object { "$($_.chave)<" + ((@($_.outras | ForEach-Object { $_.chave }) | Sort-Object) -join ',') }) | Sort-Object) -join ';' }
$copia = { param($x) @(foreach ($m in $x) { $c = $m.PSObject.Copy(); $c.outras = @(); $c }) }
$a1 = & $sig @(Join-NoticiasDuplicadas (& $copia $base)); [array]::Reverse($base); $a2 = & $sig @(Join-NoticiasDuplicadas (& $copia $base))
Check 'duplicate grouping does not depend on the order of the input (ties in score and date broken by the title key)' ($a1 -eq $a2 -and $a1 -match '<t\d+') "$a1 | $a2"

# --- iShares: mudança de formato do ficheiro dita na fonte ---
${function:Get-Url} = { param([string]$Url, [string]$UserAgent, [int]$Timeout) return $Script:Ficheiros['251861'] }
$Script:Ficheiros['251861'] = (iShares 'IE00B4K48X80' $linhasEU).Replace('>Market Currency<', '>Currency X<')
$r = Get-PesosETF $eunk
Check 'iShares file without the Market Currency column: weights still read, the source row names the missing column and the Unavailable breakdown' ($r.aoVivo -and $Script:Fontes[-1].estado -eq 'ok' -and $Script:Fontes[-1].erro -match 'no column Market Currency' -and $Script:Fontes[-1].erro -match 'currency breakdown Unavailable') $Script:Fontes[-1].erro

Write-Host 'Static checks'
Check 'TLS is not pinned to 1.2 when Windows chooses (SystemDefault)' ($texto -notmatch '\]::SecurityProtocol -bor \[Net\.SecurityProtocolType\]::Tls12')
Check 'the "<" escape is built from parts' ($texto.Contains("'\' + 'u003c'"))
Check 'PowerShell 7 accepts the SEC User-Agent with an e-mail (header check skipped only on 6+)' ($texto -match "PSVersion\.Major -ge 6\) \{ \`$extra\.SkipHeaderValidation = \`$true \}" -and $texto -match "'en-US,en;q=0\.8,pt-PT;q=0\.6' \} @extra")
Check 'script file has a UTF-8 BOM' (([IO.File]::ReadAllBytes($scriptPath)[0..2] -join ',') -eq '239,187,191')

Write-Host ''
if ($Script:Falhas) { Write-Host "$($Script:Falhas) of $($Script:Total) checks FAILED" -ForegroundColor Red; exit 1 }
Write-Host "All $($Script:Total) checks passed" -ForegroundColor Green
