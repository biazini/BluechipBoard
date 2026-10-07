<#
.SYNOPSIS
    Bluechip Board: news and prices for Apple (AAPL), NVIDIA (NVDA), Alphabet (GOOGL),
    iShares Core S&P 500 UCITS ETF (Acc) - IE00B5BMR087 (SXR8), iShares Core MSCI Europe (EUNK),
    iShares Core MSCI EM IMI (IS3N), iShares Core MSCI Japan IMI (EUNN) and Bitcoin (BTC, in euros).

.DESCRIPTION
    Collects news from public sources (Google News, Yahoo Finance, company newsrooms, the Federal Reserve,
    the ECB, crypto press and, optionally, SEC EDGAR), 1-year prices, long-term price histories, EUR/USD,
    the ETF holdings and Bitcoin indicators (sentiment, dominance, network and halving).
    Classifies each story by potential impact (material, important, moderate, noise) with keyword rules,
    and builds a local website (a single HTML file) with tabs, filters, charts and your portfolio.

    It is an educational, rule-based tool: not financial advice.

.PARAMETER Dias
    News window, in days (default 7). Alias: -Days.

.PARAMETER Pasta
    Folder where the website and data are saved (default: the folder of this script). Alias: -Folder.

.PARAMETER EmailSEC
    Your e-mail, sent in the User-Agent header to the SEC (an SEC access rule).
    Without it, the SEC sources are skipped. Alias: -SecEmail.

.PARAMETER NaoAbrir
    Does not open the website in the browser at the end. Alias: -NoOpen.

.PARAMETER AgendarDiariamente
    Registers a Windows Task Scheduler task that runs every day at the given time. Alias: -ScheduleDaily.

.PARAMETER Hora
    Time of the daily run (default 08:30). Alias: -Time.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Bluechip-Board.ps1

.EXAMPLE
    .\Bluechip-Board.ps1 -Days 3 -SecEmail "your@email.com"

.EXAMPLE
    .\Bluechip-Board.ps1 -ScheduleDaily -Time 08:30 -SecEmail "your@email.com"
#>
[CmdletBinding()]
param(
    [Alias('Days')][ValidateRange(1, 30)][int]$Dias = 7,
    [Alias('Folder')][string]$Pasta = '',
    [Alias('SecEmail')][string]$EmailSEC = '',
    [Alias('NoOpen')][switch]$NaoAbrir,
    [Alias('ScheduleDaily')][switch]$AgendarDiariamente,
    [Alias('Time')][string]$Hora = '08:30'
)

# Mensagens em inglês na consola, incluindo os erros do próprio Windows/.NET (ex.: falhas de rede), seja qual for a língua do Windows
try { [Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; [Globalization.CultureInfo]::DefaultThreadCurrentUICulture = 'en-US' } catch { }

if (-not $Pasta) { $Pasta = $PSScriptRoot }   # por defeito, os resultados ficam junto do script
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # acelera Invoke-WebRequest no Windows PowerShell 5.1
# TLS: when .NET already lets Windows choose (SystemDefault = 0, which allows TLS 1.3), leave it alone; forcing Tls12 there
# would pin the connection to TLS 1.2. Only add TLS 1.2 on old .NET setups that list protocols explicitly without it.
try {
    $protocolo = [int][Net.ServicePointManager]::SecurityProtocol
    if ($protocolo -ne 0 -and -not ($protocolo -band 3072)) { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]($protocolo -bor 3072) }
} catch { }

# Browser identity sent to news and price sites (some refuse non-browser clients). Update the Chrome version now and then:
# $Script:UAChrome is the major version and $Script:UAData the day it was current (Chrome ships a new one every 4 weeks), so a
# maintenance reminder appears when it is about 8 versions behind.
$Script:UAChrome = 154; $Script:UAData = '2026-10-07'
$Script:UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/$($Script:UAChrome).0.0.0 Safari/537.36"
$Script:Inv = [Globalization.CultureInfo]::InvariantCulture
$Script:Agora = [DateTimeOffset]::UtcNow
$Script:Fontes = New-Object System.Collections.Generic.List[object]

# ============================================================================
# 1. CONFIGURAÇÃO (podes editar)
# ============================================================================

$Ativos = @(
    # Nasdaq: alternativa ao Yahoo para o preço de 1 ano (validada contra a execução anterior; ver Get-SerieNasdaq)
    @{ Id = 'AAPL';  Nome = 'Apple';    Yahoo = 'AAPL';    Stooq = 'aapl.us'; Nasdaq = 'AAPL';  Moeda = 'USD' },
    @{ Id = 'NVDA';  Nome = 'NVIDIA';   Yahoo = 'NVDA';    Stooq = 'nvda.us'; Nasdaq = 'NVDA';  Moeda = 'USD' },
    @{ Id = 'GOOGL'; Nome = 'Alphabet'; Yahoo = 'GOOGL';   Stooq = 'googl.us'; Nasdaq = 'GOOGL'; Moeda = 'USD' },
    @{ Id = 'SXR8';  Nome = 'iShares Core S&P 500 (SXR8)'; Yahoo = 'SXR8.DE'; Stooq = 'sxr8.de'; Moeda = 'EUR' },
    # Outros ETF da iShares, na Xetra em euros (a mesma bolsa e moeda do SXR8). Metadados e holdings em $ETFs, abaixo.
    @{ Id = 'EUNK';  Nome = 'iShares Core MSCI Europe (EUNK)'; Yahoo = 'EUNK.DE'; Stooq = ''; Moeda = 'EUR' },
    @{ Id = 'IS3N';  Nome = 'iShares Core MSCI EM IMI (IS3N)'; Yahoo = 'IS3N.DE'; Stooq = ''; Moeda = 'EUR' },
    @{ Id = 'EUNN';  Nome = 'iShares Core MSCI Japan IMI (EUNN)'; Yahoo = 'EUNN.DE'; Stooq = ''; Moeda = 'EUR' },
    # Bitcoin em euros: negoceia 24 h por dia, 7 dias por semana (um ponto por dia, em UTC); alternativa: Kraken (BTC/EUR)
    @{ Id = 'BTC';   Nome = 'Bitcoin';  Yahoo = 'BTC-EUR'; Stooq = '';  Kraken = 'XBTEUR'; Moeda = 'EUR' }
)
$Mercado = @(
    @{ Id = 'GSPC'; Nome = 'S&P 500';                  Yahoo = '^GSPC'; Stooq = '^spx' },
    @{ Id = 'VIX';  Nome = 'VIX (volatility)';         Yahoo = '^VIX';  Stooq = '' },
    @{ Id = 'TNX';  Nome = '10-year Treasury yield';   Yahoo = '^TNX';  Stooq = '' }
)

# Pesos de referência no ETF (dados BlackRock de 30 jul 2026), usados se o download ao vivo falhar
$PesosReferencia = @{ AAPL = 7.81; NVDA = 7.46; GOOGL = 5.52; Data = '2026-07-30' }
# Ficheiro "Detailed Holdings and Analytics" do site da iShares (Excel 2003 XML; a folha "Holdings" tem as posições)
$UrlPesosETF = 'https://www.blackrock.com/varnish-api/uk-retail01-product-data/product-data/api/v1/get-fund-document?appType=PRODUCT_PAGE&appSubType=ISHARES&targetSite=ishares-uk&locale=en_GB&portfolioId=253743&component=fundDownloadV2&userType=individual'
# O mesmo ficheiro para outro fundo: muda só o portfolioId (o número da página do fundo no site da iShares UK)
function New-UrlPesosETF([int]$PortfolioId) { $UrlPesosETF -replace 'portfolioId=\d+', "portfolioId=$PortfolioId" }

# ETF acompanhados: metadados, ficheiro de holdings de cada fundo e pesos de referência (usados se o download falhar).
# Id = o mesmo de $Ativos. Empresas = as empresas do painel que o fundo deve conter: os pesos delas têm de ser
# plausíveis (no SXR8, AAPL, NVDA e GOOGL); nos outros fundos não estão lá (0%, confirmado nos ficheiros de 2 out 2026).
# Indice = nome curto do índice, usado no site. Acumulacao = $true: os dividendos ficam no preço (sem dividendos projetados).
$ETFs = @(
    @{ Id = 'SXR8'; Fundo = 'iShares Core S&P 500 UCITS ETF USD (Acc)'; Isin = 'IE00B5BMR087'; Indice = 'S&P 500'; Benchmark = 'S&P 500 Index'
       MoedaBase = 'USD'; Domicilio = 'IE'; Acumulacao = $true; Gestora = 'iShares (BlackRock)'; Bolsa = 'DE'; Inicio = '2010-05-19'
       PaisIndice = 'United States'   # sem repartição por país no ficheiro da iShares: o país do índice, assinalado no site (aproximado)
       Pagina = 'https://www.ishares.com/uk/individual/en/products/253743/'; Url = $UrlPesosETF; Referencia = $PesosReferencia; Empresas = @('AAPL', 'NVDA', 'GOOGL') },
    @{ Id = 'EUNK'; Chip = 'Europe'; Fundo = 'iShares Core MSCI Europe UCITS ETF EUR (Acc)'; Isin = 'IE00B4K48X80'; Indice = 'MSCI Europe'; Benchmark = 'MSCI Europe Index'
       MoedaBase = 'EUR'; Domicilio = 'IE'; Acumulacao = $true; Gestora = 'iShares (BlackRock)'; Bolsa = 'DE'; Inicio = '2009-09-25'
       Pagina = 'https://www.ishares.com/uk/individual/en/products/251861/'; Url = (New-UrlPesosETF 251861); Referencia = @{ AAPL = 0; NVDA = 0; GOOGL = 0; Data = '2026-10-02' }; Empresas = @() },
    @{ Id = 'IS3N'; Chip = 'Emerging markets'; Fundo = 'iShares Core MSCI EM IMI UCITS ETF USD (Acc)'; Isin = 'IE00BKM4GZ66'; Indice = 'MSCI EM IMI'; Benchmark = 'MSCI Emerging Markets Investable Market Index (Net)'
       MoedaBase = 'USD'; Domicilio = 'IE'; Acumulacao = $true; Gestora = 'iShares (BlackRock)'; Bolsa = 'DE'; Inicio = '2014-05-30'
       Pagina = 'https://www.ishares.com/uk/individual/en/products/264659/'; Url = (New-UrlPesosETF 264659); Referencia = @{ AAPL = 0; NVDA = 0; GOOGL = 0; Data = '2026-10-02' }; Empresas = @() },
    @{ Id = 'EUNN'; Chip = 'Japan'; Fundo = 'iShares Core MSCI Japan IMI UCITS ETF USD (Acc)'; Isin = 'IE00B4L5YX21'; Indice = 'MSCI Japan IMI'; Benchmark = 'MSCI Japan Investable Market Index (Net)'
       MoedaBase = 'USD'; Domicilio = 'IE'; Acumulacao = $true; Gestora = 'iShares (BlackRock)'; Bolsa = 'DE'; Inicio = '2009-09-25'
       PaisIndice = 'Japan'           # idem (índice de um só país)
       Pagina = 'https://www.ishares.com/uk/individual/en/products/251867/'; Url = (New-UrlPesosETF 251867); Referencia = @{ AAPL = 0; NVDA = 0; GOOGL = 0; Data = '2026-10-02' }; Empresas = @() }
)

# Ativos com dividendos projetados na carteira. Os ETF de acumulação ficam de fora (os dividendos são reinvestidos no
# preço; Get-Dividendos ignora qualquer ativo marcado Acumulacao em $ETFs) e a Bitcoin não paga dividendos.
$AtivosDividendos = @('AAPL', 'NVDA', 'GOOGL')

# Calendário (o texto aparece no site, em inglês): atualiza as datas quando as empresas as confirmarem.
# imp: 'Very high', 'High' ou 'Medium'. Estado: C = confirmada, E = estimada.
# Reuniões dos bancos centrais (C), das páginas oficiais, lidas a 6 out 2026: federalreserve.gov/monetarypolicy/fomccalendars.htm
# e ecb.europa.eu/press/calendars/mgcgc. A data é a da decisão (2.º dia; no BCE, com a conferência de imprensa). As do BCE
# ficam com imp 'Medium' para não criarem alertas novos no Overview (os eventos 'High' geram um alerta 7 dias antes).
$Calendario = @(
    @{ d = '2026-10-07'; e = 'MKT';   ev = 'Minutes of the Fed''s September meeting';            imp = 'Medium';    st = 'E' },
    @{ d = '2026-10-28'; e = 'MKT';   ev = 'Fed decision (27–28 Oct meeting)';                   imp = 'High';      st = 'C' },
    @{ d = '2026-10-28'; e = 'GOOGL'; ev = 'Q3 2026 earnings';                                   imp = 'High';      st = 'E' },
    @{ d = '2026-10-29'; e = 'AAPL';  ev = 'FQ4 2026 earnings and holiday-quarter guidance';      imp = 'High';      st = 'E' },
    @{ d = '2026-10-29'; e = 'MKT';   ev = 'ECB monetary policy decision (28–29 Oct meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2026-11-18'; e = 'NVDA';  ev = 'Q3 FY2027 earnings';                                 imp = 'Very high'; st = 'E' },
    @{ d = '2026-12-09'; e = 'MKT';   ev = 'Fed decision with projections (8–9 Dec)';            imp = 'High';      st = 'C' },
    @{ d = '2026-12-17'; e = 'MKT';   ev = 'ECB monetary policy decision (16–17 Dec meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2027-01-27'; e = 'MKT';   ev = 'Fed decision (26–27 Jan meeting)';                   imp = 'High';      st = 'C' },
    @{ d = '2027-01-28'; e = 'AAPL';  ev = 'Holiday-quarter earnings';                           imp = 'Very high'; st = 'E' },
    @{ d = '2027-02-03'; e = 'GOOGL'; ev = 'Q4 2026 earnings and 2027 CapEx';                    imp = 'Very high'; st = 'E' },
    @{ d = '2027-02-04'; e = 'MKT';   ev = 'ECB monetary policy decision (3–4 Feb meeting)';     imp = 'Medium';    st = 'C' },
    @{ d = '2027-02-24'; e = 'NVDA';  ev = 'Q4 FY2027 earnings and FY2028 outlook';              imp = 'Very high'; st = 'E' },
    @{ d = '2027-03-16'; e = 'NVDA';  ev = 'GTC, annual conference';                             imp = 'High';      st = 'E' },
    @{ d = '2027-03-17'; e = 'MKT';   ev = 'Fed decision with projections (16–17 Mar)';          imp = 'High';      st = 'C' },
    @{ d = '2027-03-18'; e = 'MKT';   ev = 'ECB monetary policy decision (17–18 Mar meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2027-04-28'; e = 'MKT';   ev = 'Fed decision (27–28 Apr meeting)';                   imp = 'High';      st = 'C' },
    @{ d = '2027-04-29'; e = 'MKT';   ev = 'ECB monetary policy decision (28–29 Apr meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2027-06-09'; e = 'MKT';   ev = 'Fed decision with projections (8–9 Jun)';            imp = 'High';      st = 'C' },
    @{ d = '2027-06-10'; e = 'MKT';   ev = 'ECB monetary policy decision (9–10 Jun meeting)';    imp = 'Medium';    st = 'C' },
    @{ d = '2027-06-30'; e = 'TU';    ev = 'Deadline to file your 2026 IRS tax return (Portugal)'; imp = 'High';    st = 'C' },
    @{ d = '2027-07-22'; e = 'MKT';   ev = 'ECB monetary policy decision (21–22 Jul meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2027-07-28'; e = 'MKT';   ev = 'Fed decision (27–28 Jul meeting)';                   imp = 'High';      st = 'C' },
    @{ d = '2027-09-09'; e = 'MKT';   ev = 'ECB monetary policy decision (8–9 Sep meeting)';     imp = 'Medium';    st = 'C' },
    @{ d = '2027-09-15'; e = 'MKT';   ev = 'Fed decision with projections (14–15 Sep)';          imp = 'High';      st = 'C' },
    @{ d = '2027-10-27'; e = 'MKT';   ev = 'Fed decision (26–27 Oct meeting)';                   imp = 'High';      st = 'C' },
    @{ d = '2027-10-28'; e = 'MKT';   ev = 'ECB monetary policy decision (27–28 Oct meeting)';   imp = 'Medium';    st = 'C' },
    @{ d = '2027-12-08'; e = 'MKT';   ev = 'Fed decision with projections (7–8 Dec)';            imp = 'High';      st = 'C' },
    @{ d = '2027-12-16'; e = 'MKT';   ev = 'ECB monetary policy decision (15–16 Dec meeting)';   imp = 'Medium';    st = 'C' }
)

# Horário das bolsas, na hora local de cada bolsa (o site converte para a hora de Lisboa, com as mudanças de hora).
# Os feriados habituais e as sessões mais curtas são calculados no site pelas regras de cada bolsa (servem para qualquer ano).
# As listas abaixo juntam-se a esses e só precisam de ser alteradas para exceções (ex.: um encerramento extraordinário).
$Bolsas = @(
    @{ id = 'US'; curto = 'US'; nome = 'New York exchanges'; detalhe = 'NYSE and Nasdaq: AAPL, NVDA, GOOGL'; tz = 'America/New_York'; abre = '09:30'; fecha = '16:00'
       feriados = @('2026-01-01', '2026-01-19', '2026-02-16', '2026-04-03', '2026-05-25', '2026-06-19', '2026-07-03', '2026-09-07', '2026-11-26', '2026-12-25',
                    '2027-01-01', '2027-01-18', '2027-02-15', '2027-03-26', '2027-05-31', '2027-06-18', '2027-07-05', '2027-09-06', '2027-11-25', '2027-12-24')
       curtos = @{ '2026-11-27' = '13:00'; '2026-12-24' = '13:00'; '2027-11-26' = '13:00' } },
    @{ id = 'DE'; curto = 'Xetra'; nome = 'Xetra (Frankfurt)'; detalhe = 'Deutsche Börse: SXR8, EUNK, IS3N, EUNN'; tz = 'Europe/Berlin'; abre = '09:00'; fecha = '17:30'
       feriados = @('2026-01-01', '2026-04-03', '2026-04-06', '2026-05-01', '2026-12-24', '2026-12-25', '2026-12-31',
                    '2027-01-01', '2027-03-26', '2027-03-29', '2027-12-24', '2027-12-31')
       curtos = @{} }
)

function New-GNewsUrl([string]$Consulta, [string]$Lingua = 'en') {
    $q = [uri]::EscapeDataString("$Consulta when:${Dias}d")
    if ($Lingua -eq 'pt') { return "https://news.google.com/rss/search?q=$q&hl=pt-PT&gl=PT&ceid=PT:pt-150" }
    return "https://news.google.com/rss/search?q=$q&hl=en-US&gl=US&ceid=US:en"
}
function New-YahooRss([string]$Simbolo) { "https://feeds.finance.yahoo.com/rss/2.0/headline?s=$([uri]::EscapeDataString($Simbolo))&region=US&lang=en-US" }

$Feeds = @(
    # Notícias por empresa (Google News agrega Reuters, Bloomberg, CNBC, etc.)
    @{ Nome = 'Google News: Apple (stock)';        Url = (New-GNewsUrl '"Apple" (stock OR shares OR AAPL)');              Dica = 'AAPL' },
    @{ Nome = 'Google News: Apple (products and regulation)'; Url = (New-GNewsUrl 'Apple (iPhone OR "App Store" OR Siri OR antitrust OR China)'); Dica = 'AAPL' },
    @{ Nome = 'Google News: NVIDIA (stock)';       Url = (New-GNewsUrl 'Nvidia (stock OR shares OR NVDA)');               Dica = 'NVDA' },
    @{ Nome = 'Google News: NVIDIA (chips, China, AI)'; Url = (New-GNewsUrl 'Nvidia (export OR China OR Blackwell OR Rubin OR hyperscaler OR capex)'); Dica = 'NVDA' },
    @{ Nome = 'Google News: Alphabet (stock)';     Url = (New-GNewsUrl '(Alphabet OR Google) (stock OR shares OR GOOGL)'); Dica = 'GOOGL' },
    @{ Nome = 'Google News: Alphabet (regulation, AI, Cloud)'; Url = (New-GNewsUrl 'Google (antitrust OR Gemini OR "Google Cloud" OR Waymo OR YouTube)'); Dica = 'GOOGL' },
    @{ Nome = 'Google News: Reuters on the three'; Url = (New-GNewsUrl 'site:reuters.com (Nvidia OR Apple OR Alphabet OR Google)'); Dica = '' },
    # Mercado, ETF e macro
    @{ Nome = 'Google News: S&P 500 and Wall Street'; Url = (New-GNewsUrl '"S&P 500" (stocks OR index OR futures)');        Dica = 'MKT' },
    @{ Nome = 'Google News: Fed, inflation and rates'; Url = (New-GNewsUrl '"Federal Reserve" (rates OR inflation OR Warsh)'); Dica = 'MKT' },
    @{ Nome = 'Google News: ETF SXR8 / CSPX';       Url = (New-GNewsUrl '"iShares Core S&P 500" OR SXR8 OR CSPX');          Dica = 'SXR8' },
    # Os outros ETF: o próprio fundo e o mercado que o índice cobre (é isso que move o preço do ETF)
    @{ Nome = 'Google News: Europe (MSCI Europe, STOXX 600)'; Url = (New-GNewsUrl '"MSCI Europe" OR "STOXX 600" OR "European stocks" OR "European shares"'); Dica = 'EUNK' },
    @{ Nome = 'Google News: Emerging markets (MSCI EM)'; Url = (New-GNewsUrl '"MSCI Emerging Markets" OR "emerging markets stocks" OR "emerging market stocks" OR "emerging-market stocks"'); Dica = 'IS3N' },
    @{ Nome = 'Google News: Japan (Nikkei, Topix)';  Url = (New-GNewsUrl '"Nikkei 225" OR Topix OR "Japanese stocks" OR "Japanese shares" OR "MSCI Japan"'); Dica = 'EUNN' },
    @{ Nome = 'Google News PT: US stocks';  Url = (New-GNewsUrl '(Apple OR Nvidia OR Alphabet OR Google) ações' 'pt'); Dica = '' },
    @{ Nome = 'Google News PT: ECB, euro and dollar'; Url = (New-GNewsUrl '(BCE OR Fed) (juros OR euro OR dólar)' 'pt');      Dica = 'MKT' },
    # Bitcoin. Nos feeds de imprensa cripto (Exigir = 'BTC') só entram títulos que falem da Bitcoin ou do mercado cripto,
    # para não encher a lista com outras moedas.
    @{ Nome = 'Google News: Bitcoin (market)';     Url = (New-GNewsUrl 'Bitcoin (price OR BTC OR ETF OR rally OR selloff)'); Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'Google News: Bitcoin (regulation and adoption)'; Url = (New-GNewsUrl 'Bitcoin (regulation OR SEC OR "strategic reserve" OR MiCA OR treasury OR halving)'); Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'Google News PT: Bitcoin';            Url = (New-GNewsUrl '(Bitcoin OR criptomoedas) (preço OR euros OR ETF OR regulação)' 'pt'); Dica = ''; Exigir = 'BTC' },
    # sem a barra final: com ela, o CoinDesk responde 308, que o Windows PowerShell 5.1 não segue
    @{ Nome = 'CoinDesk';                           Url = 'https://www.coindesk.com/arc/outboundfeeds/rss';   Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'The Block';                          Url = 'https://www.theblock.co/rss.xml';                  Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'Cointelegraph';                      Url = 'https://cointelegraph.com/rss';                    Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'Decrypt';                            Url = 'https://decrypt.co/feed';                          Dica = ''; Exigir = 'BTC' },
    @{ Nome = 'Yahoo Finance: Bitcoin';             Url = (New-YahooRss 'BTC-USD');                           Dica = ''; Exigir = 'BTC' },
    # Yahoo Finance
    @{ Nome = 'Yahoo Finance: AAPL';  Url = (New-YahooRss 'AAPL');  Dica = 'AAPL' },
    @{ Nome = 'Yahoo Finance: NVDA';  Url = (New-YahooRss 'NVDA');  Dica = 'NVDA' },
    @{ Nome = 'Yahoo Finance: GOOGL'; Url = (New-YahooRss 'GOOGL'); Dica = 'GOOGL' },
    @{ Nome = 'Yahoo Finance: S&P 500'; Url = (New-YahooRss '^GSPC'); Dica = 'MKT' },
    # Fontes primárias
    @{ Nome = 'Apple Newsroom';             Url = 'https://www.apple.com/newsroom/rss-feed.rss';                Dica = 'AAPL' },
    @{ Nome = 'NVIDIA Newsroom';            Url = 'https://nvidianews.nvidia.com/releases.xml';                 Dica = 'NVDA' },
    @{ Nome = 'Google Blog';                Url = 'https://blog.google/rss/';                                   Dica = 'GOOGL' },
    @{ Nome = 'Federal Reserve (monetary policy)'; Url = 'https://www.federalreserve.gov/feeds/press_monetary.xml'; Dica = 'MKT' },
    @{ Nome = 'ECB (press releases)';       Url = 'https://www.ecb.europa.eu/rss/press.html';                   Dica = 'MKT' }
)
$SecEmpresas = @(
    @{ Id = 'AAPL';  Cik = '0000320193' },
    @{ Id = 'NVDA';  Cik = '0001045810' },
    @{ Id = 'GOOGL'; Cik = '0001652044' }
)
# Indicadores da área do euro: ECB Data Portal (CSV, sem chave), chaves confirmadas no portal a 6 out 2026. Min/Max: o
# intervalo plausível (fora dele o valor é ignorado e contado). A taxa de depósito só traz as datas em que mudou.
$SeriesMacro = @(
    @{ Id = 'BCE_DFR'; Nome = 'deposit facility rate'; Fluxo = 'FM'; Chave = 'B.U2.EUR.4F.KR.DFR.LEV'; Min = -2; Max = 20; Desde = '2010-01' }
    @{ Id = 'HICP_EA'; Nome = 'euro area inflation (HICP, annual rate)'; Fluxo = 'ICP'; Chave = 'M.U2.N.000000.4.ANR'; Min = -5; Max = 25; Desde = '2010-01' }
)
# Fundamentais (SEC XBRL companyfacts): para cada métrica, a unidade aceite e as tags us-gaap por ordem de preferência.
# As empresas mudam de tag ao longo dos anos (a Apple passou de SalesRevenueNet para RevenueFromContract… em 2018; a NVIDIA
# usa Revenues e mudou o capex para PaymentsToAcquireProductiveAssets): cada período usa a tag preferida que o tiver.
# PorAcao: o valor é dividido pelos splits posteriores à entrega; Acoes: multiplicado (e nunca derivado por subtração).
$MetricasSEC = [ordered]@{
    receita          = @{ Unidade = 'USD'; Tags = @('RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'SalesRevenueNet', 'RevenueFromContractWithCustomerIncludingAssessedTax') }
    lucroBruto       = @{ Unidade = 'USD'; Tags = @('GrossProfit') }
    lucroOperacional = @{ Unidade = 'USD'; Tags = @('OperatingIncomeLoss') }
    cfo              = @{ Unidade = 'USD'; Tags = @('NetCashProvidedByUsedInOperatingActivities', 'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations') }
    capex            = @{ Unidade = 'USD'; Tags = @('PaymentsToAcquirePropertyPlantAndEquipment', 'PaymentsToAcquireProductiveAssets') }
    eps              = @{ Unidade = 'USD/shares'; Tags = @('EarningsPerShareDiluted'); PorAcao = $true }
    acoes            = @{ Unidade = 'shares'; Tags = @('WeightedAverageNumberOfDilutedSharesOutstanding'); Acoes = $true }
}

# Regras de classificação (heurísticas; ajusta à vontade)
$EmpresasRe = [ordered]@{
    # "Big Apple" (New York) and other companies called Apple (Apple Hospitality, Apple Bank) are not the company
    AAPL  = '(?<!\bbig )\b(apple|aapl|iphones?|ipad|app store|siri|apple intelligence|vision pro|tim cook|john ternus|airpods|macbook)\b(?! (hospitality|bank|pie|cider|juice|orchards?))'
    NVDA  = '\b(nvidia|nvda|jensen huang|blackwell|vera rubin|rubin ultra|cuda|geforce|h20|h200|gb300|dgx)\b'
    GOOGL = '\b(alphabet|google|googl|goog|gemini|youtube|waymo|sundar pichai|deepmind|android|chrome|tpus?)\b'
    # "iShares" alone is not SXR8: iShares Bitcoin Trust (IBIT), iShares Gold, the iShares MSCI funds, etc. are other funds
    SXR8  = '\b(ishares(?!\s+(bitcoin|ethereum|ether|gold|silver|crypto|core msci|msci))|sxr8|cspx|ie00b5bmr087|s&p 500 etf|index funds?)\b'
    # The other ETFs: the fund itself (ISIN, tickers, index) and the broad stock market its index covers. A headline
    # about that market as a whole ("European shares", "Nikkei 225", "emerging-market stocks") counts; a single
    # company or a country's economy (GDP, rates) does not, so the match stays specific.
    EUNK  = '\b(ie00b4k48x80|eunk|imae|smea|msci europe|stoxx (europe )?600|european (stocks|shares|equities|bourses)|europe''s (stocks|shares|bourses)|ações europeias|bolsas europeias)\b'
    IS3N  = '\b(ie00bkm4gz66|is3n|eimi|emim|msci (emerging markets|em)( imi)?|emerging[- ]markets? (stocks|shares|equities|etfs?|funds?)|em (stocks|equities)|mercados emergentes)\b'
    EUNN  = '\b(ie00b4l5yx21|eunn|ijpa|sjpa|msci japan|nikkei (225|average|stock average|share average|index)|topix|japanese (stocks|shares|equities)|tokyo (stocks|shares|stock exchange)|japan''s (stocks|shares|stock market)|bolsa de tóquio)\b'
    # MiCA (the EU crypto regulation) only in that exact spelling, so the mineral "mica" does not count
    BTC   = '\b(bitcoins?|btc|satoshi nakamoto|halving|spot bitcoin|bitcoin etfs?|ibit|microstrategy|mstr|michael saylor|strategic bitcoin reserve|crypto (market|markets|bill|law|rules|regulation|etfs?|selloff|sell-off|rally|crash)|cryptocurrenc(y|ies)|criptomoedas?|criptoativos?|mercado cripto|(?-i:MiCA))\b'
    MKT   = '(s&p 500|s&amp;p 500|\bsp500\b|\bspx\b|wall street|\bnasdaq\b|dow jones|stock market|\bstocks\b|\bfederal reserve\b|\bfed\b|\bfomc\b|inflation|\bcpi\b|(?<!bitcoin )\btreasur(y|ies)\b(?! (firms?|compan))|\byields?\b|recession|\becb\b|\bbce\b|\beuro\b|\bdollar\b|\bdólar\b|\btariffs?\b|\bbolsas?\b|\bjuros\b|inflação)'
}
# Maiores posições de cada ETF (top 10 dos ficheiros da iShares de 5 out 2026), para ligar ao fundo as notícias que só falam
# delas. Só contam quando o título não tem nenhuma palavra de $EmpresasRe, e sempre como correspondência fraca: no máximo
# "moderada", sem bónus, com a etiqueta "Via top holding". Apple, NVIDIA e Alphabet já estão em $EmpresasRe.
# Tickers = os do ficheiro da iShares (para o lembrete de manutenção); Re = limites de palavra e exclusões para os falsos
# positivos conhecidos (outras empresas com nome parecido, sentidos comuns da palavra), cada um com um teste no Test-Engine.
# Quando o top 10 de um fundo mudar, o lembrete de manutenção nomeia a posição nova sem alias.
$AliasesPosicoes = [ordered]@{
    SXR8 = @(
        # dicas de produtos (Word, Excel…) não são notícias da empresa
        @{ Nome = 'Microsoft'; Tickers = @('MSFT'); Re = '\b(microsoft|msft|satya nadella)\b(?! (word|excel|powerpoint|outlook|paint|flight simulator)\b)' }
        # a floresta e o rio Amazonas
        @{ Nome = 'Amazon'; Tickers = @('AMZN'); Re = '\b(amazon(?! (rainforest|river|basin|jungle|forest|region|tribes?|fires?|deforestation)\b)|amzn|andy jassy|amazon web services|aws)\b' }
        # o concurso escolar "Broadcom MASTERS"
        @{ Nome = 'Broadcom'; Tickers = @('AVGO'); Re = '\b(broadcom(?! masters\b)|avgo|hock tan)\b' }
        # "meta" em português (objetivo: "meta de inflação") e "meta-analysis": só "Meta" com maiúscula
        @{ Nome = 'Meta Platforms'; Tickers = @('META'); Re = '\b(meta platforms|(?-i:Meta)(?! de\b|-analys| analys)|facebook|instagram|whatsapp|zuckerberg)\b' }
        # o mícron como unidade ("5 micron filter")
        @{ Nome = 'Micron Technology'; Tickers = @('MU'); Re = '\b(micron technology|micron(?! (thick|wide|scale|sized?|level|particles?|filters?|resolution)\b)|sanjay mehrotra)\b' }
        # Nikola Tesla e as bobinas de Tesla
        @{ Nome = 'Tesla'; Tickers = @('TSLA'); Re = '(?<!nikola )\b(tesla(?! (coils?|units?|towers?)\b)|tsla|cybertruck)\b' }
        # AMD também é a degenerescência macular (doença dos olhos)
        @{ Nome = 'AMD'; Tickers = @('AMD'); Re = '(?<!(wet|dry|neovascular) )\b(amd|advanced micro devices|lisa su)\b(?! (patients?|eye|treatment|drugs?|therapy)\b)' }
    )
    EUNK = @(
        # a ASM International é outra empresa
        @{ Nome = 'ASML'; Tickers = @('ASML'); Re = '\b(asml|christophe fouquet)\b' }
        # torneios patrocinados (golfe, râguebi)
        @{ Nome = 'HSBC'; Tickers = @('HSBA'); Re = '\bhsbc\b(?! (women|champions|sevens|svns|golf|rugby)\b)' }
        # La Roche-Posay (cosmética), Roche Bobois (mobiliário)
        @{ Nome = 'Roche'; Tickers = @('ROP', 'ROG'); Re = '(?<!la )\b(roche(?!-posay| bobois| harbou?r)|genentech)\b' }
        # o clube Racing Santander e a cidade/região
        @{ Nome = 'Banco Santander'; Tickers = @('SAN'); Re = '(?<!racing )\b(banco santander|santander(?! (province|department|airport|city)\b)|ana bot[ií]n)\b' }
        # "shell company" e outros sentidos de "shell": só "Shell plc" ou Shell seguida de termos da empresa
        @{ Nome = 'Shell'; Tickers = @('SHEL'); Re = '\b(shell plc|royal dutch shell|shell(?=(''s)? (plc|shares|profits?|earnings|results|ceo|buybacks?|lng|oil|gas)\b)|wael sawan)\b' }
        # a Sandoz separou-se da Novartis em 2023
        @{ Nome = 'Novartis'; Tickers = @('NOVN'); Re = '\b(novartis|vas narasimhan)\b' }
        # o Opel Astra
        @{ Nome = 'AstraZeneca'; Tickers = @('AZN'); Re = '\b(astrazeneca|astra zeneca|pascal soriot)\b' }
        # a marca Toll House nos EUA é da Ferrero desde 2018
        @{ Nome = 'Nestlé'; Tickers = @('NESN'); Re = '\bnestl[eé]\b(?! toll house)' }
        # Siemens Energy, Siemens Healthineers e Siemens Gamesa são outras cotadas
        @{ Nome = 'Siemens'; Tickers = @('SIE'); Re = '\bsiemens\b(?! (energy|healthineers|gamesa)\b)' }
        # a seiva ("maple sap"): só "SAP" em maiúsculas
        @{ Nome = 'SAP'; Tickers = @('SAP'); Re = '\b((?-i:SAP)|christian klein|s/4hana)\b' }
    )
    IS3N = @(
        # "Taiwan semiconductor exports" (o setor, não a empresa)
        @{ Nome = 'TSMC'; Tickers = @('2330'); Re = '\b(tsmc|taiwan semiconductor manufacturing|c\. ?c\. wei)\b' }
        # Samsung Heavy, C&T, Biologics, Life, SDI… são outras cotadas
        @{ Nome = 'Samsung Electronics'; Tickers = @('005930', '005935'); Re = '\bsamsung\b(?! (heavy|c&t|biologics|life|engineering|sdi|securities|fire|e&a)\b)' }
        # outras empresas do grupo SK
        @{ Nome = 'SK Hynix'; Tickers = @('000660'); Re = '\b(sk hynix|hynix)\b' }
        # a Tencent Music é cotada à parte
        @{ Nome = 'Tencent'; Tickers = @('700'); Re = '\b(tencent(?! music)|wechat|weixin)\b' }
        # "media tech" (o setor)
        @{ Nome = 'MediaTek'; Tickers = @('2454'); Re = '\bmedia ?tek\b' }
        # Ali Baba (o conto)
        @{ Nome = 'Alibaba'; Tickers = @('9988'); Re = '\b(alibaba|aliexpress|taobao|tmall|eddie wu)\b' }
        # a Delta Air Lines
        @{ Nome = 'Delta Electronics'; Tickers = @('2308'); Re = '\bdelta electronics\b' }
        # a China State Construction é outra empresa
        @{ Nome = 'China Construction Bank'; Tickers = @('939'); Re = '\bchina construction bank\b' }
        # a Fox Corporation
        @{ Nome = 'Hon Hai (Foxconn)'; Tickers = @('2317'); Re = '\b(hon hai|foxconn)\b' }
    )
    EUNN = @(
        # Mitsubishi Motors, Mitsubishi Corp… são outras empresas
        @{ Nome = 'Mitsubishi UFJ'; Tickers = @('8306'); Re = '\b(mitsubishi ufj|mufg)\b' }
        # a Advantech (Taiwan)
        @{ Nome = 'Advantest'; Tickers = @('6857'); Re = '\badvantest\b' }
        # a Tokyo Electron Device é cotada à parte
        @{ Nome = 'Tokyo Electron'; Tickers = @('8035'); Re = '\btokyo electron\b(?! device)' }
        # Toyota Tsusho, Toyota Boshoku e Toyota Industries são outras cotadas
        @{ Nome = 'Toyota'; Tickers = @('7203'); Re = '\b(toyota(?! (tsusho|boshoku|industries)\b)|lexus|akio toyoda)\b' }
        # a Hitachi Zosen (hoje Kanadevia) não é do grupo
        @{ Nome = 'Hitachi'; Tickers = @('6501'); Re = '\bhitachi\b(?! (zosen|construction machinery|metals|high-tech)\b)' }
        # a Sumitomo Mitsui Trust é outro grupo
        @{ Nome = 'Sumitomo Mitsui Financial'; Tickers = @('8316'); Re = '\b(sumitomo mitsui financial|smfg|smbc)\b' }
        # a SoftBank Corp (telecomunicações) é cotada à parte
        @{ Nome = 'SoftBank Group'; Tickers = @('9984'); Re = '\b(softbank(?! corp\b)|masayoshi son)\b' }
        # a Toshiba (antiga dona)
        @{ Nome = 'Kioxia'; Tickers = @('285A'); Re = '\bkioxia\b' }
        # o verbo "recruit"
        @{ Nome = 'Recruit Holdings'; Tickers = @('6098'); Re = '\b(recruit holdings|indeed\.com|glassdoor)\b' }
        # a Sony Financial Group é cotada à parte desde 2025
        @{ Nome = 'Sony'; Tickers = @('6758'); Re = '\b(sony(?! (financial|bank)\b)|playstation)\b' }
    )
}
$Temas = @(
    # "forecast" and "outlook" only count as earnings next to earnings words ("rate forecast" or "price forecast" are not earnings)
    @{ Nome = 'Earnings';               Peso = 3;   Re = '\b(earnings|quarterly results|revenues?|guidance|eps|profits?|q[1-4]|fiscal (first|second|third|fourth)|(earnings|profit|revenue|sales|annual|quarterly) (forecasts?|outlook)|forecasts? (for )?(earnings|profits?|revenues?|sales)|resultados|receitas?|lucros?)\b' },
    @{ Nome = 'Regulation and courts';  Peso = 3;   Re = '\b(antitrust|doj|justice department|ftc|european commission|european union|eu regulators?|brussels|digital markets act|dma|fined?|fines|lawsuits?|sued|court|judge|ruling|appeals?|probe|investigation|regulators?|monopoly|multas?|tribunal|concorrência|comissão europeia|cftc|mica|stablecoins?|market structure|clarity act|genius act|regulação)\b' },
    # "curb" counts only as trade curbs ("to curb inflation" is a verb about prices)
    @{ Nome = 'China and exports';      Peso = 3;   Re = '\b(china|chinese|export controls?|exports?|(export|chip|trade|tech|tighter|new|us|u\.s\.|american|chinese|beijing) curbs?|curbs? on|restrictions?|commerce department|licen[cs]es?|huawei|taiwan|tsmc|tariffs?|tarifas?|exportações)\b' },
    @{ Nome = 'AI and CapEx';           Peso = 2;   Re = '\b(ai|artificial intelligence|capex|capital expenditures?|data cent(er|re)s?|hyperscalers?|openai|anthropic|gpus?|asics?|trainium|amd|broadcom|inteligência artificial|centros de dados)\b' },
    @{ Nome = 'Capital and shareholders'; Peso = 2;   Re = '\b(buybacks?|repurchases?|dividends?|share sale|stock offering|offering|at-the-market|dilution|convertible|stock split|recompras?|dividendos?|aumento de capital)\b' },
    @{ Nome = 'Management';             Peso = 2;   Re = '\b(ceo|cfo|chief executive|steps down|resigns?|appoint(s|ed)?|succession|demite|nomeia)\b' },
    @{ Nome = 'Products and technology'; Peso = 1;   Re = '\b(launch(es|ed)?|unveil(s|ed)?|announce[sd]?|releases?|foldable|blackwell|rubin|gemini|waymo|siri|chips?|lança|apresenta)\b' },
    @{ Nome = 'Macro and rates';        Peso = 2;   Re = '\b(fed|federal reserve|fomc|rate (hike|cut)s?|interest rates?|inflation|cpi|pce|jobs report|payrolls|unemployment|(?<!bitcoin )treasur(y|ies)(?! (firms?|compan))|yields?|recession|gdp|oil|warsh|inflação|taxas? de juro|desemprego|bce|ecb)\b' },
    @{ Nome = 'Currency';               Peso = 1;   Re = '\b(dollar|euro|eur/?usd|currenc(y|ies)|forex|dólar|câmbio)\b' },
    @{ Nome = 'ETF flows';              Peso = 2;   Re = '\b(etf (inflows?|outflows?|flows?)|inflows?|outflows?|ibit|fbtc|spot (bitcoin )?etfs?|bitcoin etfs?|fluxos)\b' },
    @{ Nome = 'Institutional adoption'; Peso = 2;   Re = '\b(microstrategy|mstr|saylor|strategic (bitcoin )?reserve|reserva estratégica|bitcoin treasur(y|ies)|treasury compan(y|ies)|institutional (investors?|adoption|demand)|adoção institucional)\b' },
    @{ Nome = 'Security and fraud';     Peso = 3;   Re = '\b(hacks?|hacked|hackers?|exploits?|exploited|stolen|theft|breach|scams?|fraud|rug pull|ataque informático|roubo|burla|fraude)\b' },
    @{ Nome = 'Network and mining';     Peso = 1;   Re = '\b(miners?|mining|hash ?rate|halving|mining difficulty|lightning network|taproot|mineração|mineradores)\b' },
    @{ Nome = 'Analysts';               Peso = 0.5; Re = '\b(price targets?|upgrade[sd]?|downgrade[sd]?|ratings?|analysts?|preço-alvo)\b' }
)
# Severity needs context for "curb", "warn" and "delay": "to curb inflation", "Bill Miller warns MSCI…" or
# "Gemini 4 after months of delays" are not adverse events for the asset
$Severo  = '\b((export|chip|trade|tech|tighter|new|us|u\.s\.|american|chinese|beijing) curbs?|curbs? on|arrested|smuggl(e|es|ed|ing)|export bans?|bans?|banned|breakup|break up|divest(iture|s)?|plunge[sd]?|tumble[sd]?|crash(es|ed)?|misses|missed|cuts? (its )?(guidance|forecast|outlook)|(profit|revenue|sales|earnings) warnings?|warns? (of|on) |halt(s|ed)?|recalls?|(delays?|delayed|postpones?|postponed) ([\w-]+ ){0,2}(launch|release|rollout|shipments?|production|deliveries|ipo)|(launch|release|rollout|shipments?|production|deliveries) (is |are |was |were )?(delayed|postponed)|subpoena|indict(ed|ment)?|sell-?off|slump(s|ed)?|hacked|exploited|bankrupt(cy)?|insolven(t|cy)|collapse[sd]?|proibição|afunda|derrapa)\b'
$DecisaoJuros = '\b((raises|raised|cuts|cut|hikes|hiked|holds|held) (interest )?rates|rate (hike|cut|decision)|sobe (as )?taxas|desce (as )?taxas)\b'
$Ruido   = '\b(should you buy|stocks? to buy|buy now|millionaire|could soar|prediction|best stocks?|no-brainer|forever|top \d+ stocks?|my favorite|vale a pena comprar|price predictions?|price forecasts?|price analysis|could (hit|reach)|memecoins?|meme coins?|altcoins? to buy|next 100x|whales? (buy|sell|move)|previsão de preço)\b'
$Positivo = '\b(beats?|surge[sd]?|soar(s|ed)?|record|raises?|raised|upgrade[sd]?|strong|jumps?|rall(y|ies|ied)|gains?|approv(al|ed)|wins?|tops)\b'
$Negativo = '\b(miss(es|ed)?|plunge[sd]?|falls?|fell|drops?|slump(s|ed)?|cuts?|downgrade[sd]?|weak|lawsuits?|probe|fined?|bans?|delay(s|ed)?|warns?|loss(es)?|decline[sd]?|curbs?)\b'
$TierPrimaria  = 'sec\.gov|apple\.com|apple newsroom|nvidia\.com|nvidianews|nvidia newsroom|blog\.google|google blog|abc\.xyz|federalreserve\.gov|federal reserve|ecb\.europa\.eu|\bbce\b|bls\.gov|ishares\.com|blackrock'
$TierReferencia = 'reuters|bloomberg|financial times|\bft\.com|wall street journal|\bwsj\b|cnbc|associated press|apnews|ap news|barron|economist|marketwatch|axios|the information|nikkei|jornal de negócios|jornaldenegocios|eco\.sapo|\beco\b|expresso|público|publico\.pt|observador|dinheiro vivo|coindesk|the block|theblock'
$TierCuidado   = 'motley fool|fool\.com|24/7 wall|247wallst|seeking alpha|seekingalpha|benzinga|investorplace|zacks|tipranks|simply wall|marketbeat|insider monkey|gurufocus|finbold|coincodex|ad-hoc-news|invezz|stocktitan|coingape|u\.today|newsbtc|bitcoinist|cryptopolitan|coinpedia|crypto\.news|zycrypto|watcher\.guru|coincentral|ambcrypto|thecryptobasic|cryptonews|blockonomi|coinspeaker|crypto economy|bitcoinsistemi|coinfomania'

# ============================================================================
# 2. FUNÇÕES DE APOIO
# ============================================================================

function Write-Passo([string]$Texto) { Write-Host "  • $Texto" -ForegroundColor Cyan }

# E-mail de contacto da SEC quando não vem em -EmailSEC: a variável de ambiente BLUECHIP_SEC_EMAIL ou, se estiver vazia,
# bluechip-board.config.json na pasta do script ({ "secEmail": "…" }, fora do git). Assim a tarefa agendada e o lançador
# não o precisam de ter nos argumentos (que ficam visíveis no Agendador de Tarefas e na lista de processos).
# Devolve o e-mail (ou '') e um aviso quando o ficheiro existe mas não tem um e-mail válido.
function Get-EmailSecLocal([string]$PastaScript) {
    $origem = 'BLUECHIP_SEC_EMAIL'; $e = "$env:BLUECHIP_SEC_EMAIL".Trim(); $aviso = ''
    $cfg = Join-Path $PastaScript 'bluechip-board.config.json'
    if (-not $e -and (Test-Path -LiteralPath $cfg)) {
        $origem = 'bluechip-board.config.json'
        try { $e = "$((Get-Content -LiteralPath $cfg -Raw -Encoding UTF8 | ConvertFrom-Json).secEmail)".Trim() }
        catch { $aviso = "bluechip-board.config.json could not be read ($($_.Exception.Message)): running without the SEC sources." }
    }
    if ($e -and ($e -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or $e -eq 'your@email.com')) { $aviso = "The SEC e-mail in $origem is not valid: running without the SEC sources."; $e = '' }
    [pscustomobject]@{ email = $e; aviso = $aviso }
}

function Get-Url {
    param([string]$Url, [string]$UserAgent = $Script:UA, [int]$Timeout = 25)
    $ultimoErro = $null
    # PowerShell 7 validates the User-Agent format and rejects the "@" of the e-mail the SEC requires in it; the header
    # sent is the same, only that check is skipped (Windows PowerShell 5.1 has no such check, nor this parameter)
    $extra = @{}; if ($PSVersionTable.PSVersion.Major -ge 6) { $extra.SkipHeaderValidation = $true }
    for ($tentativa = 1; $tentativa -le 2; $tentativa++) {
        try {
            $r = Invoke-WebRequest -Uri $Url -UserAgent $UserAgent -UseBasicParsing -TimeoutSec $Timeout -MaximumRedirection 5 -Headers @{ 'Accept-Language' = 'en-US,en;q=0.8,pt-PT;q=0.6' } @extra
            return (ConvertFrom-Bytes $r.RawContentStream.ToArray() "$($r.Headers['Content-Type'])")
        } catch {
            $ultimoErro = $_
            $codigo = 0; try { $codigo = [int]$_.Exception.Response.StatusCode } catch { }
            # a permanent answer (bad request, unauthorised, forbidden, not found, gone) does not change a few seconds later:
            # no second attempt, so a dead source does not slow the run down
            if ($codigo -in 400, 401, 403, 404, 410) { break }
            if ($tentativa -lt 2) {
                # rate limit or temporary overload (429/503): wait longer before the second attempt
                Start-Sleep -Seconds $(if ($codigo -eq 429 -or $codigo -eq 503) { 8 } else { 2 })
            }
        }
    }
    throw $ultimoErro
}

# Decodes a response with the charset it declares (HTTP header or XML prolog); UTF-8 when nothing is declared
function ConvertFrom-Bytes([byte[]]$Bytes, [string]$ContentType) {
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return '' }
    $enc = $null
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $enc = [Text.Encoding]::UTF8 }
    if (-not $enc -and $ContentType -match 'charset=["'']?([\w\-]+)') { try { $enc = [Text.Encoding]::GetEncoding($Matches[1]) } catch { } }
    if (-not $enc) {
        $inicio = [Text.Encoding]::ASCII.GetString($Bytes, 0, [math]::Min(200, $Bytes.Length))
        if ($inicio -match '^\s*<\?xml[^>]*encoding=["'']([\w\-]+)') { try { $enc = [Text.Encoding]::GetEncoding($Matches[1]) } catch { } }
    }
    if (-not $enc) { $enc = [Text.Encoding]::UTF8 }
    return $enc.GetString($Bytes).TrimStart([char]0xFEFF)
}

# Parses XML from the internet without processing DTDs or external entities (protects against entity-expansion attacks)
function ConvertTo-XmlSeguro([string]$Texto) {
    $def = New-Object System.Xml.XmlReaderSettings
    $def.DtdProcessing = [System.Xml.DtdProcessing]::Ignore
    $def.XmlResolver = $null
    $def.MaxCharactersFromEntities = 1000000
    $leitor = [System.Xml.XmlReader]::Create((New-Object IO.StringReader $Texto), $def)
    try { $x = New-Object System.Xml.XmlDocument; $x.XmlResolver = $null; $x.Load($leitor) } finally { $leitor.Close() }
    return , $x   # the comma stops PowerShell from enumerating the document's child nodes
}

function Get-Texto($No) {
    if ($null -eq $No) { return '' }
    if ($No -is [string]) { return $No }
    if ($No -is [System.Xml.XmlElement]) { return $No.InnerText }
    if ($No -is [array]) { return (Get-Texto $No[0]) }
    return "$No"
}

function ConvertTo-Data([string]$Texto) {
    if ([string]::IsNullOrWhiteSpace($Texto)) { return $null }
    $t = $Texto.Trim() -replace '\sEDT$', ' -0400' -replace '\sEST$', ' -0500' -replace '\sCDT$', ' -0500' -replace '\sCST$', ' -0600' -replace '\sPDT$', ' -0700' -replace '\sPST$', ' -0800' -replace '\s(UT|UTC|GMT)$', ' +0000' -replace '\sBST$', ' +0100' -replace '\sCEST$', ' +0200' -replace '\sCET$', ' +0100' -replace '\sEEST$', ' +0300' -replace '\sEET$', ' +0200' -replace '\sJST$', ' +0900'
    $d = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($t, $Script:Inv, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$d)) { return $d.ToUniversalTime() }
    return $null
}

function Add-Fonte($Nome, $Tipo, $Url, $Estado, $Itens, $Ms, $Erro) {
    $Script:Fontes.Add([pscustomobject]@{ nome = $Nome; tipo = $Tipo; url = $Url; estado = $Estado; itens = $Itens; ms = $Ms; erro = $Erro })
}

function Get-Dominio([string]$Url) {
    try { return ([uri]$Url).Host -replace '^www\.', '' } catch { return '' }
}

# Lê um feed RSS ou Atom e devolve itens normalizados
function Read-Feed {
    param([hashtable]$Feed, [string]$UserAgent = $Script:UA, [string]$Tipo = 'rss')
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lista = New-Object System.Collections.Generic.List[object]
    try {
        $texto = Get-Url -Url $Feed.Url -UserAgent $UserAgent
        $x = ConvertTo-XmlSeguro $texto
        $limite = $Script:Agora.AddDays(-$Dias)
        $futuro = $Script:Agora.AddHours(12)   # a date this far ahead is malformed (wrong time zone or typo)
        $semData = 0                            # items without a valid date cannot be placed in the news window: skipped and counted
        if ($x.rss) {
            foreach ($i in @($x.rss.channel.item)) {
                if ($null -eq $i) { continue }
                $titulo = [Net.WebUtility]::HtmlDecode((Get-Texto $i.title)).Trim()
                $link = (Get-Texto $i.link).Trim()
                $fonte = Get-Texto $i.source
                if (-not $fonte -and $titulo -match '\s[-–]\s([^-–]{2,60})$') { $fonte = $Matches[1] }
                if ($fonte -and $titulo.EndsWith(" - $fonte")) { $titulo = $titulo.Substring(0, $titulo.Length - $fonte.Length - 3) }
                $srcUrl = ''; if ($i.source -is [System.Xml.XmlElement]) { $srcUrl = $i.source.GetAttribute('url') }
                $dom = if ($srcUrl) { Get-Dominio $srcUrl } else { Get-Dominio $link }
                if (-not $fonte) { $fonte = if ($Feed.Nome -match '^Yahoo') { 'Yahoo Finance' } elseif ($Feed.Nome -notmatch '^Google News') { $Feed.Nome } else { $dom } }
                $data = ConvertTo-Data (Get-Texto $i.pubDate)
                if (-not $data -or $data -gt $futuro) { $semData++; continue }
                if ($data -lt $limite) { continue }
                $lista.Add([pscustomobject]@{ titulo = $titulo; link = $link; fonte = $fonte.Trim(); dominio = $dom; data = $data; feed = $Feed.Nome; dica = $Feed.Dica; exigir = "$($Feed.Exigir)"; extraTema = ''; extraPeso = 0 })
            }
        } elseif ($x.feed) {
            foreach ($e in @($x.feed.entry)) {
                if ($null -eq $e) { continue }
                $titulo = [Net.WebUtility]::HtmlDecode((Get-Texto $e.title)).Trim()
                $link = ''
                foreach ($l in @($e.link)) { if ($l -is [System.Xml.XmlElement]) { $h = $l.GetAttribute('href'); if ($h -and (-not $link -or $l.GetAttribute('rel') -eq 'alternate')) { $link = $h } } }
                $data = ConvertTo-Data (Get-Texto $e.updated); if (-not $data) { $data = ConvertTo-Data (Get-Texto $e.published) }
                if (-not $data -or $data -gt $futuro) { $semData++; continue }
                if ($data -lt $limite) { continue }
                $extraTema = ''; $extraPeso = 0
                if ($Tipo -eq 'sec') {
                    $form = ''; if ($e.category -is [System.Xml.XmlElement]) { $form = $e.category.GetAttribute('term') }
                    if ($form -match '^(3|4|5|144)(/A)?$') { continue }   # transações de insiders: demasiado ruído
                    switch -Regex ($form) {
                        '^8-K'                { $extraTema = 'Official filing (8-K)'; $extraPeso = 3 }
                        '^(10-Q|10-K)'        { $extraTema = 'Earnings'; $extraPeso = 3 }
                        '^(424B|S-3|FWP|S-8)' { $extraTema = 'Capital and shareholders'; $extraPeso = 3 }
                        default               { $extraTema = 'SEC filing'; $extraPeso = 1 }
                    }
                    $titulo = "SEC $form – $titulo"
                }
                $lista.Add([pscustomobject]@{ titulo = $titulo; link = $link; fonte = $Feed.Nome; dominio = (Get-Dominio $link); data = $data; feed = $Feed.Nome; dica = $Feed.Dica; exigir = "$($Feed.Exigir)"; extraTema = $extraTema; extraPeso = $extraPeso })
            }
        } else {
            throw 'Unrecognised feed format'
        }
        Add-Fonte $Feed.Nome $Tipo $Feed.Url 'ok' $lista.Count $sw.ElapsedMilliseconds $(if ($semData) { "$semData item(s) skipped: missing or invalid date" } else { '' })
    } catch {
        Add-Fonte $Feed.Nome $Tipo $Feed.Url 'error' 0 $sw.ElapsedMilliseconds ($_.Exception.Message -replace '\s+', ' ')
    }
    return , $lista
}

# Classifica uma notícia: empresas, temas, nível de impacto, sentimento e tipo de fonte
function Measure-Noticia($N) {
    $txt = "$($N.titulo)"
    # (listas do PowerShell e operadores -ccontains, sensíveis a maiúsculas como List.Contains: no PowerShell 7 cada
    # chamada a um método .NET custa ~10 µs, ver $DupVazias)
    $emp = @(foreach ($k in $EmpresasRe.Keys) { if ($txt -match $EmpresasRe[$k]) { $k } })
    # Without a keyword, the story goes to the feed's asset. Aggregated feeds (Google News, Yahoo Finance) return many
    # general stories ("China's AI agents…" in an Apple search), so there this is only a weak match: no company bonus
    # and at most "moderate". Primary feeds (newsrooms, SEC, Fed, ECB) keep the full assignment.
    $soPeloFeed = $false; $viaPosicao = @()
    if ($emp.Count -eq 0 -and $N.dica) { $emp = @("$($N.dica)"); $soPeloFeed = ("$($N.feed)" -match '^(Google News|Yahoo Finance)') }
    # Sem palavra de nenhum ativo e num feed sem ativo próprio (feeds gerais, que antes deixavam cair a notícia): uma das
    # maiores posições de um fundo ($AliasesPosicoes) liga a notícia a esse fundo, como correspondência fraca (no máximo
    # "moderada", sem bónus). Os feeds com ativo próprio mantêm a atribuição de sempre.
    if ($emp.Count -eq 0) {
        foreach ($f in $AliasesPosicoes.Keys) {
            foreach ($a in @($AliasesPosicoes[$f])) {
                if ($a -and $txt -match $a.Re) { if ($emp -cnotcontains $f) { $emp += "$f" }; if ($viaPosicao -cnotcontains $a.Nome) { $viaPosicao += "$($a.Nome)" } }
            }
        }
    }
    if ($emp.Count -eq 0) { return $null }
    if ($N.exigir -and $emp -cnotcontains $N.exigir) { return $null }   # feed temático: o título tem de ser sobre o ativo
    if ($emp -ccontains 'SXR8' -and $emp -cnotcontains 'MKT') { $emp += 'MKT' }

    # Themes: the strongest counts in full and each extra one at half, up to 5 points. (Adding them all, up to 7,
    # let a headline that merely mentions AI, chips and China reach "material" with no adverse event.)
    $achados = @(); $pesos = @()
    foreach ($t in $Temas) { if ($txt -match $t.Re) { $achados += "$($t.Nome)"; $pesos += [double]$t.Peso } }
    if ($N.extraTema -and $achados -cnotcontains $N.extraTema) { $achados = @("$($N.extraTema)") + $achados; $pesos += [double]$N.extraPeso }
    # o mais forte conta inteiro e cada um dos outros a metade (= máximo + (soma − máximo) ÷ 2; os pesos são múltiplos de 0,5)
    $score = 0.0
    if ($pesos.Count) { $mx = $pesos[0]; $soma = 0.0; foreach ($w in $pesos) { $soma += $w; if ($w -gt $mx) { $mx = $w } }; $score = $mx + ($soma - $mx) / 2 }
    if ($score -gt 5) { $score = 5 }
    $eSevero = $txt -match $Severo
    if ($eSevero) { $score += 3 }
    if ($txt -match $DecisaoJuros) { $score += 2 }
    $eRuido = $txt -match $Ruido

    $src = ('{0} {1}' -f $N.fonte, $N.dominio).ToLowerInvariant()
    $tier = if ($src -match $TierPrimaria) { 'primaria' } elseif ($src -match $TierCuidado) { 'cuidado' } elseif ($src -match $TierReferencia) { 'referencia' } else { 'outra' }
    switch ($tier) { 'primaria' { $score += 1.5 } 'referencia' { $score += 1 } 'cuidado' { $score -= 1.5 } }
    $bonus = $false; foreach ($e in $emp) { if ($e -in 'AAPL', 'NVDA', 'GOOGL', 'BTC') { $bonus = $true } }
    if (-not $soPeloFeed -and -not $viaPosicao.Count -and $bonus) { $score += 1 }
    if ($eRuido) { $score -= 3 }

    $nivel = if ($eRuido -and $tier -ne 'primaria') { 'white' } elseif ($score -ge 7) { 'red' } elseif ($score -ge 4) { 'orange' } elseif ($score -ge 2) { 'yellow' } else { 'white' }
    if (($soPeloFeed -or $viaPosicao.Count) -and $nivel -in 'red', 'orange') { $nivel = 'yellow' }
    $p = [regex]::Matches($txt, $Positivo, 'IgnoreCase').Count
    $q = [regex]::Matches($txt, $Negativo, 'IgnoreCase').Count
    $sent = if ($p -gt $q) { 'positivo' } elseif ($q -gt $p) { 'negativo' } elseif ($p -gt 0) { 'misto' } else { 'neutro' }

    $chave = (($txt.ToLowerInvariant() -replace '[^\p{L}\p{N}]+', ' ').Trim())
    if ($chave.Length -gt 90) { $chave = $chave.Substring(0, 90) }
    return [pscustomobject]@{
        chave = $chave; titulo = $txt; link = $N.link; fonte = $N.fonte; dominio = $N.dominio
        data = if ($N.data) { $N.data.ToString('o') } else { $null }
        feed = $N.feed; empresas = @($emp); temas = @($achados); nivel = $nivel; score = [math]::Round($score, 1)
        severo = [bool]$eSevero; sentimento = $sent; tier = $tier; novo = $false; outras = @(); soFeed = $soPeloFeed; viaPosicao = @($viaPosicao)
    }
}

# Lembrete de manutenção: posições do top 10 atual de cada fundo (ficheiro da iShares desta execução) sem alias em
# $AliasesPosicoes. Apple, NVIDIA e Alphabet não precisam (estão em $EmpresasRe). Com os pesos de referência (sem ficheiro
# atual) não há top 10, e o lembrete não dispara.
function Get-PosicoesSemAlias($PesosEtfs) {
    $r = New-Object System.Collections.Generic.List[string]
    foreach ($id in @($PesosEtfs.Keys)) {
        $p = $PesosEtfs[$id]
        if (-not $p -or -not $p.aoVivo) { continue }
        $al = @($AliasesPosicoes[$id])
        foreach ($h in @($p.top10)) {
            if (-not $h) { continue }
            $t = "$($h.t)"; $n = "$($h.n)"
            if ($t -in 'AAPL', 'NVDA', 'GOOGL') { continue }
            if (@($al | Where-Object { $_ -and ($t -in @($_.Tickers) -or $n -match $_.Re) }).Count) { continue }
            $r.Add("$id $n ($t)")
        }
    }
    return @($r)
}

# Instante em UTC como texto ISO (aaaa-mm-ddThh:mm:ssZ), ou $null. Aceita texto e as datas em que o PowerShell 7
# transforma os textos ISO ao ler JSON (ConvertFrom-Json), para os ficheiros guardados manterem sempre o mesmo formato.
function ConvertTo-IsoUtc($Valor) {
    $d = if ($Valor -is [DateTimeOffset]) { $Valor } elseif ($Valor -is [datetime]) { [DateTimeOffset]$Valor } elseif ($Valor) { ConvertTo-Data "$Valor" } else { $null }
    if (-not $d) { return $null }
    return $d.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $Script:Inv)
}

# ----------------------------------------------------------------------------
# Histórico de notícias (noticias-historico.json): as notícias "material" e "important" de cada execução, guardadas
# $HistoricoDias dias, para o site as pôr no gráfico e explicar os grandes movimentos. Só títulos públicos, sem dados
# pessoais. Escrito no fim da execução, depois do vistos.json (uma execução que falha a meio não o altera).
# ----------------------------------------------------------------------------
$HistoricoDias = 400
$HistoricoMax = 15000   # notícias no máximo (hoje entram ~20 por dia: ~8 000 em 400 dias)
function ConvertTo-NoticiaHistorico($N) {
    if (-not $N) { return $null }
    $d = ConvertTo-IsoUtc $N.data; $t = "$($N.titulo)"
    if (-not $d -or -not $t -or "$($N.nivel)" -notin 'red', 'orange') { return $null }
    $s = 0.0; [void][double]::TryParse("$($N.score)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$s)
    $chave = if ($N.chave) { "$($N.chave)" } else { (($t.ToLowerInvariant() -replace '[^\p{L}\p{N}]+', ' ').Trim()) }
    [pscustomobject]@{
        chave = $chave; titulo = $t; link = "$($N.link)"; fonte = "$($N.fonte)"; data = $d
        empresas = @(@($N.empresas) | Where-Object { $_ } | ForEach-Object { "$_" }); nivel = "$($N.nivel)"
        score = [math]::Round($s, 1); temas = @(@($N.temas) | Where-Object { $_ } | ForEach-Object { "$_" })
    }
}
# Ficheiro em falta: começa vazio, com um aviso. Corrompido: começa vazio, com um aviso, e o ficheiro antigo é guardado
# como .bad quando o novo for escrito (no fim da execução).
function Read-HistoricoNoticias([string]$Caminho) {
    $nome = Split-Path $Caminho -Leaf
    if (-not (Test-Path -LiteralPath $Caminho)) {
        return [pscustomobject]@{ inicio = $null; itens = @(); estado = 'new'; aviso = "News history file not found ($nome): starting a new one." }
    }
    try {
        $o = Get-Content -LiteralPath $Caminho -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $o -or -not $o.PSObject.Properties['noticias']) { throw 'it is not a news history file' }
        $itens = @(foreach ($n in @($o.noticias)) { $x = ConvertTo-NoticiaHistorico $n; if ($x) { $x } })
        return [pscustomobject]@{ inicio = (ConvertTo-IsoUtc $o.inicio); itens = $itens; estado = 'ok'; aviso = '' }
    } catch {
        return [pscustomobject]@{ inicio = $null; itens = @(); estado = 'corrupted'; aviso = "The news history file ($nome) could not be read ($($_.Exception.Message)): starting a new one. The old file will be kept as $nome.bad." }
    }
}
# Junta as notícias desta execução ao histórico: uma por título (fica a de maior pontuação), só as dos últimos $Dias dias.
# O início do histórico nunca é anterior a esse corte (o que estava antes foi apagado).
function Merge-HistoricoNoticias($Historico, $Noticias, [DateTimeOffset]$Agora, [int]$Dias = $HistoricoDias) {
    $corte = $Agora.AddDays(-$Dias)
    $porChave = [ordered]@{}
    $todas = @(@($Historico.itens) + @(foreach ($n in @($Noticias)) { ConvertTo-NoticiaHistorico $n }))
    foreach ($n in $todas) {
        if (-not $n) { continue }
        $t = ConvertTo-Data $n.data
        if (-not $t -or $t -lt $corte) { continue }
        if (-not $porChave.Contains($n.chave) -or $n.score -gt $porChave[$n.chave].score) { $porChave[$n.chave] = $n }
    }
    $inicio = if ($Historico.inicio) { $Historico.inicio } else { ConvertTo-IsoUtc $Agora }
    if ((ConvertTo-Data $inicio) -lt $corte) { $inicio = ConvertTo-IsoUtc $corte }
    $itens = @($porChave.Values | Sort-Object data, chave)
    # Limite de segurança ($HistoricoMax notícias, cerca de 2 anos ao ritmo de hoje): o ficheiro vai para o site e para cada
    # cópia do Archive, e o Windows PowerShell 5.1 lê-o inteiro. Acima dele ficam as mais recentes, e o início passa a ser a
    # data da mais antiga que ficou (o site nunca diz que o histórico começa antes do que tem).
    if ($itens.Count -gt $HistoricoMax) {
        $itens = @($itens | Select-Object -Last $HistoricoMax)
        $inicio = $itens[0].data
    }
    [pscustomobject]@{ inicio = $inicio; itens = $itens; estado = $Historico.estado; aviso = $Historico.aviso }
}

# ----------------------------------------------------------------------------
# Datas passadas de resultados (SEC): os 8-K com o item 2.02 ("Results of Operations and Financial Condition") e o
# último 10-Q/10-K. Lista e datas de entrega: data.sec.gov/submissions (JSON). Hora de aceitação: feed Atom de 8-K do
# EDGAR, que a dá com o fuso (o acceptanceDateTime do JSON não é fiável: na Apple vem várias horas ao lado da hora de
# aceitação que o EDGAR mostra). Sem hora verificada, a hora e a sessão ficam vazias: nunca se inventam.
# ----------------------------------------------------------------------------
$Script:FusoNY = $null
$MaxCabecalhos = 12   # pedidos de cabeçalhos de entregas do EDGAR por empresa e execução (só para 8-K ainda sem hora)
function Get-FusoNY { if (-not $Script:FusoNY) { $Script:FusoNY = [TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time') }; return $Script:FusoNY }
# Primeira sessão em $Data ou depois (ou só depois, com -Depois). $Sessoes: datas com fecho no histórico da empresa,
# por ordem; para além do fim do histórico, os dias úteis sem os feriados indicados.
function Get-ProximaSessao([string]$Data, [bool]$Depois, [string[]]$Sessoes, [string[]]$Feriados) {
    # dentro do período do histórico: as datas do histórico (antes do seu início, ou depois do fim, os dias úteis)
    if ($Sessoes -and $Sessoes.Count -and [string]::CompareOrdinal($Data, $Sessoes[-1]) -le 0 -and [string]::CompareOrdinal($Data, $Sessoes[0]) -ge 0) {
        $i = [Array]::BinarySearch($Sessoes, $Data, [StringComparer]::Ordinal)
        $k = if ($i -ge 0) { $(if ($Depois) { $i + 1 } else { $i }) } else { -bnot $i }
        if ($k -lt $Sessoes.Count) { return $Sessoes[$k] }
    }
    $dt = [datetime]::ParseExact($Data, 'yyyy-MM-dd', $Script:Inv)
    if ($Depois) { $dt = $dt.AddDays(1) }
    for ($j = 0; $j -lt 10; $j++) {
        $iso = $dt.ToString('yyyy-MM-dd', $Script:Inv)
        if ($dt.DayOfWeek -ne [DayOfWeek]::Saturday -and $dt.DayOfWeek -ne [DayOfWeek]::Sunday -and $iso -notin @($Feriados)) { return $iso }
        $dt = $dt.AddDays(1)
    }
    return $null
}
# Sessão em que o mercado reage a um anúncio, em hora de Nova Iorque (com as mudanças de hora dos EUA): aceite depois
# das 16:00, a sessão seguinte; antes das 09:30, a própria sessão; durante a sessão, a própria sessão, com nota.
function Get-SessaoReacao([DateTimeOffset]$Aceite, [string[]]$Sessoes = @(), [string[]]$Feriados = @()) {
    $ny = [TimeZoneInfo]::ConvertTime($Aceite, (Get-FusoNY))
    $min = $ny.Hour * 60 + $ny.Minute; $dia = $ny.ToString('yyyy-MM-dd', $Script:Inv)
    $quando = if ($min -ge 960) { 'after' } elseif ($min -lt 570) { 'before' } else { 'during' }
    # num dia sem sessão (fim de semana ou feriado, por exemplo a Sexta-feira Santa, em que a SEC recebe entregas) não
    # há "durante" nem "depois do fecho": a reação é a sessão seguinte
    if ((Get-ProximaSessao $dia $false $Sessoes $Feriados) -ne $dia) { $quando = 'closed' }
    $sessao = Get-ProximaSessao $dia ($quando -eq 'after') $Sessoes $Feriados
    [pscustomobject]@{
        horaNY = $ny.ToString('yyyy-MM-dd HH:mm', $Script:Inv); quando = $quando; sessao = $sessao
        nota = $(if ($quando -eq 'during') { 'Accepted during the session (09:30-16:00 New York): the reaction is counted from that same session.' } elseif ($quando -eq 'closed') { 'Accepted on a day the exchange was closed: the reaction is counted from the next session.' } else { '' })
    }
}
# Hora de aceitação de uma entrega pelo cabeçalho do próprio arquivo do EDGAR (…-index-headers.html, "ACCEPTANCE-DATETIME",
# em hora de Nova Iorque): a fonte oficial para as entregas que já não estão no feed Atom (só as 40 mais recentes) ou se o
# feed deixar de existir. Devolve a hora em UTC, ou $null (sem a confirmar, nada é inventado).
function Get-AceiteCabecalho([string]$Cik, [string]$Acc, [string]$UserAgent) {
    if ($Acc -notmatch '^\d{10}-\d{2}-\d{6}$') { return $null }
    $u = "https://www.sec.gov/Archives/edgar/data/$([int64]$Cik)/$($Acc -replace '-', '')/$Acc-index-headers.html"
    $t = Get-Url $u -UserAgent $UserAgent
    if ($t -notmatch "<ACCESSION-NUMBER>\s*$([regex]::Escape($Acc))\b") { throw "the EDGAR header is not for $Acc" }
    $m = [regex]::Match($t, '<ACCEPTANCE-DATETIME>\s*(\d{14})\b')
    if (-not $m.Success) { return $null }
    $local = [datetime]::ParseExact($m.Groups[1].Value, 'yyyyMMddHHmmss', $Script:Inv)
    $fuso = Get-FusoNY
    if ($fuso.IsInvalidTime($local)) { return $null }
    return [DateTimeOffset]::new($local, $fuso.GetUtcOffset($local)).ToUniversalTime()
}
# $Conhecidas: número de acesso → hora de aceitação já verificada numa execução anterior (as entregas não mudam)
function Get-ResultadosSEC([hashtable]$Empresa, [string]$UserAgent, [string[]]$Sessoes = @(), [string[]]$Feriados = @(), [hashtable]$Conhecidas = @{}) {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $cik = "$($Empresa.Cik)"
    $u = "https://data.sec.gov/submissions/CIK$cik.json"; $nome = "SEC: past earnings dates ($($Empresa.Id))"
    try {
        $j = (Get-Url $u -UserAgent $UserAgent) | ConvertFrom-Json
        $cikR = [int64]0
        if (-not $j -or -not [int64]::TryParse("$($j.cik)", [ref]$cikR) -or $cikR -ne [int64]$cik) { throw "SEC returned CIK '$($j.cik)' instead of $cik" }
        $rc = $j.filings.recent
        $fm = @($rc.form); $ac = @($rc.accessionNumber); $it = @($rc.items); $fd = @($rc.filingDate); $rd = @($rc.reportDate)
        if (-not $fm.Count -or $ac.Count -ne $fm.Count -or $it.Count -ne $fm.Count -or $fd.Count -ne $fm.Count -or $rd.Count -ne $fm.Count) { throw 'Unexpected SEC response (filing lists missing or of different lengths)' }
        $dia = { param($v) if ($v -is [datetime]) { $v.ToString('yyyy-MM-dd', $Script:Inv) } else { "$v" } }
        $limite = $Script:Agora.UtcDateTime.AddYears(-5).ToString('yyyy-MM-dd', $Script:Inv)
        $oitos = New-Object System.Collections.Generic.List[object]; $ultimo = $null
        for ($i = 0; $i -lt $fm.Count; $i++) {
            $f = "$($fm[$i])"; $dEnt = & $dia $fd[$i]
            if ($dEnt -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
            if ($f -in '10-Q', '10-K' -and (-not $ultimo -or $dEnt -gt $ultimo.data)) { $ultimo = [pscustomobject]@{ form = $f; data = $dEnt; periodo = (& $dia $rd[$i]); acc = "$($ac[$i])" } }
            if ($f -eq '8-K' -and ",$($it[$i])," -match ',2\.02,' -and $dEnt -ge $limite) { $oitos.Add([pscustomobject]@{ acc = "$($ac[$i])"; entrega = $dEnt }) }
        }
        # horas de aceitação (com fuso) do feed Atom de 8-K; se falhar, ficam as já conhecidas e as outras vazias
        $horas = @{}; $erroHoras = ''
        Start-Sleep -Milliseconds 400
        $ua = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=$cik&type=8-K&dateb=&owner=include&count=40&output=atom"
        try {
            $x = ConvertTo-XmlSeguro (Get-Url $ua -UserAgent $UserAgent)
            foreach ($e in @($x.feed.entry)) {
                $a = (Get-Texto $e.content.'accession-number').Trim(); $t = ConvertTo-Data (Get-Texto $e.updated)
                if ($a -and $t) { $horas[$a] = $t }
            }
            if (-not $horas.Count) { $erroHoras = 'no entries in the 8-K feed' }
        } catch { $erroHoras = $_.Exception.Message }
        # sem hora no feed nem numa execução anterior: o cabeçalho da entrega no arquivo do EDGAR (no máximo $MaxCabecalhos
        # pedidos por empresa e execução, 400 ms entre eles; as horas encontradas ficam guardadas e não voltam a ser pedidas)
        $doCab = 0; $erroCab = ''; $pedidos = 0
        foreach ($o in $oitos) {
            if ($horas.ContainsKey($o.acc) -or $Conhecidas.ContainsKey($o.acc)) { continue }
            if ($pedidos -ge $MaxCabecalhos) { break }
            $pedidos++; Start-Sleep -Milliseconds 400
            try { $tc = Get-AceiteCabecalho $cik $o.acc $UserAgent; if ($tc) { $horas[$o.acc] = $tc; $doCab++ } } catch { $erroCab = $_.Exception.Message }
        }
        $semHora = 0
        $lista = @(foreach ($o in $oitos) {
            $t = if ($horas.ContainsKey($o.acc)) { $horas[$o.acc] } elseif ($Conhecidas.ContainsKey($o.acc)) { ConvertTo-Data "$($Conhecidas[$o.acc])" } else { $null }
            if ($t) {
                $s = Get-SessaoReacao $t $Sessoes $Feriados
                [pscustomobject]@{ acc = $o.acc; entrega = $o.entrega; aceite = (ConvertTo-IsoUtc $t); horaNY = $s.horaNY; quando = $s.quando; sessao = $s.sessao; nota = $s.nota }
            } else {
                $semHora++
                [pscustomobject]@{ acc = $o.acc; entrega = $o.entrega; aceite = $null; horaNY = $null; quando = $null; sessao = $null; nota = 'Acceptance time unavailable (not in the EDGAR 8-K feed or the filing header): no reaction session.' }
            }
        })
        $notas = @()
        if ($erroHoras) { $notas += "8-K feed: $erroHoras" }
        if ($doCab) { $notas += "$doCab acceptance time(s) from the EDGAR filing headers" }
        if ($erroCab) { $notas += "filing headers: $erroCab" }
        if ($semHora) { $notas += "$semHora 8-K without an acceptance time" }
        $nota = $notas -join '; '
        Add-Fonte $nome 'sec' $u 'ok' $lista.Count $sw.ElapsedMilliseconds $nota
        return [pscustomobject]@{ id = $Empresa.Id; estado = 'ok'; fonte = 'SEC EDGAR'; obtidoEm = $Script:Agora.ToString('o'); resultados = $lista; ultimoRelatorio = $ultimo; erro = ''; nota = $nota }
    } catch {
        Add-Fonte $nome 'sec' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message
        return [pscustomobject]@{ id = $Empresa.Id; estado = 'error'; fonte = ''; obtidoEm = $Script:Agora.ToString('o'); resultados = @(); ultimoRelatorio = $null; erro = $_.Exception.Message; nota = '' }
    }
}

# ----------------------------------------------------------------------------
# Fundamentais (SEC XBRL companyfacts). Para cada período (início|fim) guarda-se o valor da entrega mais recente (v, f:
# data "filed"; as reexpressões ganham) e o primeiro valor publicado (v0, só se for diferente; p: a primeira data "filed").
# Os comparativos voltam a ser entregues um ano depois: sem p, o separador da tarefa 14 teria um ano de atraso.
# Trimestres: o facto de ~90 dias; senão a diferença de dois acumulados com o mesmo início (o fluxo de caixa só vem
# acumulado no ano; o 4.º trimestre = ano − 9 meses), marcada d. Nada é interpolado: um trimestre sem dados fica vazio.
# ----------------------------------------------------------------------------
$FundamentaisAnos = 11
# fator dos splits posteriores a uma data (as entregas posteriores a um split já vêm ajustadas por ele)
function Get-FatorSplit($Splits, [string]$Depois) {
    $k = 1.0
    foreach ($s in @($Splits)) { $q = @($s); $r = 0.0; if ($q.Count -ge 2 -and "$($q[0])" -gt $Depois -and [double]::TryParse("$($q[1])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$r) -and $r -gt 0) { $k *= $r } }
    return $k
}
function Get-FatosMetrica($Json, [hashtable]$Def, $Splits) {
    $por = @{}; $usadas = New-Object System.Collections.Generic.List[string]; $outrasUnid = New-Object System.Collections.Generic.List[string]
    $iso = '^\d{4}-\d{2}-\d{2}$'
    foreach ($tag in $Def.Tags) {
        $no = $Json.facts.'us-gaap'.$tag
        if (-not $no -or -not $no.units) { continue }
        $lista = $no.units.($Def.Unidade)
        if (-not $lista) { foreach ($u in $no.units.PSObject.Properties) { $outrasUnid.Add("$tag in $($u.Name)") }; continue }
        $usou = $false
        foreach ($x in @($lista)) {
            if (-not $x) { continue }
            $ini = "$($x.start)"; $fim = "$($x.end)"; $fil = "$($x.filed)"
            if ($ini -notmatch $iso -or $fim -notmatch $iso -or $fil -notmatch $iso -or "$($x.form)" -notmatch '^10-[QK]') { continue }
            $v = 0.0; try { $v = [double]$x.val } catch { continue }
            if ([double]::IsNaN($v) -or [double]::IsInfinity($v)) { continue }
            $dias = ([datetime]::ParseExact($fim, 'yyyy-MM-dd', $Script:Inv) - [datetime]::ParseExact($ini, 'yyyy-MM-dd', $Script:Inv)).Days
            if (-not (($dias -ge 80 -and $dias -le 100) -or ($dias -ge 170 -and $dias -le 190) -or ($dias -ge 260 -and $dias -le 285) -or ($dias -ge 350 -and $dias -le 380))) { continue }
            if ($Def.PorAcao) { $v = $v / (Get-FatorSplit $Splits $fil) } elseif ($Def.Acoes) { $v = $v * (Get-FatorSplit $Splits $fil) }
            $chave = "$ini|$fim"
            if ($por.ContainsKey($chave)) {
                $e = $por[$chave]
                if ($e.tag -ne $tag) { continue }   # o período já vem de uma tag preferida
                if ($fil -gt $e.f) { $e.v = $v; $e.f = $fil }
                if ($fil -lt $e.p) { $e.v0 = $v; $e.p = $fil }
            } else { $por[$chave] = [pscustomobject]@{ tag = $tag; inicio = $ini; fim = $fim; dias = $dias; v = $v; f = $fil; v0 = $v; p = $fil } }
            $usou = $true
        }
        if ($usou) { $usadas.Add($tag) }
    }
    return [pscustomobject]@{ por = $por; usadas = @($usadas); outrasUnid = @($outrasUnid) }
}
function Get-TrimestresMetrica($Por, [bool]$Derivar = $true) {
    $q = @{}; $L = @($Por.Values)
    foreach ($x in $L) { if ($x.dias -le 100) { $q[$x.fim] = [pscustomobject]@{ inicio = $x.inicio; fim = $x.fim; v = $x.v; f = $x.f; v0 = $x.v0; p = $x.p; d = $false } } }
    if ($Derivar) {
        foreach ($g in @($L | Group-Object inicio)) {
            $acum = @($g.Group | Sort-Object dias)
            for ($i = 1; $i -lt $acum.Count; $i++) {
                $dep = $acum[$i]; $ant = $acum[$i - 1]; $dd = $dep.dias - $ant.dias
                if ($dd -lt 80 -or $dd -gt 100 -or $q.ContainsKey($dep.fim)) { continue }
                $ini = [datetime]::ParseExact($ant.fim, 'yyyy-MM-dd', $Script:Inv).AddDays(1).ToString('yyyy-MM-dd', $Script:Inv)
                $q[$dep.fim] = [pscustomobject]@{ inicio = $ini; fim = $dep.fim; v = $dep.v - $ant.v; f = $(if ($dep.f -gt $ant.f) { $dep.f } else { $ant.f }); v0 = $dep.v0 - $ant.v0; p = $(if ($dep.p -gt $ant.p) { $dep.p } else { $ant.p }); d = $true }
            }
        }
    }
    return $q
}
# um valor guardado: v e f (entrega mais recente), p (primeira publicação), v0 só se diferente, d só se derivado
function ConvertTo-ValorFund($x, [int]$Casas) {
    if (-not $x) { return $null }
    $o = [ordered]@{ v = [math]::Round([double]$x.v, $Casas); f = $x.f; p = $x.p }
    if ([math]::Abs([double]$x.v0 - [double]$x.v) -gt 1e-9 * [math]::Max(1.0, [math]::Abs([double]$x.v))) { $o.v0 = [math]::Round([double]$x.v0, $Casas) }
    if ($x.d) { $o.d = $true }
    return [pscustomobject]$o
}
# os fundamentais de uma resposta companyfacts já lida (sem rede): trimestres dos últimos $FundamentaisAnos anos e TTM
function ConvertTo-Fundamentais($Json, $Splits) {
    $metr = [ordered]@{}; $tags = [ordered]@{}; $faltam = [ordered]@{}; $inicios = @{}
    foreach ($k in $MetricasSEC.Keys) {
        $def = $MetricasSEC[$k]; $r = Get-FatosMetrica $Json $def $Splits
        $tags[$k] = @($r.usadas)
        if (-not $r.por.Count) { $faltam[$k] = $(if (@($r.outrasUnid).Count) { "no fact in $($def.Unidade) ($(@($r.outrasUnid) -join ', '))" } else { "no fact for $($def.Tags -join ', ')" }) }
        foreach ($x in $r.por.Values) { if ($x.dias -ge 170) { $inicios[$x.inicio] = $true } }
        $metr[$k] = Get-TrimestresMetrica $r.por (-not $def.Acoes)
    }
    $limite = $Script:Agora.UtcDateTime.AddYears(-$FundamentaisAnos).ToString('yyyy-MM-dd', $Script:Inv)
    $fins = @(@($metr.Values | ForEach-Object { @($_.Keys) }) | Where-Object { $_ -ge $limite } | Sort-Object -Unique)
    $comecos = @($inicios.Keys | Sort-Object)
    $casas = @{ receita = 0; lucroBruto = 0; lucroOperacional = 0; cfo = 0; capex = 0; eps = 6; acoes = 0 }
    $tri = @(foreach ($e in $fins) {
        $m = [ordered]@{}; $ini = $null
        foreach ($k in $metr.Keys) { $x = $metr[$k][$e]; $m[$k] = ConvertTo-ValorFund $x $casas[$k]; if ($x -and -not $ini) { $ini = $x.inicio } }
        $c = $metr.cfo[$e]; $x = $metr.capex[$e]
        $m.fcf = if ($c -and $x) { ConvertTo-ValorFund ([pscustomobject]@{ v = $c.v - $x.v; f = $(if ($c.f -gt $x.f) { $c.f } else { $x.f }); v0 = $c.v0 - $x.v0; p = $(if ($c.p -gt $x.p) { $c.p } else { $x.p }); d = ($c.d -or $x.d) }) 0 } else { $null }
        # ano fiscal: o início acumulado mais recente até 380 dias antes do fim do trimestre
        $s = @($comecos | Where-Object { $_ -le $e -and ([datetime]::ParseExact($e, 'yyyy-MM-dd', $Script:Inv) - [datetime]::ParseExact($_, 'yyyy-MM-dd', $Script:Inv)).Days -le 380 }) | Select-Object -Last 1
        $ano = $null; $nq = $null
        if ($s) { $dS = [datetime]::ParseExact($s, 'yyyy-MM-dd', $Script:Inv); $nq = [int][math]::Round(($([datetime]::ParseExact($e, 'yyyy-MM-dd', $Script:Inv)) - $dS).Days / 91.3); $ano = 'FY' + $dS.AddDays(364).Year; if ($nq -lt 1 -or $nq -gt 4) { $nq = $null } }
        [pscustomobject]@{ fim = $e; inicio = $ini; ano = $ano; q = $nq; m = [pscustomobject]$m }
    })
    # TTM do trimestre mais recente com 4 trimestres consecutivos (fins com 80–100 dias de distância), todos com valor
    $ttm = $null
    if ($tri.Count) {
        $ult = $tri[-1]; $ttm = [ordered]@{ fim = $ult.fim }
        foreach ($k in @('receita', 'lucroBruto', 'lucroOperacional', 'cfo', 'capex', 'fcf', 'eps')) {
            $ok = $tri.Count -ge 4; $soma = 0.0; $f = ''
            if ($ok) { for ($i = $tri.Count - 4; $i -lt $tri.Count; $i++) {
                $val = $tri[$i].m.$k
                if (-not $val) { $ok = $false; break }
                if ($i -gt $tri.Count - 4 -and ([datetime]::ParseExact($tri[$i].fim, 'yyyy-MM-dd', $Script:Inv) - [datetime]::ParseExact($tri[$i - 1].fim, 'yyyy-MM-dd', $Script:Inv)).Days -notin 80..100) { $ok = $false; break }
                $soma += $val.v; if ($val.f -gt $f) { $f = $val.f } } }
            $ttm[$k] = if ($ok) { [pscustomobject]@{ v = [math]::Round($soma, $casas[$(if ($k -eq 'fcf') { 'cfo' } else { $k })]); f = $f } } else { $null }
        }
        $ac = @($tri | Where-Object { $_.m.acoes }) | Select-Object -Last 1
        $ttm.acoes = if ($ac) { [pscustomobject]@{ v = $ac.m.acoes.v; f = $ac.m.acoes.f; fim = $ac.fim } } else { $null }
        $ttm = [pscustomobject]$ttm
    }
    return [pscustomobject]@{ tags = [pscustomobject]$tags; faltam = [pscustomobject]$faltam; trimestres = $tri; ttm = $ttm }
}
# ----------------------------------------------------------------------------
# Indicadores da área do euro (ECB Data Portal, CSV "dataonly": KEY, dimensões, TIME_PERIOD, OBS_VALUE). Uma página HTML
# (erro do servidor) nunca é lida como dados; ".", "NaN" e vazio contam como em falta (nunca 0).
# ----------------------------------------------------------------------------
function ConvertFrom-CsvSerie([string]$Texto, [string]$Chave, [double]$Min, [double]$Max) {
    $t = "$Texto".TrimStart([char]0xFEFF).Trim()
    if ($t -match '^<' -or $t -match '(?i)<html') { throw 'HTML page instead of CSV (the source returned an error page)' }
    $linhas = @($t -split "`r?`n" | Where-Object { $_ })
    if (-not $linhas.Count) { throw 'Empty response' }
    $h = @($linhas[0] -split ','); $iK = [array]::IndexOf($h, 'KEY'); $iP = [array]::IndexOf($h, 'TIME_PERIOD'); $iV = [array]::IndexOf($h, 'OBS_VALUE')
    if ($iK -lt 0 -or $iP -lt 0 -or $iV -lt 0) { throw "Unexpected CSV header: $($linhas[0].Substring(0, [math]::Min(80, $linhas[0].Length)))" }
    $lidos = New-Object System.Collections.Generic.List[object]; $faltas = 0; $invalidos = 0
    foreach ($l in @($linhas | Select-Object -Skip 1)) {
        $c = @($l -split ',')
        if ($c.Count -le [math]::Max($iP, $iV)) { $invalidos++; continue }
        if ($c[$iK] -ne $Chave) { throw "The response has another series (key '$($c[$iK])' instead of $Chave)" }
        $p = $c[$iP].Trim()
        $d = if ($p -match '^\d{4}-\d{2}-\d{2}$') { $p } elseif ($p -match '^\d{4}-\d{2}$') { "$p-01" } else { $null }
        if (-not $d) { $invalidos++; continue }
        $s = $c[$iV].Trim(); $v = 0.0
        if ($s -in '', '.', 'NaN') { $faltas++; continue }
        if (-not [double]::TryParse($s, [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$v) -or $v -lt $Min -or $v -gt $Max) { $invalidos++; continue }
        $lidos.Add([pscustomobject]@{ d = $d; v = $v })
    }
    if (-not $lidos.Count) { throw 'The series has no value' }
    $pts = @(foreach ($x in @($lidos | Sort-Object d)) { , @("$($x.d)", [double]$x.v) })
    return [pscustomobject]@{ pontos = $pts; faltas = $faltas; invalidos = $invalidos }
}
function Get-SerieMacro([hashtable]$S) {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $nome = "ECB: $($S.Nome)"
    $u = "https://data-api.ecb.europa.eu/service/data/$($S.Fluxo)/$($S.Chave)?format=csvdata&detail=dataonly&startPeriod=$($S.Desde)"
    try {
        $r = ConvertFrom-CsvSerie (Get-Url $u) "$($S.Fluxo).$($S.Chave)" $S.Min $S.Max
        $nota = @(); if ($r.faltas) { $nota += "$($r.faltas) missing value(s)" }; if ($r.invalidos) { $nota += "$($r.invalidos) invalid value(s) ignored" }
        Add-Fonte $nome 'macro' $u 'ok' @($r.pontos).Count $sw.ElapsedMilliseconds ($nota -join '; ')
        return [pscustomobject]@{ id = $S.Id; nome = $S.Nome; estado = 'ok'; fonte = 'ECB Data Portal'; obtidoEm = $Script:Agora.ToString('o'); freq = $S.Chave.Substring(0, 1); pontos = $r.pontos; erro = '' }
    } catch {
        Add-Fonte $nome 'macro' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message
        return [pscustomobject]@{ id = $S.Id; nome = $S.Nome; estado = 'error'; fonte = ''; obtidoEm = $Script:Agora.ToString('o'); freq = $S.Chave.Substring(0, 1); pontos = @(); erro = $_.Exception.Message }
    }
}
# Lembrete de manutenção: a versão do Chrome no User-Agent ($Script:UAChrome, atual em $Script:UAData) com cerca de 8 versões
# de atraso (uma a cada 4 semanas). Um navegador muito antigo é mais facilmente recusado pelos sites.
function Get-LembreteUA([int]$Versao, [string]$Desde, [DateTimeOffset]$Agora) {
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($Desde, 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return "The date of the browser identity (`$Script:UAData = '$Desde') is not yyyy-MM-dd." }
    $atraso = [int][math]::Floor(($Agora.UtcDateTime - $d).TotalDays / 28)
    if ($atraso -lt 8) { return $null }
    return "The browser identity sent to the sites says Chrome $Versao, about $atraso versions behind the current Chrome (a new one every 4 weeks): set `$Script:UAChrome to the current version and `$Script:UAData to today."
}
# Lembrete de manutenção: sem uma decisão da Fed, ou do BCE, a mais de 60 dias no calendário manual
function Get-LembreteReunioes($Cal, [DateTimeOffset]$Agora) {
    $lim = $Agora.UtcDateTime.Date.AddDays(60); $falta = @()
    foreach ($p in @(@('Fed', '^Fed decision'), @('ECB', '^ECB monetary policy decision'))) {
        $tem = @($Cal | Where-Object { $x = [datetime]::MinValue; "$($_.ev)" -match $p[1] -and [datetime]::TryParseExact("$($_.d)", 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$x) -and $x -gt $lim }).Count
        if (-not $tem) { $falta += $p[0] }
    }
    if (-not $falta.Count) { return $null }
    return "The manual calendar (`$Calendario) has no $($falta -join ' or ') rate decision more than 60 days ahead: add the next meetings from the official $($falta -join ' and ') calendar."
}
# Uma empresa: só com -EmailSEC (mesmo User-Agent); só descarrega quando a SEC mostra um 10-Q/10-K mais recente do que o
# da recolha anterior ($Relatorio = resultadosSec[id].ultimoRelatorio desta execução). Se falhar: a execução anterior até
# 120 dias, marcada como tal; senão, "error" (ou "skipped" sem e-mail), sem números.
function Get-FundamentaisEmpresa([hashtable]$Empresa, [string]$UserAgent, $Splits, $Anterior, $Relatorio) {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $cik = "$($Empresa.Cik)"
    $u = "https://data.sec.gov/api/xbrl/companyfacts/CIK$cik.json"; $nome = "SEC: fundamentals ($($Empresa.Id))"
    $relData = if ($Relatorio -and "$($Relatorio.data)" -match '^\d{4}-\d{2}-\d{2}$') { "$($Relatorio.data)" } else { '' }
    $copia = { param($A, $Estado, $Fonte, $Erro, $Nota) [pscustomobject]@{ id = $Empresa.Id; estado = $Estado; fonte = $Fonte; obtidoEm = (ConvertTo-IsoUtc $A.obtidoEm); relatorio = $A.relatorio; tags = $A.tags; faltam = $A.faltam; trimestres = @($A.trimestres); ttm = $A.ttm; erro = $Erro; nota = $Nota } }
    if ($UserAgent -and $relData -and $Anterior -and "$($Anterior.estado)" -eq 'ok' -and $Anterior.relatorio -and "$($Anterior.relatorio.data)" -ge $relData) {
        Add-Fonte $nome 'sec' $u 'ok (cached)' @($Anterior.trimestres).Count 0 "Not downloaded: no 10-Q/10-K newer than $($Anterior.relatorio.data)."
        return (& $copia $Anterior 'ok' "$($Anterior.fonte)" '' "Facts collected $((ConvertTo-IsoUtc $Anterior.obtidoEm).Substring(0, 10)); no newer 10-Q/10-K since.")
    }
    $erro = ''
    if ($UserAgent) {
        try {
            $j = (Get-Url $u -UserAgent $UserAgent -Timeout 60) | ConvertFrom-Json
            $cikR = [int64]0
            if (-not $j -or -not [int64]::TryParse("$($j.cik)", [ref]$cikR) -or $cikR -ne [int64]$cik) { throw "SEC returned CIK '$($j.cik)' instead of $cik" }
            if (-not $j.facts -or -not $j.facts.'us-gaap') { throw 'Unexpected SEC response (no us-gaap facts)' }
            $r = ConvertTo-Fundamentais $j $Splits
            if (-not @($r.trimestres).Count) { throw 'No quarterly facts in the SEC response' }
            Add-Fonte $nome 'sec' $u 'ok' @($r.trimestres).Count $sw.ElapsedMilliseconds $(if (@($r.faltam.PSObject.Properties).Count) { "Unavailable: $(@($r.faltam.PSObject.Properties | ForEach-Object { $_.Name }) -join ', ')" } else { '' })
            return [pscustomobject]@{ id = $Empresa.Id; estado = 'ok'; fonte = 'SEC XBRL (companyfacts)'; obtidoEm = $Script:Agora.ToString('o'); relatorio = $(if ($relData) { $Relatorio } else { $null }); tags = $r.tags; faltam = $r.faltam; trimestres = $r.trimestres; ttm = $r.ttm; erro = ''; nota = '' }
        } catch { $erro = $_.Exception.Message; Add-Fonte $nome 'sec' $u 'error' 0 $sw.ElapsedMilliseconds $erro }
    } else { $erro = 'Run with -EmailSEC "your@email.com" to enable it: the SEC requires a contact in each request.' }
    $quando = if ($Anterior) { ConvertTo-Data (ConvertTo-IsoUtc $Anterior.obtidoEm) } else { $null }
    if ($Anterior -and "$($Anterior.estado)" -in 'ok', 'previous run' -and $quando -and ($Script:Agora - $quando).TotalDays -le 120) {
        return (& $copia $Anterior 'previous run' "previous run (retrieved $($quando.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv)))" $erro '')
    }
    return [pscustomobject]@{ id = $Empresa.Id; estado = $(if ($UserAgent) { 'error' } else { 'skipped' }); fonte = ''; obtidoEm = $Script:Agora.ToString('o'); relatorio = $null; tags = $null; faltam = $null; trimestres = @(); ttm = $null; erro = $erro; nota = '' }
}

# ----------------------------------------------------------------------------
# Notícias duplicadas: o mesmo acontecimento contado por várias fontes com títulos diferentes
# (ex.: "California tech executive arrested for smuggling $300M in Nvidia AI chips" e
#  "California Man Charged With Smuggling $300 Million Nvidia AI Chips to China").
# Compara as palavras dos títulos, dando mais peso às palavras raras na recolha (nomes, valores),
# e agrupa as notícias do mesmo acontecimento; fica a de maior pontuação e as restantes passam a
# "outras fontes" dessa notícia.
# ----------------------------------------------------------------------------
$DupVazias = @('a','an','the','and','or','but','of','to','in','on','at','for','with','by','from','as','is','are','was','were','be','been','being','it','its','this','that','these','those',
    'into','over','under','after','before','about','amid','against','between','than','then','so','not','no','new','says','say','said','report','reports','reportedly','could','would',
    'will','can','may','might','should','has','have','had','do','does','did','up','out','more','most','less','very','just','also','how','why','what','when','who','which','while',
    'de','da','do','das','dos','e','o','os','as','um','uma','uns','umas','em','no','na','nos','nas','por','para','com','sem','que','se','ao','aos','sao','foi','como','mais','menos','sobre',
    'vs','via','here','there','their','his','her','he','she','they','we','you','your','our','us','watch','update','exclusive','breaking','video','live','analysis','opinion',
    'stock','stocks','share','shares','acoes','acao','today','hoje','best','top','guide','ranking','things','ways','reasons')
# Tabela do PowerShell consultada pelo índice ($DupVazias[$w]), não um HashSet: no PowerShell 7 cada chamada a um método
# .NET (HashSet.Contains, List.Add, String.EndsWith…) passa pelo registo AMSI do Windows e custa cerca de 10 µs, o que nos
# ciclos sobre milhares de palavras somava segundos. Os índices e os operadores (-match, -replace) não têm esse custo.
$vaziasH = @{}; foreach ($w in $DupVazias) { $vaziasH[$w] = $true }; $DupVazias = $vaziasH
$DupSinonimos = @{}
foreach ($grupo in @(
        @('launch','unveil','announce','introduce','debut','release','roll','lanca','apresenta','anuncia'),
        @('charge','arrest','indict','accuse','acusa','detido','preso'),
        @('rise','climb','jump','gain','surge','rally','soar','advance','sobe','subindo','avanca','dispara'),
        @('fall','drop','slide','slump','sink','tumble','decline','plunge','dip','cai','caindo','recua','afunda'),
        @('buyback','repurchase','recompra'),
        @('smuggle','smuggl','contrabando'),
        @('chip','semiconductor','chipset','processor'),
        @('ceo','executive','chief','boss','executivo'),
        @('allow','let','permit','approve','permite'),
        @('weigh','consider','mull','pondera'))) {
    foreach ($w in $grupo) { $DupSinonimos[$w] = $grupo[0] }
}

function Get-PalavrasTitulo([string]$Titulo) {
    $s = $Titulo.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
    $s = $s -replace "[’‘``]", "'" -replace "'s\b", ''
    # valores: "$300 million", "$300M" e "300 milhões" ficam todos "300m" (o mesmo para mil milhões / "b")
    $s = $s -replace '\$\s?(\d+(?:[.,]\d+)?)\s?(billion|bn|b|bilioes|mil milhoes)\b', '${1}b ' -replace '\$\s?(\d+(?:[.,]\d+)?)\s?(million|mln|m|milhoes)\b', '${1}m '
    $s = $s -replace '\b(\d+(?:[.,]\d+)?)\s?(billion|bn|bilioes|mil milhoes)\b', '${1}b ' -replace '\b(\d+(?:[.,]\d+)?)\s?(million|mln|milhoes)\b', '${1}m '
    $s = $s -replace '(\d),(\d)', '$1.$2'
    $s = $s -replace '\s[-–—|]\s[^-–—|]{2,40}$', '' -replace '\s(por|by)\s(reuters|investing\.com)$', ''
    # Devolve as palavras distintas pela ordem em que aparecem (a ordem decide os empates do agrupamento). Só operadores e
    # índices dentro do ciclo (ver $DupVazias): o sufixo sai com uma expressão regular equivalente ao ciclo anterior (o
    # primeiro de ings, ing, edly, ed, es, s que deixe pelo menos 4 letras), as pontas "." e "-" com outra.
    $vistas = @{}
    $palavras = @(foreach ($m in [regex]::Matches($s, '[a-z0-9][a-z0-9.\-]*')) {
        $w = $m.Value -replace '^[.\-]+|[.\-]+$', ''
        if (($w.Length -lt 2 -and $w -notmatch '^\d+$') -or $DupVazias[$w]) { continue }
        if ($w -notmatch '\d') {
            $w = $w -replace '^(.{4,}?)(?:ings|ing|edly|ed|es|s)$', '$1'
            $sin = $DupSinonimos[$w]; if ($sin) { $w = $sin }
        }
        if (-not $vistas[$w]) { $vistas[$w] = $true; $w }
    })
    return , $palavras
}

function Join-NoticiasDuplicadas($Lista) {
    # desempate pela chave (o título normalizado, único): sem ele a ordem dos empates dependia da ordem da tabela de origem
    # (aleatória entre execuções no PowerShell 7) e do algoritmo de ordenação (diferente no 5.1), e os grupos mudavam com ela
    $L = @($Lista | Sort-Object -Property @{ Expression = 'score'; Descending = $true }, @{ Expression = 'data'; Descending = $true }, @{ Expression = 'chave'; Descending = $false })
    $n = $L.Count
    if ($n -lt 2) { return $L }
    # (só índices e operadores nos ciclos: ver $DupVazias. $tokH[i] = as palavras do título i, para consultas pelo índice;
    # $dia[i] = o instante em ticks, para comparar sem chamar métodos)
    $tok = New-Object object[] $n; $tokH = New-Object object[] $n; $nums = New-Object object[] $n; $emp = New-Object object[] $n; $t = New-Object object[] $n; $dia = New-Object object[] $n
    $df = @{}; $post = @{}
    for ($i = 0; $i -lt $n; $i++) {
        $tok[$i] = Get-PalavrasTitulo $L[$i].titulo
        $th = @{}; foreach ($w in $tok[$i]) { $th[$w] = $true }; $tokH[$i] = $th
        $nums[$i] = @(foreach ($w in $tok[$i]) { if ($w -match '\d') { $w } })
        $emp[$i] = @(foreach ($e in @($L[$i].empresas)) { if ($e -ne 'MKT') { $e } })
        $t[$i] = if ($L[$i].data) { [DateTimeOffset]::Parse($L[$i].data, $Script:Inv) } else { $null }
        $dia[$i] = if ($t[$i]) { $t[$i].UtcTicks } else { $null }
        foreach ($w in $tok[$i]) {
            if ($df[$w]) { $df[$w]++; $post[$w] += , $i } else { $df[$w] = 1; $post[$w] = @($i) }
        }
    }
    $idf = @{}; foreach ($w in $df.Keys) { $idf[$w] = [math]::Log(($n + 1) / ($df[$w] + 0.5)) }
    $massa = New-Object double[] $n
    for ($i = 0; $i -lt $n; $i++) { foreach ($w in $tok[$i]) { $massa[$i] += $idf[$w] } }
    $raro = [math]::Max(4, 0.015 * $n)   # palavra "rara": aparece em até 1,5% dos títulos

    # Medidas de semelhança entre duas notícias, ou $null quando não podem ser a mesma
    $medir = {
        param($a, $b)
        if ($null -ne $dia[$a] -and $null -ne $dia[$b] -and ($dia[$a] - $dia[$b] -gt 2592000000000 -or $dia[$b] - $dia[$a] -gt 2592000000000)) { return $null }      # mais de 3 dias de distância (em ticks: exato)
        if ($emp[$a].Count -and $emp[$b].Count) { $comum = $false; foreach ($e in $emp[$a]) { if ($emp[$b] -contains $e) { $comum = $true } }; if (-not $comum) { return $null } }   # empresas diferentes
        $soA = 0; foreach ($x in $nums[$a]) { if ($nums[$b] -notcontains $x) { $soA++ } }
        $soB = 0; foreach ($x in $nums[$b]) { if ($nums[$a] -notcontains $x) { $soB++ } }
        if ($soA -and $soB) { return $null }                                                               # números diferentes (datas, valores)
        $k = 0; $ms = 0.0; $nr = 0
        $hb = $tokH[$b]; foreach ($w in $tok[$a]) { if ($hb[$w]) { $k++; $ms += $idf[$w]; if ($df[$w] -le $raro) { $nr++ } } }
        if (-not $k) { return $null }
        @{ n = $k; ms = $ms; nr = $nr; wj = $ms / ($massa[$a] + $massa[$b] - $ms); ov = $ms / $(if ($massa[$a] -lt $massa[$b]) { $massa[$a] } else { $massa[$b] }) }
    }
    $forte = { param($f) $f -and $f.n -ge 3 -and ($f.wj -ge 0.5 -or ($f.ov -ge 0.75 -and $f.wj -ge 0.3 -and $f.ms -ge 11 -and $f.n -ge 4) -or ($f.nr -ge 2 -and $f.ov -ge 0.6 -and $f.wj -ge 0.2 -and $f.n -ge 4)) }
    $fraca = { param($f) $f -and ($f.wj -ge 0.2 -or ($f.nr -ge 2 -and $f.ov -ge 0.5)) }

    # Pares candidatos: só os que partilham 3 ou mais palavras (índice invertido, evita comparar tudo com tudo)
    $viz = New-Object object[] $n
    for ($i = 0; $i -lt $n; $i++) { $viz[$i] = New-Object 'System.Collections.Generic.HashSet[int]' }
    $conta = New-Object int[] $n; $tocados = New-Object int[] $n
    for ($i = 0; $i -lt $n; $i++) {
        $nt = 0   # $tocados[0..nt-1]: as notícias seguintes que partilham palavras com i, pela ordem em que aparecem
        foreach ($w in $tok[$i]) { foreach ($j in $post[$w]) { if ($j -gt $i) { if ($conta[$j] -eq 0) { $tocados[$nt] = $j; $nt++ }; $conta[$j]++ } } }
        for ($q = 0; $q -lt $nt; $q++) {
            $j = $tocados[$q]
            if ($conta[$j] -ge 3 -and (& $forte (& $medir $i $j))) { [void]$viz[$i].Add($j); [void]$viz[$j].Add($i) }
            $conta[$j] = 0
        }
    }

    # Agrupamento: cada notícia (da mais pontuada para a menos) junta-se ao grupo cuja notícia principal
    # é semelhante, ou com o qual é semelhante a pelo menos um terço dos membros (mínimo 2)
    $grupoDe = New-Object int[] $n; for ($i = 0; $i -lt $n; $i++) { $grupoDe[$i] = -1 }
    $grupos = New-Object System.Collections.Generic.List[object]
    # Time span of each group (earliest and latest story). A group never spans more than 3 days: pairs are compared
    # one at a time, so without this limit a chain A~B~C of related stories could join events a week apart.
    $gMin = New-Object System.Collections.Generic.List[object]; $gMax = New-Object System.Collections.Generic.List[object]
    $cabe = { param($g, $ta, $tb)   # $ta/$tb: earliest and latest dates being added (or $null)
        $lo = $gMin[$g]; $hi = $gMax[$g]
        if ($ta) { if (-not $lo -or $ta -lt $lo) { $lo = $ta } }
        if ($tb) { if (-not $hi -or $tb -gt $hi) { $hi = $tb } }
        -not ($lo -and $hi) -or ($hi - $lo).TotalDays -le 3 }
    for ($i = 0; $i -lt $n; $i++) {
        $melhor = -1; $melhorK = 0.0
        foreach ($g in @($viz[$i] | Where-Object { $grupoDe[$_] -ge 0 } | ForEach-Object { $grupoDe[$_] } | Select-Object -Unique)) {
            $membros = $grupos[$g]
            if (-not (& $cabe $g $t[$i] $t[$i])) { continue }
            $k = @($membros | Where-Object { $viz[$i].Contains($_) }).Count
            if ($viz[$i].Contains($membros[0]) -or $k -ge [math]::Max(2, [math]::Ceiling($membros.Count / 3))) {
                if ($k / $membros.Count -gt $melhorK) { $melhorK = $k / $membros.Count; $melhor = $g }
            }
        }
        if ($melhor -ge 0) {
            $grupos[$melhor].Add($i); $grupoDe[$i] = $melhor
            if ($t[$i]) { if (-not $gMin[$melhor] -or $t[$i] -lt $gMin[$melhor]) { $gMin[$melhor] = $t[$i] }; if (-not $gMax[$melhor] -or $t[$i] -gt $gMax[$melhor]) { $gMax[$melhor] = $t[$i] } }
        }
        else { $novo = New-Object System.Collections.Generic.List[int]; $novo.Add($i); $grupoDe[$i] = $grupos.Count; $grupos.Add($novo); $gMin.Add($t[$i]); $gMax.Add($t[$i]) }
    }
    # Junta grupos que são o mesmo acontecimento (ficaram separados pela ordem de chegada)
    do {
        $mudou = $false
        for ($i = 0; $i -lt $n -and -not $mudou; $i++) {
            foreach ($j in $viz[$i]) {
                $ga = $grupoDe[$i]; $gb = $grupoDe[$j]
                if ($ga -eq $gb) { continue }
                $A = $grupos[$ga]; $B = $grupos[$gb]
                if (-not (& $cabe $ga $gMin[$gb] $gMax[$gb])) { continue }   # together they would span more than 3 days
                $cruz = 0; foreach ($x in $A) { foreach ($y in $B) { if ($viz[$x].Contains($y)) { $cruz++ } } }
                if ($viz[$A[0]].Contains($B[0]) -or $cruz -ge [math]::Max(2, [math]::Ceiling([math]::Min($A.Count, $B.Count) / 2)) -or (& $fraca (& $medir $A[0] $B[0]))) {
                    foreach ($y in $B) { $A.Add($y); $grupoDe[$y] = $ga }
                    if ($gMin[$gb] -and (-not $gMin[$ga] -or $gMin[$gb] -lt $gMin[$ga])) { $gMin[$ga] = $gMin[$gb] }
                    if ($gMax[$gb] -and (-not $gMax[$ga] -or $gMax[$gb] -gt $gMax[$ga])) { $gMax[$ga] = $gMax[$gb] }
                    $A.Sort(); $B.Clear(); $mudou = $true; break
                }
            }
        }
    } while ($mudou)

    $saida = New-Object System.Collections.Generic.List[object]
    foreach ($g in $grupos) {
        if ($g.Count -eq 0) { continue }
        $principal = $L[$g[0]]
        if ($g.Count -gt 1) {
            $outras = @(foreach ($x in @($g)[1..($g.Count - 1)]) { $o = $L[$x]; [pscustomobject]@{ titulo = $o.titulo; link = $o.link; fonte = $o.fonte; tier = $o.tier; data = $o.data; chave = $o.chave } })
            $principal.outras = $outras
            $principal.empresas = @(@(foreach ($x in $g) { @($L[$x].empresas) }) | Select-Object -Unique)
        }
        $saida.Add($principal)
    }
    return $saida.ToArray()
}

# Pontos isolados impossíveis: um fecho mais de 50 % acima (ou abaixo) dos dois vizinhos, quando os vizinhos concordam
# entre si (até 10 %). É um erro do fornecedor (um "tick" errado), não um movimento: um desdobramento ou uma queda real
# mudam o nível de forma duradoura e os vizinhos já não concordam. O último ponto não tem vizinho seguinte e fica sempre.
function Remove-PicoIsolado($Pontos) {
    $P = @($Pontos); $fora = 0
    if ($P.Count -lt 3) { return [pscustomobject]@{ pontos = $P; removidos = 0 } }
    $out = New-Object System.Collections.Generic.List[object]; $out.Add($P[0])
    for ($i = 1; $i -lt $P.Count - 1; $i++) {
        $a = [double]$out[$out.Count - 1][1]; $v = [double]$P[$i][1]; $b = [double]$P[$i + 1][1]
        if ($a -gt 0 -and $b -gt 0 -and [math]::Abs($b / $a - 1) -le 0.10 -and (($v / $a -gt 1.5 -and $v / $b -gt 1.5) -or ($v / $a -lt 1 / 1.5 -and $v / $b -lt 1 / 1.5))) { $fora++; continue }
        $out.Add($P[$i])
    }
    $out.Add($P[-1])
    return [pscustomobject]@{ pontos = @(foreach ($x in $out) { , @($x[0], $x[1]) }); removidos = $fora }
}

# Alternativa ao Yahoo para as ações dos EUA: o histórico diário público da Nasdaq (o mesmo serviço das datas de resultados),
# 1 ano, em USD. Só é aceite se coincidir com a série da execução anterior nas datas em comum (pelo menos 20, mediana
# dentro de 1 % e cada data dentro de 3 %): assim um histórico com outro ajuste a desdobramentos, ou de outro título, nunca
# passa por bom. Sem execução anterior para comparar, não é usado (nada é inventado). Uma sessão de hoje ainda aberta
# (antes das 16:15 de Nova Iorque) fica de fora: não é um fecho.
function Get-SerieNasdaq([hashtable]$Ativo, $Referencia) {
    $u = "https://api.nasdaq.com/api/quote/$([uri]::EscapeDataString($Ativo.Nasdaq))/historical?assetclass=stocks&fromdate=$($Script:Agora.UtcDateTime.AddYears(-1).AddDays(-7).ToString('yyyy-MM-dd', $Script:Inv))&limit=9999&todate=$($Script:Agora.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv))"
    $j = (Get-Url $u) | ConvertFrom-Json
    if (-not $j.data) { throw "Nasdaq: $(@($j.status.bCodeMessage | ForEach-Object { $_.errorMessage }) -join ' ')".Trim() }
    if ("$($j.data.symbol)" -ne $Ativo.Nasdaq) { throw "Nasdaq returned '$($j.data.symbol)' instead of $($Ativo.Nasdaq)" }
    $porData = @{}; $inv = 0; $ny = [TimeZoneInfo]::ConvertTime($Script:Agora, (Get-FusoNY))
    foreach ($l in @($j.data.tradesTable.rows)) {
        $d = [datetime]::MinValue; $c = 0.0
        if (-not $l -or -not [datetime]::TryParseExact("$($l.date)", 'MM/dd/yyyy', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$d) -or
            -not [double]::TryParse(("$($l.close)" -replace '[$,\s]', ''), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$c) -or $c -le 0 -or [double]::IsInfinity($c)) { $inv++; continue }
        $iso = $d.ToString('yyyy-MM-dd', $Script:Inv)
        if ($iso -gt $ny.ToString('yyyy-MM-dd', $Script:Inv) -or ($iso -eq $ny.ToString('yyyy-MM-dd', $Script:Inv) -and ($ny.Hour * 60 + $ny.Minute) -lt 975)) { continue }
        $porData[$iso] = [math]::Round($c, 4)
    }
    $pts = @(foreach ($k in @($porData.Keys | Sort-Object)) { , @($k, $porData[$k]) })
    if ($pts.Count -lt 5) { throw 'Nasdaq returned no valid prices' }
    $ref = @{}; foreach ($p in @($Referencia)) { $q = @($p); if ($q.Count -ge 2) { $ref["$($q[0])"] = [double]$q[1] } }
    $razoes = @(foreach ($p in $pts) { if ($ref.ContainsKey($p[0]) -and $ref[$p[0]] -gt 0) { $p[1] / $ref[$p[0]] } })
    if ($razoes.Count -lt 20) { throw "Nasdaq prices could not be checked against the previous run ($($razoes.Count) dates in common, at least 20 needed): not used" }
    $ord = @($razoes | Sort-Object); $med = $ord[[math]::Floor(($ord.Count - 1) / 2)]
    if ([math]::Abs($med - 1) -gt 0.01 -or @($razoes | Where-Object { [math]::Abs($_ - 1) -gt 0.03 }).Count) { throw ('Nasdaq prices do not match the previous run''s (median ratio {0}): not used' -f $med.ToString('0.####', $Script:Inv)) }
    return [pscustomobject]@{ moeda = 'USD'; ultimo = $pts[-1][1]; pontos = $pts; fonte = 'Nasdaq'; splits = @(); parcial = $false; hora = $null; url = $u; invalidos = $inv }
}

# Série diária de 1 ano: Yahoo Finance (não oficial) com Stooq como alternativa.
# Com -Desde (data Unix), devolve o histórico diário completo desde essa data, só do Yahoo (as alternativas não têm tanto histórico).
function Get-Serie([hashtable]$Ativo, [int64]$Desde = 0, $Referencia = $null) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $nome = if ($Desde) { "Price history: $($Ativo.Nome)" } else { "Prices: $($Ativo.Nome)" }
    try {
        $periodo = if ($Desde) { "period1=$Desde&period2=$($Script:Agora.ToUnixTimeSeconds())" } else { 'range=1y' }
        $u = "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($Ativo.Yahoo))?$periodo&interval=1d&includePrePost=false&events=split"
        $j = (Get-Url $u) | ConvertFrom-Json
        # Response checks: a failure must not turn into plausible-looking but wrong data
        if ($j.chart.error) { throw "Yahoo error: $($j.chart.error.description)" }
        $res = @($j.chart.result)[0]
        if (-not $res -or -not $res.meta) { throw 'Unexpected Yahoo response (no chart result)' }
        if ("$($res.meta.symbol)" -and "$($res.meta.symbol)" -ne $Ativo.Yahoo) { throw "Yahoo returned $($res.meta.symbol) instead of $($Ativo.Yahoo)" }
        if ($Ativo.Moeda -and "$($res.meta.currency)" -and "$($res.meta.currency)" -ne $Ativo.Moeda) { throw "Unexpected currency $($res.meta.currency) (expected $($Ativo.Moeda))" }
        $ts = @($res.timestamp); $cl = @($res.indicators.quote[0].close)
        if ($ts.Count -ne $cl.Count) { throw 'Timestamps and prices do not match' }
        # O dia de cada cotação conta no fuso da bolsa: no câmbio, o Yahoo marca a sessão às 00:00 de Londres,
        # que no verão são 23:00 UTC do dia anterior (em UTC as datas ficavam um dia atrasadas).
        $tzWin = @{ 'America/New_York' = 'Eastern Standard Time'; 'America/Chicago' = 'Central Standard Time'; 'Europe/London' = 'GMT Standard Time'; 'Europe/Berlin' = 'W. Europe Standard Time' }["$($res.meta.exchangeTimezoneName)"]
        $tzBolsa = $null; if ($tzWin) { try { $tzBolsa = [TimeZoneInfo]::FindSystemTimeZoneById($tzWin) } catch { } }
        $diaBolsa = { param($instante) if ($tzBolsa) { [TimeZoneInfo]::ConvertTime($instante, $tzBolsa).ToString('yyyy-MM-dd', $Script:Inv) } else { $instante.AddSeconds([int]$res.meta.gmtoffset).UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv) } }
        $porData = @{}; $invalidos = 0
        for ($i = 0; $i -lt $ts.Count; $i++) {
            if ($null -eq $cl[$i]) { continue }
            $v = [double]$cl[$i]
            if ([double]::IsNaN($v) -or [double]::IsInfinity($v) -or $v -le 0) { $invalidos++; continue }   # a price is never zero or negative
            $porData[(& $diaBolsa ([DateTimeOffset]::FromUnixTimeSeconds([int64]$ts[$i])))] = [math]::Round($v, 4)
        }
        $pts = @(foreach ($k in @($porData.Keys | Sort-Object)) { , @($k, $porData[$k]) })
        # um ponto isolado muito fora dos vizinhos (erro do fornecedor, não um movimento real) sai e é contado
        $limpo = Remove-PicoIsolado $pts; $pts = $limpo.pontos; $invalidos += $limpo.removidos
        if ($pts.Count -lt 5) { throw 'Empty series' }
        # Stock splits reported by Yahoo (its prices are already adjusted for them): the site uses them to keep
        # purchases registered before a split in the right number of shares
        $splits = @()
        if ($res.events -and $res.events.splits) {
            $lidos = @(foreach ($p in @($res.events.splits.PSObject.Properties)) {
                $s = $p.Value; $num = 0.0; $den = 0.0
                if (-not $s -or -not $s.date -or -not [double]::TryParse("$($s.numerator)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$num) -or -not [double]::TryParse("$($s.denominator)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$den)) { continue }
                if ($num -le 0 -or $den -le 0 -or $num -eq $den) { continue }
                $f = $num / $den; if ($f -lt 0.001 -or $f -gt 1000) { continue }
                , @((& $diaBolsa ([DateTimeOffset]::FromUnixTimeSeconds([int64]$s.date))), [math]::Round($f, 6))
            })
            # (recriados num ciclo: os pares que saem de um pipeline o Windows PowerShell 5.1 escreve como {value, Count})
            $splits = @(foreach ($x in @($lidos | Sort-Object { $_[0] })) { , @("$($x[0])", [double]$x[1]) })
        }
        # Is the last point an intraday price (session still open when the data was collected)? Then it is not a close.
        $parcial = $false
        try {
            $sessao = $res.meta.currentTradingPeriod.regular; $ini = [int64]$sessao.start; $fim = [int64]$sessao.end; $hora = [int64]$res.meta.regularMarketTime
            if ($fim -gt 0 -and $hora -ge $ini -and $hora -lt $fim -and $Script:Agora.ToUnixTimeSeconds() -lt $fim) { $parcial = $true }
        } catch { }
        # Depois do fecho, o Yahoo às vezes ainda não tem a vela diária dessa sessão (visto a 7 out 2026 à 01:17: séries até 5 out,
        # com o fecho de 6 out só no regularMarketPrice/regularMarketTime). Com a sessão terminada, esse preço é o fecho oficial
        # do dia: entra como último ponto, se for plausível (a 50 % do anterior), e a fonte diz de onde veio.
        $notaFecho = ''
        if (-not $parcial -and $res.meta.regularMarketTime) {
            $tM = [DateTimeOffset]::FromUnixTimeSeconds([int64]$res.meta.regularMarketTime); $diaM = & $diaBolsa $tM; $pM = 0.0
            if ([string]::CompareOrdinal($diaM, $pts[-1][0]) -gt 0 -and $tM -le $Script:Agora.AddMinutes(5) -and [double]::TryParse("$($res.meta.regularMarketPrice)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$pM) -and $pM -gt 0 -and [math]::Abs($pM / [double]$pts[-1][1] - 1) -le 0.5) {
                $notaFecho = "close of $diaM from the quote (the daily series ended on $($pts[-1][0]))"
                $pts = @($pts) + @(, @($diaM, [math]::Round($pM, 4)))
            }
        }
        $horaUltimo = $null; if ($res.meta.regularMarketTime) { $horaUltimo = [DateTimeOffset]::FromUnixTimeSeconds([int64]$res.meta.regularMarketTime).ToString('o') }
        # o último ponto não pode ser comparado com o seguinte: um salto enorme fica assinalado na fonte (não é apagado)
        $notas = @(); if ($invalidos) { $notas += "$invalidos invalid price(s) ignored" }; if ($notaFecho) { $notas += $notaFecho }
        $salto = [double]$pts[-1][1] / [double]$pts[-2][1] - 1
        if ([math]::Abs($salto) -gt 0.5) { $notas += ('latest price {0:+0;-0}% from the previous close: check it' -f ($salto * 100)) }
        Add-Fonte $nome 'prices' $u 'ok' $pts.Count $sw.ElapsedMilliseconds ($notas -join '; ')
        return [pscustomobject]@{ moeda = "$($res.meta.currency)"; ultimo = [double]$res.meta.regularMarketPrice; pontos = $pts; fonte = 'Yahoo Finance'; splits = $splits; parcial = $parcial; hora = $horaUltimo }
    } catch {
        $erroYahoo = $_.Exception.Message
        if ($Desde) { Add-Fonte $nome 'prices' $Ativo.Yahoo 'error' 0 $sw.ElapsedMilliseconds $erroYahoo; return $null }
        if ($Ativo.Kraken) {
            # Kraken: velas diárias em UTC (até 720); a última é o dia em curso, com o preço atual como fecho
            try {
                $u3 = "https://api.kraken.com/0/public/OHLC?pair=$($Ativo.Kraken)&interval=1440"
                $k = (Get-Url $u3) | ConvertFrom-Json
                if (@($k.error).Count -and "$($k.error)") { throw "$($k.error)" }
                $par = @($k.result.PSObject.Properties | Where-Object { $_.Name -ne 'last' })[0]
                $desde = $Script:Agora.AddYears(-1).ToUnixTimeSeconds(); $invK = 0
                # o mesmo controlo do Yahoo: um preço nunca é zero, negativo ou não numérico
                $pts = @(foreach ($v in @($par.Value)) { $c = 0.0
                    if ([int64]$v[0] -lt $desde) { continue }
                    if (-not [double]::TryParse("$($v[4])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$c) -or [double]::IsNaN($c) -or [double]::IsInfinity($c) -or $c -le 0) { $invK++; continue }
                    , @([DateTimeOffset]::FromUnixTimeSeconds([int64]$v[0]).ToString('yyyy-MM-dd', $Script:Inv), [math]::Round($c, 4)) })
                if ($pts.Count -lt 5) { throw 'Kraken returned no data' }
                $limpo = Remove-PicoIsolado $pts; $pts = $limpo.pontos; $invK += $limpo.removidos
                Add-Fonte $nome 'prices' $u3 'ok (Kraken fallback)' $pts.Count $sw.ElapsedMilliseconds ("Yahoo failed: $erroYahoo" + $(if ($invK) { " | $invK invalid price(s) ignored" } else { '' }))
                return [pscustomobject]@{ moeda = $Ativo.Moeda; ultimo = $pts[-1][1]; pontos = $pts; fonte = 'Kraken'; splits = @(); parcial = $true; hora = $Script:Agora.ToString('o') }
            } catch {
                Add-Fonte $nome 'prices' $Ativo.Yahoo 'error' 0 $sw.ElapsedMilliseconds "Yahoo: $erroYahoo | Kraken: $($_.Exception.Message)"
                return $null
            }
        }
        # ações dos EUA: o histórico da Nasdaq, validado contra a execução anterior (ver Get-SerieNasdaq)
        $erros = "Yahoo: $erroYahoo"
        if ($Ativo.Nasdaq) {
            try {
                $n = Get-SerieNasdaq $Ativo $Referencia
                Add-Fonte $nome 'prices' $n.url 'ok (Nasdaq fallback)' $n.pontos.Count $sw.ElapsedMilliseconds ("Yahoo failed: $erroYahoo" + $(if ($n.invalidos) { " | $($n.invalidos) invalid price(s) ignored" } else { '' }))
                return $n
            } catch { $erros += " | Nasdaq: $($_.Exception.Message)" }
        }
        if (-not $Ativo.Stooq) { Add-Fonte $nome 'prices' $Ativo.Yahoo 'error' 0 $sw.ElapsedMilliseconds $(if ($Ativo.Nasdaq) { $erros } else { $erroYahoo }); return $null }
        try {
            $u2 = "https://stooq.com/q/d/l/?s=$($Ativo.Stooq)&i=d"
            $csvTexto = Get-Url $u2
            if ($csvTexto -match '^\s*<') { throw 'Stooq returned a web page instead of CSV (it blocks automated requests)' }
            $linhas = @($csvTexto | ConvertFrom-Csv)
            if ($linhas.Count -lt 5 -or -not $linhas[0].Close) { throw 'Stooq returned no data' }
            $desde = $Script:Agora.UtcDateTime.AddYears(-1).ToString('yyyy-MM-dd', $Script:Inv)
            $pts = @(foreach ($l in $linhas) { $c = 0.0
                if ("$($l.Date)" -notmatch '^\d{4}-\d{2}-\d{2}$' -or $l.Date -lt $desde) { continue }
                if (-not [double]::TryParse("$($l.Close)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$c) -or $c -le 0) { continue }
                , @($l.Date, [math]::Round($c, 4)) })
            if ($pts.Count -lt 5) { throw 'Stooq returned no valid prices' }
            Add-Fonte $nome 'prices' $u2 'ok (Stooq fallback)' $pts.Count $sw.ElapsedMilliseconds "Yahoo failed: $erroYahoo"
            return [pscustomobject]@{ moeda = $(if ($Ativo.Moeda) { $Ativo.Moeda } else { '' }); ultimo = $pts[-1][1]; pontos = $pts; fonte = 'Stooq'; splits = @(); parcial = $false; hora = $null }
        } catch {
            Add-Fonte $nome 'prices' $Ativo.Yahoo 'error' 0 $sw.ElapsedMilliseconds "$erros | Stooq: $($_.Exception.Message)"
            return $null
        }
    }
}

function Get-TaxaBCE {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $u = 'https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml'
    try {
        $x = ConvertTo-XmlSeguro (Get-Url $u)
        $pts = @(foreach ($dia in @($x.Envelope.Cube.Cube)) {
            $usd = @($dia.Cube) | Where-Object { $_.currency -eq 'USD' } | Select-Object -First 1
            if ($usd) { , @($dia.time, [double]::Parse($usd.rate, $Script:Inv)) }
        })
        $pts = @(foreach ($x in @($pts | Sort-Object { $_[0] })) { , @("$($x[0])", [double]$x[1]) })   # pares simples, não {value, Count}
        Add-Fonte 'ECB: EUR/USD reference rate (90 days)' 'currency' $u 'ok' $pts.Count $sw.ElapsedMilliseconds ''
        return $pts
    } catch {
        Add-Fonte 'ECB: EUR/USD reference rate (90 days)' 'currency' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message
        return @()
    }
}

# Metadados do ETF (de $ETFs) juntos aos pesos, para o site: vêm da configuração, com ou sem download
function Add-InfoETF($R, [hashtable]$Etf) {
    $info = [ordered]@{ id = $Etf.Id; chip = $Etf.Chip; fundo = $Etf.Fundo; isin = $Etf.Isin; indice = $Etf.Indice; benchmark = $Etf.Benchmark; moedaBase = $Etf.MoedaBase
        domicilio = $Etf.Domicilio; acumulacao = [bool]$Etf.Acumulacao; gestora = $Etf.Gestora; bolsa = $Etf.Bolsa; inicio = $Etf.Inicio; pagina = $Etf.Pagina; empresas = @($Etf.Empresas); paisIndice = "$($Etf.PaisIndice)" }
    foreach ($k in $info.Keys) { $R | Add-Member -NotePropertyName $k -NotePropertyValue $info[$k] -Force }
    return $R
}

# Agregados da exposição de um fundo, a partir do ficheiro da iShares (só agregados, nunca as posições):
# - setor e moeda subjacente (Market Currency) com os pesos de todas as linhas da folha Holdings; o que não é ação
#   (caixa, futuros, câmbio, fundos monetários) conta como "Cash/Other";
# - país, do bloco "Geography/Locations" da folha "Exposure Breakdowns" quando existe; senão, para um índice de um só
#   país (PaisIndice em $ETFs), esse país com o peso das ações, assinalado como "index" (o site diz que é aproximado).
# Cada agregado tem de somar cerca de 100 % (95 a 105); senão fica $null ("Unavailable").
function Get-AgregadosETF($Cab, $Csv, [int]$ColPeso, [hashtable]$Etf, [string]$Texto) {
    $colSetor = [array]::IndexOf($Cab, 'Sector'); $colClasse = [array]::IndexOf($Cab, 'Asset Class'); $colMoeda = [array]::IndexOf($Cab, 'Market Currency')
    $lista = { param($h) @(foreach ($k in @($h.Keys)) { [pscustomobject]@{ n = $k; w = [math]::Round([double]$h[$k], 2) } }) | Sort-Object w -Descending }
    $bate = { param($l) $t = (@($l) | Measure-Object w -Sum).Sum; $t -ge 95 -and $t -le 105 }
    $set = [ordered]@{}; $moe = [ordered]@{}; $acoes = 0.0; $outros = 0.0
    foreach ($l in $Csv) {
        $w = 0.0; if (-not [double]::TryParse(($l[$ColPeso] -replace ',', ''), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$w)) { continue }
        $acao = $colClasse -lt 0 -or $l[$colClasse] -eq 'Equity'
        if ($acao) { $acoes += $w } else { $outros += $w }
        $s = if (-not $acao) { 'Cash/Other' } elseif ($colSetor -ge 0 -and $l[$colSetor]) { $l[$colSetor] } else { 'Other' }
        $m = if (-not $acao) { 'Cash/Other' } elseif ($colMoeda -ge 0 -and $l[$colMoeda]) { $l[$colMoeda] } else { 'Other' }
        $set[$s] = [double]$set[$s] + $w; $moe[$m] = [double]$moe[$m] + $w
    }
    $paises = $null; $fonte = $null
    $exp = [regex]::Match($Texto, '(?s)<ss:Worksheet ss:Name="Exposure Breakdowns">(.*?)</ss:Worksheet>')
    if ($exp.Success) {
        $ls = @(foreach ($row in [regex]::Matches($exp.Groups[1].Value, '(?s)<ss:Row[^>]*>(.*?)</ss:Row>')) { , @([regex]::Matches($row.Groups[1].Value, '(?s)<ss:Data[^>]*>(.*?)</ss:Data>') | ForEach-Object { [Net.WebUtility]::HtmlDecode($_.Groups[1].Value).Trim() }) })
        $i0 = -1; for ($i = 0; $i -lt $ls.Count; $i++) { if (@($ls[$i])[0] -eq 'Geography/Locations') { $i0 = $i; break } }
        if ($i0 -ge 0) {
            $geo = [ordered]@{}
            for ($i = $i0 + 1; $i -lt $ls.Count; $i++) {
                $q = @($ls[$i]); if ($q.Count -lt 2) { if ($geo.Count) { break } else { continue } }
                if ($q[0] -in 'as of', 'Type') { continue }
                $w = 0.0; if (-not [double]::TryParse(($q[1] -replace ',', ''), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$w)) { continue }
                $n = if ($q[0] -eq 'Cash and/or Derivatives') { 'Cash/Other' } else { $q[0] }
                $geo[$n] = [double]$geo[$n] + $w
            }
            if ($geo.Count) { $paises = @(& $lista $geo); $fonte = 'file' }
        }
    }
    if (-not $paises -and $Etf.PaisIndice -and $acoes -gt 0) {
        $paises = @([pscustomobject]@{ n = $Etf.PaisIndice; w = [math]::Round($acoes, 2) })
        if ($outros -gt 0) { $paises += [pscustomobject]@{ n = 'Cash/Other'; w = [math]::Round($outros, 2) } }
        $fonte = 'index'
    }
    $setL = @(& $lista $set); $moeL = @(& $lista $moe)
    $pOut = $null; $fOut = $null; $sOut = $null; $mOut = $null
    if ($paises -and (& $bate $paises)) { $pOut = $paises; $fOut = $fonte }
    if ($setL.Count -and (& $bate $setL)) { $sOut = $setL }
    if ($moeL.Count -and (& $bate $moeL)) { $mOut = $moeL }
    return [pscustomobject]@{ paises = $pOut; setores = $sOut; moedas = $mOut; fontePaises = $fOut }
}

# Holdings de um ETF da iShares (sem argumento: o SXR8, como antes). Cada fundo tem o seu ficheiro e os seus pesos de
# referência; uma falha num fundo não afeta os outros.
function Get-PesosETF([hashtable]$Etf = $null) {
    if (-not $Etf) { $Etf = @($ETFs | Where-Object { $_.Id -eq 'SXR8' })[0] }
    if (-not $Etf) { $Etf = @{ Id = 'SXR8'; Isin = 'IE00B5BMR087'; Url = $UrlPesosETF; Referencia = $PesosReferencia; Empresas = @('AAPL', 'NVDA', 'GOOGL') } }
    $nomeFonte = "iShares: ETF holdings ($($Etf.Id))"; $urlEtf = $Etf.Url; $ref = $Etf.Referencia; $esperadas = @($Etf.Empresas)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $texto = Get-Url $urlEtf
        # O ficheiro é do fundo certo? (a folha "Key Facts" tem o ISIN; se a folha mudar de formato, este teste é saltado)
        $isinFicheiro = [regex]::Match($texto, '(?s)<ss:Data[^>]*>ISIN</ss:Data>.{0,400}?<ss:Data[^>]*>([A-Z]{2}[A-Z0-9]{9}\d)</ss:Data>')
        if ($Etf.Isin -and $isinFicheiro.Success -and $isinFicheiro.Groups[1].Value -ne $Etf.Isin) { throw "The iShares file is for $($isinFicheiro.Groups[1].Value), not $($Etf.Isin)" }
        # O ficheiro não é XML válido (tem "S&P" sem escape), por isso é lido com expressões regulares
        $folha = [regex]::Match($texto, '(?s)<ss:Worksheet ss:Name="Holdings">(.*?)</ss:Worksheet>')
        if (-not $folha.Success) { throw 'Sheet "Holdings" not found in the iShares file' }
        $linhas = @(foreach ($r in [regex]::Matches($folha.Groups[1].Value, '(?s)<ss:Row[^>]*>(.*?)</ss:Row>')) {
            , @([regex]::Matches($r.Groups[1].Value, '(?s)<ss:Data[^>]*>(.*?)</ss:Data>') | ForEach-Object { [Net.WebUtility]::HtmlDecode($_.Groups[1].Value).Trim() })
        })
        $dataRef = ''
        foreach ($l in $linhas) { if ($l.Count -ge 2 -and $l[0] -eq 'as of') { $dataRef = $l[1]; break } }
        $inicio = -1
        for ($i = 0; $i -lt $linhas.Count; $i++) { if ($linhas[$i] -contains 'Weight (%)') { $inicio = $i; break } }
        if ($inicio -lt 0) { throw 'Holdings header not found' }
        $colTicker = [array]::IndexOf($linhas[$inicio], 'Issuer Ticker'); if ($colTicker -lt 0) { $colTicker = [array]::IndexOf($linhas[$inicio], 'Ticker') }
        $colPeso = [array]::IndexOf($linhas[$inicio], 'Weight (%)')
        if ($colTicker -lt 0) { throw 'Ticker column not found' }
        $csv = @($linhas[($inicio + 1)..($linhas.Count - 1)] | Where-Object { $_.Count -gt $colPeso })
        $peso = { param($t) ($csv | Where-Object { $_[$colTicker] -eq $t } | ForEach-Object { [double]::Parse(($_[$colPeso] -replace ',', ''), $Script:Inv) } | Measure-Object -Sum).Sum }
        # Shape checks: the weights of all holdings must add up to about 100% and each of the three must be plausible
        $somaTotal = 0.0; foreach ($l in $csv) { $w = 0.0; if ([double]::TryParse(($l[$colPeso] -replace ',', ''), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$w)) { $somaTotal += $w } }
        if ($somaTotal -lt 80 -or $somaTotal -gt 120) { throw ('Holdings weights add up to {0:N1}%, not about 100%: the file format may have changed' -f $somaTotal) }
        # "as of" date of the file (e.g. "01/Oct/2026"), kept also in ISO so the site can tell how old it is
        $dataIso = ''; $dt = [datetime]::MinValue
        if ($dataRef -and [datetime]::TryParseExact($dataRef, [string[]]@('dd/MMM/yyyy', 'MMM dd, yyyy', 'dd MMM yyyy', 'yyyy-MM-dd', 'MM/dd/yyyy'), [Globalization.CultureInfo]::GetCultureInfo('en-US'), [Globalization.DateTimeStyles]::None, [ref]$dt)) { $dataIso = $dt.ToString('yyyy-MM-dd', $Script:Inv) }
        $r = [pscustomobject]@{ AAPL = [math]::Round((& $peso 'AAPL'), 2); NVDA = [math]::Round((& $peso 'NVDA'), 2); GOOGL = [math]::Round(((& $peso 'GOOGL') + (& $peso 'GOOG')), 2); data = $dataRef; dataIso = $dataIso; aoVivo = $true; fonte = 'iShares (BlackRock)'; top10 = @(); setores = @(); posicoes = 0 }
        # Plausibility: the companies this fund is expected to hold must have a weight between 0 and 25%; in the other
        # funds the three companies may be absent (0%), but never negative or above 25%
        if ($esperadas.Count -and (($esperadas | ForEach-Object { $r.$_ } | Measure-Object -Sum).Sum) -le 0) { throw 'All weights are zero' }
        foreach ($k in 'AAPL', 'NVDA', 'GOOGL') {
            if ($k -in $esperadas) { if ($r.$k -le 0 -or $r.$k -gt 25) { throw "Implausible weight for ${k}: $($r.$k)%" } }
            elseif ($r.$k -lt 0 -or $r.$k -gt 25) { throw "Implausible weight for ${k}: $($r.$k)%" }
        }
        $colMoeda = [array]::IndexOf($linhas[$inicio], 'Market Currency')
        # Maiores posições e setores (só ações; as classes A e C da Alphabet contam como uma empresa)
        $colNome = [array]::IndexOf($linhas[$inicio], 'Name'); $colSetor = [array]::IndexOf($linhas[$inicio], 'Sector'); $colClasse = [array]::IndexOf($linhas[$inicio], 'Asset Class')
        if ($colNome -ge 0 -and $colSetor -ge 0) {
            $titulo = (Get-Culture).TextInfo
            $acoes = @(foreach ($l in $csv) {
                if ($colClasse -ge 0 -and $l[$colClasse] -ne 'Equity') { continue }
                $w = 0.0; if (-not [double]::TryParse(($l[$colPeso] -replace ',', ''), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$w)) { continue }
                $tk = $l[$colTicker]
                $nm = $titulo.ToTitleCase($l[$colNome].ToLowerInvariant()) -replace '\.Com\b', '.com' -replace '(\s+(Inc|Corp|Co|Ltd|Plc|Class [A-C]))+\.?$', ''
                $fixo = @{ NVDA = 'NVIDIA'; AMD = 'AMD'; AMZN = 'Amazon'; JPM = 'JPMorgan Chase'; LLY = 'Eli Lilly'; 'BRKB' = 'Berkshire Hathaway'; XOM = 'Exxon Mobil'; UNH = 'UnitedHealth'; V = 'Visa'; MA = 'Mastercard' }[$tk]
                # (só para ações americanas: noutros mercados os mesmos tickers curtos podem ser outras empresas)
                if ($fixo -and ($colMoeda -lt 0 -or $l[$colMoeda] -eq 'USD')) { $nm = $fixo }   # (sem ??: o Windows PowerShell 5.1 não o conhece)
                if ($tk -in 'GOOGL', 'GOOG') { $tk = 'GOOGL'; $nm = 'Alphabet (A + C)' }
                [pscustomobject]@{ t = $tk; n = $nm; s = $l[$colSetor]; w = $w }
            })
            $r.posicoes = @($acoes | Where-Object { $_.t -ne 'GOOGL' }).Count + [int](@($acoes | Where-Object { $_.t -eq 'GOOGL' }).Count -gt 0)
            $r.top10 = @($acoes | Group-Object t | ForEach-Object { [pscustomobject]@{ t = $_.Name; n = $_.Group[0].n; s = $_.Group[0].s; w = [math]::Round(($_.Group | Measure-Object w -Sum).Sum, 2) } } |
                Sort-Object w -Descending | Select-Object -First 10)
            $r.setores = @($acoes | Group-Object s | ForEach-Object { [pscustomobject]@{ s = $_.Name; w = [math]::Round(($_.Group | Measure-Object w -Sum).Sum, 2) } } | Sort-Object w -Descending)
        }
        $r | Add-Member -NotePropertyName agregados -NotePropertyValue (Get-AgregadosETF $linhas[$inicio] $csv $colPeso $Etf $texto) -Force
        # mudança de formato do ficheiro: as colunas que faltam dizem-se na fonte (o que depende delas fica "Unavailable")
        $faltaCol = @('Name', 'Sector', 'Asset Class', 'Market Currency' | Where-Object { [array]::IndexOf($linhas[$inicio], $_) -lt 0 })
        $notaEtf = @(); if ($faltaCol.Count) { $notaEtf += "file layout changed: no column $($faltaCol -join ', ')" }
        if (-not $isinFicheiro.Success) { $notaEtf += 'no ISIN found in the file, so the fund could not be confirmed' }
        foreach ($k in 'paises', 'setores', 'moedas') { if ($null -eq $r.agregados.$k) { $notaEtf += "$(@{ paises = 'country'; setores = 'sector'; moedas = 'currency' }[$k]) breakdown Unavailable" } }
        Add-Fonte $nomeFonte 'ETF' $urlEtf 'ok' $csv.Count $sw.ElapsedMilliseconds ($notaEtf -join '; ')
        return (Add-InfoETF $r $Etf)
    } catch {
        Add-Fonte $nomeFonte 'ETF' $urlEtf 'error (using reference weights)' 0 $sw.ElapsedMilliseconds $_.Exception.Message
        $r = [pscustomobject]@{ AAPL = $ref.AAPL; NVDA = $ref.NVDA; GOOGL = $ref.GOOGL; data = $ref.Data; dataIso = $ref.Data; aoVivo = $false; fonte = "Reference weights in the script (BlackRock, $($ref.Data))"; top10 = @(); setores = @(); posicoes = 0; agregados = $null }
        return (Add-InfoETF $r $Etf)
    }
}

# Próxima data de resultados de cada empresa (Nasdaq, com dados da Zacks). O texto diz "is estimated to report"
# enquanto a data é uma previsão; quando a empresa a anuncia, deixa de dizer "estimated".
function Get-DatasResultados {
    $saida = New-Object System.Collections.Generic.List[object]
    foreach ($id in 'AAPL', 'NVDA', 'GOOGL') {
        $sw = [Diagnostics.Stopwatch]::StartNew(); $u = "https://api.nasdaq.com/api/analyst/$id/earnings-date"
        try {
            $j = (Get-Url $u) | ConvertFrom-Json
            $txt = "$($j.data.reportText) $($j.data.announcement)"
            $m = [regex]::Match($txt, '(\d{1,2})/(\d{1,2})/(\d{4})')
            if (-not $m.Success) { throw 'No date in the response' }
            $d = [datetime]::new([int]$m.Groups[3].Value, [int]$m.Groups[1].Value, [int]$m.Groups[2].Value)
            # Plausibility: the next earnings date is between yesterday and about 6 months ahead
            $distancia = ($d - $Script:Agora.UtcDateTime.Date).TotalDays
            if ($distancia -lt -2 -or $distancia -gt 200) { throw "Implausible date in the response: $($d.ToString('yyyy-MM-dd', $Script:Inv))" }
            $saida.Add([pscustomobject]@{ e = $id; d = $d.ToString('yyyy-MM-dd', $Script:Inv); st = $(if ($txt -match 'estimated') { 'E' } else { 'C' }) })
            Add-Fonte "Nasdaq: next earnings date ($id)" 'calendar' $u 'ok' 1 $sw.ElapsedMilliseconds ''
        } catch { Add-Fonte "Nasdaq: next earnings date ($id)" 'calendar' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message }
    }
    return $saida.ToArray()
}

# Dividendos por ação (USD) e rendimento, a partir dos eventos de dividendo do gráfico do Yahoo: o mesmo pedido público
# das cotações (o quoteSummary e o quote do Yahoo pedem autenticação desde 2023). Cada evento é a data ex-dividendo e o
# valor por ação, já ajustado aos desdobramentos posteriores, como os preços (ou seja, em ações de hoje).
# Dividendo anual = último pagamento × pagamentos por ano (a taxa "indicada"); a soma dos últimos 12 meses fica ao lado.
# O rendimento é calculado aqui (dividendo anual ÷ preço da mesma resposta) e guardado em pontos percentuais
# (0.32 = 0,32%): não depende do formato decimal ou percentual de nenhum fornecedor.
# Uma falha só afeta este bloco: o resto do site é construído na mesma, e nada é inventado.
function Get-Dividendo([hashtable]$Ativo) {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $nome = "Dividends: $($Ativo.Nome)"
    $u = "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($Ativo.Yahoo))?range=2y&interval=1mo&events=div"
    $num = { param($x, [ref]$v) [double]::TryParse("$x", [Globalization.NumberStyles]::Float, $Script:Inv, $v) -and -not [double]::IsNaN($v.Value) -and -not [double]::IsInfinity($v.Value) }
    try {
        $j = (Get-Url $u) | ConvertFrom-Json
        if ($j.chart.error) { throw "Yahoo error: $($j.chart.error.description)" }
        $res = @($j.chart.result)[0]
        if (-not $res -or -not $res.meta) { throw 'Unexpected Yahoo response (no chart result)' }
        if ("$($res.meta.symbol)" -ne $Ativo.Yahoo) { throw "Yahoo returned '$($res.meta.symbol)' instead of $($Ativo.Yahoo)" }
        if ("$($res.meta.currency)" -ne 'USD') { throw "Unexpected currency '$($res.meta.currency)' (expected USD)" }
        $preco = 0.0
        if (-not (& $num $res.meta.regularMarketPrice ([ref]$preco)) -or $preco -le 0) { throw 'No valid price in the response' }
        $agoraS = $Script:Agora.ToUnixTimeSeconds(); $gmt = 0; [void][int]::TryParse("$($res.meta.gmtoffset)", [ref]$gmt)
        $lidos = New-Object System.Collections.Generic.List[object]; $invalidos = 0
        if ($res.events -and $res.events.dividends) {
            foreach ($p in @($res.events.dividends.PSObject.Properties)) {
                $x = $p.Value; $v = 0.0; $t = [int64]0
                # valor por pagamento: positivo e abaixo de 25% do preço (acima disso é um erro de dados); data não muito no futuro
                if (-not $x -or -not [int64]::TryParse("$($x.date)", [ref]$t) -or -not (& $num $x.amount ([ref]$v)) -or $v -le 0 -or $v -gt $preco * 0.25 -or $t -gt $agoraS + 45 * 86400) { $invalidos++; continue }
                $lidos.Add([pscustomobject]@{ t = $t; v = $v })
            }
        }
        if (-not $lidos.Count) { throw 'No dividend payments in the provider data for the last 2 years' }
        $ord = @($lidos | Sort-Object t)
        # (pares recriados num ciclo simples: os que saem de um pipeline o Windows PowerShell 5.1 escreve como {value, Count})
        $pagamentos = @(foreach ($x in $ord) { , @([DateTimeOffset]::FromUnixTimeSeconds($x.t).AddSeconds($gmt).UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv), [math]::Round([double]$x.v, 6)) })
        # pagamentos por ano: pelo intervalo mediano entre pagamentos (mensal, trimestral, semestral ou anual)
        $freq = $null
        if ($ord.Count -ge 2) {
            $gaps = @(@(for ($i = 1; $i -lt $ord.Count; $i++) { ($ord[$i].t - $ord[$i - 1].t) / 86400.0 }) | Sort-Object)
            $mediana = $gaps[[math]::Floor(($gaps.Count - 1) / 2)]
            if ($mediana -gt 0) { $alvo = 365.25 / $mediana; $freq = @(12, 4, 2, 1) | Sort-Object { [math]::Abs([math]::Log($_ / $alvo)) } | Select-Object -First 1 }
        }
        $ultimo = $ord[-1]; $diasUltimo = ($agoraS - $ultimo.t) / 86400.0
        $ttm = 0.0; foreach ($x in $ord) { if ($x.t -gt $agoraS - 365 * 86400 -and $x.t -le $agoraS) { $ttm += $x.v } }
        $anual = $null; $rend = $null; $nota = ''
        if (-not $freq) { $nota = 'Only one dividend payment in the provider data: the payment frequency is unknown, so no annual figure is projected.' }
        elseif ($diasUltimo -gt (365.25 / $freq) * 1.5 + 31) { $nota = "The latest dividend is from $($pagamentos[-1][0]), longer ago than the usual interval: it may have been suspended, so no annual figure is projected." }
        else {
            $anual = [math]::Round($ultimo.v * $freq, 6); $rend = [math]::Round($anual / $preco * 100, 4)
            if ($rend -lt 0 -or $rend -gt 25) { throw ('Implausible dividend yield: {0}%' -f $rend.ToString($Script:Inv)) }
        }
        Add-Fonte $nome 'dividends' $u 'ok' $pagamentos.Count $sw.ElapsedMilliseconds $(if ($invalidos) { "$invalidos invalid payment(s) ignored" } else { '' })
        return [pscustomobject]@{
            id = $Ativo.Id; simbolo = $Ativo.Yahoo; moeda = 'USD'; estado = 'ok'; fonte = 'Yahoo Finance'; obtidoEm = $Script:Agora.ToString('o')
            anualPorAcao = $anual; ttmPorAcao = [math]::Round($ttm, 6); frequencia = $freq; rendimentoPct = $rend
            ultimo = @($pagamentos[-1][0], $pagamentos[-1][1]); pagamentos = $pagamentos
            preco = [math]::Round($preco, 4); precoData = $(if ($res.meta.regularMarketTime) { [DateTimeOffset]::FromUnixTimeSeconds([int64]$res.meta.regularMarketTime).ToString('o') } else { $null })
            nota = $nota; erro = ''
        }
    } catch {
        Add-Fonte $nome 'dividends' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message
        return [pscustomobject]@{ id = $Ativo.Id; simbolo = $Ativo.Yahoo; moeda = 'USD'; estado = 'error'; fonte = ''; obtidoEm = $Script:Agora.ToString('o'); erro = $_.Exception.Message }
    }
}
function Get-Dividendos {
    $r = [ordered]@{}
    foreach ($id in $AtivosDividendos) {
        # um ETF de acumulação não distribui: nunca entra nos dividendos projetados, mesmo que alguém o junte à lista
        if (@($ETFs | Where-Object { $_.Id -eq $id -and $_.Acumulacao }).Count) { continue }
        $a = $Ativos | Where-Object { $_.Id -eq $id } | Select-Object -First 1
        if ($a) { $r[$id] = Get-Dividendo $a }
    }
    return $r
}

# Indicadores da Bitcoin: sentimento (Fear & Greed), mercado e domínio (CoinGecko), rede e halving (mempool.space).
# Cada parte é independente: se uma fonte falhar, as outras continuam e o site mostra o que houver.
function Get-DadosBitcoin {
    $r = [ordered]@{ sentimento = $null; mercado = $null; rede = $null }

    $sw = [Diagnostics.Stopwatch]::StartNew(); $u = 'https://api.alternative.me/fng/?limit=365&format=json'
    try {
        $j = (Get-Url $u) | ConvertFrom-Json
        $serie = @(foreach ($x in @($j.data)) { , @([DateTimeOffset]::FromUnixTimeSeconds([int64]$x.timestamp).ToString('yyyy-MM-dd', $Script:Inv), [int]$x.value) })
        if (-not $serie.Count) { throw 'No values' }
        $hoje = @($j.data)[0]
        if ([int]$hoje.value -lt 0 -or [int]$hoje.value -gt 100) { throw "Index out of range: $($hoje.value)" }   # the index is 0–100
        $serie = @(foreach ($x in $serie) { if ($x[1] -ge 0 -and $x[1] -le 100) { , $x } })
        [array]::Reverse($serie)
        $r.sentimento = [ordered]@{ valor = [int]$hoje.value; classe = "$($hoje.value_classification)"; serie = $serie }
        Add-Fonte 'Alternative.me: Bitcoin Fear & Greed Index' 'bitcoin' $u 'ok' $serie.Count $sw.ElapsedMilliseconds ''
    } catch { Add-Fonte 'Alternative.me: Bitcoin Fear & Greed Index' 'bitcoin' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message }

    $sw = [Diagnostics.Stopwatch]::StartNew(); $u = 'https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=eur,usd&include_market_cap=true&include_24hr_change=true&include_24hr_vol=true'
    try {
        $p = ((Get-Url $u) | ConvertFrom-Json).bitcoin
        if (-not $p.eur -or [double]$p.eur -le 0) { throw 'No price' }
        # campos em falta ou inválidos ficam $null (o site mostra "—" ou a variação desde as 00:00 UTC): [double]$null seria 0,
        # e uma variação de 0,00% ou um domínio de 0% pareceriam dados verdadeiros
        $num = { param($x, [double]$Min = [double]::MinValue) $v = 0.0; if ($null -ne $x -and "$x" -ne '' -and [double]::TryParse("$x", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$v) -and -not [double]::IsNaN($v) -and -not [double]::IsInfinity($v) -and $v -ge $Min) { $v } else { $null } }
        $v24 = & $num $p.eur_24h_change -100
        $m = [ordered]@{ eur = [double]$p.eur; usd = (& $num $p.usd 0); var24 = $(if ($null -ne $v24) { [math]::Round($v24, 2) } else { $null }); capEur = (& $num $p.eur_market_cap 0); volEur = (& $num $p.eur_24h_vol 0); dominio = $null; capTotalEur = $null }
        $sw2 = [Diagnostics.Stopwatch]::StartNew(); $u2 = 'https://api.coingecko.com/api/v3/global'
        try {
            $g = ((Get-Url $u2) | ConvertFrom-Json).data
            $dom = & $num $g.market_cap_percentage.btc 0
            if ($null -eq $dom -or $dom -gt 100) { throw "Unexpected Bitcoin dominance: '$($g.market_cap_percentage.btc)'" }
            $m.dominio = [math]::Round($dom, 2); $m.capTotalEur = & $num $g.total_market_cap.eur 0
            Add-Fonte 'CoinGecko: Bitcoin dominance' 'bitcoin' $u2 'ok' 1 $sw2.ElapsedMilliseconds ''
        } catch { Add-Fonte 'CoinGecko: Bitcoin dominance' 'bitcoin' $u2 'error' 0 $sw2.ElapsedMilliseconds $_.Exception.Message }
        $r.mercado = $m
        Add-Fonte 'CoinGecko: Bitcoin price, market cap and volume' 'bitcoin' $u 'ok' 1 $sw.ElapsedMilliseconds ''
    } catch { Add-Fonte 'CoinGecko: Bitcoin price, market cap and volume' 'bitcoin' $u 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message }

    $sw = [Diagnostics.Stopwatch]::StartNew(); $u = 'https://mempool.space/api/blocks/tip/height'
    try {
        $alturaTxt = (Get-Url $u).Trim(); $altura = [int64]0
        # the chain only grows: a height below 900,000 (passed in 2025) or not a number means a bad response
        if (-not [int64]::TryParse($alturaTxt, [ref]$altura) -or $altura -lt 900000) { throw "Unexpected block height: $($alturaTxt.Substring(0, [math]::Min(40, $alturaTxt.Length)))" }
        $ajuste = $null; try { $ajuste = (Get-Url 'https://mempool.space/api/v1/difficulty-adjustment') | ConvertFrom-Json } catch { }
        $hash = $null; try { $hash = ((Get-Url 'https://mempool.space/api/v1/mining/hashrate/1m') | ConvertFrom-Json).currentHashrate } catch { }
        $epoca = [math]::Floor($altura / 210000)
        $proximo = ($epoca + 1) * 210000
        $msBloco = if ($ajuste -and $ajuste.timeAvg -gt 0) { [double]$ajuste.timeAvg } else { 600000 }   # média recente (só para mostrar)
        # Halvings already passed (fixed list; any later one is looked up once on mempool.space)
        $passados = @(Get-Halvings $epoca)
        # Forecast pace: average block time since the last halving. It is a long, stable window (the average of the last
        # few days swung so much that the date jumped weeks between runs) and, unlike a fixed 10 min, it follows the
        # usual trend of blocks coming slightly faster while the hash rate grows. Kept within 8–12 min.
        $msPrevisao = 600000
        $ultimoH = $passados | Where-Object { $_.altura -eq $epoca * 210000 } | Select-Object -First 1
        if ($ultimoH -and ($altura - $ultimoH.altura) -ge 2016) { $msPrevisao = ($Script:Agora - $ultimoH.t).TotalMilliseconds / ($altura - $ultimoH.altura) }
        $msPrevisao = [math]::Min(720000, [math]::Max(480000, $msPrevisao))
        $previsto = $Script:Agora.AddMilliseconds(($proximo - $altura) * $msPrevisao)
        $r.rede = [ordered]@{
            altura = $altura; halvingAltura = $proximo; blocosFalta = $proximo - $altura; minBloco = [math]::Round($msBloco / 60000, 1); minBlocoEpoca = [math]::Round($msPrevisao / 60000, 2)
            halvings = @(foreach ($h in $passados) { , @($h.t.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv), $h.n) })
            halvingPrevisto = $previsto.ToString('o'); recompensa = 50 / [math]::Pow(2, $epoca); recompensaNova = 50 / [math]::Pow(2, $epoca + 1)
            hashrateEH = $(if ($hash) { [math]::Round([double]$hash / 1e18, 0) } else { $null })
            ajusteDif = $(if ($ajuste -and $null -ne $ajuste.difficultyChange -and "$($ajuste.difficultyChange)" -ne '') { [math]::Round([double]$ajuste.difficultyChange, 2) } else { $null })
            ajusteData = $(if ($ajuste -and $ajuste.estimatedRetargetDate) { [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$ajuste.estimatedRetargetDate).ToString('o') } else { $null })
        }
        $r.rede.evento = [ordered]@{ d = $previsto.ToString('yyyy-MM-dd', $Script:Inv); e = 'BTC'; imp = 'High'; st = 'E'
            ev = 'Bitcoin halving (block {0}): the block reward drops from {1} to {2} BTC' -f $proximo.ToString('N0', $Script:Inv), $r.rede.recompensa.ToString('0.#####', $Script:Inv), $r.rede.recompensaNova.ToString('0.#####', $Script:Inv) }
        Add-Fonte 'mempool.space: Bitcoin network and halving' 'bitcoin' 'https://mempool.space/' 'ok' 1 $sw.ElapsedMilliseconds ''
    } catch { Add-Fonte 'mempool.space: Bitcoin network and halving' 'bitcoin' 'https://mempool.space/' 'error' 0 $sw.ElapsedMilliseconds $_.Exception.Message }
    return $r
}

# Halvings: block and UTC time of the ones that already happened. Halvings after these are looked up on mempool.space
# (block at height 210,000 × n), so the list does not need editing when a new one happens.
$HalvingsConhecidos = @(
    @{ n = 1; altura = 210000; t = '2012-11-28T15:24:38Z' },
    @{ n = 2; altura = 420000; t = '2016-07-09T16:46:13Z' },
    @{ n = 3; altura = 630000; t = '2020-05-11T19:23:43Z' },
    @{ n = 4; altura = 840000; t = '2024-04-20T00:09:27Z' }
)
function Get-Halvings([int]$Epoca) {
    $out = New-Object System.Collections.Generic.List[object]
    for ($n = 1; $n -le $Epoca; $n++) {
        $c = $HalvingsConhecidos | Where-Object { $_.n -eq $n } | Select-Object -First 1
        if ($c) { $out.Add([pscustomobject]@{ n = $n; altura = $c.altura; t = [DateTimeOffset]::Parse($c.t, $Script:Inv) }); continue }
        try {
            $hashBloco = (Get-Url "https://mempool.space/api/block-height/$($n * 210000)").Trim()
            if ($hashBloco -notmatch '^[0-9a-f]{64}$') { throw 'bad hash' }
            $b = (Get-Url "https://mempool.space/api/block/$hashBloco") | ConvertFrom-Json
            $out.Add([pscustomobject]@{ n = $n; altura = $n * 210000; t = [DateTimeOffset]::FromUnixTimeSeconds([int64]$b.timestamp) })
        } catch { }   # without it, the chart just has no mark for that halving
    }
    return $out.ToArray()
}
# Approximate next halving when mempool.space is unavailable: from the last known one, 210,000 blocks of 10 minutes each
function Get-HalvingAproximado {
    $u = $HalvingsConhecidos[-1]; $n = $u.n; $alt = $u.altura; $t = [DateTimeOffset]::Parse($u.t, $Script:Inv)
    do { $n++; $alt += 210000; $t = $t.AddMinutes(210000 * 10) } while ($t -lt $Script:Agora)
    $antes = 50 / [math]::Pow(2, $n - 1); $depois = 50 / [math]::Pow(2, $n)
    return [pscustomobject]@{ d = $t.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv); e = 'BTC'; imp = 'High'; st = 'E'
        ev = 'Bitcoin halving (block {0}), approximate date: the block reward drops from {1} to {2} BTC' -f $alt.ToString('N0', $Script:Inv), $antes.ToString('0.#####', $Script:Inv), $depois.ToString('0.#####', $Script:Inv) }
}

# ============================================================================
# 3. MODELO DO SITE (HTML + CSS + JavaScript num único ficheiro)
# ============================================================================
function Get-Plantilla {
return @'
<!doctype html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Bluechip Board: Apple, NVIDIA, Alphabet, ETFs and Bitcoin</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Cdefs%3E%3ClinearGradient id='g' x1='0' y1='0' x2='1' y2='1'%3E%3Cstop offset='0' stop-color='%236FA3FF'/%3E%3Cstop offset='1' stop-color='%232A5BD7'/%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width='32' height='32' rx='8' fill='url(%23g)'/%3E%3Cg fill='%23fff' fill-opacity='.26'%3E%3Crect x='7' y='19' width='4' height='6' rx='1'/%3E%3Crect x='14' y='15.5' width='4' height='9.5' rx='1'/%3E%3Crect x='21' y='12' width='4' height='13' rx='1'/%3E%3C/g%3E%3Cpath d='M6.5 19.5 12 14l4.2 3.2 8.6-8.6M19.8 8.3h5.3v5.3' fill='none' stroke='%23fff' stroke-width='2.4' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;500;600;700&family=Newsreader:opsz,wght@6..72,400;6..72,500&display=swap" rel="stylesheet">
<style>
/* ---------- design tokens ---------- */
/* Camadas, da mais funda para a mais elevada: --surface-3 (calhas dos filtros) < --bg (página) < --band (cabeçalho e panorama)
   < --surface (cartões) < --surface-2 (cabeçalhos de tabela, painéis internos) < --sel (opção escolhida).
   Os neutros são carvão com um toque de verde; o verde forte fica reservado ao destaque. */
:root{--bg:#10172A;--band:#151E33;--surface:#1A243C;--surface-2:#222D48;--surface-3:#0C1322;--hover:#202B46;--sel:#2D3B5C;
--ink:#F0F4FA;--text:#CDD6E4;--muted:#99A6BB;--faint:#76849B;--rule:#2B3856;--rule-2:#3C4B6D;
--accent:#4F8CFF;--accent-ink:#061021;--focus:#7AA8FF;--bar:#3F74DB;
--aapl:#A08CFF;--nvda:#9BD66B;--googl:#4FC6EE;--spx:#E57AA6;--btc:#F7931A;--macro:#96A3B9;--ref:#E9A25E;
--eunk:#7FA7F2;--is3n:#45C4A6;--eunn:#F2858A;
--red:#F06A5F;--orange:#F09A4A;--yellow:#E2BE4C;--grey:#7F8CA2;--pos:#4FCB8F;--neg:#F06A5F;
--red-bg:rgba(240,106,95,.13);--orange-bg:rgba(240,154,74,.13);--yellow-bg:rgba(226,190,76,.13);--grey-bg:rgba(127,140,162,.14);--pos-bg:rgba(79,203,143,.13);--neg-bg:rgba(240,106,95,.13);--guilloche:#4F8CFF;
--glass:rgba(21,30,51,.92);--tip-bg:#283453;--tip-head:#94A1B6;--seg-shadow:inset 0 1px 0 rgba(255,255,255,.06),0 1px 2px rgba(0,0,0,.5),0 0 0 1px var(--rule-2);
--r-s:6px;--r-m:10px;--r-l:14px;
--sh-1:inset 0 1px 0 rgba(255,255,255,.035),0 1px 3px rgba(0,0,0,.45);--sh-2:inset 0 1px 0 rgba(255,255,255,.05),0 10px 26px rgba(0,0,0,.45);--sh-3:0 16px 40px rgba(0,0,0,.6);
--gut:32px;--max:2400px;--ease:cubic-bezier(.2,.7,.2,1);
--font-ui:"Hanken Grotesk","Segoe UI",Roboto,Helvetica,Arial,sans-serif;--font-read:"Newsreader",Georgia,"Times New Roman",serif;
color-scheme:dark;box-sizing:border-box}
*,*::before,*::after{box-sizing:inherit}
html{scroll-padding-top:calc(var(--topbar-h,64px) + 16px)}
body{margin:0;display:flex;flex-direction:column;min-height:100vh;background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:15px;line-height:1.55;font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased;
padding-bottom:env(safe-area-inset-bottom,0px)}
a{color:var(--ink);text-underline-offset:3px;text-decoration-color:var(--rule-2)}
a:hover{text-decoration-color:currentColor}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px;border-radius:4px}
h1,h2,h3,h4{color:var(--ink);margin:0}
button{font:inherit;color:inherit}
.i{width:16px;height:16px;flex:none;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;vertical-align:-3px}
.sr{position:absolute!important;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0}
.muted{color:var(--muted)}.small{font-size:.86rem}
.pos{color:var(--pos)}.neg{color:var(--neg)}
.l-red{--lc:var(--red);--lbg:var(--red-bg)}.l-orange{--lc:var(--orange);--lbg:var(--orange-bg)}.l-yellow{--lc:var(--yellow);--lbg:var(--yellow-bg)}.l-white{--lc:var(--grey);--lbg:var(--grey-bg)}

/* ---------- buttons, tags, controls ---------- */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;font-size:.84rem;font-weight:600;padding:7px 12px;min-height:34px;border-radius:8px;border:1px solid var(--rule-2);background:var(--surface);color:var(--ink);cursor:pointer;white-space:nowrap;
transition:background .15s,border-color .15s,box-shadow .15s,transform .15s}
.btn:hover{background:var(--surface-2);border-color:var(--faint)}
.btn:active{transform:translateY(1px)}
.btn .i{width:15px;height:15px;transition:transform .15s var(--ease)}
.btn:hover .i-go{transform:translateX(2px)}
.tag{display:inline-flex;align-items:center;gap:4px;font-size:.7rem;font-weight:650;line-height:1.5;padding:1px 8px;border-radius:999px;background:var(--surface-2);color:var(--muted);border:1px solid var(--rule);white-space:nowrap}
.tag .i{width:12px;height:12px}
.t-pos{color:var(--pos);background:var(--pos-bg);border-color:transparent}
.t-warn{color:var(--orange);background:var(--orange-bg);border-color:transparent}
.co{display:inline-flex;align-items:center;gap:6px;white-space:nowrap;font-size:.86rem;font-weight:600;color:var(--text)}
.co::before{content:"";width:8px;height:8px;border-radius:50%;background:var(--c);flex:none}
.co-chip{font-size:.72rem;font-weight:650;padding:1px 8px 1px 7px;border-radius:999px;background:var(--surface-2);background:color-mix(in srgb,var(--c) 12%,var(--surface));border:1px solid var(--rule);border-color:color-mix(in srgb,var(--c) 30%,var(--surface));margin-right:6px;vertical-align:1px;gap:5px}
.co-chip::before{width:6px;height:6px}
.f-label{display:inline-flex;align-items:center;gap:6px;font-size:.7rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--muted)}
.f-label .i{width:13px;height:13px;color:var(--faint)}
.field{display:flex;flex-direction:column;gap:6px;min-width:170px}
select,.search input{font:inherit;font-size:.9rem;color:var(--text);background:var(--surface);border:1px solid var(--rule-2);border-radius:8px;height:36px;padding:0 10px;width:100%;
transition:border-color .15s,box-shadow .15s}
select{appearance:none;-webkit-appearance:none;padding-right:32px;cursor:pointer;
background:var(--surface) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2394A1B6' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E") no-repeat right 10px center/14px}
select:hover,.search input:hover{border-color:var(--faint)}
select:focus-visible,.search input:focus-visible{outline:none;border-color:var(--focus);box-shadow:0 0 0 3px rgba(79,140,255,.25)}
.search{position:relative}.search .i{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--faint);pointer-events:none}
.search input{padding-left:34px}
.seg{display:inline-flex;gap:2px;padding:3px;background:var(--surface-3);border:1px solid var(--rule);border-radius:10px;max-width:100%;overflow-x:auto;scrollbar-width:none}
.seg::-webkit-scrollbar{display:none}
.seg button{display:inline-flex;align-items:center;gap:6px;border:0;background:transparent;color:var(--muted);font-size:.84rem;font-weight:600;padding:5px 12px;min-height:30px;border-radius:7px;cursor:pointer;white-space:nowrap;transition:background .15s,color .15s,box-shadow .15s}
.seg button:hover{color:var(--ink)}
.seg button[aria-pressed="true"]{background:var(--sel);color:var(--ink);box-shadow:var(--seg-shadow)}
.seg button .i{width:14px;height:14px;color:var(--lc,currentColor)}
.seg .sw-c{display:inline-block;width:8px;height:8px;border-radius:2px;background:var(--c)}
.switch{position:relative;display:inline-flex;align-items:center;gap:10px;font-size:.88rem;color:var(--text);cursor:pointer;min-height:36px;user-select:none}
.switch input{position:absolute;opacity:0;width:1px;height:1px}
.switch .sw{position:relative;width:34px;height:20px;border-radius:999px;background:var(--rule-2);flex:none;transition:background .2s}
.switch .sw::after{content:"";position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:var(--ink);box-shadow:0 1px 2px rgba(0,0,0,.4);transition:transform .2s var(--ease)}
.switch input:checked+.sw{background:var(--accent)}
.switch input:checked+.sw::after{transform:translateX(14px)}
.switch input:focus-visible+.sw{outline:2px solid var(--focus);outline-offset:2px}
.switch input:disabled~*{opacity:.5;cursor:not-allowed}

/* ---------- top bar ---------- */
.topbar{position:sticky;top:0;z-index:30;background:var(--glass);-webkit-backdrop-filter:saturate(1.4) blur(10px);backdrop-filter:saturate(1.4) blur(10px);border-bottom:1px solid var(--rule-2);padding-top:env(safe-area-inset-top,0px)}
.topbar-in{max-width:var(--max);margin:0 auto;display:grid;grid-template-columns:auto minmax(0,1fr) auto;align-items:center;gap:10px 24px;padding:10px var(--gut)}
.brand{display:flex;align-items:center;gap:12px;color:var(--ink);min-width:0;padding-right:20px;border-right:1px solid var(--rule)}
.brand-mark{width:36px;height:36px;border-radius:10px;display:block;flex:none;box-shadow:0 4px 14px rgba(79,140,255,.3)}
.brand-mark .logo{display:block;width:100%;height:100%}
.brand-mark .i{width:20px;height:20px;stroke-width:2.2}
.brand-txt{display:flex;flex-direction:column;gap:3px;line-height:1.1;min-width:0}
.brand-txt b{font-family:var(--font-read);font-weight:500;font-size:1.22rem;letter-spacing:-.01em;white-space:nowrap}
.brand-txt small{display:flex;align-items:center;gap:6px;font-size:.66rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--faint);white-space:nowrap}
.brand-txt small i{width:3px;height:3px;border-radius:50%;background:var(--faint)}
.cofilter{justify-self:center;display:inline-flex;gap:2px;padding:3px;background:var(--surface-3);border:1px solid var(--rule);border-radius:12px;max-width:100%;overflow-x:auto;scrollbar-width:none}
.cofilter::-webkit-scrollbar{display:none}
.chip{display:inline-flex;align-items:center;gap:8px;padding:4px 12px 4px 4px;min-height:36px;border:0;border-radius:9px;background:transparent;color:var(--muted);font-size:.88rem;font-weight:600;cursor:pointer;white-space:nowrap;
transition:background .15s,color .15s,box-shadow .15s}
.chip .ci{width:28px;height:28px;border-radius:7px;display:grid;place-items:center;color:var(--c);background:var(--surface-2);background:color-mix(in srgb,var(--c) 14%,var(--surface));transition:background .15s,color .15s}
.chip .ci .i{width:15px;height:15px}
.chip .tkr{font-size:.68rem;font-weight:700;letter-spacing:.04em;color:var(--faint)}
.chip:hover{color:var(--ink);background:rgba(255,255,255,.04)}
.chip[aria-pressed="true"]{background:var(--sel);color:var(--ink);box-shadow:var(--seg-shadow)}
.chip[aria-pressed="true"] .ci{background:var(--c);color:var(--accent-ink)}

/* ---------- hero / market snapshot ---------- */
.hero{position:sticky;top:calc(var(--topbar-h,64px) + 20px);max-height:calc(100vh - var(--topbar-h,64px) - 40px);overflow:hidden auto;scrollbar-width:thin;scrollbar-color:var(--rule-2) transparent;background:var(--band);border:1px solid var(--rule);border-radius:var(--r-l);box-shadow:var(--sh-1)}
.guilloche{position:absolute;inset:0;width:100%;height:100%;opacity:.035;pointer-events:none}
.hero-in{position:relative;padding:20px 16px 16px}
.hero-head{display:flex;flex-direction:column;gap:12px;margin-bottom:16px;padding:0 2px}
.eyebrow{display:flex;align-items:center;gap:7px;margin:0 0 6px;font-size:.72rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}
.eyebrow .i{width:14px;height:14px}
.hero h1{font-family:var(--font-read);font-weight:500;font-size:1.45rem;line-height:1.15;letter-spacing:-.015em}
.hero-meta{display:flex;flex-direction:column;gap:5px;font-size:.8rem;color:var(--muted)}
.hero-meta span{display:inline-flex;align-items:center;gap:6px}.hero-meta .i{width:14px;height:14px;color:var(--faint)}
.hero-meta b{color:var(--ink);font-weight:650}
.tickers{display:grid;grid-template-columns:minmax(0,1fr);gap:10px}
.tk{position:relative;display:grid;grid-template-columns:auto minmax(0,1fr) auto;grid-template-areas:"badge id price" "badge eur chg" "spark spark spark" "rows rows rows" "range range range";row-gap:2px;
column-gap:10px;align-items:center;background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-l);padding:16px 16px 14px;box-shadow:var(--sh-1);overflow:hidden;
transition:opacity .25s,box-shadow .2s,transform .2s}
.tk::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--c)}
.tk:hover{box-shadow:var(--sh-2)}
.tk.dim{opacity:.38}
.tk-badge{grid-area:badge;width:34px;height:34px;border-radius:9px;display:grid;place-items:center;color:var(--c);background:var(--surface-2);background:color-mix(in srgb,var(--c) 14%,var(--surface))}
.tk-badge .i{width:17px;height:17px}
.tk-id{grid-area:id;display:flex;flex-direction:column;min-width:0;line-height:1.25}
.tk-name{font-weight:700;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tk-sym{font-size:.72rem;color:var(--muted);letter-spacing:.04em;font-weight:600}
.asof{letter-spacing:0;font-weight:500}.asof.old{color:var(--orange);font-weight:700}
.td-sub{display:block;font-size:.72rem;color:var(--faint);font-weight:500}.td-sub.old{color:var(--orange);font-weight:700}
.mk-c.old{color:var(--orange);font-weight:700}.ev-mark{cursor:help}.fund-co{margin-bottom:22px}.fund-co .sec-h .muted{font-weight:500;text-transform:none;letter-spacing:0;margin-left:6px}.fund-co sup{color:var(--faint);margin-left:1px}.mv-news{margin:0;padding-left:1.1em}.mv-news li{margin:2px 0}.mv-news a{color:var(--ink)}
.row-sold td{opacity:.55}.tag.t-sell{color:var(--neg);background:var(--neg-bg);border-color:transparent}
#bkState .callout{margin-top:12px}
.chg{grid-area:chg;display:inline-flex;align-items:center;gap:4px;font-size:.8rem;font-weight:700;padding:3px 8px;border-radius:999px;white-space:nowrap;justify-self:end}
.chg.pos{background:var(--pos-bg)}.chg.neg{background:var(--neg-bg)}
.chg .i{width:14px;height:14px}
.tk-price{grid-area:price;font-size:1.9rem;font-weight:600;color:var(--ink);letter-spacing:-.025em;line-height:1.1;margin:14px 0 2px;white-space:nowrap}
.tk-price .cur{font-size:1rem;font-weight:500;color:var(--muted);margin-right:2px;vertical-align:.5em}
.tk-eur{grid-area:eur;font-size:.78rem;color:var(--muted)}
.tk-spark{grid-area:spark;height:40px;margin:22px 0 8px;position:relative}
.tk-spark svg{display:block;width:100%;height:100%}
.tk-spark .spark-l{position:absolute;right:0;top:-16px;font-size:.62rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--faint)}
.tk-rows{grid-area:rows;display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:0;padding-top:10px;border-top:1px solid var(--rule)}
.tk-rows dt{font-size:.7rem;color:var(--muted);font-weight:600}
.tk-rows dd{margin:0;font-weight:700;font-size:.92rem;color:var(--ink)}
.tk-rows dd.pos{color:var(--pos)}.tk-rows dd.neg{color:var(--neg)}
.range{grid-area:range;margin-top:10px}
.range-bar{position:relative;height:4px;border-radius:2px;background:var(--surface-3)}
.range-fill{position:absolute;left:0;top:0;bottom:0;border-radius:2px;background:var(--c);opacity:.35}
.range-mk{position:absolute;top:50%;width:10px;height:10px;border-radius:50%;background:var(--surface);border:2px solid var(--c);transform:translate(-50%,-50%)}
.range-lbl{display:flex;justify-content:space-between;gap:6px;margin-top:6px;font-size:.68rem;color:var(--faint)}
.tk-empty{grid-column:1/-1;margin:12px 0 0;font-size:.86rem;color:var(--muted)}

.tk .tk-price{margin:0;font-size:1.45rem;text-align:right;justify-self:end}
.tk .tk-price .cur{font-size:.85rem}
.tk .tk-eur{font-size:.72rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tk .tk-spark{height:32px;margin:20px 0 6px}
.tk .tk-rows{padding-top:8px}
/* ordem do panorama: o S&P 500 primeiro (tk-lead), a Bitcoin no fim (tk-wide, já é o último) */
.tk-lead{order:-1}
/* ---------- layout & navigation ---------- */
.layout{max-width:var(--max);margin:0 auto;display:grid;grid-template-columns:clamp(400px,22vw,540px) minmax(0,1fr);align-items:start;gap:32px;width:100%;flex:1 0 auto;padding:20px var(--gut) 72px}
/* ecrãs largos: panorama mais largo e em grelha (S&P 500 a toda a largura; as três ações e os três ETF regionais em três
   colunas; Bitcoin a toda a largura). Os cartões estreitos põem o preço por baixo do nome para caberem */
@media (min-width:1600px){
 /* margens laterais finas; a largura do panorama (--side) cresce por escalões (ver também 2400px e 3000px) */
 :root{--gut:32px;--max:3200px;--side:clamp(560px,29vw,820px)}
 .layout{grid-template-columns:var(--side) minmax(0,1fr)}
 .tickers{grid-template-columns:repeat(6,minmax(0,1fr))}
 .tk{grid-column:span 2}
 .tk-lead,.tk-wide{grid-column:1/-1}
 .tk:not(.tk-lead):not(.tk-wide){grid-template-columns:auto minmax(0,1fr) auto;grid-template-areas:"badge id id" "price price chg" "eur eur eur" "spark spark spark" "rows rows rows" "range range range";column-gap:8px;padding:14px 12px 12px}
 .tk:not(.tk-lead):not(.tk-wide) .tk-price{justify-self:start;text-align:left;font-size:1.3rem;margin-top:12px}
 .tk:not(.tk-lead):not(.tk-wide) .chg{margin-top:12px;padding:2px 6px}
 .tk:not(.tk-lead):not(.tk-wide) .tk-name{white-space:normal;overflow:visible}
 .tk:not(.tk-lead):not(.tk-wide) .tk-sym{white-space:normal}
 .tk:not(.tk-lead):not(.tk-wide) .tk-eur.nota{display:none}
 .tk:not(.tk-lead):not(.tk-wide) .tk-spark{height:26px;margin:16px 0 6px}
 .tk:not(.tk-lead):not(.tk-wide) .tk-rows{gap:6px}
 .tk:not(.tk-lead):not(.tk-wide) .tk-rows dt{font-size:.66rem}
 .tk:not(.tk-lead):not(.tk-wide) .lg{display:none}
 .tk-lead .tk-rows dt,.tk-wide .tk-rows dt{white-space:nowrap}
 .tk:not(.tk-lead):not(.tk-wide) .range-lbl span:nth-child(2){display:none}
 /* os dois cartões a toda a largura ficam na horizontal (mais baixos): nome, gráfico e preço numa linha, os números por baixo */
 .tk-lead,.tk-wide{grid-template-columns:auto minmax(150px,.8fr) minmax(0,1.6fr) auto;grid-template-areas:"badge id spark price" "badge eur spark chg" "rows rows range range";column-gap:14px;padding:14px 16px 12px}
 .tk-lead .tk-spark,.tk-wide .tk-spark{margin:0;height:44px}
 .tk-lead .tk-spark .spark-l,.tk-wide .tk-spark .spark-l{display:none}
 .tk-lead .tk-rows,.tk-wide .tk-rows{margin-top:10px}
 .tk-lead .range,.tk-wide .range{margin-top:10px;padding-top:10px;border-top:1px solid var(--rule)}
 .hero-meta{flex-direction:row;flex-wrap:wrap;gap:4px 18px}
 /* o panorama ocupa a altura do ecrã: as linhas de cartões repartem o espaço por igual e os gráficos crescem com elas
    (num ecrã mais baixo do que o conteúdo, ficam no tamanho mínimo e o panorama tem scroll, como antes) */
 .hero{height:calc(100vh - var(--topbar-h,64px) - 40px)}
 .hero-in{display:grid;grid-template-rows:auto minmax(0,1fr);min-height:100%}
 /* S&P 500, ações, ETF, Bitcoin, EUR/USD: as filas dos cartões estreitos levam mais espaço, mas os largos também crescem,
    para os gráficos não ficarem desproporcionados (muito altos e estreitos) num ecrã alto */
 .tickers{grid-template-rows:minmax(min-content,1fr) minmax(min-content,1.6fr) minmax(min-content,1.6fr) minmax(min-content,1fr) minmax(min-content,1fr)}
 .tk-lead,.tk-wide{grid-template-rows:1fr 1fr auto}
 .tk-lead .tk-id,.tk-wide .tk-id,.tk-lead .tk-price,.tk-wide .tk-price{align-self:end}
 .tk-lead .tk-eur,.tk-wide .tk-eur,.tk-lead .chg,.tk-wide .chg{align-self:start}
 .tk:not(.tk-lead):not(.tk-wide){grid-template-rows:auto auto auto minmax(26px,1fr) auto auto}
 .tk:not(.tk-lead):not(.tk-wide) .tk-spark{height:auto;min-height:26px;align-self:stretch}
 .tk:not(.tk-lead):not(.tk-wide) .tk-name{font-size:1.02rem}
 .tk:not(.tk-lead):not(.tk-wide) .tk-price{font-size:1.5rem}
 .tk-lead .tk-spark,.tk-wide .tk-spark{height:auto;min-height:48px;align-self:stretch}
 /* o gráfico preenche o espaço do cartão em vez de o definir (senão a proporção do SVG esticava os cartões largos) */
 .tk-spark svg{position:absolute;inset:0}}
@media (min-width:1501px) and (max-width:2100px){.chip .tkr{display:none}}
/* a partir de 2400px a coluna da marca no cabeçalho tem a largura do panorama: o "All" dos filtros fica alinhado com o
   início da coluna das secções (abaixo disso não há espaço para os filtros todos e o cabeçalho fica centrado, como antes) */
@media (min-width:2400px){
 :root{--side:clamp(860px,36vw,1100px)}
 .topbar-in{grid-template-columns:var(--side) minmax(0,1fr) auto;column-gap:32px}
 .cofilter{justify-self:start}}
@media (min-width:2400px) and (max-width:2999px){.chip .tkr{display:none}}
.content{min-width:0}
.rail{position:sticky;top:var(--topbar-h,64px);z-index:20;display:flex;gap:4px;overflow-x:auto;scrollbar-width:none;margin-bottom:24px;padding:6px;background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-l);box-shadow:var(--sh-1)}
.rail-label{margin:0 0 6px 10px;font-size:.68rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--faint)}
.rail a{position:relative;flex:1 0 auto;justify-content:center;white-space:nowrap;display:flex;align-items:center;gap:9px;padding:8px 14px;min-height:40px;border-radius:8px;color:var(--text);text-decoration:none;font-size:.92rem;font-weight:500;transition:background .15s,color .15s}
.rail a .i{color:var(--faint);transition:color .15s}
.rail a:hover{background:var(--hover);color:var(--ink)}
.rail a:hover .i{color:var(--ink)}
.rail a[aria-current="page"]{background:var(--sel);color:var(--ink);font-weight:700;box-shadow:inset 0 0 0 1px var(--rule-2)}
.rail a[aria-current="page"] .i{color:var(--ink)}
.rail a[aria-current="page"]::before{content:"";position:absolute;left:14px;right:14px;bottom:0;height:2px;border-radius:2px 2px 0 0;background:var(--accent)}
.rail::-webkit-scrollbar{display:none}
.cnt{margin-left:2px;min-width:22px;padding:1px 7px;border-radius:999px;font-size:.7rem;font-weight:700;text-align:center;background:var(--surface-2);color:var(--muted)}
.cnt:empty{display:none}
.cnt.warn{background:var(--orange-bg);color:var(--orange)}.cnt.err{background:var(--red-bg);color:var(--red)}
main{min-width:0;outline:none}
[hidden]{display:none!important}
.panel{animation:rise .28s var(--ease)}
@keyframes rise{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.panel-head{display:flex;align-items:center;gap:12px;padding-bottom:14px;border-bottom:1px solid var(--rule)}
.ph-ic{width:38px;height:38px;border-radius:10px;display:grid;place-items:center;background:var(--surface);border:1px solid var(--rule);color:var(--ink);box-shadow:var(--sh-1);flex:none}
.ph-ic .i{width:18px;height:18px}
h2{font-family:var(--font-read);font-weight:500;font-size:clamp(1.6rem,2.6vw,2.1rem);line-height:1.15;letter-spacing:-.01em}
.lead{font-family:var(--font-read);font-size:1.12rem;line-height:1.55;color:var(--text);max-width:72ch;margin:16px 0 22px}
.sec-h{display:flex;align-items:center;gap:8px;margin:38px 0 14px;font-size:.74rem;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:var(--muted)}
.sec-h .i{width:15px;height:15px;color:var(--faint)}
.sec-h::after{content:"";flex:1;height:1px;background:var(--rule);order:1;margin-left:6px}.sec-h .btn{order:2;text-transform:none;letter-spacing:0;font-size:.8rem;min-height:30px;padding:4px 10px}
.note{display:flex;gap:6px;align-items:flex-start;margin:10px 0 0;font-size:.8rem;color:var(--muted)}
.note .i{width:14px;height:14px;margin-top:2px;color:var(--faint)}

/* ---------- cards, lists, kpis ---------- */
.card,.list,.fig,.kpis,.filters,.tbl-wrap{background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-l);box-shadow:var(--sh-1)}
.list{overflow:hidden}
.empty{padding:22px 18px;color:var(--muted);font-size:.9rem;display:flex;align-items:center;gap:8px}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));overflow:hidden}
.kpi{padding:16px 18px;border-left:1px solid var(--rule);min-width:0}
.kpi:first-child{border-left:0}
.kpi .k{display:flex;align-items:center;gap:6px;font-size:.7rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--muted)}
.kpi .k .i{width:14px;height:14px;color:var(--faint)}
.kpi .v{margin-top:6px;font-size:1.7rem;font-weight:600;line-height:1.15;letter-spacing:-.02em;color:var(--ink)}
.kpi .v.pos{color:var(--pos)}.kpi .v.neg{color:var(--neg)}
.kpi .s{margin-top:2px;font-size:.78rem;color:var(--muted)}
.kpi.k-red .k .i{color:var(--red)}.kpi.k-orange .k .i{color:var(--orange)}
.sev-ic{width:30px;height:30px;border-radius:8px;display:grid;place-items:center;flex:none;color:var(--lc);background:var(--lbg)}
.sev-ic .i{width:16px;height:16px}

/* alerts */
.alert{position:relative;display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:14px;align-items:center;padding:14px 16px 14px 18px;border-bottom:1px solid var(--rule);transition:background .15s}
.alert:last-child{border-bottom:0}
.alert::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--lc)}
.alert:hover{background:var(--hover)}
.alert-top{display:flex;flex-wrap:wrap;align-items:center;gap:6px 10px;margin-bottom:3px}
.sev-txt{font-size:.68rem;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--lc)}
.alert-msg{margin:0;color:var(--text);font-size:.93rem}
.alert-ctx{display:flex;gap:6px;align-items:flex-start;margin:6px 0 0;font-size:.8rem;color:var(--muted);line-height:1.45}
.alert-ctx .i{width:13px;height:13px;margin-top:2px;color:var(--faint)}
.pol-form{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px 16px;margin:4px 0 12px}
.pol-form .field{min-width:0}
.pol-form textarea.in{height:auto;min-height:62px;padding:8px 10px;line-height:1.4;resize:vertical}
.pol-form .btn{justify-self:start}
@media (max-width:720px){.pol-form{grid-template-columns:minmax(0,1fr)}}

/* news */
.news{position:relative;display:grid;grid-template-columns:auto minmax(0,1fr);gap:12px;padding:14px 16px;border-bottom:1px solid var(--rule);transition:background .15s}
.news:last-child{border-bottom:0}
.news:hover{background:var(--hover)}
.news .sev-ic{width:28px;height:28px;margin-top:1px}
.news .sev-ic .i{width:15px;height:15px}
.news-line{line-height:1.45}
.news-t{font-weight:650;color:var(--ink);text-decoration:none;font-size:.97rem}
.news-t:hover{text-decoration:underline;text-decoration-color:var(--rule-2)}
.news-t .ext{width:13px;height:13px;margin-left:5px;color:var(--faint);vertical-align:-1px;transition:color .15s,transform .15s}
.news-t:hover .ext{color:var(--ink);transform:translate(1px,-1px)}
.meta{display:flex;flex-wrap:wrap;align-items:center;gap:4px 14px;margin-top:6px;font-size:.79rem;color:var(--muted)}
.meta>span{display:inline-flex;align-items:center;gap:5px}
.meta .i{width:13px;height:13px;color:var(--faint)}
.m-sev{font-size:.68rem;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--lc)}
.m-score{font-weight:700;color:var(--text)}
.dups{margin-top:8px}
.dups summary{display:inline-flex;align-items:center;gap:6px;max-width:100%;list-style:none;cursor:pointer;font-size:.78rem;font-weight:650;color:var(--muted);padding:3px 10px;border-radius:999px;background:var(--surface-2);border:1px solid var(--rule);transition:background .15s,color .15s}
.dups summary::-webkit-details-marker{display:none}
.dups summary:hover{color:var(--ink);background:var(--hover)}
.dups summary .i{width:13px;height:13px;flex:none}
.dups-src{font-weight:500;color:var(--faint);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
.dups summary .chev{transition:transform .2s var(--ease)}.dups[open] summary .chev{transform:rotate(180deg)}
.dups ul{list-style:none;margin:8px 0 2px;padding:0 0 0 12px;border-left:2px solid var(--rule);display:grid;gap:7px}
.dups li{display:flex;flex-direction:column;font-size:.84rem;line-height:1.4}
.dups li a{color:var(--text);text-decoration:none;font-weight:550}.dups li a:hover{text-decoration:underline}
.dups li a .ext{width:12px;height:12px;margin-left:4px;color:var(--faint);vertical-align:-1px}
.dups li>span{font-size:.74rem;color:var(--faint)}
.s-pos .i{color:var(--pos)}.s-neg .i{color:var(--neg)}
.filters{padding:14px 16px;margin-bottom:12px}
.f-row{display:flex;flex-wrap:wrap;align-items:flex-end;gap:12px 14px}
.f-row+.f-row{margin-top:14px;padding-top:14px;border-top:1px dashed var(--rule)}
.f-search{flex:1 1 260px}
.list-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:0 0 8px}
.list-bar .note{margin:0}
.more{margin:12px 0 0;width:100%}

/* figures & charts */
.fig{padding:18px 20px 14px;margin:14px 0}
.fig-lg{padding:20px 22px 16px}
.fig-head{display:flex;justify-content:space-between;align-items:flex-start;gap:8px 16px;flex-wrap:wrap;margin-bottom:10px}
.fig-title{margin:0;font-weight:700;color:var(--ink);font-size:.98rem}
.fig-sub{margin:2px 0 0;font-size:.8rem;color:var(--muted)}
.fig svg:not(.i){display:block;width:100%;height:auto;overflow:visible}
.mk{text-align:right}
.mk-v{display:block;font-size:1.35rem;font-weight:600;color:var(--ink);letter-spacing:-.02em;line-height:1.1}
.mk-c{font-size:.76rem;font-weight:650}
.legend{display:flex;flex-wrap:wrap;gap:6px 18px;margin:12px 0 0;font-size:.8rem;color:var(--muted)}
.legend span{display:inline-flex;align-items:center;gap:7px}
.legend i{width:10px;height:10px;border-radius:3px;background:var(--c);display:inline-block}
.legend.lines i{width:16px;height:3px;border-radius:2px}
.legend b{color:var(--ink);font-weight:700}
.svg-label{font-family:var(--font-ui);font-size:11px;fill:var(--faint)}
.svg-val{font-family:var(--font-ui);font-size:11.5px;fill:var(--text);font-weight:650}
.svg-end{font-family:var(--font-ui);font-size:11.5px;font-weight:700}
.svg-grid{stroke:var(--rule);stroke-width:1}
.svg-axis{stroke:var(--rule-2);stroke-width:1}
.svg-line{fill:none;stroke-width:2;stroke-linejoin:round;stroke-linecap:round}
.fig svg rect[data-tip]{transition:opacity .15s}.fig svg rect[data-tip]:hover{opacity:.8}
#tip{position:fixed;z-index:60;pointer-events:none;background:var(--tip-bg);color:var(--ink);border:1px solid var(--rule-2);padding:7px 10px;border-radius:8px;font-size:.8rem;max-width:320px;line-height:1.45;white-space:pre-line;box-shadow:var(--sh-3)}
#tip.rich{white-space:normal;padding:9px 12px}
.tt-h{margin-bottom:5px;font-size:.72rem;font-weight:700;letter-spacing:.04em;color:var(--tip-head)}
.tt-r{display:flex;align-items:center;gap:8px;white-space:nowrap}
.tt-r i{width:8px;height:8px;border-radius:50%;flex:none}
.tt-r b{margin-left:auto;padding-left:16px;font-weight:700}
.toolbar{display:flex;flex-wrap:wrap;align-items:flex-end;gap:12px 24px;margin:0 0 4px}
.tb-group{display:flex;flex-direction:column;gap:6px;max-width:100%}
.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
.grid3 .fig{margin:0}
.grid3 .fig-head{flex-wrap:nowrap}.grid3 .fig-head>div:first-child{min-width:0}.mk{flex:none}

/* tables */
.tbl-wrap{overflow:auto;margin:0 0 6px}
.tbl-wrap.tall{max-height:72vh}
table{border-collapse:separate;border-spacing:0;width:100%;min-width:640px;font-size:.88rem}
th,td{padding:11px 14px;border-bottom:1px solid var(--rule);text-align:left;vertical-align:middle}
thead th{position:sticky;top:0;z-index:2;background:var(--surface-2);color:var(--muted);font-size:.68rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;white-space:nowrap;border-bottom:1px solid var(--rule-2)}
th:first-child,td:first-child{position:sticky;left:0;z-index:1;background:var(--surface)}
thead th:first-child{z-index:3;background:var(--surface-2)}
tbody tr:last-child td{border-bottom:0}
tbody td{transition:background .12s}
tbody tr:hover td{background:var(--hover)}
td.num,th.num{text-align:right;white-space:nowrap}
td.num{font-weight:600;color:var(--ink)}
td.num.pos{color:var(--pos)}td.num.neg{color:var(--neg)}
td.dim{color:var(--muted);font-weight:500}
.below{display:inline-block;margin-left:6px;font-size:.66rem;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:var(--orange);background:var(--orange-bg);padding:0 6px;border-radius:999px;vertical-align:1px}
.st{display:inline-flex;align-items:center;gap:6px;font-weight:650;font-size:.82rem}
.st .i{width:15px;height:15px}
.st-ok{color:var(--pos)}.st-warn{color:var(--orange)}.st-err{color:var(--red)}.st-skip{color:var(--muted)}
.src-link{display:inline-flex;align-items:center;gap:5px;font-weight:600;text-decoration:none}
.src-link .i{width:12px;height:12px;color:var(--faint)}
.src-link:hover{text-decoration:underline}
.detail{font-size:.8rem;color:var(--muted);min-width:220px}

/* calendar */
.cal-date{display:flex;align-items:center;gap:10px}
.cal-d{min-width:2ch;font-size:1.5rem;font-weight:600;line-height:1;color:var(--ink);text-align:right;letter-spacing:-.02em}
.cal-m{font-size:.68rem;font-weight:700;line-height:1.3;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
.cal-ev{font-weight:600;color:var(--ink);margin-bottom:4px}
.imp{display:inline-flex;align-items:center;gap:8px;font-size:.84rem;font-weight:600;white-space:nowrap}
.imp .bars{display:inline-flex;align-items:flex-end;gap:2px}
.imp .bars i{width:4px;border-radius:1px;background:var(--rule-2)}
.imp .bars i:nth-child(1){height:7px}.imp .bars i:nth-child(2){height:10px}.imp .bars i:nth-child(3){height:13px}
.imp[data-l="3"] .bars i.on{background:var(--red)}.imp[data-l="2"] .bars i.on{background:var(--orange)}.imp[data-l="1"] .bars i.on{background:var(--yellow)}
.when{display:inline-flex;align-items:center;gap:5px;font-size:.8rem;font-weight:650;padding:2px 9px;border-radius:999px;background:var(--surface-2);color:var(--muted);white-space:nowrap}
.when .i{width:13px;height:13px}
.when.soon{background:var(--orange-bg);color:var(--orange)}.when.today{background:var(--red-bg);color:var(--red)}
tr.past td>*{opacity:.55}

/* ETF */
.etf-total{text-align:right}
.etf-total .big{font-size:2rem;font-weight:600;color:var(--ink);letter-spacing:-.02em;line-height:1}
.etf-total .s{font-size:.76rem;color:var(--muted)}
.dist{display:flex;gap:2px;height:38px;border-radius:9px;overflow:hidden;background:var(--surface-3)}
.dist>div{min-width:0;transition:filter .15s}.dist>div:hover{filter:brightness(1.1)}
.etf-legend{list-style:none;margin:16px 0 0;padding:0;display:grid;gap:2px}
.etf-legend li{display:grid;grid-template-columns:12px minmax(90px,150px) minmax(0,1fr) 70px;align-items:center;gap:12px;padding:7px 0;border-bottom:1px solid var(--rule);font-size:.88rem}
.etf-legend li:last-child{border-bottom:0}
.etf-legend .dot{width:10px;height:10px;border-radius:3px;background:var(--c)}
.etf-legend .nm{font-weight:600;color:var(--ink)}
.etf-legend .bar{height:6px;border-radius:3px;background:var(--surface-2);overflow:hidden}
.etf-legend .bar i{display:block;height:100%;background:var(--c);border-radius:3px}
.etf-legend b{text-align:right;color:var(--ink)}

/* sources & method */
.method{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.method li{background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-m);padding:14px 16px;font-size:.88rem;color:var(--text)}
.method li b{display:flex;align-items:center;gap:7px;color:var(--ink);margin-bottom:4px}
.method li b .i{width:15px;height:15px;color:var(--faint)}
.lvl-in{display:inline-flex;align-items:center;gap:4px;font-weight:650;color:var(--lc)}
.lvl-in .i{width:13px;height:13px}
.callout{display:flex;gap:12px;align-items:flex-start;margin:16px 0;padding:13px 16px;max-width:84ch;font-size:.88rem;background:var(--surface);border:1px solid var(--rule);border-left:3px solid var(--ref);border-radius:var(--r-m)}
.callout .i{width:18px;height:18px;margin-top:1px;color:var(--ref)}
.callout.warn{border-left-color:var(--orange);background:var(--orange-bg)}.callout.warn .i{color:var(--orange)}
footer{background:var(--band);border-top:1px solid var(--rule-2);color:var(--muted);font-size:.8rem}
.mkts{display:flex;gap:8px}
.mkt{display:flex;flex-direction:column;gap:2px;padding:5px 12px;border-radius:10px;background:var(--surface-3);border:1px solid var(--rule);line-height:1.2;white-space:nowrap;cursor:default}
.mkt-top{display:flex;align-items:center;gap:7px;font-size:.74rem}
.mkt-top b{color:var(--ink);font-weight:700}
.mkt-dot{width:7px;height:7px;border-radius:50%;background:var(--grey);flex:none}
.mkt.open .mkt-dot{background:var(--pos);box-shadow:0 0 0 3px var(--pos-bg)}
.mkt-st{margin-left:auto;padding-left:8px;font-size:.62rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#CC4A40}
.mkt.open .mkt-st{color:#2E9E68}
.mkt-when{font-size:.75rem;color:var(--text)}
.mkt-when small{color:var(--faint);font-size:.68rem;margin-left:4px}

/* Bitcoin */
.btc-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-top:14px}
.btc-grid .fig{margin:0;display:flex;flex-direction:column}
.big-v{display:flex;align-items:baseline;gap:10px;margin:2px 0 12px}
.big-v b{font-size:2.3rem;font-weight:600;line-height:1;letter-spacing:-.02em;color:var(--ink)}
.big-v span{font-size:.9rem;font-weight:700}
.fng-bar{position:relative;height:8px;border-radius:4px;background:linear-gradient(90deg,var(--red),var(--orange) 30%,var(--yellow) 50%,#9BD66B 70%,var(--pos))}
.fng-bar i{position:absolute;top:50%;width:14px;height:14px;border-radius:50%;background:var(--surface);border:3px solid var(--ink);transform:translate(-50%,-50%);box-shadow:0 1px 3px rgba(0,0,0,.5)}
.fng-lbl{display:flex;justify-content:space-between;margin-top:6px;font-size:.68rem;color:var(--faint)}
.prog{height:8px;border-radius:4px;background:var(--surface-3);overflow:hidden;margin:4px 0 6px}
.prog i{display:block;height:100%;border-radius:4px;background:var(--btc)}
.kv{list-style:none;margin:10px 0 0;padding:0;display:grid}
.kv li{display:flex;justify-content:space-between;align-items:baseline;gap:12px;padding:8px 0;border-top:1px solid var(--rule);font-size:.86rem;color:var(--muted)}
.kv li b{color:var(--ink);font-weight:650;text-align:right}
#ch-fng{margin-top:auto;padding-top:8px}

/* formulários, carteira, simuladores, compras de Bitcoin */
.in{font:inherit;font-size:.9rem;color:var(--ink);background:var(--surface-2);border:1px solid var(--rule-2);border-radius:8px;height:36px;padding:0 10px;width:100%;color-scheme:dark;transition:border-color .15s,box-shadow .15s}
.in:hover{border-color:var(--faint)}
.in:focus-visible{outline:none;border-color:var(--focus);box-shadow:0 0 0 3px rgba(79,140,255,.25)}
.in-num{width:140px;text-align:right}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin:14px 0}
.dd-grid{align-items:start}
.expo-grid{margin:12px 0}
.expo-grid .fig{margin:0}
.dd-grid .tbl-wrap{margin-top:0}
details.roll>summary{display:flex;flex-direction:column;gap:3px;cursor:pointer;list-style:none}
details.roll>summary::-webkit-details-marker{display:none}
details.roll>summary .fig-title::before{content:"▸ ";color:var(--faint)}
details.roll[open]>summary .fig-title::before{content:"▾ "}
details.roll[open]>summary{margin-bottom:12px}
details.roll .toolbar{margin:14px 0 4px}
@media (max-width:1400px){.dd-grid{grid-template-columns:minmax(0,1fr)}}
.grid2 .fig{margin:0}
.sim-form,.lot-form{display:flex;flex-wrap:wrap;align-items:flex-end;gap:12px 14px;margin:4px 0 14px}
.sim-form .field{flex:0 1 210px;min-width:160px}
.lot-form .field{flex:1 1 180px;min-width:150px}
.tgt-in{width:92px;max-width:100%;height:32px;text-align:right}
.tgt-sum.bad{color:var(--neg);font-weight:700}
.sim-kpis{margin-bottom:12px;background:var(--surface-2);box-shadow:none}
.btn-pri{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.btn-pri:hover{background:#6A9DFF;border-color:#6A9DFF}
.btn-ic{min-height:30px;padding:4px 8px;color:var(--muted)}
.btn-ic:hover{color:var(--red)}
.card{padding:16px 18px;margin:14px 0}
.card .kpis{margin:4px 0 14px;background:var(--surface-2);box-shadow:none}
.card .kpis:empty{display:none}
.form-msg{min-height:1.3em;margin:0 0 10px;font-size:.84rem;color:var(--muted)}
.lot-st{display:flex;flex-direction:column;gap:6px;min-width:240px}
.prog.sm{height:4px;margin:0}
.pf-src{font-size:.72rem;color:var(--faint);margin-top:3px;white-space:normal}
.pf-src:empty{display:none}
tfoot th,tfoot td{padding:11px 14px;border-top:1px solid var(--rule-2);font-weight:700;color:var(--ink);background:var(--surface-2);text-align:left}
tfoot th:first-child{position:sticky;left:0;z-index:1;background:var(--surface-2)}
tfoot td.num{text-align:right}
td small{color:var(--muted);font-weight:500;margin-left:4px}
.tools{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:18px 0 6px}
.tools label.btn{cursor:pointer}
.foot{max-width:var(--max);margin:0 auto;padding:22px var(--gut) 28px;display:flex;flex-wrap:wrap;justify-content:center;align-items:center;gap:10px 28px;text-align:center}
.foot-brand{display:flex;align-items:center;gap:10px;color:var(--ink);font-weight:650}
.foot-brand .brand-mark{width:26px;height:26px;border-radius:7px;box-shadow:none}
.foot-meta{display:flex;flex-wrap:wrap;justify-content:center;gap:4px 18px}
.foot-meta span{display:inline-flex;align-items:center;gap:6px}.foot-meta .i{width:13px;height:13px;color:var(--faint)}
.foot-meta button{font:inherit;color:var(--text);background:none;border:0;padding:0;cursor:pointer;text-decoration:underline;text-decoration-color:var(--rule-2);text-underline-offset:3px}
.foot-meta button:hover{color:var(--ink);text-decoration-color:currentColor}

/* ---------- responsive ---------- */
@media (min-width:2200px){body{font-size:16px}}
@media (min-width:3000px){:root{--max:4400px;--gut:44px;--side:clamp(1100px,40vw,1700px)}body{font-size:17px}}
@media (min-width:1561px) and (max-width:2399px){.rail{gap:2px}.rail a{padding:8px 10px;gap:7px;font-size:.88rem}}
@media (max-width:1560px){.layout{grid-template-columns:380px minmax(0,1fr);gap:24px}.rail{gap:2px}.rail a{padding:8px 10px;gap:7px;font-size:.88rem}}
@media (max-width:1400px){.layout{grid-template-columns:320px minmax(0,1fr)}.rail a{padding:8px 8px;gap:6px}.rail a>.i{display:none}}
@media (max-width:1320px){.chip .tkr{display:none}}
@media (max-width:1500px){
 .topbar-in{grid-template-columns:minmax(0,1fr) auto}
 .cofilter{grid-column:1/-1;grid-row:2;justify-self:stretch}
 .chip{flex:1 0 auto;justify-content:center}
 .brand{border-right:0;padding-right:0}}
@media (max-width:760px){.mkts{grid-column:1/-1;grid-row:3}.mkt{flex:1}}
@media (max-width:1180px){
 .layout{grid-template-columns:minmax(0,1fr);gap:20px}
 .hero{position:relative;top:auto;max-height:none;overflow:hidden}
 .hero-head{flex-direction:row;flex-wrap:wrap;justify-content:space-between;align-items:flex-end}
 .hero-meta{flex-direction:row;flex-wrap:wrap;gap:6px 18px}
 .tickers{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:980px){
 .layout{gap:16px;padding-top:16px}
 .rail{position:sticky;top:var(--topbar-h,100px);z-index:20;flex-direction:row;gap:4px;overflow-x:auto;scrollbar-width:none;margin:0 calc(-1 * var(--gut)) 18px;padding:8px var(--gut);border:0;border-radius:0;box-shadow:none;
  background:var(--glass);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);border-bottom:1px solid var(--rule)}
 .rail::-webkit-scrollbar{display:none}
 .rail-label{display:none}
 .rail a{flex:none;padding:7px 12px}
 .rail a[aria-current="page"]::before{display:none}
 .grid3,.btc-grid,.grid2{grid-template-columns:minmax(0,1fr)}
 .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}
 .kpi:nth-child(3){border-left:0}.kpi:nth-child(n+3){border-top:1px solid var(--rule)}
 .method{grid-template-columns:minmax(0,1fr)}
 .foot{flex-direction:column}}
@media (max-width:720px){
 #tbl-cal{min-width:0}
 #tbl-cal thead{display:none}
 #tbl-cal,#tbl-cal tbody{display:block}
 #tbl-cal tr{display:grid;grid-template-columns:auto minmax(0,1fr);column-gap:16px;row-gap:6px;padding:14px 16px;border-bottom:1px solid var(--rule)}
 #tbl-cal tr:last-child{border-bottom:0}
 #tbl-cal td{position:static;border:0;padding:0;background:transparent}
 #tbl-cal td:first-child{grid-row:span 4;align-self:start;padding-top:2px}
 #tbl-cal td[colspan]{grid-column:1/-1;grid-row:auto}
 #tbl-cal .cal-date{flex-direction:column;align-items:center;gap:2px;width:52px;padding:8px 0;border-radius:10px;background:var(--surface-2);text-align:center}
 #tbl-cal .cal-d{text-align:center}
 .etf-legend li{grid-template-columns:12px minmax(0,1fr) 64px}.etf-legend .bar{display:none}}
@media (max-width:640px){
 :root{--gut:16px}
 body{font-size:14.5px}
 .topbar-in{padding-top:8px;padding-bottom:8px}
 .brand-txt small{display:none}.brand{padding-right:0;border-right:0}.brand-txt b{font-size:1.1rem}
 .chip{padding-right:10px}
 .hero-in{padding:16px 12px 12px}
 .hero-head{margin-bottom:14px}
 .tickers{grid-template-columns:minmax(0,1fr);gap:8px}
 .tk{grid-template-columns:auto minmax(0,1fr) auto;grid-template-areas:"badge id price" "badge eur chg" "rows rows rows";row-gap:2px;padding:12px 14px}
 .tk-price{margin:0;font-size:1.3rem;text-align:right;justify-self:end}
 .tk-price .cur{font-size:.8rem}
 .tk-eur{font-size:.72rem}
 .tk-spark,.range{display:none}
 .tk-rows{margin-top:10px;padding-top:8px}
 .lead{font-size:1.04rem}
 .kpi{padding:12px 14px}.kpi .v{font-size:1.35rem}
 .alert{grid-template-columns:auto minmax(0,1fr);padding:12px 14px 12px 16px}
 .alert .btn{grid-column:2;justify-self:start}
 .news{padding:12px 14px;gap:10px}
 .filters{padding:12px}
 .field{min-width:0;flex:1 1 calc(50% - 7px)}
 .f-lvl,.f-search{flex-basis:100%}
 .f-lvl .seg{width:100%}
 .toolbar{gap:12px}
 .tb-group,.tb-group .seg{width:100%}
 .tb-group .seg button{flex:1 0 auto;justify-content:center}
 .fig{padding:14px 14px 12px}
 .fig-lg{padding:16px 14px 12px}
}
@media (pointer:coarse){.seg button{min-height:38px}.chip{min-height:40px}.btn{min-height:40px}}
@media (prefers-reduced-motion: reduce){*,*::before,*::after{transition:none!important;animation:none!important}}
</style>
</head>
<body>
<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;width:0;height:0;overflow:hidden" aria-hidden="true" focusable="false">
<symbol id="i-layers" viewBox="0 0 24 24"><path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/><path d="m22 17.65-9.17 4.16a2 2 0 0 1-1.66 0L2 17.65"/><path d="m22 12.65-9.17 4.16a2 2 0 0 1-1.66 0L2 12.65"/></symbol>
<symbol id="i-apple" viewBox="0 0 24 24"><path d="M12 20.94c1.5 0 2.75 1.06 4 1.06 3 0 6-8 6-12.22A4.91 4.91 0 0 0 17 5c-2.22 0-4 1.44-5 2-1-.56-2.78-2-5-2a4.9 4.9 0 0 0-5 4.78C2 14 5 22 8 22c1.25 0 2.5-1.06 4-1.06Z"/><path d="M10 2c1 .5 2 2 2 5"/></symbol>
<symbol id="i-cpu" viewBox="0 0 24 24"><rect width="16" height="16" x="4" y="4" rx="2"/><rect width="6" height="6" x="9" y="9" rx="1"/><path d="M15 2v2M15 20v2M2 15h2M2 9h2M20 15h2M20 9h2M9 2v2M9 20v2"/></symbol>
<symbol id="i-search" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></symbol>
<symbol id="i-pie" viewBox="0 0 24 24"><path d="M21 12c.552 0 1.005-.449.95-.998a10 10 0 0 0-8.953-8.951c-.55-.055-.998.398-.998.95v8a1 1 0 0 0 1 1z"/><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/></symbol>
<symbol id="i-landmark" viewBox="0 0 24 24"><path d="M3 22h18M6 18v-7M10 18v-7M14 18v-7M18 18v-7"/><path d="M12 2 20 7H4z"/></symbol>
<symbol id="i-user" viewBox="0 0 24 24"><path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></symbol>
<symbol id="i-dashboard" viewBox="0 0 24 24"><rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/></symbol>
<symbol id="i-newspaper" viewBox="0 0 24 24"><path d="M4 22h16a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2H8a2 2 0 0 0-2 2v16a2 2 0 0 1-2 2Zm0 0a2 2 0 0 1-2-2v-9c0-1.1.9-2 2-2h2"/><path d="M18 14h-8M15 18h-5"/><path d="M10 6h8v4h-8z"/></symbol>
<symbol id="i-chart" viewBox="0 0 24 24"><path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="m19 9-5 5-4-4-3 3"/></symbol>
<symbol id="i-bars" viewBox="0 0 24 24"><path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9M13 17V5M8 17v-3"/></symbol>
<symbol id="i-globe" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 2a14.5 14.5 0 0 0 0 20 14.5 14.5 0 0 0 0-20M2 12h20"/></symbol>
<symbol id="i-euro" viewBox="0 0 24 24"><path d="M4 10h12M4 14h9"/><path d="M19 6a7.7 7.7 0 0 0-5.2-2A7.9 7.9 0 0 0 6 12c0 4.4 3.5 8 7.8 8 2 0 3.8-.8 5.2-2"/></symbol>
<symbol id="i-swap" viewBox="0 0 24 24"><path d="M8 3 4 7l4 4M4 7h16M16 21l4-4-4-4M20 17H4"/></symbol>
<symbol id="i-calendar" viewBox="0 0 24 24"><path d="M8 2v4M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/></symbol>
<symbol id="i-database" viewBox="0 0 24 24"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v14a9 3 0 0 0 18 0V5"/><path d="M3 12a9 3 0 0 0 18 0"/></symbol>
<symbol id="i-book" viewBox="0 0 24 24"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></symbol>
<symbol id="i-octagon" viewBox="0 0 24 24"><path d="M12 16h.01M12 8v4"/><path d="M15.31 2a2 2 0 0 1 1.42.59l4.68 4.68A2 2 0 0 1 22 8.69v6.62a2 2 0 0 1-.59 1.42l-4.68 4.68a2 2 0 0 1-1.42.59H8.69a2 2 0 0 1-1.42-.59l-4.68-4.68A2 2 0 0 1 2 15.31V8.69a2 2 0 0 1 .59-1.42l4.68-4.68A2 2 0 0 1 8.69 2z"/></symbol>
<symbol id="i-triangle" viewBox="0 0 24 24"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4M12 17h.01"/></symbol>
<symbol id="i-circle-alert" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></symbol>
<symbol id="i-info" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></symbol>
<symbol id="i-circle-minus" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M8 12h8"/></symbol>
<symbol id="i-check" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></symbol>
<symbol id="i-x" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="m15 9-6 6M9 9l6 6"/></symbol>
<symbol id="i-external" viewBox="0 0 24 24"><path d="M15 3h6v6M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h3"/></symbol>
<symbol id="i-arrow-right" viewBox="0 0 24 24"><path d="M5 12h14M12 5l7 7-7 7"/></symbol>
<symbol id="i-chevron-down" viewBox="0 0 24 24"><path d="m6 9 6 6 6-6"/></symbol>
<symbol id="i-sliders" viewBox="0 0 24 24"><path d="M21 4h-7M10 4H3M21 12h-9M8 12H3M21 20h-5M12 20H3M14 2v4M8 10v4M16 18v4"/></symbol>
<symbol id="i-sort" viewBox="0 0 24 24"><path d="m21 16-4 4-4-4M17 20V4M3 8l4-4 4 4M7 4v16"/></symbol>
<symbol id="i-clock" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></symbol>
<symbol id="i-tag" viewBox="0 0 24 24"><path d="M12.59 2.59A2 2 0 0 0 11.17 2H4a2 2 0 0 0-2 2v7.17a2 2 0 0 0 .59 1.42l8.7 8.7a2.43 2.43 0 0 0 3.42 0l6.58-6.58a2.43 2.43 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r="1"/></symbol>
<symbol id="i-trend-up" viewBox="0 0 24 24"><path d="M22 7 13.5 15.5 8.5 10.5 2 17"/><path d="M16 7h6v6"/></symbol>
<symbol id="i-trend-down" viewBox="0 0 24 24"><path d="M22 17 13.5 8.5 8.5 13.5 2 7"/><path d="M16 17h6v-6"/></symbol>
<symbol id="i-minus" viewBox="0 0 24 24"><path d="M5 12h14"/></symbol>
<symbol id="i-rss" viewBox="0 0 24 24"><path d="M4 11a9 9 0 0 1 9 9M4 4a16 16 0 0 1 16 16"/><circle cx="5" cy="19" r="1"/></symbol>
<symbol id="i-sparkle" viewBox="0 0 24 24"><path d="M9.94 15.5A2 2 0 0 0 8.5 14.06l-6.13-1.58a.5.5 0 0 1 0-.96L8.5 9.94A2 2 0 0 0 9.94 8.5l1.58-6.14a.5.5 0 0 1 .96 0l1.58 6.14a2 2 0 0 0 1.44 1.44l6.14 1.58a.5.5 0 0 1 0 .96l-6.14 1.58a2 2 0 0 0-1.44 1.44l-1.58 6.14a.5.5 0 0 1-.96 0z"/></symbol>
<symbol id="i-text-search" viewBox="0 0 24 24"><path d="M21 6H3M10 12H3M10 18H3"/><circle cx="17" cy="15" r="3"/><path d="m21 19-1.9-1.9"/></symbol>
<symbol id="i-bell" viewBox="0 0 24 24"><path d="M10.27 21a2 2 0 0 0 3.46 0"/><path d="M3.26 15.33A1 1 0 0 0 4 17h16a1 1 0 0 0 .74-1.67C19.41 13.96 18 12.5 18 8A6 6 0 0 0 6 8c0 4.5-1.41 5.96-2.74 7.33"/></symbol>
<symbol id="i-zap" viewBox="0 0 24 24"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"/></symbol>
<symbol id="i-gauge" viewBox="0 0 24 24"><path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/></symbol>
<symbol id="i-activity" viewBox="0 0 24 24"><path d="M22 12h-2.48a2 2 0 0 0-1.93 1.46l-2.35 8.36a.25.25 0 0 1-.48 0L9.24 2.18a.25.25 0 0 0-.48 0l-2.35 8.36A2 2 0 0 1 4.49 12H2"/></symbol>
<symbol id="i-refresh" viewBox="0 0 24 24"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></symbol>
<symbol id="i-bitcoin" viewBox="0 0 24 24"><path d="M7 5h6.5a3.5 3.5 0 0 1 0 7H7zM7 12h7.5a3.5 3.5 0 0 1 0 7H7zM7 5v14M5.5 5H7M5.5 19H7M10 3v2M13.5 3v2M10 19v2M13.5 19v2"/></symbol>
<defs><linearGradient id="bbLogoG" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#6FA3FF"/><stop offset="1" stop-color="#2A5BD7"/></linearGradient></defs>
<symbol id="i-logo" viewBox="0 0 32 32"><rect width="32" height="32" rx="8" fill="url(#bbLogoG)"/><rect x=".5" y=".5" width="31" height="31" rx="7.5" fill="none" stroke="#fff" stroke-opacity=".18"/><g fill="#fff" fill-opacity=".26"><rect x="7" y="19" width="4" height="6" rx="1"/><rect x="14" y="15.5" width="4" height="9.5" rx="1"/><rect x="21" y="12" width="4" height="13" rx="1"/></g><path d="M6.5 19.5 12 14l4.2 3.2 8.6-8.6M19.8 8.3h5.3v5.3" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></symbol>
<symbol id="i-wallet" viewBox="0 0 24 24"><path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1"/><path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4"/></symbol>
<symbol id="i-plus" viewBox="0 0 24 24"><path d="M5 12h14M12 5v14"/></symbol>
<symbol id="i-trash" viewBox="0 0 24 24"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></symbol>
<symbol id="i-download" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5M12 15V3"/></symbol>
<symbol id="i-upload" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5M12 3v12"/></symbol>
<symbol id="i-calculator" viewBox="0 0 24 24"><rect width="16" height="20" x="4" y="2" rx="2"/><path d="M8 6h8M16 14v4M16 10h.01M12 10h.01M8 10h.01M12 14h.01M8 14h.01M12 18h.01M8 18h.01"/></symbol>
<symbol id="i-box" viewBox="0 0 24 24"><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5M12 22V12"/></symbol>
<symbol id="i-scale" viewBox="0 0 24 24"><path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1ZM2 16l3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1ZM7 21h10M12 3v18M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/></symbol>
</svg>

<header class="topbar" id="topbar"><div class="topbar-in">
  <div class="brand"><span class="brand-mark" aria-hidden="true"><svg class="logo" viewBox="0 0 32 32"><use href="#i-logo"/></svg></span><span class="brand-txt"><b>Bluechip Board</b><small>Daily board<i></i>AAPL · NVDA · GOOGL · ETFs · BTC</small></span></div>
  <div class="cofilter" role="group" aria-label="Filter by asset">
    <button class="chip" type="button" data-co="all" aria-pressed="true" style="--c:var(--ink)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg></span>All</button>
    <button class="chip" type="button" data-co="AAPL" aria-pressed="false" style="--c:var(--aapl)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-apple"/></svg></span>Apple<span class="tkr">AAPL</span></button>
    <button class="chip" type="button" data-co="NVDA" aria-pressed="false" style="--c:var(--nvda)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-cpu"/></svg></span>NVIDIA<span class="tkr">NVDA</span></button>
    <button class="chip" type="button" data-co="GOOGL" aria-pressed="false" style="--c:var(--googl)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-search"/></svg></span>Alphabet<span class="tkr">GOOGL</span></button>
    <button class="chip" type="button" data-co="MKT" aria-pressed="false" style="--c:var(--spx)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-pie"/></svg></span>ETF &amp; market<span class="tkr">SXR8</span></button>
    <button class="chip" type="button" data-co="BTC" aria-pressed="false" style="--c:var(--btc)"><span class="ci"><svg class="i" aria-hidden="true"><use href="#i-bitcoin"/></svg></span>Bitcoin<span class="tkr">BTC</span></button>
  </div>
  <div class="mkts" id="mkts" role="group" aria-label="Exchange hours, Lisbon time"></div>
</div></header>

<div class="layout">
<aside class="hero" aria-labelledby="heroTitle">
  <svg class="guilloche" id="guilloche" aria-hidden="true" preserveAspectRatio="none" viewBox="0 0 1200 340"></svg>
  <div class="hero-in">
    <div class="hero-head">
      <div><p class="eyebrow"><svg class="i" aria-hidden="true"><use href="#i-activity"/></svg>Market snapshot</p><h1 id="heroTitle">Apple, NVIDIA, Alphabet, iShares ETFs and Bitcoin</h1></div>
      <div class="hero-meta" id="heroSub"></div>
    </div>
    <div class="tickers" id="tickers"></div>
  </div>
</aside>
<div class="content">
<nav class="rail" id="rail" aria-label="Sections">
  <a href="#overview" data-tab="overview"><svg class="i" aria-hidden="true"><use href="#i-dashboard"/></svg><span>Overview</span><span class="cnt warn" id="cnt-alertas"></span></a>
  <a href="#portfolio" data-tab="portfolio"><svg class="i" aria-hidden="true"><use href="#i-wallet"/></svg><span>Portfolio</span></a>
  <a href="#news" data-tab="news"><svg class="i" aria-hidden="true"><use href="#i-newspaper"/></svg><span>News</span><span class="cnt" id="cnt-noticias"></span></a>
  <a href="#prices" data-tab="prices"><svg class="i" aria-hidden="true"><use href="#i-chart"/></svg><span>Prices</span></a>
<a href="#fundamentals" data-tab="fundamentals"><svg class="i" aria-hidden="true"><use href="#i-bars"/></svg><span>Fundamentals</span></a>
  <a href="#fx" data-tab="fx"><svg class="i" aria-hidden="true"><use href="#i-globe"/></svg><span>Currency &amp; Macroeconomics</span></a>
  <a href="#calendar" data-tab="calendar"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg><span>Calendar</span></a>
  <a href="#etf" data-tab="etf"><svg class="i" aria-hidden="true"><use href="#i-pie"/></svg><span>ETFs</span></a>
  <a href="#bitcoin" data-tab="bitcoin"><svg class="i" aria-hidden="true"><use href="#i-bitcoin"/></svg><span>Bitcoin</span></a>
  <a href="#sources" data-tab="sources"><svg class="i" aria-hidden="true"><use href="#i-database"/></svg><span>Sources &amp; method</span><span class="cnt err" id="cnt-fontes"></span></a>
</nav>
<main id="main" tabindex="-1">
<div id="stale"></div>

<section class="panel" data-panel="overview" aria-labelledby="h-resumo">
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-dashboard"/></svg></span><h2 id="h-resumo">Overview</h2></div>
  <p class="lead" id="resumoLead"></p>
  <div class="kpis" id="resumoStats"></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-bell"/></svg>Alerts</h3>
  <div class="list" id="alertas"></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Alerts come from simple rules on prices, currency, the calendar and the news. They are not buy or sell signals.</p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-zap"/></svg>News with the highest potential impact<button class="btn" type="button" data-goto="news">See all<svg class="i i-go" aria-hidden="true"><use href="#i-arrow-right"/></svg></button></h3>
  <div class="list" id="topNews"></div>
  <div class="fig"><div class="fig-head"><div><p class="fig-title">News by asset and level</p><p class="fig-sub">Number of stories in the collected window, excluding those classed as noise.</p></div></div><div id="ch-niveis"></div></div>
</section>

<section class="panel" data-panel="portfolio" aria-labelledby="h-portfolio" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-wallet"/></svg></span><h2 id="h-portfolio">Portfolio</h2></div>
  <p class="lead">Record what you buy and sell, in euros. Values follow the latest prices every time the script runs. Nothing you type is sent over the internet: it is kept in this browser and, once you save it, in bluechip-board-backup.json. The script copies that file into bluechip-board.html so it loads by itself, so do not share that page if you want to keep your portfolio private.</p>
  <div class="kpis" id="pfKpis"></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-sliders"/></svg>Your holdings</h3>
  <div class="tbl-wrap"><table id="tbl-pf"></table></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Quantities and amounts invested are your purchases minus your sales, matched first in, first out (each sale uses up your oldest purchases first). Bitcoin purchases are shared with the 365-day counter of the Bitcoin section: add them in either place.</p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-scale"/></svg>Your return · compared with the same money in SXR8</h3>
  <div class="kpis" id="retKpis"></div>
  <p class="note" id="retNote"></p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-pie"/></svg>Target allocation · next contribution</h3>
  <div class="card">
    <form class="lot-form" id="tgtForm" novalidate>
      <div class="field"><label class="f-label" for="tgtBand"><svg class="i" aria-hidden="true"><use href="#i-sliders"/></svg>Tolerance band (± percentage points)</label><input class="in" type="number" id="tgtBand" min="0" max="50" step="any" placeholder="5"></div>
      <div class="field"><label class="f-label" for="tgtMonthly"><svg class="i" aria-hidden="true"><use href="#i-euro"/></svg>Monthly amount (€)</label><input class="in" type="number" id="tgtMonthly" min="0" step="any" placeholder="0"></div>
      <button class="btn btn-pri" type="submit" id="tgtSubmit"><svg class="i" aria-hidden="true"><use href="#i-check"/></svg>Save targets</button>
    </form>
    <p class="form-msg" id="tgtMsg" aria-live="polite">Type a target weight for each asset in the table (they must add up to 100%), the tolerance band and the amount you plan to invest each month, then save.</p>
    <div class="tbl-wrap"><table id="tbl-tgt"></table></div>
    <div id="tgtOut"></div>
  </div>
  <p class="note" id="tgtNote"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Rebalancing by selling realises taxable capital gains; rebalancing with new contributions does not. The split works out how far each asset would be below its target once the next contribution is added, and shares the contribution in proportion to those gaps. Assets without a current price are left out. It is arithmetic on today's prices, not advice.</span></p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-book"/></svg>Investment policy</h3>
  <div class="card">
    <form class="pol-form" id="polForm" novalidate>
      <div class="field"><label class="f-label" for="pol-horizon">Horizon</label><textarea class="in" id="pol-horizon" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <div class="field"><label class="f-label" for="pol-allocation">Allocation</label><textarea class="in" id="pol-allocation" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <div class="field"><label class="f-label" for="pol-monthly">Monthly amount</label><textarea class="in" id="pol-monthly" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <div class="field"><label class="f-label" for="pol-sell">What would make me sell</label><textarea class="in" id="pol-sell" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <div class="field"><label class="f-label" for="pol-drop20">My rule for a 20% drop</label><textarea class="in" id="pol-drop20" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <div class="field"><label class="f-label" for="pol-drop30">My rule for a 30% drop</label><textarea class="in" id="pol-drop30" rows="2" maxlength="2000" placeholder="In your own words"></textarea></div>
      <button class="btn btn-pri" type="submit" id="polSubmit"><svg class="i" aria-hidden="true"><use href="#i-check"/></svg>Save policy</button>
    </form>
    <p class="form-msg" id="polMsg" aria-live="polite">Write down your own rules while markets are calm. Your rules for a 20% and a 30% drop are shown next to the drop alerts (Overview), so you can read them when prices fall. It stays on this computer, like the rest of your data.</p>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-zap"/></svg>Stress test · past market falls</h3>
  <p class="note" id="stressLead"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span><b>Hypothetical scenario, not a forecast.</b> What your holdings today would lose if each asset repeated its own move, in euros, between the peak and the trough of three S&amp;P 500 falls (close of each date, or the last close before it).</span></p>
  <div id="stressOut"></div>
  <div class="tbl-wrap"><table id="tbl-stress"></table></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Dividends · projected for the next 12 months</h3>
  <div class="kpis" id="divKpis"></div>
  <div class="tbl-wrap"><table id="tbl-div"></table></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span id="divNetNote"></span></p>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Projection = the shares you hold now (the same quantities as Your holdings: purchases minus sales, oldest first, adjusted for stock splits) × the annual dividend per share, which is the latest dividend × the number of payments a year (from Yahoo Finance, in dollars). It is converted to euros at the latest EUR/USD rate, because the payments are still to come. Before tax: US withholding and Portuguese tax are not deducted. The ETFs (SXR8, EUNK, IS3N, EUNN) are not included: they are accumulating ETFs, so their dividends are reinvested in their price. Bitcoin pays no dividends.</span></p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-plus"/></svg>Purchases and sales · stocks, ETF and Bitcoin</h3>
  <div class="card">
    <form class="lot-form" id="buyForm" novalidate>
      <div class="field"><label class="f-label" for="buyAsset"><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg>Asset</label><select id="buyAsset"></select></div>
      <div class="field"><label class="f-label" for="buyType"><svg class="i" aria-hidden="true"><use href="#i-swap"/></svg>Type</label><select id="buyType"><option value="buy">Purchase</option><option value="sell">Sale</option></select></div>
      <div class="field"><label class="f-label" for="buyDate"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Date</label><input class="in" type="date" id="buyDate" required></div>
      <div class="field"><label class="f-label" for="buyQty"><svg class="i" aria-hidden="true"><use href="#i-tag"/></svg>Quantity</label><input class="in" type="number" id="buyQty" min="0" step="any" value="1" required></div>
      <div class="field"><label class="f-label" for="buyPrice"><svg class="i" aria-hidden="true"><use href="#i-euro"/></svg>Price per unit (€)</label><input class="in" type="number" id="buyPrice" min="0" step="any" placeholder="Filled in from the date" required></div>
      <div class="field"><label class="f-label" for="buyFee"><svg class="i" aria-hidden="true"><use href="#i-euro"/></svg>Fee (€, optional)</label><input class="in" type="number" id="buyFee" min="0" step="any" placeholder="Not recorded"></div>
      <button class="btn btn-pri" type="submit" id="buySubmit"><svg class="i" aria-hidden="true"><use href="#i-plus"/></svg>Add purchase</button>
    </form>
    <p class="form-msg" id="buyMsg" aria-live="polite">Pick the asset and the date: the price is filled in with that day's closing price, in euros (US stocks are converted at that day's EUR/USD rate; Bitcoin uses the BTC/EUR price at the end of the day, UTC).</p>
    <div class="tbl-wrap"><table id="tbl-buys"></table></div>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-calculator"/></svg>Before you sell · simulation</h3>
  <div class="card">
    <form class="lot-form" id="sellForm" novalidate>
      <div class="field"><label class="f-label" for="sellAsset"><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg>Asset</label><select id="sellAsset"></select></div>
      <div class="field"><label class="f-label" for="sellQty"><svg class="i" aria-hidden="true"><use href="#i-tag"/></svg>Quantity</label><input class="in" type="number" id="sellQty" min="0" step="any" value="1"></div>
      <button class="btn btn-pri" type="submit"><svg class="i" aria-hidden="true"><use href="#i-calculator"/></svg>Simulate</button>
    </form>
    <p class="form-msg">What selling today, at the latest price, would mean: which purchases it uses up (oldest first) and the gain. Nothing is saved.</p>
    <div id="sellOut"></div>
    <div class="tbl-wrap"><table id="tbl-sell"></table></div>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-calculator"/></svg>Export for taxes · IRS, Anexo J</h3>
  <div class="card">
    <form class="lot-form" id="taxForm" novalidate>
      <div class="field"><label class="f-label" for="taxYear"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Tax year</label><select id="taxYear"></select></div>
      <button class="btn btn-pri" type="submit" id="taxExport"><svg class="i" aria-hidden="true"><use href="#i-download"/></svg>Export for Taxes (Anexo J)</button>
    </form>
    <p class="form-msg" id="taxMsg" aria-live="polite">Creates two CSV files from your sales in the chosen year: AnexoJ_Stocks_ETFs (Quadro 9.2A) and AnexoJ_Crypto (Quadro 9.4A). Everything is calculated in this browser: nothing is sent anywhere and your portfolio is not changed.</p>
    <form class="lot-form" id="div8aForm" novalidate>
      <div class="field"><label class="f-label" for="divYear"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Tax year (dividends)</label><select id="divYear"></select></div>
      <button class="btn btn-pri" type="submit" id="divExport"><svg class="i" aria-hidden="true"><use href="#i-download"/></svg>Export dividends (Quadro 8A)</button>
    </form>
    <p class="form-msg" id="divMsg" aria-live="polite">Creates AnexoJ_Dividends: one row per Apple, NVIDIA and Alphabet dividend in the chosen year (by ex-dividend date), with the shares you held before the ex-date. Amounts in dollars; euros only as a reference at that day's EUR/USD rate; US tax withheld estimated at 15% (assumes form W-8BEN).</p>
  </div>
  <div class="callout warn"><svg class="i" aria-hidden="true"><use href="#i-triangle"/></svg><div>These files help you fill in the return; the Portal das Finanças does not import them. Each row is one sale matched to one purchase (oldest first), with the dates and euro amounts that go into Anexo J. Amounts in euros come from the prices you entered; the dollar columns are only a reference at that day's EUR/USD rate. Add what the page cannot know: the counterparty country, fees and any tax paid abroad. Rows marked REVIEW need checking. Not tax advice: confirm with the official instructions or an accountant.</div></div>
  <div class="grid2">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Allocation</p><p class="fig-sub">Share of each asset in the value of the portfolio.</p></div></div><div id="ch-pf-alloc"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Real exposure to each company</p><p class="fig-sub" id="pfExpSub">What you hold directly plus what you own through the ETF, using the iShares weights.</p></div></div><div id="ch-pf-exp"></div></div>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-globe"/></svg>Real exposure · country, sector and currency</h3>
  <p class="note" id="expoLead"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Your holdings at today's values, looking through each ETF with the country, sector and currency breakdown of its latest iShares file. Apple, NVIDIA and Alphabet count as United States and USD; Bitcoin counts as Crypto.</span></p>
  <div id="expoOut"></div>
  <div class="grid3 expo-grid">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Country</p><p class="fig-sub">Where the companies are based.</p></div></div><div id="ch-expo-paises"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Sector</p><p class="fig-sub">What the companies do.</p></div></div><div id="ch-expo-setores"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Underlying currency</p><p class="fig-sub">The currencies the holdings trade in.</p></div></div><div id="ch-expo-moedas"></div></div>
  </div>
  <p class="note" id="expoNote"></p>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-scale"/></svg>Concentration by company</h3>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Only where the company is unambiguous: Apple, NVIDIA and Alphabet (directly plus through the ETFs), and each fund's top 10 on its own. Companies are never added up across funds by name.</span></p>
  <div id="concOut"></div>
  <div class="tbl-wrap"><table id="tbl-conc"></table></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-database"/></svg>Backup</h3>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Your purchases live in this browser. Save them to <b>bluechip-board-backup.json</b> in the project folder (the first time, choose the BluechipBoard folder; Chrome remembers it): the script includes that file in the site, so your data loads by itself when you open it, even if the browser was cleared.</p>
  <div class="tools"><button class="btn btn-pri" type="button" id="bkExport"><svg class="i" aria-hidden="true"><use href="#i-download"/></svg>Save to project folder</button><label class="btn" for="bkImport"><svg class="i" aria-hidden="true"><use href="#i-upload"/></svg>Restore backup</label><input type="file" id="bkImport" accept=".json,application/json" hidden><span class="form-msg" id="bkMsg" aria-live="polite" style="margin:0"></span></div>
  <div id="bkState" aria-live="polite"></div>
</section>

<section class="panel" data-panel="news" aria-labelledby="h-noticias" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-newspaper"/></svg></span><h2 id="h-noticias">News</h2></div>
  <p class="lead">Each story is classified automatically by its potential impact on the fundamentals, based on the headline, its themes and the quality of the source. Always check the primary source before acting.</p>
  <div class="filters" role="group" aria-label="News filters">
    <div class="f-row">
      <div class="field f-lvl"><span class="f-label" id="lbl-lvl"><svg class="i" aria-hidden="true"><use href="#i-sliders"/></svg>Level</span>
        <div class="seg" role="group" aria-labelledby="lbl-lvl" id="f-lvl">
          <button type="button" aria-pressed="true" data-lvl="all">All</button><button type="button" aria-pressed="false" data-lvl="red" style="--lc:var(--red)"><svg class="i" aria-hidden="true"><use href="#i-octagon"/></svg>Material</button><button type="button" aria-pressed="false" data-lvl="orange" style="--lc:var(--orange)"><svg class="i" aria-hidden="true"><use href="#i-triangle"/></svg>Important</button><button type="button" aria-pressed="false" data-lvl="yellow" style="--lc:var(--yellow)"><svg class="i" aria-hidden="true"><use href="#i-circle-alert"/></svg>Moderate</button><button type="button" aria-pressed="false" data-lvl="white" style="--lc:var(--grey)"><svg class="i" aria-hidden="true"><use href="#i-circle-minus"/></svg>Noise</button>
        </div></div>
      <div class="field"><label class="f-label" for="f-tema"><svg class="i" aria-hidden="true"><use href="#i-tag"/></svg>Theme</label><select id="f-tema"><option value="">All</option></select></div>
      <div class="field"><label class="f-label" for="f-tier"><svg class="i" aria-hidden="true"><use href="#i-rss"/></svg>Source</label><select id="f-tier"><option value="">All</option><option value="primaria">Primary sources</option><option value="referencia">Leading press</option><option value="outra">Other</option><option value="cuidado">Handle with care</option></select></div>
    </div>
    <div class="f-row">
      <div class="field f-search"><label class="f-label" for="f-q"><svg class="i" aria-hidden="true"><use href="#i-text-search"/></svg>Search</label><div class="search"><svg class="i" aria-hidden="true"><use href="#i-text-search"/></svg><input type="search" id="f-q" placeholder="e.g. antitrust, China" autocomplete="off"></div></div>
      <div class="field"><label class="f-label" for="f-dias"><svg class="i" aria-hidden="true"><use href="#i-clock"/></svg>Period</label><select id="f-dias"><option value="">Whole window</option><option value="1">Last 24 h</option><option value="3">Last 3 days</option></select></div>
      <div class="field"><label class="f-label" for="f-ord"><svg class="i" aria-hidden="true"><use href="#i-sort"/></svg>Sort by</label><select id="f-ord"><option value="score">Potential impact</option><option value="data">Most recent</option><option value="rel">Relevance to me</option></select></div>
      <label class="switch"><input type="checkbox" id="f-novas"><span class="sw" aria-hidden="true"></span><span>Only new since the last run</span></label>
    </div>
  </div>
  <div class="list-bar"><p class="note" id="newsCount" aria-live="polite"></p></div>
  <p class="note" id="newsRelNote" hidden></p>
  <div class="list" id="newsList"></div>
  <button class="btn more" id="moreBtn" type="button"></button>
  <div class="fig"><div class="fig-head"><div><p class="fig-title">Themes in the filtered news</p><p class="fig-sub">A story can have several themes.</p></div></div><div id="ch-temas"></div></div>
</section>

<section class="panel" data-panel="prices" aria-labelledby="h-cotacoes" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-chart"/></svg></span><h2 id="h-cotacoes">Prices</h2></div>
  <p class="lead" id="cotLead"></p>
  <div class="toolbar">
    <div class="tb-group"><span class="f-label" id="lbl-per"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Period</span><div class="seg" role="group" aria-labelledby="lbl-per" id="f-per"><button type="button" aria-pressed="false" data-per="1M">1 month</button><button type="button" aria-pressed="false" data-per="3M">3 months</button><button type="button" aria-pressed="false" data-per="6M">6 months</button><button type="button" aria-pressed="true" data-per="1A">1 year</button></div></div>
    <div class="tb-group"><span class="f-label" id="lbl-cur"><svg class="i" aria-hidden="true"><use href="#i-swap"/></svg>Currency</span><div class="seg" role="group" aria-labelledby="lbl-cur" id="f-cur"><button type="button" aria-pressed="true" data-cur="EUR">€ In euros</button><button type="button" aria-pressed="false" data-cur="USD">$ In dollars</button></div></div>
  </div>
  <div class="fig fig-lg"><div class="fig-head"><div><p class="fig-title">Relative performance</p><p class="fig-sub" id="perfSub"></p></div></div><div id="ch-perf"></div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-bars"/></svg>Returns</h3>
  <div class="tbl-wrap"><table id="tbl-ret"></table></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-gauge"/></svg>Risk and trend indicators</h3>
  <div class="tbl-wrap"><table id="tbl-ind"></table></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Stocks and the ETF count trading sessions (1 week = 5, 1 month = 21, 1 year = 252). Bitcoin trades 24 hours a day, 7 days a week: for it, 1 session is 1 day (UTC), 1 week is 7 days, and the averages and volatility count calendar days. Under each price, “close” gives the date of that closing price and “live” the time of an intraday price taken while the session was still open.</p>
  <div class="fig"><div class="fig-head"><div><p class="fig-title">Distance from the 52-week high</p><p class="fig-sub">In %, in the selected currency. Drops of 20% or more are worth checking against expected earnings.</p></div></div><div id="ch-dist"></div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-zap"/></svg>Big moves explained</h3>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span id="movesNote"></span></p>
  <div class="tbl-wrap"><table id="tbl-moves"></table></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Earnings reactions</h3>
  <div class="tbl-wrap"><table id="tbl-earn"></table></div>
  <div id="earnCases"></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Moves in dollars, from the long-term daily closes (adjusted for stock splits). <b>Reaction session</b>: the first session after the earnings release was accepted by the SEC (New York time: after 16:00, the next session; before 09:30 or during the session, that same session), compared with the close before it. <b>After 5 sessions</b>: the close of the 5th session, the reaction session being the 1st. Releases whose acceptance time is not known are not counted. Past reactions say nothing certain about the next one.</span></p>
</section>

<section class="panel" data-panel="fundamentals" aria-labelledby="h-fund" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-bars"/></svg></span><h2 id="h-fund">Fundamentals</h2></div>
  <p class="lead">Apple, NVIDIA and Alphabet from their own SEC filings (10-Q and 10-K): revenue and its growth, margins, free cash flow, earnings per share and the share count, with the valuation today and over 10 years. Each figure shows its quarter and the date it was published at the SEC.</p>
  <div id="fundOut"></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span><b>Valuation gives context; it does not predict the short term:</b> a high or low P/E says little about the next weeks or months. The 10-year history uses, at each month-end, only the quarters already published by then (by their SEC filing date), so later revisions are not used before they existed. P/E = price ÷ diluted EPS of the last 4 quarters; P/FCF = price × diluted shares ÷ free cash flow (operating cash flow − capex) of the last 4 quarters. Prices in dollars, adjusted for stock splits, like the EPS and the shares. * = derived: the 4th quarter is the year minus 9 months, and cash flow quarters come from year-to-date figures.</span></p>
</section>

<section class="panel" data-panel="fx" aria-labelledby="h-cambio" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-globe"/></svg></span><h2 id="h-cambio">Currency &amp; Macroeconomics</h2></div>
  <p class="lead">Your return in euros is the stock's return in dollars combined with the move of the dollar against the euro: when EUR/USD rises, the dollar weakens and your return in euros shrinks.</p>
  <div class="kpis" id="fxKpis"></div>
  <div class="fig fig-lg"><div class="fig-head"><div><p class="fig-title">EUR/USD</p><p class="fig-sub">Dollars per euro, over the period selected in Prices.</p></div></div><div id="ch-fx"></div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-swap"/></svg>How much of the return came from currency</h3>
  <div class="tbl-wrap"><table id="tbl-dec"></table></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-landmark"/></svg>US market</h3>
  <div class="grid3">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">S&amp;P 500</p><p class="fig-sub">Benchmark index.</p></div><div class="mk" id="k-gspc"></div></div><div id="ch-gspc"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">VIX</p><p class="fig-sub">Above 20–25 signals market stress.</p></div><div class="mk" id="k-vix"></div></div><div id="ch-vix"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">10-year Treasury yield</p><p class="fig-sub">In %. Higher yields weigh on expensive stocks.</p></div><div class="mk" id="k-tnx"></div></div><div id="ch-tnx"></div></div>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-euro"/></svg>Euro area · ECB</h3>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span><b>Thermometers, not clocks:</b> these indicators describe conditions; they do not say when markets will move. Data from the ECB Data Portal, each with its date.</span></p>
  <div class="grid2">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">ECB deposit facility rate</p><p class="fig-sub">In %, over 10 years. The rate the ECB pays on banks' overnight deposits: its main policy rate. Changes take many months to reach the economy.</p></div><div class="mk" id="k-dfr"></div></div><div id="ch-dfr"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Euro area inflation (HICP)</p><p class="fig-sub">Change in consumer prices over 12 months, in %, over 10 years. The ECB aims at 2% over the medium term; one month says little about the next.</p></div><div class="mk" id="k-hicp"></div></div><div id="ch-hicp"></div></div>
  </div>
  <p class="note" id="macroNao"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><span>Not shown: the US yield curve (10-year minus 3-month), the high-yield credit spread and the Chicago Fed financial conditions index, because their source (FRED) refuses this script's automated downloads; and the S&amp;P 500 CAPE, because no public CSV source has current data. For context: an inverted yield curve preceded US recessions with long and variable lags, and has given false alarms; a high valuation has gone with lower returns over the following 10 years, not over the next month.</span></p>
</section>

<section class="panel" data-panel="calendar" aria-labelledby="h-calendario" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-calendar"/></svg></span><h2 id="h-calendario">Calendar</h2></div>
  <p class="lead">Dates that tend to move these assets. Estimated dates should be updated in the script once the companies confirm them. The Bitcoin halving date is forecast on every run from the network's current pace. Earnings dates marked Nasdaq are updated automatically on every run.</p>
  <div class="toolbar">
    <div class="field"><label class="f-label" for="c-imp"><svg class="i" aria-hidden="true"><use href="#i-sliders"/></svg>Importance</label><select id="c-imp"><option value="">All</option><option>Very high</option><option>High</option><option>Medium</option></select></div>
    <label class="switch"><input type="checkbox" id="c-past"><span class="sw" aria-hidden="true"></span><span>Show past events</span></label>
  </div>
  <div class="tbl-wrap" style="margin-top:14px"><table id="tbl-cal"></table></div>
</section>

<section class="panel" data-panel="etf" aria-labelledby="h-etf" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-pie"/></svg></span><h2 id="h-etf">iShares Core S&amp;P 500</h2></div>
  <div class="toolbar"><div class="tb-group"><span class="f-label" id="lbl-etf"><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg>Fund</span><div class="seg" role="group" aria-labelledby="lbl-etf" id="f-etf"></div></div></div>
  <p class="lead" id="etfLead"></p>
  <div class="fig fig-lg" id="fig-etf-w"><div class="fig-head"><div><p class="fig-title">Weight of the three companies in the ETF</p><p class="fig-sub" id="etfSub"></p></div><div class="etf-total" id="etfTotal"></div></div><div id="ch-etf"></div></div>
  <div class="grid2">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Top 10 holdings</p><p class="fig-sub" id="top10Sub"></p></div></div><div id="ch-top10"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Sectors</p><p class="fig-sub">Weight of each sector in the ETF.</p></div></div><div id="ch-sect"></div></div>
  </div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-chart"/></svg>Long-term history</h3>
  <div class="toolbar"><div class="tb-group"><span class="f-label" id="lbl-elp"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Period</span><div class="seg" role="group" aria-labelledby="lbl-elp" id="f-elp"><button type="button" aria-pressed="false" data-lp="5">5 years</button><button type="button" aria-pressed="true" data-lp="10">10 years</button><button type="button" aria-pressed="false" data-lp="max" id="elp-max">Since 2010</button></div></div><div class="tb-group"><span class="f-label" id="lbl-elog"><svg class="i" aria-hidden="true"><use href="#i-chart"/></svg>Scale</span><div class="seg" role="group" aria-labelledby="lbl-elog" id="f-elog"><button type="button" aria-pressed="true" data-log="0">Linear</button><button type="button" aria-pressed="false" data-log="1">Log</button></div></div></div>
  <div class="fig fig-lg"><div class="fig-head"><div><p class="fig-title" id="etfHistTitle">SXR8 price in euros</p><p class="fig-sub">Daily close on Xetra. Dividends are reinvested in the price (accumulating ETF).</p></div><div class="mk" id="k-etf-hist"></div></div><div id="ch-etf-hist"></div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-trend-down"/></svg><span id="etfDropsTitle">Biggest drops since 2010</span></h3>
  <div class="grid2 dd-grid">
    <div><div class="tbl-wrap"><table id="tbl-drops-etf"></table></div>
    <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Drops of 10% or more from a previous high, and how long the price took to get back there.</p></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Below its previous high</p><p class="fig-sub">Underwater chart: how far below its highest close so far the price was on each day (price ÷ running maximum − 1), in euros.</p></div><div class="mk" id="k-etf-dd"></div></div><div id="ch-etf-dd"></div></div>
  </div>
  <details class="fig roll" id="roll-etf"><summary><span class="fig-title">Rolling returns · 1, 3 and 5 years</span><span class="fig-sub">Every period of that length in the history of the chosen fund. Calculated when you open it.</span></summary><div id="roll-etf-body"></div></details>
  <div class="fig" id="sim-etf"></div>
  <div id="sim-etf-outros"></div>
  <div class="fig" id="strat-etf"></div>
</section>

<section class="panel" data-panel="bitcoin" aria-labelledby="h-bitcoin" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-bitcoin"/></svg></span><h2 id="h-bitcoin">Bitcoin</h2></div>
  <p class="lead" id="btcLead"></p>
  <div class="kpis" id="btcKpis"></div>
  <div class="toolbar" style="margin-top:22px">
    <div class="tb-group"><span class="f-label" id="lbl-bper"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Period</span><div class="seg" role="group" aria-labelledby="lbl-bper"><button type="button" aria-pressed="false" data-per="1M">1 month</button><button type="button" aria-pressed="false" data-per="3M">3 months</button><button type="button" aria-pressed="false" data-per="6M">6 months</button><button type="button" aria-pressed="true" data-per="1A">1 year</button></div></div>
  </div>
  <div class="fig fig-lg"><div class="fig-head"><div><p class="fig-title">Bitcoin price in euros</p><p class="fig-sub">BTC/EUR, one point per day (UTC), with the 50- and 200-day moving averages.</p></div><div class="mk" id="k-btc"></div></div><div id="ch-btc"></div></div>
  <div class="btc-grid">
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Market sentiment</p><p class="fig-sub">Fear &amp; Greed Index (0 = extreme fear, 100 = extreme greed).</p></div></div><div id="btcSent"></div><div id="ch-fng"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Next halving</p><p class="fig-sub">Every 210,000 blocks, the issuance of new bitcoins is cut in half.</p></div></div><div id="btcHalv"></div></div>
    <div class="fig"><div class="fig-head"><div><p class="fig-title">Network and market</p><p class="fig-sub">Blockchain data (mempool.space) and market data (CoinGecko).</p></div></div><div id="btcRede"></div></div>
  </div>
  <div class="fig"><div class="fig-head"><div><p class="fig-title">Correlation with the other assets</p><p class="fig-sub">Daily returns on the days both traded, last 3 months, in euros. Close to 1: they move together; close to 0: no relationship.</p></div></div><div id="ch-corr"></div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-chart"/></svg>Long-term history</h3>
  <div class="fig fig-lg"><div class="fig-head"><div><p class="fig-title">Bitcoin price in euros since 2014</p><p class="fig-sub">Logarithmic scale: the same distance means the same % move. The dashed lines mark the halvings.</p></div></div><div id="ch-btc-hist"></div></div>
  <div class="fig"><div class="fig-head"><div><p class="fig-title">Drop from the all-time high</p><p class="fig-sub">How far below its highest price so far Bitcoin was on each day, in euros. The history starts in September 2014, after the late-2013 peak, so until early 2017 the drop is measured from the highest price since September 2014, not from that earlier peak.</p></div><div class="mk" id="k-btc-dd"></div></div><div id="ch-btc-dd"></div></div>
  <div class="tbl-wrap"><table id="tbl-drops-btc"></table></div>
  <p class="note"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg>Drops of 40% or more from a previous high, and how long the price took to get back there.</p>
  <details class="fig roll" id="roll-btc"><summary><span class="fig-title">Rolling returns · 1, 3 and 5 years</span><span class="fig-sub">Every period of that length in Bitcoin's history, in calendar days. Calculated when you open it.</span></summary><div id="roll-btc-body"></div></details>
  <div class="fig" id="sim-btc"></div>
  <div class="fig" id="strat-btc"></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-clock"/></svg>365-day tax counter</h3>
  <div class="card">
    <form class="lot-form" id="lotForm" novalidate>
      <div class="field"><label class="f-label" for="lotDate"><svg class="i" aria-hidden="true"><use href="#i-calendar"/></svg>Purchase date</label><input class="in" type="date" id="lotDate" required></div>
      <div class="field"><label class="f-label" for="lotQty"><svg class="i" aria-hidden="true"><use href="#i-bitcoin"/></svg>Amount (BTC)</label><input class="in" type="number" id="lotQty" min="0" step="any" placeholder="e.g. 0.05" required></div>
      <div class="field"><label class="f-label" for="lotCost"><svg class="i" aria-hidden="true"><use href="#i-euro"/></svg>Total paid (€)</label><input class="in" type="number" id="lotCost" min="0" step="any" placeholder="e.g. 3500" required></div>
      <button class="btn btn-pri" type="submit"><svg class="i" aria-hidden="true"><use href="#i-plus"/></svg>Add purchase</button>
    </form>
    <p class="form-msg" id="lotMsg" aria-live="polite"></p>
    <div class="kpis" id="lotSum"></div>
    <div class="tbl-wrap"><table id="tbl-lots"></table></div>
  </div>
  <div class="callout"><svg class="i" aria-hidden="true"><use href="#i-info"/></svg><div>Taxes in Portugal: as a rule, gains on crypto-assets held for 365 days or more are not taxed; below that, they are taxed at 28% and declared in the IRS return: Annex G for a Portuguese platform, Annex J (Quadro 9.4A) for a foreign one. Sales after 365 days are exempt but still declared (Annex G1). The Portfolio section can export your sales for Annex J. Each purchase counts on its own: the 365 days run from its own date. Sales recorded in the Portfolio section use up your oldest purchases first (first in, first out). The counter treats a purchase as tax-free from the day it completes 365 days; if you plan to sell close to that day, leave a margin of a day or two. Confirm your situation on the Portal das Finanças or with an accountant. Your purchases are kept in this browser and in the backup file (see Backup in the Portfolio section).</div></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-zap"/></svg>Bitcoin news with the highest potential impact<button class="btn" type="button" id="btcAllNews">See all<svg class="i i-go" aria-hidden="true"><use href="#i-arrow-right"/></svg></button></h3>
  <div class="list" id="btcNews"></div>
</section>

<section class="panel" data-panel="sources" aria-labelledby="h-fontes" hidden>
  <div class="panel-head"><span class="ph-ic" aria-hidden="true"><svg class="i"><use href="#i-database"/></svg></span><h2 id="h-fontes">Sources &amp; method</h2></div>
  <p class="lead" id="srcLead"></p>
  <div class="tbl-wrap tall"><table id="tbl-src"></table></div>
  <div id="manut"></div>
  <h3 class="sec-h"><svg class="i" aria-hidden="true"><use href="#i-book"/></svg>How the classification works</h3>
  <ul class="method">
    <li><b><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg>Assets</b>Keywords in the headline (for example, iPhone for Apple, Blackwell for NVIDIA, Gemini for Alphabet, spot ETF or halving for Bitcoin). Stories from a company's own newsroom without keywords are assigned to that company. Stories from a search feed (Google News, Yahoo Finance) without any keyword are only a weak match: they are tagged “Feed match only”, get no company bonus and are at most moderate. From crypto news sites, only headlines about Bitcoin or the crypto market are kept.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-tag"/></svg>Themes and weights</b>Earnings, regulation and courts, China and exports, security and fraud (3 points); AI and CapEx, capital and shareholders, management, macro and rates, ETF flows, institutional adoption (2); products, currency, network and mining (1); analysts (0.5). The strongest theme counts in full and each extra one at half, up to 5 points.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-triangle"/></svg>Severity</b>Adverse events such as a ban, export curbs, a guidance cut, a plunge or a delayed launch add 3 points.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-rss"/></svg>Source</b>Primary sources +1.5; leading press +1; sites to handle with care −1.5. Headlines like “stocks to buy now” are classed as noise.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-sliders"/></svg>Levels</b><span class="lvl-in l-red"><svg class="i" aria-hidden="true"><use href="#i-octagon"/></svg>Material</span> from 7 points, <span class="lvl-in l-orange"><svg class="i" aria-hidden="true"><use href="#i-triangle"/></svg>Important</span> from 4, <span class="lvl-in l-yellow"><svg class="i" aria-hidden="true"><use href="#i-circle-alert"/></svg>Moderate</span> from 2, <span class="lvl-in l-white"><svg class="i" aria-hidden="true"><use href="#i-circle-minus"/></svg>Noise</span> below that.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-layers"/></svg>Duplicates</b>The same event reported by several sources appears only once, with the other sources listed underneath. Two headlines are grouped when they share enough words, with more weight on rare ones (names and amounts such as “California”, “smuggling” or “$300M”), were published less than 3 days apart, cover the same asset and have no conflicting numbers (dates, amounts). The story with the most points is kept.</li>
    <li><b><svg class="i" aria-hidden="true"><use href="#i-scale"/></svg>Sentiment</b>A simple count of positive and negative words in the headline; only a hint.</li>
  </ul>
  <div class="callout warn"><svg class="i" aria-hidden="true"><use href="#i-triangle"/></svg><div>Automatic keyword classification: it gets irony, ambiguous headlines and stories whose impact depends on numbers wrong. Prices from Yahoo Finance (unofficial), from Kraken for Bitcoin if Yahoo fails, or else from the previous run (shown with its date); currency from Yahoo and the ECB; Bitcoin indicators from CoinGecko, Alternative.me and mempool.space. An educational tool, not financial advice.</div></div>
</section>
</main>
</div>
</div>
<div id="tip" role="tooltip" hidden></div>
<footer><div class="foot" id="foot"></div></footer>
<script id="dados" type="application/json">__DADOS_JSON__</script>
<script>
(function(){
'use strict';
/* dados embutidos: se não puderem ser lidos, a página diz porquê em vez de ficar em branco */
let D;try{D=JSON.parse(document.getElementById('dados').textContent);if(!D||typeof D!=='object')throw new Error('empty');}
catch(e){const m=document.getElementById('main');if(m)m.insertAdjacentHTML('afterbegin','<div class="callout warn"><div>The data embedded in this page could not be read ('+String(e&&e.message||e).replace(/[<>&"]/g,'')+'). Run the script again to rebuild it.</div></div>');return;}
const arr=x=>Array.isArray(x)?x:(x==null?[]:[x]);
const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>Array.from(r.querySelectorAll(s));
const nf=(v,d=1)=>(v==null||!isFinite(v))?'—':Number(v).toLocaleString('en-GB',{minimumFractionDigits:d,maximumFractionDigits:d}).replace(/^-(?=0(\.0+)?$)/,'').replace(/^-/,'−'); /* −0.0 arredondado passa a 0.0 */
const pct=(v,d=1)=>(v==null||!isFinite(v))?'—':(v>0?'+':'')+nf(v,d)+'%';
const cls=v=>v==null?'':(v>=0?'pos':'neg');
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const safe=u=>/^https?:\/\//i.test(u||'')?u:'#';
const GEN=Date.parse(D.geradoEm)||Date.now(),DAY=864e5;
const MES=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],WD=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
/* datas e horas sempre na hora de Lisboa, seja qual for o fuso do computador */
const LIS='Europe/Lisbon';
/* tzParts: um formatador por fuso, criado uma vez (criá-lo custa ~0,1 ms e a página formata milhares de datas ao abrir) */
const TZF={};
const fdt=t=>{if(t==null||!isFinite(t))return'—';const p=tzParts(t,LIS);return(+p.day)+' '+MES[+p.month-1]+' '+p.year;};  /* data inválida (ex.: 2026-13-45 num backup antigo): '—' em vez de parar a página */
const ic=(n,c)=>`<svg class="i${c?' '+c:''}" aria-hidden="true" focusable="false"><use href="#i-${n}"/></svg>`;
const CO={AAPL:{n:'Apple',c:'var(--aapl)',i:'apple'},NVDA:{n:'NVIDIA',c:'var(--nvda)',i:'cpu'},GOOGL:{n:'Alphabet',c:'var(--googl)',i:'search'},SXR8:{n:'ETF SXR8',c:'var(--spx)',i:'pie'},EUNK:{n:'ETF EUNK',c:'var(--eunk)',i:'pie'},IS3N:{n:'ETF IS3N',c:'var(--is3n)',i:'pie'},EUNN:{n:'ETF EUNN',c:'var(--eunn)',i:'pie'},BTC:{n:'Bitcoin',c:'var(--btc)',i:'bitcoin'},MKT:{n:'Market & macro',c:'var(--macro)',i:'landmark'},GSPC:{n:'S&P 500',c:'var(--macro)',i:'landmark'},TU:{n:'You',c:'var(--muted)',i:'user'}};
const LV={red:{n:'Material',o:4,c:'var(--red)',i:'octagon'},orange:{n:'Important',o:3,c:'var(--orange)',i:'triangle'},yellow:{n:'Moderate',o:2,c:'var(--yellow)',i:'circle-alert'},white:{n:'Noise',o:1,c:'var(--grey)',i:'circle-minus'}};
const AL=Object.assign({},LV,{white:Object.assign({},LV.white,{n:'Info',i:'info'})});
const TIER={primaria:'Primary source',referencia:'Leading press',outra:'Other source',cuidado:'Handle with care'};
const TABN={overview:'overview',portfolio:'portfolio',news:'news',prices:'prices',fundamentals:'fundamentals',fx:'currency',calendar:'calendar',etf:'ETF',bitcoin:'Bitcoin',sources:'sources'};
const SENT={positivo:'positive',negativo:'negative',misto:'mixed',neutro:'neutral'};
const coTag=(c,chip)=>CO[c]?`<span class="co${chip?' co-chip':''}" style="--c:${CO[c].c}">${esc(CO[c].n)}</span>`:'';
const S0={co:'all',per:'1A',cur:'EUR',lvl:'all',lim:60};const st=Object.assign({},S0);
const dias=n=>n+(n===1?' day':' days');
const ago=t=>{if(!t)return'no date';const m=(GEN-t)/6e4;if(m<60)return`${Math.max(1,Math.round(m))} min ago`;const h=m/60;if(h<24)return`${Math.round(h)} h ago`;return`${dias(Math.round(h/24))} ago`;};
/* o PowerShell 5.1 às vezes serializa um par [data, valor] como {value:[data, valor], Count:2}: aceita as duas formas */
const pair=p=>Array.isArray(p)?p:(p&&Array.isArray(p.value)?p.value:[]);
/* um valor em falta (null, '') fica de fora: +null daria 0, um preço de zero que pareceria uma queda de 100 % */
const toPts=a=>arr(a).map(pair).filter(p=>p[1]!=null&&p[1]!==''&&typeof p[1]!=='boolean').map(p=>[Date.parse(p[0]+'T00:00:00Z'),+p[1]]).filter(p=>isFinite(p[0])&&isFinite(p[1]));
const NEWS=arr(D.noticias).map(n=>{const o=arr(n.outras);return Object.assign({},n,{empresas:arr(n.empresas),temas:arr(n.temas),viaPosicao:arr(n.viaPosicao).map(String),outras:o,t:n.data?Date.parse(n.data):null,busca:[n.titulo,n.fonte].concat(o.map(x=>x.titulo+' '+x.fonte)).join(' ').toLowerCase()});});
const ATIVOS=arr(D.ativos).map(a=>Object.assign({},a,{pts:toPts(a.pontos)}));
const MERC=arr(D.mercado).map(a=>Object.assign({},a,{pts:toPts(a.pontos)}));
/* ETF: os da configuração do script ($ETFs → D.etfs; o SXR8 também em D.etf, o formato anterior). O SXR8 vem sempre
   primeiro e é o ETF por omissão. Um ETF novo na configuração aparece em todo o site sem mudar este código. */
const ETFD=Object.assign({},D.etf&&typeof D.etf==='object'?{SXR8:D.etf}:{},D.etfs&&typeof D.etfs==='object'&&!Array.isArray(D.etfs)?D.etfs:{});
const ETF_IDS=[...new Set(['SXR8'].concat(Object.keys(ETFD),ATIVOS.filter(a=>a.tipo==='ETF').map(a=>a.id)))].filter(id=>/^[A-Za-z0-9._-]{1,15}$/.test(id)&&(id==='SXR8'||ATIVOS.some(a=>a.id===id)));   /* ids simples: entram em atributos e chaves */
const ehEtf=id=>ETF_IDS.includes(id),etfInfo=id=>ETFD[id]&&typeof ETFD[id]==='object'?ETFD[id]:{};
const etfIdx=id=>etfInfo(id).indice||(id==='SXR8'?'S&P 500':id),etfAcum=id=>etfInfo(id).acumulacao!==false;
ETF_IDS.forEach(id=>{if(!CO[id])CO[id]={n:'ETF '+id,c:'var(--macro)',i:'pie'};});
/* EUR/USD para converter: a série de 1 ano do Yahoo; se falhar, o histórico longo (também do Yahoo, pedido à parte);
   só em último caso as taxas do BCE (90 dias). Nunca se usa uma taxa fora do período que a série cobre. */
const BCE=toPts(D.fx&&D.fx.bce),FXH=toPts(D.historico&&D.historico.FX);let FX=toPts(D.fx&&D.fx.yahoo),FXSRC=FX.length>=5?((D.fx&&D.fx.fonte)||'Yahoo Finance'):'';
if(FX.length<5){if(FXH.length>=5){FX=FXH.filter(p=>p[0]>=FXH[FXH.length-1][0]-400*DAY);FXSRC='long-term history';}else if(BCE.length){FX=BCE;FXSRC='ECB (90 days)';}else{FX=[];FXSRC='';}}
const matchCo=(list,co)=>co==='all'||list.includes(co)||(co==='MKT'&&(list.includes('SXR8')||list.includes('GSPC')));

/* ---------- finance helpers ---------- */
/* taxa do dia t (ou do dia útil anterior, até 5 dias antes); null fora do período coberto pela série */
const FXGAP=5*DAY;
function fxPt(t,S){S=S||FX;if(!S.length||t<S[0][0]-FXGAP)return null;let lo=0,hi=S.length-1;while(lo<hi){const m=(lo+hi+1)>>1;if(S[m][0]<=t)lo=m;else hi=m-1;}const p=S[lo];return t-p[0]>FXGAP?null:p;}
function fxAt(t,S){const p=fxPt(t,S);return p?p[1]:null;}  /* fxPt dá também a data da taxa usada */
/* converte uma série entre USD e EUR com a taxa de cada dia. Sem taxa para o último dia, ou numa moeda desconhecida,
   devolve [] (sem dados): mostrar dólares como euros, ou o preço de há dias como o último, seria pior do que nada */
function inCur(a,cur){const m=(a.moeda||'').toUpperCase();if(!a.pts.length||m===cur)return a.pts;
 const k=m==='USD'&&cur==='EUR'?-1:m==='EUR'&&cur==='USD'?1:0;if(!k)return[];
 const o=[];for(const p of a.pts){const f=fxAt(p[0]);if(f)o.push([p[0],k<0?p[1]/f:p[1]*f]);}
 return o.length&&o[o.length-1][0]===a.pts[a.pts.length-1][0]?o:[];}
/* ativos que negoceiam ao fim de semana (Bitcoin) contam dias de calendário: 1 semana = 7 pontos, 1 ano = 365; ações e câmbio contam sessões */
function stats(pts){if(!pts||pts.length<2)return null;const c=pts.map(p=>p[1]),n=c.length,last=c[n-1],back=k=>n>k?(last/c[n-1-k]-1)*100:null;
 const cont=pts.slice(-30).some(p=>{const w=new Date(p[0]).getUTCDay();return w===0||w===6;}),K=cont?{w:7,m:30,q:91,a:365}:{w:5,m:21,q:63,a:252};
 /* só há "1 ano", "desde o início do ano" e "52 semanas" quando a série cobre esse período (com a taxa do BCE de 90 dias,
    por exemplo, não cobre): sem isso o valor seria de um período mais curto com o nome errado */
 const span=pts[n-1][0]-pts[0][0],full=span>=350*DAY;
 const yr=new Date(pts[n-1][0]).getUTCFullYear(),yi=pts.findIndex(p=>new Date(p[0]).getUTCFullYear()===yr);const ytd=yi>0?(last/c[yi-1]-1)*100:null;
 const hi=full?Math.max(...c):null,lo=full?Math.min(...c):null,ma=k=>n>=k?c.slice(n-k).reduce((a,b)=>a+b,0)/k:null;
 const r=[];for(let i=Math.max(1,n-30);i<n;i++)r.push(Math.log(c[i]/c[i-1]));const mu=r.reduce((a,b)=>a+b,0)/(r.length||1);const vol=r.length>5?Math.sqrt(r.reduce((a,b)=>a+(b-mu)*(b-mu),0)/(r.length-1))*Math.sqrt(K.a)*100:null;
 let pk=c[0],mdd=0;for(const v of c){if(v>pk)pk=v;mdd=Math.min(mdd,(v/pk-1)*100);}
 return{last,cont,full,d1:back(1),w1:back(K.w),m1:back(K.m),m3:back(K.q),ytd,y1:full?(last/c[0]-1)*100:null,hi,lo,dist:full?(last/hi-1)*100:null,ma50:ma(50),ma200:ma(200),vol,mdd:full?mdd:null};}
const PER={'1M':31,'3M':92,'6M':183,'1A':372};
function slice(pts){if(!pts.length)return pts;const t1=pts[pts.length-1][0];return pts.filter(p=>p[0]>=t1-PER[st.per]*DAY);}

/* ---------- tooltip & charts ---------- */
const tip=$('#tip');
function showTip(t,x,y,rich){tip.classList.toggle('rich',!!rich);if(rich)tip.innerHTML=t;else tip.textContent=t;tip.hidden=false;const w=tip.offsetWidth,h=tip.offsetHeight;let L=x+14,T=y-h-12;if(L+w>innerWidth-8)L=x-w-14;if(L<8)L=8;if(T<8)T=y+18;tip.style.left=L+'px';tip.style.top=T+'px';}
function hideTip(){tip.hidden=true;}
document.addEventListener('pointermove',e=>{const t=e.target&&e.target.closest?e.target.closest('[data-tip]'):null;if(t)showTip(t.getAttribute('data-tip'),e.clientX,e.clientY);});
document.addEventListener('pointerout',e=>{const t=e.target&&e.target.closest?e.target.closest('[data-tip]'):null;if(t)hideTip();});
addEventListener('scroll',hideTip,{passive:true});
const NS='http://www.w3.org/2000/svg';
function S(tag,a,p){const e=document.createElementNS(NS,tag);for(const k in a)e.setAttribute(k,a[k]);if(p)p.appendChild(e);return e;}
function T(p,x,y,txt,a){const t=S('text',Object.assign({x,y,class:'svg-label'},a||{}),p);t.textContent=txt;return t;}
function ticks(mn,mx,n){const span=(mx-mn)||1,raw=span/n,mag=Math.pow(10,Math.floor(Math.log10(raw))),f=raw/mag,s=(f<1.5?1:f<3?2:f<7?5:10)*mag,lo=Math.floor(mn/s)*s,hi=Math.ceil(mx/s)*s,o=[];for(let v=lo;v<=hi+s/2;v+=s)o.push(+v.toFixed(8));return o;}
const cw=box=>Math.max(260,Math.round(box.clientWidth||640));
function legend(box,items,kind){const d=document.createElement('div');d.className='legend'+(kind?' '+kind:'');d.innerHTML=items.map(i=>`<span><i style="--c:${i.c}"></i>${esc(i.n)}${i.v!=null?` <b>${esc(i.v)}</b>`:''}</span>`).join('');box.appendChild(d);}
function empty(box,msg){box.innerHTML=`<p class="empty">${ic('info')}${esc(msg||'No data to show.')}</p>`;}
function lineMulti(box,series,o){if(!box)return;o=o||{};box.innerHTML='';series=series.filter(s=>s.pts&&s.pts.length>1);if(!series.length)return empty(box);
 const W=cw(box),H=Math.round(Math.min(o.h||300,Math.max(190,W*.6))),ml=46,mr=o.end?62:12,mt=12,mb=30,all=series.flatMap(s=>s.pts);const x0=Math.min(...all.map(p=>p[0])),x1=Math.max(...all.map(p=>p[0]));const LG=!!o.log,tv=v=>LG?Math.log10(Math.max(v,1e-9)):v;let y0=Math.min(...all.map(p=>tv(p[1]))),y1=Math.max(...all.map(p=>tv(p[1])));if(o.ref!=null){y0=Math.min(y0,o.ref);y1=Math.max(y1,o.ref);}const pad=(y1-y0)*.08||1;y0-=pad;y1+=pad;
 const svg=S('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':o.label||series.map(s=>s.n).join(', ')},box),X=v=>ml+(v-x0)/((x1-x0)||1)*(W-ml-mr),Y=v=>mt+(1-(v-y0)/(y1-y0))*(H-mt-mb),Yv=v=>Y(tv(v)),dec=o.dec==null?1:o.dec,tdec=o.tdec==null?dec:o.tdec;
 (LG?logTicks(y0,y1):ticks(y0,y1,4)).forEach(t=>{if(t<y0||t>y1)return;S('line',{x1:ml,x2:W-mr,y1:Y(t),y2:Y(t),class:'svg-grid'},svg);T(svg,ml-8,Y(t)+4,LG?curto(Math.pow(10,t)):nf(t,tdec),{'text-anchor':'end'});});
 S('line',{x1:ml,x2:W-mr,y1:H-mb,y2:H-mb,class:'svg-axis'},svg);
 const nx=W<520?2:4;for(let i=0;i<=nx;i++){const t=x0+(x1-x0)*i/nx,d=new Date(t);T(svg,X(t),H-mb+18,d.getUTCDate()+' '+MES[d.getUTCMonth()]+((x1-x0)>200*DAY?' '+String(d.getUTCFullYear()).slice(2):''),{'text-anchor':i===0?'start':i===nx?'end':'middle'});}
 if(o.ref!=null){S('line',{x1:ml,x2:W-mr,y1:Y(o.ref),y2:Y(o.ref),style:'stroke:var(--ref);stroke-width:1.2;stroke-dasharray:5 4'},svg);if(o.refLabel)T(svg,W-mr-4,Y(o.ref)-6,o.refLabel,{'text-anchor':'end',style:'fill:var(--ref);font-weight:600'});}
 arr(o.marks).forEach(m=>{if(m.t<x0||m.t>x1)return;S('line',{x1:X(m.t),x2:X(m.t),y1:mt,y2:H-mb,style:'stroke:var(--btc);stroke-width:1;stroke-dasharray:4 4;opacity:.6'},svg);T(svg,X(m.t)+5,mt+11,m.l,{style:'fill:var(--btc);font-weight:600'});});
 const path=s=>s.pts.map((p,i)=>(i?'L':'M')+X(p[0]).toFixed(1)+','+Yv(p[1]).toFixed(1)).join('');
 if(o.area&&series.length===1){const s=series[0],by=o.base!=null?Yv(o.base):H-mb;S('path',{d:path(s)+`L${X(s.pts[s.pts.length-1][0]).toFixed(1)},${by.toFixed(1)}L${X(s.pts[0][0]).toFixed(1)},${by.toFixed(1)}Z`,style:`fill:${s.c};opacity:.07`},svg);}
 series.forEach(s=>S('path',{d:path(s),class:'svg-line',style:`stroke:${s.c}`},svg));
 if(o.end){const E=series.map(s=>{const p=s.pts[s.pts.length-1];return{s,x:X(p[0]),py:Yv(p[1]),y:Yv(p[1]),v:p[1]};}).sort((a,b)=>a.y-b.y);
  for(let i=1;i<E.length;i++)if(E[i].y-E[i-1].y<14)E[i].y=E[i-1].y+14;const over=E.length?E[E.length-1].y-(H-mb):0;if(over>0)E.forEach(e=>e.y-=over);
  E.forEach(e=>{S('circle',{cx:e.x,cy:e.py,r:3.2,style:`fill:${e.s.c};stroke:var(--surface);stroke-width:1.5`},svg);const t=S('text',{x:W-mr+8,y:e.y+4,class:'svg-end',style:`fill:${e.s.c}`},svg);t.textContent=o.end(e.v);});}
 const g=S('line',{y1:mt,y2:H-mb,style:'stroke:var(--faint);stroke-dasharray:3 3;display:none'},svg),dots=series.map(s=>S('circle',{r:4.5,style:`fill:${s.c};stroke:var(--surface);stroke-width:2;display:none`},svg)),ov=S('rect',{x:ml,y:mt,width:W-ml-mr,height:H-mt-mb,style:'fill:transparent;cursor:crosshair'},svg);
 const near=(pts,x)=>{let b=pts[0];for(const p of pts)if(Math.abs(p[0]-x)<Math.abs(b[0]-x))b=p;return b;};
 function mv(ev){const r=svg.getBoundingClientRect(),xv=x0+((ev.clientX-r.left)/r.width*W-ml)/(W-ml-mr)*(x1-x0);let hd='';const R=[];series.forEach((s,i)=>{const b=near(s.pts,xv);dots[i].setAttribute('cx',X(b[0]));dots[i].setAttribute('cy',Yv(b[1]));dots[i].style.display='';R.push({n:s.n,c:s.c,v:b[1]});if(i===0){g.setAttribute('x1',X(b[0]));g.setAttribute('x2',X(b[0]));g.style.display='';hd=fdt(b[0]);}});
  R.sort((a,b)=>b.v-a.v);showTip(`<div class="tt-h">${esc(hd)}</div>`+R.map(x=>`<div class="tt-r"><i style="background:${x.c}"></i><span>${esc(x.n)}</span><b>${nf(x.v,dec)}${esc(o.unit||'')}</b></div>`).join(''),ev.clientX,ev.clientY,true);}
 /* marcadores de eventos (Prices): por cima da camada do cursor, para o rato os encontrar; o título vai em data-tip (texto) */
 arr(o.events).forEach(e=>{if(!(e.t>=x0&&e.t<=x1))return;S('circle',{cx:X(e.t),cy:H-mb-6-(e.row||0)*10,r:4.2,class:'ev-mark','data-k':e.k||'','data-tip':e.tip,style:`fill:${e.c};stroke:var(--surface);stroke-width:1.5`},svg);});
 ov.addEventListener('pointermove',mv);ov.addEventListener('pointerdown',mv);ov.addEventListener('pointerleave',()=>{g.style.display='none';dots.forEach(d=>d.style.display='none');hideTip();});
 if(series.length>1)legend(box,series.map(s=>({n:s.n,c:s.c,v:o.legendVal?o.legendVal(s):null})),'lines');}
function barChart(box,o){if(!box)return;box.innerHTML='';const vals=o.series.flatMap(s=>s.values).filter(v=>v!=null&&isFinite(v));if(!vals.length)return empty(box);const W=cw(box),H=Math.round(Math.min(o.h||250,Math.max(190,W*.6))),ml=46,mr=12,mt=20,mb=36;let mn=Math.min(0,...vals),mx=Math.max(0,...vals);const tk=ticks(mn<0?mn*1.12:0,mx>0?mx*1.12:0.1,4),y0=tk[0],y1=tk[tk.length-1];
 const svg=S('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':o.label||''},box),pw=W-ml-mr,ph=H-mt-mb,Y=v=>mt+ph-(v-y0)/((y1-y0)||1)*ph,dec=o.dec==null?1:o.dec,u=o.unit||'';
 tk.forEach(t=>{S('line',{x1:ml,x2:W-mr,y1:Y(t),y2:Y(t),class:'svg-grid'},svg);T(svg,ml-8,Y(t)+4,nf(t,o.tdec||0)+u,{'text-anchor':'end'});});
 S('line',{x1:ml,x2:W-mr,y1:Y(0),y2:Y(0),class:'svg-axis'},svg);const gw=pw/o.cats.length,ns=o.series.length,bw=Math.min(54,gw*.72/ns);
 o.cats.forEach((c,i)=>{const gx=ml+gw*i+gw/2;T(svg,gx,H-mb+20,c,{'text-anchor':'middle',style:'fill:var(--text);font-weight:600'});o.series.forEach((s,j)=>{const v=s.values[i];if(v==null||!isFinite(v))return;const x=gx-ns*bw/2+j*bw,y=Math.min(Y(v),Y(0)),h=Math.max(1,Math.abs(Y(v)-Y(0)));S('rect',{x:x+1.5,y,width:bw-3,height:h,rx:3,style:`fill:${(s.colors&&s.colors[i])||s.c}`,'data-tip':`${s.n?s.n+', ':''}${c}: ${nf(v,dec)}${u}`},svg);const t=S('text',{x:x+bw/2,y:v>=0?y-5:y+h+13,'text-anchor':'middle',class:'svg-val'},svg);t.textContent=nf(v,dec)+u;});});
 if(ns>1)legend(box,o.series.map(s=>({n:s.n,c:s.c})));}
function hbars(box,rows,label){if(!box)return;box.innerHTML='';if(!rows.length)return empty(box);const lbl=r=>String(r.t!=null?r.t:r.v),W=cw(box),rh=28,mt=6,ml=Math.round(Math.min(190,W*.42)),mr=Math.max(50,Math.max(...rows.map(r=>lbl(r).length))*7+12),H=mt+rows.length*rh+6,mx=Math.max(...rows.map(r=>r.v))||1;const svg=S('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':label||'Themes'},box),X=v=>ml+v/mx*(W-ml-mr);
 rows.forEach((r,i)=>{const y=mt+i*rh;T(svg,ml-10,y+rh/2+4,r.l,{'text-anchor':'end',style:'fill:var(--text)'});S('rect',{x:ml,y:y+6,width:W-ml-mr,height:rh-12,rx:3,style:'fill:var(--surface-2)'},svg);S('rect',{x:ml,y:y+6,width:Math.max(2,X(r.v)-ml),height:rh-12,rx:3,style:`fill:${r.c||'var(--bar)'}`,'data-tip':`${r.l}: ${lbl(r)}`},svg);const t=S('text',{x:X(r.v)+6,y:y+rh/2+4,class:'svg-val'},svg);t.textContent=lbl(r);});}
/* escala logarítmica: linhas em 1, 2 e 5 de cada potência de 10 (só as potências quando são muitas) */
function logTicks(a,b){const o=[];for(let e=Math.floor(a);e<=Math.ceil(b);e++)for(const m of [1,2,5]){const v=e+Math.log10(m);if(v>=a&&v<=b)o.push(v);}return o.length>9?o.filter(v=>Math.abs(v-Math.round(v))<1e-9):o;}
const curto=v=>{const g=(x,s)=>nf(x,Math.abs(x-Math.round(x))<1e-9?0:1)+s;return v>=1e6?g(v/1e6,'M'):v>=1e3?g(v/1e3,'k'):nf(v,0);};

/* ---------- frescura dos preços: data do último preço e sessões em falta ---------- */
/* Cada série é comparada com a última sessão já terminada da sua bolsa quando o site foi gerado (fins de semana,
   feriados e sessões curtas pelas regras de cada bolsa). Assim um preço antigo nunca passa por atual. */
const BOLSA_DE={AAPL:'US',NVDA:'US',GOOGL:'US',GSPC:'US',VIX:'US',TNX:'US',SXR8:'DE'};
ATIVOS.forEach(a=>{if(!BOLSA_DE[a.id]&&a.bolsa)BOLSA_DE[a.id]=String(a.bolsa);});   /* a bolsa de cada ETF vem do script ($ETFs) */
const FXB={id:'FX',tz:'Europe/London',abre:'00:00',fecha:'22:00',feriados:[],curtos:{}}; /* EUR/USD: dias úteis, fecho de Londres */
const fdS=iso=>{const q=String(iso).split('-');return(+q[2])+' '+MES[+q[1]-1];};
function sessaoFechada(b,t){const p=tzParts(t,b.tz);for(let i=0;i<20;i++){const dt=Date.UTC(+p.year,+p.month-1,+p.day-i),iso=isoU(dt),y=+iso.slice(0,4);if(wdU(dt)===0||wdU(dt)===6||ehFeriado(b,iso,y))continue;const q=iso.split('-').map(Number),f=zoned(q[0],q[1],q[2],fechoDe(b,iso,y),b.tz);if(f<=t)return{iso,f};}return null;}
function sessoesDepois(b,isoA,isoB){let n=0;for(let t=Date.parse(isoA+'T00:00:00Z')+DAY,e=Date.parse(isoB+'T00:00:00Z');t<=e;t+=DAY){const iso=isoU(t);if(wdU(t)===0||wdU(t)===6||ehFeriado(b,iso,+iso.slice(0,4)))continue;n++;}return n;}
function frescura(id,pts){if(!pts||!pts.length)return null;const last=pts[pts.length-1][0],iso=isoU(last);
 if(id==='BTC'){const falta=Math.max(0,Math.round((Date.parse(isoU(GEN)+'T00:00:00Z')-DAY-last)/DAY));return{iso,falta,velho:falta>=1,un:'day'};}
 const b=id==='FX'?FXB:arr(D.bolsas).find(x=>x.id===BOLSA_DE[id]);if(!b)return{iso,falta:0,velho:false,un:'session'};
 const e=sessaoFechada(b,GEN);if(!e||iso>=e.iso)return{iso,falta:0,velho:false,un:'session'};
 const falta=sessoesDepois(b,iso,e.iso);
 /* uma sessão em falta é normal logo a seguir ao fecho; ao fim de 6 h (ou 2 sessões) já é atraso da fonte. No câmbio,
    que tem dias sem cotação no Natal e no Ano Novo, só 2 ou mais. */
 return{iso,falta,velho:falta>=2||(id!=='FX'&&falta===1&&GEN-e.f>6*36e5),un:'session'};}
const antigo=a=>/previous run/i.test(a.fonte||'');
function asOf(a,fr){if(!fr)return'';const via=antigo(a)?' · previous run':(a.fonte&&a.fonte!=='Yahoo Finance'?' · via '+a.fonte:'');
 if(a.parcial&&a.hora&&isFinite(Date.parse(a.hora)))return`live ${hmL(Date.parse(a.hora))}`+via;return(a.id==='BTC'?'day ':'close ')+fdS(fr.iso)+via;}

/* ---------- hero ---------- */
function guilloche(){let d='';for(let k=0;k<24;k++){let p='';const A=16+k*2,ph=k*.41,y0=30+k*12;for(let x=0;x<=1200;x+=10){const y=y0+A*Math.sin(x/96+ph)+A*.45*Math.sin(x/38-ph*1.7);p+=(x?'L':'M')+x+','+y.toFixed(1);}d+=`<path d="${p}" style="fill:none;stroke:var(--guilloche);stroke-width:.8"/>`;}$('#guilloche').innerHTML=d;}
/* lugar de cada cartão no panorama: o S&P 500 (SXR8) em cima e a Bitcoin em baixo, a toda a largura; os outros em três colunas */
const HERO_LUGAR={SXR8:' tk-lead',BTC:' tk-wide',FX:' tk-wide'};
function heroColor(a){return CO[a.id]?CO[a.id].c:'var(--macro)';}
function spark(pts,c){const p=pts.length?pts.filter(x=>x[0]>=pts[pts.length-1][0]-92*DAY):pts; /* 3 meses de calendário: ~64 sessões nas ações, ~92 dias na Bitcoin */if(p.length<2)return'';const v=p.map(x=>x[1]),mn=Math.min(...v),mx=Math.max(...v),W=200,H=40,X=i=>i/(p.length-1)*W,Y=y=>H-3-(y-mn)/((mx-mn)||1)*(H-6);
 const d=p.map((x,i)=>(i?'L':'M')+X(i).toFixed(1)+','+Y(x[1]).toFixed(1)).join('');
 return`<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true" focusable="false"><path d="${d}L${W},${H}L0,${H}Z" style="fill:${c};opacity:.08"/><path d="${d}" style="fill:none;stroke:${c};stroke-width:1.7;stroke-linejoin:round;vector-effect:non-scaling-stroke"/></svg>`;}
function renderHero(){const fx=FX.length?FX[FX.length-1][1]:null;
 $('#heroSub').innerHTML=`<span>${ic('clock')}Generated on <b>${fdt(GEN)}</b> at <b>${hmL(GEN)}</b> Lisbon time</span><span>${ic('newspaper')}News from the last ${esc(D.dias)} days · 1 year of prices</span><span>${ic('euro')}EUR/USD <b>${fx?nf(fx,4):'unavailable'}</b></span>`;
 $('#tickers').innerHTML=ATIVOS.map(a=>{const s=stats(a.pts),m=(a.moeda||'').toUpperCase(),sym=m==='EUR'?'€':'$',c=heroColor(a),nm=ehEtf(a.id)?'ETF '+etfIdx(a.id):a.nome,fr=frescura(a.id,a.pts),ao=asOf(a,fr);
  const head=`<span class="tk-badge">${ic((CO[a.id]||{}).i||'chart')}</span><div class="tk-id"><span class="tk-name">${ehEtf(a.id)?'<span class="lg">ETF </span>'+esc(etfIdx(a.id)):esc(nm)}</span><span class="tk-sym">${esc(a.simbolo)}${m?'<span class="lg"> · '+esc(m)+'</span>':''}${ao?` · <span class="asof${fr&&(fr.velho||antigo(a))?' old':''}" title="${fr&&fr.velho?`${fr.falta} ${fr.un}${fr.falta>1?'s':''} behind: the source may be lagging`:'Date of the latest price'}">${esc(ao)}</span>`:''}</span></div>`;
  if(!s)return`<article class="tk${HERO_LUGAR[a.id]||''}" data-tk="${a.id}" style="--c:${c}">${head}<p class="tk-empty">No prices. See the Sources section.</p></article>`;
  const btc=a.id==='BTC',bm=btc&&D.bitcoin&&D.bitcoin.mercado,v24=bm&&bm.var24!=null&&isFinite(bm.var24)?+bm.var24:null,chg=v24!=null?v24:s.d1,chgT=v24!=null?'over the last 24 hours':btc?'since 00:00 UTC':a.parcial?'so far today (session open)':'in the last session';
  const fxe=lastFxPara(a),eur=btc?`${fx?`≈ $${nf(s.last*fx,0)} · `:''}market open 24/7`:m==='USD'?(fxe?`≈ €${nf(s.last/fxe,2)} at the rate of that day`:'No EUR/USD rate for that day'):'Priced in euros (Xetra)';
  const range=s.hi!=null?(()=>{const rp=Math.max(0,Math.min(100,(s.last-s.lo)/((s.hi-s.lo)||1)*100));return`<div class="range" data-tip="52-week range: ${sym}${nf(s.lo,2)} – ${sym}${nf(s.hi,2)}"><div class="range-bar"><span class="range-fill" style="width:${rp}%"></span><span class="range-mk" style="left:${rp}%"></span></div><div class="range-lbl"><span>${sym}${nf(s.lo,2)}</span><span>52-week low and high</span><span>${sym}${nf(s.hi,2)}</span></div></div>`;})():'';
  return`<article class="tk${HERO_LUGAR[a.id]||''}" data-tk="${a.id}" style="--c:${c}" aria-label="${esc(nm)}">${head}<span class="chg ${cls(chg)}" title="Change ${chgT}">${ic(chg!=null&&chg<0?'trend-down':'trend-up')}${pct(chg)}<span class="sr"> ${chgT}</span></span>`+
   `<div class="tk-price"><span class="cur">${sym}</span>${nf(s.last,2)}</div><div class="tk-eur${!btc&&m!=='USD'?' nota':''}">${eur}</div><div class="tk-spark">${spark(a.pts,c)}<span class="spark-l">3 months</span></div>`+
   `<dl class="tk-rows"><div><dt>1 month</dt><dd class="${cls(s.m1)}">${pct(s.m1)}</dd></div><div><dt title="Distance from the 52-week high">From <span class="lg">52-week </span>high</dt><dd>${pct(s.dist)}</dd></div></dl>`+range+`</article>`;}).join('')+cartaoFx();}
/* EUR/USD no fim do panorama: a mesma série usada nas conversões (FX, com as alternativas e a data de cada taxa).
   O gráfico completo continua em Currency & Macroeconomics. */
function cartaoFx(){const a={id:'FX',nome:'EUR/USD',simbolo:'EURUSD=X',pts:FX,parcial:!!(D.fx&&D.fx.parcial),hora:null,fonte:/^Yahoo/.test(FXSRC)?'Yahoo Finance':FXSRC},s=stats(FX),c='var(--ink)',fr=FX.length?frescura('FX',FX):null,ao=asOf(a,fr),bl=BCE.length?BCE[BCE.length-1]:null;
 const head=`<span class="tk-badge">${ic('euro')}</span><div class="tk-id"><span class="tk-name">EUR/USD</span><span class="tk-sym">EURUSD=X${ao?` · <span class="asof${fr&&(fr.velho||antigo(a))?' old':''}" title="${fr&&fr.velho?`${fr.falta} ${fr.un}${fr.falta>1?'s':''} behind: the source may be lagging`:'Date of the latest rate'}">${esc(ao)}</span>`:''}</span></div>`;
 if(!s)return`<article class="tk${HERO_LUGAR.FX}" data-tk="FX" style="--c:${c}">${head}<p class="tk-empty">No EUR/USD rate in this run: see the Sources section.</p></article>`;
 const range=s.hi!=null?(()=>{const rp=Math.max(0,Math.min(100,(s.last-s.lo)/((s.hi-s.lo)||1)*100));return`<div class="range" data-tip="52-week range: ${nf(s.lo,4)} – ${nf(s.hi,4)}"><div class="range-bar"><span class="range-fill" style="width:${rp}%"></span><span class="range-mk" style="left:${rp}%"></span></div><div class="range-lbl"><span>${nf(s.lo,4)}</span><span>52-week low and high</span><span>${nf(s.hi,4)}</span></div></div>`;})():'';
 return`<article class="tk${HERO_LUGAR.FX}" data-tk="FX" style="--c:${c}" aria-label="EUR/USD"><span class="chg ${cls(s.d1)}" title="Change in the last session">${ic(s.d1!=null&&s.d1<0?'trend-down':'trend-up')}${pct(s.d1)}<span class="sr"> in the last session</span></span>`+head+
  `<div class="tk-price">${nf(s.last,4)}</div><div class="tk-eur">dollars per euro${bl?` · ECB reference ${nf(bl[1],4)}`:''}</div><div class="tk-spark">${spark(FX,c)}<span class="spark-l">3 months</span></div>`+
  `<dl class="tk-rows"><div><dt>1 month</dt><dd class="${cls(s.m1)}">${pct(s.m1)}</dd></div><div><dt>Year to date</dt><dd class="${cls(s.ytd)}">${pct(s.ytd)}</dd></div></dl>`+range+`</article>`;}
/* taxa EUR/USD do dia do último preço de um ativo (null se a série de câmbio não chegar a esse dia) */
function lastFxPara(a){return a.pts.length?fxAt(a.pts[a.pts.length-1][0]):null;}

/* ---------- alerts ---------- */
/* datas do calendário: dia civil (meia-noite UTC dessa data), comparado com o dia de hoje em Lisboa */
function calendar(){return arr(D.calendario).filter(e=>/^\d{4}-\d{2}-\d{2}$/.test(e.d||'')).map(e=>Object.assign({},e,{t:Date.parse(e.d+'T00:00:00Z')})).sort((a,b)=>a.t-b.t);}
const diasAte=t=>Math.round((t-hojeL())/DAY);
/* classes do índice Fear & Greed (Alternative.me): 0–24 medo extremo, 25–46 medo, 47–54 neutro, 55–75 ganância, 76–100 ganância extrema */
const fgClasse=v=>v<=24?'xf':v<=46?'f':v<=54?'n':v<=75?'g':'xg';
function alerts(){const out=[],add=(l,co,txt,tab)=>out.push({l,co,txt,tab});
 ATIVOS.forEach(a=>{const s=stats(a.pts),nm=ehEtf(a.id)?`The ${a.id} ETF`:a.nome;if(!s){add('yellow',a.id,`No prices for ${a.nome}: see the Sources section.`,'sources');return;}
  const fr=frescura(a.id,a.pts);
  if(antigo(a))add('orange',a.id,`${nm}: prices could not be downloaded in this run, so the previous run's data is shown (latest price from ${fdS(fr.iso)}). The figures for ${a.nome} are not current.`,'sources');
  else if(fr&&fr.velho)add('orange',a.id,`${nm}: the latest price is from ${fdS(fr.iso)}, ${fr.falta} ${fr.un}${fr.falta>1?'s':''} behind. The source may be lagging, so the figures for ${a.nome} are not current.`,'sources');
  /* a Bitcoin é várias vezes mais volátil do que as ações: os limiares são mais largos para não alertar todos os dias */
  const btc=a.id==='BTC',L=btc?{d:6,o:-30,y:-15}:{d:4,o:-20,y:-10},bm=btc&&D.bitcoin&&D.bitcoin.mercado,v24=bm&&bm.var24!=null&&isFinite(bm.var24)?+bm.var24:null;
  /* dados atrasados ou da execução anterior: o movimento não é de hoje, e a mensagem diz de quando é (a variação de 24 h da
     CoinGecko é sempre atual) */
  const velho=antigo(a)||!!(fr&&fr.velho),mv=btc&&v24!=null?v24:s.d1,sess=btc&&v24!=null?'over the last 24 hours':velho&&fr?`on ${fdS(fr.iso)} (the latest price available, not today)`:btc?'since 00:00 UTC':(a.parcial?'so far today':'in the last session');
  if(mv!=null&&Math.abs(mv)>=L.d)add('orange',a.id,`${nm} ${mv>0?'rose':'fell'} ${nf(Math.abs(mv))}% ${sess}. Look for the cause in ${velho&&!(btc&&v24!=null)?'the news of that day':"today's news"}.`,'news');
  if(s.dist!=null){
   if(s.dist<=L.o)add('orange',a.id,btc?`${nm} is ${nf(Math.abs(s.dist))}% below its 52-week high. Look for the cause (regulation, ETF flows, rates) in the news.`:`${nm} is ${nf(Math.abs(s.dist))}% below its 52-week high. Check whether expected earnings have also fallen (volatility or deterioration?).`,btc?'bitcoin':'prices');
   else if(s.dist<=L.y)add('yellow',a.id,`${nm} is ${nf(Math.abs(s.dist))}% below its 52-week high.`,btc?'bitcoin':'prices');
   if(s.dist<=L.y)out[out.length-1].ctx={id:a.id,d:-s.dist};  /* contexto (regra do utilizador e histórico), desenhado em renderAlerts */
   else if(s.dist>-1.5)add('white',a.id,`${nm} is close to its 52-week high.`,btc?'bitcoin':'prices');}
  if(s.ma200!=null&&s.last<s.ma200)add('yellow',a.id,`${nm} is below its ${btc?'200-day':'200-session'} average.`,btc?'bitcoin':'prices');});
 const fg=D.bitcoin&&D.bitcoin.sentimento;if(fg&&fg.valor!=null&&isFinite(fg.valor)){const k=fgClasse(+fg.valor);if(k==='xf')add('yellow','BTC',`Bitcoin sentiment is at extreme fear (${fg.valor}/100): this tends to coincide with sharp drops and high volatility.`,'bitcoin');else if(k==='xg')add('yellow','BTC',`Bitcoin sentiment is at extreme greed (${fg.valor}/100): euphoria tends to come with abrupt corrections.`,'bitcoin');}
 /* câmbio: sem taxa não há valores em euros das ações americanas; com uma alternativa, diz qual */
 if(!FX.length)add('red','MKT','No EUR/USD rate in this run: the euro values of the US stocks, and of your portfolio, cannot be calculated and are left blank.','sources');
 else if(!/^Yahoo/.test(FXSRC))add('orange','MKT',`The EUR/USD series from Yahoo failed: conversions use the ${FXSRC}${FXSRC.startsWith('ECB')?', so euro figures older than 90 days (1 year, year to date, 52-week high) are not shown':''}.`,'sources');
 else{const ff=frescura('FX',FX);if(ff&&ff.velho)add('orange','MKT',`The latest EUR/USD rate is from ${fdS(ff.iso)}, ${ff.falta} sessions behind: euro conversions may not be current.`,'sources');}
 const f=stats(FX);if(f&&f.m1!=null&&Math.abs(f.m1)>=2)add('yellow','MKT',`EUR/USD ${f.m1>0?'rose':'fell'} ${nf(Math.abs(f.m1))}% in a month: the dollar ${f.m1>0?'weakened, which reduces':'strengthened, which boosts'} the euro return on US stocks.`,'fx');
 const vix=MERC.find(m=>m.id==='VIX'),vs=vix?stats(vix.pts):null;if(vs){if(vs.last>=25)add('orange','MKT',`VIX at ${nf(vs.last)}: high volatility in the US market.`,'fx');else if(vs.last>=20)add('yellow','MKT',`VIX at ${nf(vs.last)}: some market stress.`,'fx');}
 calendar().forEach(e=>{const dd=diasAte(e.t);if(dd>=0&&dd<=7&&/high/i.test(e.imp))add('orange',e.e,`${dd===0?'Today':dd===1?'Tomorrow':`In ${dd} days`} (${fdt(e.t)}): ${e.ev}.`,'calendar');});
 ['AAPL','NVDA','GOOGL','BTC','MKT'].concat(ETF_IDS.filter(x=>x!=='SXR8')).forEach(c=>{const r=NEWS.filter(n=>n.nivel==='red'&&n.t&&GEN-n.t<=3*DAY&&matchCo(n.empresas,c));if(r.length)add('red',c,`${r.length} ${r.length>1?'stories':'story'} classed as material in the last 72 h. The most relevant: “${r.sort((a,b)=>b.score-a.score)[0].titulo}”.`,'news');});
 return out.sort((a,b)=>LV[b.l].o-LV[a.l].o);}
const alFilter=a=>st.co==='all'||a.co===st.co||(st.co==='MKT'&&(a.co==='SXR8'||a.co==='MKT'));
/* duração legível (a mesma escala da tabela das maiores quedas) */
const durQ=ms=>{const m=Math.round(ms/(30.44*DAY));return m<1?'under 1 month':m<24?m+(m===1?' month':' months'):nf(m/12,1)+' years';};
/* Linha de contexto de um alerta de queda: a regra da política do utilizador para essa profundidade (30 %, ou senão 20 %)
   e o histórico do ativo (HIST, na moeda do preço): quantas quedas desde um pico iguais ou maiores do que a atual houve
   (quedas()), em quanto tempo recuperaram (mediana, do fundo de volta ao pico) e quantas ainda não recuperaram. Devolve HTML
   com o texto do utilizador escapado. */
function contextoQueda(c){const h=HIST[c.id]||[],P=limpaPolitica(store.get('policy',null)),Dp=Math.floor(c.d*10+1e-9)/10,nm=nomePF(c.id),out=[];
 const r=P&&Dp>=30&&P.drop30.trim()?['30',P.drop30]:P&&Dp>=20&&P.drop20.trim()?['20',P.drop20]:null;
 if(r){const tx=r[1].trim();out.push(`Your rule for a ${r[0]}% drop: “${esc(tx)}”${/[.!?]$/.test(tx)?'':'.'}`);}
 if(h.length<2)out.push('No long-term price history in this run to compare with.');
 else{const Q=quedas(h,Dp),rec=Q.filter(q=>q.rec),nr=Q.length-rec.length,ult=Q.length>0&&!Q[Q.length-1].rec,per=`between ${fdt(h[0][0])} and ${fdt(h[h.length-1][0])}`;
  if(!Q.length)out.push(`In its history ${per}, ${esc(nm)} never fell ${nf(Dp,1)}% or more from a peak.`);
  else{const T=rec.map(q=>q.rec[0]-q.tr[0]).sort((a,b)=>a-b),med=T.length?(T.length%2?T[(T.length-1)/2]:(T[T.length/2-1]+T[T.length/2])/2):null;
   out.push(`In its history ${per}, ${esc(nm)} fell ${nf(Dp,1)}% or more from a peak ${Q.length} time${Q.length===1?'':'s'}: ${rec.length} recovered${med!=null?`, in a median of ${durQ(med)} from the bottom back to the peak`:''}${nr?`; ${nr} ${nr===1?'has':'have'} not recovered yet${ult?' (including the current one)':''}`:''}.`);}
  out.push(`This history (since ${new Date(h[0][0]).getUTCFullYear()}) is mostly a rising market, so past recoveries may have been quicker than in a long bear market.`);}
 return out.join(' ');}
function renderAlerts(){const al=alerts().filter(alFilter);
 $('#cnt-alertas').textContent=al.filter(a=>a.l!=='white').length||'';
 $('#alertas').innerHTML=al.length?al.map(a=>{const L=AL[a.l];return`<div class="alert l-${a.l}"><span class="sev-ic" title="${L.n}">${ic(L.i)}</span><div><div class="alert-top"><span class="sev-txt">${L.n}</span>${coTag(a.co,true)}</div><p class="alert-msg">${esc(a.txt)}</p>${a.ctx?`<p class="alert-ctx">${ic('book')}<span>${contextoQueda(a.ctx)}</span></p>`:''}</div><button class="btn" data-goto="${a.tab}" type="button" aria-label="See in the ${esc(TABN[a.tab]||a.tab)} section">See ${esc(TABN[a.tab]||'')}${ic('arrow-right','i-go')}</button></div>`;}).join(''):`<p class="empty">${ic('check')}No alerts for this filter.</p>`;}
function renderStats(){const al=alerts().filter(alFilter),N=NEWS.filter(n=>matchCo(n.empresas,st.co)),F=arr(D.fontes),ok=F.filter(x=>/^ok/.test(x.estado)).length;
 const k=(i,lab,v,s,c)=>`<div class="kpi ${c||''}"><div class="k">${ic(i)}${lab}</div><div class="v">${v}</div><div class="s">${s}</div></div>`,hi=al.filter(a=>a.l==='red'||a.l==='orange').length;
 $('#resumoStats').innerHTML=k('bell','Alerts',al.filter(a=>a.l!=='white').length,hi?`${hi} of high severity`:'none of high severity',hi?'k-orange':'')+
  k('octagon','Material',N.filter(n=>n.nivel==='red').length,'stories in the window','k-red')+k('triangle','Important',N.filter(n=>n.nivel==='orange').length,'stories in the window','k-orange')+
  k('database','Sources',`${ok}/${F.length}`,ok===F.length?'all responded':`${F.length-ok} with problems`);}

/* ---------- news ---------- */
function newsItem(n,rel){const lv=LV[n.nivel]||LV.white,lc=LV[n.nivel]?n.nivel:'white',cos=n.empresas.filter(c=>c!=='MKT'||n.empresas.length===1);
 const tc=n.tier==='cuidado'?'t-warn':n.tier==='primaria'?'t-pos':'',sc=n.sentimento==='positivo'?['trend-up','s-pos']:n.sentimento==='negativo'?['trend-down','s-neg']:['minus',''];
 return`<article class="news l-${lc}"><span class="sev-ic" title="${lv.n}">${ic(lv.i)}</span><div><div class="news-line">${cos.map(c=>coTag(c,true)).join('')}<a class="news-t" href="${esc(safe(n.link))}" target="_blank" rel="noopener noreferrer">${esc(n.titulo)}${ic('external','ext')}<span class="sr"> (opens in a new window)</span></a></div>`+
  `<div class="meta"><span class="m-sev">${lv.n}</span><span>${ic('rss')}${esc(n.fonte||n.dominio)}</span>${TIER[n.tier]?`<span class="tag ${tc}">${TIER[n.tier]}</span>`:''}<span title="${n.t?fdt(n.t):''}">${ic('clock')}${ago(n.t)}</span>${n.temas.length?`<span>${ic('tag')}${esc(n.temas.join(', '))}</span>`:''}<span class="${sc[1]}">${ic(sc[0])}Sentiment ${esc(SENT[n.sentimento]||n.sentimento)}</span><span class="m-score">${nf(n.score)} points</span>${n.soFeed?`<span class="tag" title="No keyword for the asset in the headline: it came from that asset's search feed, so it is at most moderate">Feed match only</span>`:''}${n.viaPosicao&&n.viaPosicao.length?`<span class="tag" title="Matched only through a top holding of the fund (${esc(n.viaPosicao.join(', '))}), so it is at most moderate">Via top holding</span>`:''}${rel!=null?`<span class="m-score" title="Relevance to me: points + 3 × the share of your portfolio exposed to this story's asset">relevance ${nf(rel)}</span>`:''}${n.novo?`<span class="tag t-pos">${ic('sparkle')}New</span>`:''}</div>${dupList(n)}</div></article>`;}
function dupList(n){if(!n.outras.length)return'';const N=n.outras.length;
 return`<details class="dups"><summary>${ic('layers')}Also reported by ${N} other source${N>1?'s':''}<span class="dups-src">${esc(n.outras.slice(0,4).map(o=>o.fonte).join(' · '))}${N>4?' …':''}</span>${ic('chevron-down','chev')}</summary><ul>${n.outras.map(o=>`<li><a href="${esc(safe(o.link))}" target="_blank" rel="noopener noreferrer">${esc(o.titulo)}${ic('external','ext')}<span class="sr"> (opens in a new window)</span></a><span>${esc(o.fonte)}${TIER[o.tier]?' · '+TIER[o.tier]:''} · ${ago(o.data?Date.parse(o.data):null)}</span></li>`).join('')}</ul></details>`;}
/* "Relevance to me": fração da carteira exposta ao ativo de cada notícia, calculada aqui a partir de pfDados() (carteira(),
   o único FIFO). Ações: direto + através de cada ETF (pesos etfInfo, o mesmo look-through de "Real exposure"); ETF e
   Bitcoin: o seu valor; Market: ações e ETF (tudo menos a Bitcoin). Sem posições com valor: null (a opção fica desativada).
   A pontuação e o nível vêm do script e não mudam: a relevância só ordena. */
function expoNoticias(R){R=R||pfDados();const tv=R.reduce((s,x)=>s+(x.v||0),0);if(!(tv>0))return null;const v=id=>(R.find(x=>x.id===id)||{}).v||0,F={};
 R.forEach(x=>{F[x.id]=(x.v||0)/tv;});
 ['AAPL','NVDA','GOOGL'].forEach(k=>{F[k]=(v(k)+ETF_IDS.reduce((s,id)=>s+v(id)*(+etfInfo(id)[k]||0)/100,0))/tv;});
 F.MKT=(tv-v('BTC'))/tv;return F;}
/* relevância = pontuação + 3 × fração exposta (soma dos ativos da notícia, no máximo 1; Market só quando é o único, como nas etiquetas) */
const relevancia=(n,F)=>n.score+3*Math.min(1,n.empresas.filter(c=>c!=='MKT'||n.empresas.length===1).reduce((s,c)=>s+(F&&F[c]||0),0));
function filteredNews(R){const tema=$('#f-tema').value,tier=$('#f-tier').value,dias=+$('#f-dias').value||0,q=$('#f-q').value.trim().toLowerCase(),nov=$('#f-novas').checked,ord=$('#f-ord').value;
 return NEWS.filter(n=>matchCo(n.empresas,st.co)&&(st.lvl==='all'||n.nivel===st.lvl)&&(!tema||n.temas.includes(tema))&&(!tier||n.tier===tier)&&(!dias||(n.t&&GEN-n.t<=dias*DAY))&&(!q||n.busca.includes(q))&&(!nov||n.novo))
  .sort((a,b)=>ord==='data'?((b.t||0)-(a.t||0)):ord==='rel'&&R?(R.get(b)-R.get(a)||b.score-a.score||(b.t||0)-(a.t||0)):(b.score-a.score||(b.t||0)-(a.t||0)));}
function renderNews(){const F=expoNoticias(),op=$('#f-ord option[value="rel"]'),rn=$('#newsRelNote');
 if(op){op.disabled=!F;if(!F&&$('#f-ord').value==='rel')$('#f-ord').value='score';}
 const R=F&&$('#f-ord').value==='rel'?new Map(NEWS.map(n=>[n,relevancia(n,F)])):null;
 if(rn){rn.hidden=!(!F||R);rn.textContent=!F?'Relevance to me needs holdings: add your purchases in Portfolio to sort the stories by your exposure.':'Relevance to me = the potential impact points + 3 × the share of your portfolio exposed to the story\'s asset (directly or through the ETFs; Market = your stocks and ETFs). The levels and points do not change.';}
 const f=filteredNews(R);$('#cnt-noticias').textContent=NEWS.filter(n=>matchCo(n.empresas,st.co)&&(n.nivel==='red'||n.nivel==='orange')).length||'';
 $('#newsCount').textContent=`${f.length} stories match these filters, out of ${NEWS.length} collected${D.duplicadas?` (${D.duplicadas} duplicates grouped)`:''}.${D.primeiraExecucao?' This is the first run: marking new stories starts with the next one.':''}`;
 $('#newsList').innerHTML=f.length?f.slice(0,st.lim).map(n=>newsItem(n,R?R.get(n):null)).join(''):`<p class="empty">${ic('text-search')}No stories match the filters. Widen the period or clear the search.</p>`;
 $('#moreBtn').hidden=f.length<=st.lim;$('#moreBtn').innerHTML=`${ic('chevron-down')}Show more <span class="muted">(${Math.max(0,f.length-st.lim)} remaining)</span>`;
 drawTemas(f);
 const top=NEWS.filter(n=>matchCo(n.empresas,st.co)&&n.nivel!=='white'&&n.t&&GEN-n.t<=3*DAY).sort((a,b)=>b.score-a.score).slice(0,8);
 $('#topNews').innerHTML=top.length?top.map(n=>newsItem(n)).join(''):`<p class="empty">${ic('info')}No relevant news in the last 72 h for this filter.</p>`;
 drawNiveis();
 const r=NEWS.filter(n=>n.nivel==='red').length,o=NEWS.filter(n=>n.nivel==='orange').length;
 $('#resumoLead').textContent=`${NEWS.length} stories were collected from ${arr(D.fontes).filter(x=>/^ok/.test(x.estado)).length} sources: ${r} classed as material and ${o} as important. Start with the alerts and confirm each relevant story at the original source.`;}

/* ---------- prices ---------- */
function rows(){const list=ATIVOS.slice();const g=MERC.find(m=>m.id==='GSPC');if(g)list.push(Object.assign({},g,{id:'GSPC',moeda:g.moeda||'USD'}));return list.filter(a=>st.co==='all'||a.id===st.co||(st.co==='MKT'&&(a.id==='SXR8'||a.id==='GSPC')));}
function colorOf(a){return a.id==='GSPC'?'var(--macro)':heroColor(a);}
const nmOf=a=>ehEtf(a.id)?'ETF '+a.id:a.nome;
function renderPrices(){const cur=st.cur,R=rows();$('#cotLead').textContent=`In ${cur==='EUR'?'euros, which is what counts for you: the return includes the currency effect':'dollars: the companies\' return without the currency effect (Bitcoin and the ETFs, priced in euros, are converted)'}. The S&P 500 is shown as a benchmark.`;
 $('#perfSub').textContent=`Indexed to 100 at the start of the period, in ${cur==='EUR'?'euros':'dollars'}. Hover to see the values. Markers at the bottom: earnings (bottom row), rate decisions and material news (top row); hover a marker for the headline.`;
 lineMulti($('#ch-perf'),R.map(a=>{const p=slice(inCur(a,cur));return{n:nmOf(a),c:colorOf(a),pts:p.length?p.map(x=>[x[0],x[1]/p[0][1]*100]):[]};}),{ref:100,refLabel:'base 100',dec:1,tdec:0,h:340,label:'Relative performance, indexed to 100',end:v=>pct(v-100),legendVal:s=>pct(s.pts[s.pts.length-1][1]-100),events:eventosGrafico(R)});
 const sym=cur==='EUR'?'€':'$',first=a=>`<td><span class="co" style="--c:${colorOf(a)}">${esc(nmOf(a))}</span></td>`;
 $('#tbl-ret').innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">Last</th><th class="num" scope="col">1 session</th><th class="num" scope="col">1 week</th><th class="num" scope="col">1 month</th><th class="num" scope="col">3 months</th><th class="num" scope="col">Year to date</th><th class="num" scope="col">1 year</th></tr></thead><tbody>'+R.map(a=>{const s=stats(inCur(a,cur));if(!s)return`<tr>${first(a)}<td colspan="7" class="dim">${a.pts.length?'No EUR/USD rate for these dates':'No data'}</td></tr>`;const unit=a.id==='GSPC'?'':sym,fr=frescura(a.id,a.pts),ao=asOf(a,fr);return`<tr>${first(a)}<td class="num">${unit}${nf(s.last,2)}${ao?`<span class="td-sub${fr&&(fr.velho||antigo(a))?' old':''}">${esc(ao)}</span>`:''}</td>${[s.d1,s.w1,s.m1,s.m3,s.ytd,s.y1].map(v=>`<td class="num ${cls(v)}">${pct(v)}</td>`).join('')}</tr>`;}).join('')+'</tbody>';
 $('#tbl-ind').innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">52-week high</th><th class="num" scope="col">Distance</th><th class="num" scope="col">50-session avg.</th><th class="num" scope="col">200-session avg.</th><th class="num" scope="col">30-session volatility</th><th class="num" scope="col">Max drawdown, 1 year</th></tr></thead><tbody>'+R.map(a=>{const s=stats(inCur(a,cur));if(!s)return'';const unit=a.id==='GSPC'?'':sym;return`<tr>${first(a)}<td class="num">${s.hi!=null?unit+nf(s.hi,2):'—'}</td><td class="num ${cls(s.dist)}">${pct(s.dist)}</td><td class="num">${s.ma50!=null?unit+nf(s.ma50,2):'—'}</td><td class="num">${s.ma200!=null?unit+nf(s.ma200,2)+(s.last<s.ma200?'<span class="below">below</span>':''):'—'}</td><td class="num">${s.vol!=null?nf(s.vol)+'%':'—'}</td><td class="num neg">${pct(s.mdd)}</td></tr>`;}).join('')+'</tbody>';
 const ok=R.map(a=>({a,s:stats(inCur(a,cur))})).filter(x=>x.s&&x.s.dist!=null);barChart($('#ch-dist'),{cats:ok.map(x=>nmOf(x.a)),series:[{n:'Distance',c:'var(--ink)',colors:ok.map(x=>colorOf(x.a)),values:ok.map(x=>x.s.dist)}],unit:'%',label:'Distance from the 52-week high'});
 drawMoves();drawEarn();}

/* ---------- Prices: eventos no gráfico, grandes movimentos explicados e reação aos resultados ---------- */
/* Dados da tarefa 11 (script): histórico de notícias (material e important, 400 dias) e datas passadas de resultados (SEC).
   Sem eles, cada parte diz "Unavailable" com a razão; o resto do separador não muda. */
const HN=(()=>{const h=D.historicoNoticias&&typeof D.historicoNoticias==='object'&&!Array.isArray(D.historicoNoticias)?D.historicoNoticias:null,ini=h?Date.parse(h.inicio):NaN;
 return{ok:!!h&&isFinite(ini),ini:isFinite(ini)?ini:null,L:h?arr(h.noticias).filter(n=>n&&n.titulo&&isFinite(Date.parse(n.data))).map(n=>({titulo:String(n.titulo),link:String(n.link||''),fonte:String(n.fonte||''),t:Date.parse(n.data),empresas:arr(n.empresas).map(String),temas:arr(n.temas).map(String),nivel:String(n.nivel||'')})):[]};})();
/* limiares de movimento diário: os mesmos dos alertas (alerts(): 4% ações e ETF, 6% Bitcoin) */
const MOV_LIM=id=>id==='BTC'?6:4;
/* decisões de juros: eventos do calendário ("Fed decision…") e títulos com a mesma regra de $DecisaoJuros no script */
const JUROS_EV=/\b(fed|fomc|ecb|bce)\b[^()]*\bdecision\b|\brate decision\b/i,JUROS_RE=/\b((raises|raised|cuts|cut|hikes|hiked|holds|held) (interest )?rates|rate (hike|cut|decision)|sobe (as )?taxas|desce (as )?taxas)\b/i;
/* Sessão de um instante t para um ativo, no fuso da sua bolsa: depois do fecho (ou de uma sessão curta) conta a sessão
   seguinte; fim de semana ou feriado passa para a sessão seguinte da série (sess: instantes das datas da série, por ordem).
   Bitcoin: o dia UTC. Formatadores e fechos guardados em cache (o histórico pode ter milhares de notícias). */
const SESS_F={},SESS_FECHO=new Map(),SESS_C=new Map();
function sessaoDe(id,t,sess){const bid=id==='BTC'?'UTC':(BOLSA_DE[id]||'US'),ck=bid+':'+t;let k=SESS_C.get(ck);
 if(k==null){let iso=isoU(t);const b=bid==='UTC'?null:arr(D.bolsas).find(x=>x.id===bid);
  if(b){const f=SESS_F[b.tz]||(SESS_F[b.tz]=new Intl.DateTimeFormat('en-CA',{timeZone:b.tz,year:'numeric',month:'2-digit',day:'2-digit'})),d0=f.format(new Date(t)),fk=bid+d0;let fe=SESS_FECHO.get(fk);
   if(fe==null){const q=d0.split('-').map(Number);fe=zoned(q[0],q[1],q[2],fechoDe(b,d0,q[0]),b.tz);SESS_FECHO.set(fk,fe);}iso=t>=fe?isoU(Date.parse(d0+'T00:00:00Z')+DAY):d0;}
  k=Date.parse(iso+'T00:00:00Z');SESS_C.set(ck,k);}
 if(!sess)return k;let lo=0,hi=sess.length;while(lo<hi){const m=(lo+hi)>>1;if(sess[m]<k)lo=m+1;else hi=m;}return lo<sess.length?sess[lo]:null;}
/* notícias de um ativo: as suas, mais as de Market para as ações dos EUA e o SXR8 */
const NEWS_DE=id=>n=>n.empresas.includes(id)||((BOLSA_DE[id]==='US'||id==='SXR8')&&n.empresas.includes('MKT'));
/* sessões do período escolhido com um movimento igual ou acima do limiar (preço na moeda do ativo, sem o preço intradiário),
   com as notícias do histórico dessa sessão e da anterior */
function grandesMovimentos(){const out=[],ini=HN.ini!=null?Date.parse(isoU(HN.ini)+'T00:00:00Z'):null;
 rows().forEach(a=>{if(a.id==='GSPC')return;const pts=a.parcial?a.pts.slice(0,-1):a.pts;if(pts.length<2)return;const S=slice(pts),t0=S.length?S[0][0]:Infinity,lim=MOV_LIM(a.id),sess=pts.map(p=>p[0]);let NS=null;
  const noticias=(t1,t2)=>{if(!NS)NS=HN.L.filter(NEWS_DE(a.id)).map(n=>({n,s:sessaoDe(a.id,n.t,sess)}));return NS.filter(x=>x.s===t1||x.s===t2).map(x=>x.n).sort((p,q)=>(q.nivel==='red')-(p.nivel==='red')||q.t-p.t);};
  for(let j=1;j<pts.length;j++){if(pts[j][0]<t0)continue;const mv=(pts[j][1]/pts[j-1][1]-1)*100;if(!isFinite(mv)||Math.abs(mv)<lim)continue;
   out.push({id:a.id,t:pts[j][0],mv,antes:ini!=null&&pts[j-1][0]<ini,news:HN.ok?noticias(pts[j][0],pts[j-1][0]):[]});}});
 return out.sort((x,y)=>y.t-x.t||String(x.id).localeCompare(String(y.id)));}
function drawMoves(){const box=$('#tbl-moves'),nt=$('#movesNote');if(!box)return;const M=grandesMovimentos(),V=M.slice(0,30);
 nt.textContent=(HN.ok?`News history since ${fdt(HN.ini)}: the material and important stories of each session and of the session before, in the exchange's time zone (a story after the close counts for the next session).`:'News history: Unavailable (this data file has no news history; the script keeps it from its next run). The moves are listed without their news.')+` Moves of ${MOV_LIM('AAPL')}% or more for stocks and ETFs, ${MOV_LIM('BTC')}% for Bitcoin (the alert thresholds), in each asset's own currency.${M.length>30?` Showing the 30 most recent of ${M.length}.`:''}`;
 const li=n=>`<li><span class="tag ${n.nivel==='red'?'t-warn':''}">${n.nivel==='red'?'Material':'Important'}</span> <a href="${esc(safe(n.link))}" target="_blank" rel="noopener noreferrer">${esc(n.titulo)}</a> <span class="muted">${esc(n.fonte)}</span></li>`;
 box.innerHTML='<thead><tr><th scope="col">Session</th><th scope="col">Asset</th><th class="num" scope="col">Move</th><th scope="col">News of this session and the one before</th></tr></thead><tbody>'+(V.length?V.map(m=>`<tr data-co="${esc(m.id)}"><td>${fdt(m.t)}</td><td>${coTag(m.id)}</td><td class="num ${cls(m.mv)}">${pct(m.mv,2)}</td><td>${!HN.ok?'<span class="dim">Unavailable</span>':m.news.length?'<ul class="mv-news">'+m.news.slice(0,4).map(li).join('')+(m.news.length>4?`<li class="muted">and ${m.news.length-4} more</li>`:'')+'</ul>'+(m.antes?'<span class="td-sub">Partly before the news history</span>':''):`<span class="dim">${m.antes?'Before the news history':'No material or important story in the history for these sessions'}</span>`}</td></tr>`).join(''):`<tr><td colspan="4" class="dim">No session with a move at or above the threshold in this period.</td></tr>`)+'</tbody>';}
/* reação aos resultados (USD, histórico longo): sessão de reação vs fecho anterior; 5 sessões = fecho da 5.ª (a de reação é a 1.ª) */
function reacoes(id){const r=D.resultadosSec&&typeof D.resultadosSec==='object'?D.resultadosSec[id]:null;
 if(!r||typeof r!=='object')return{ok:false,motivo:'no SEC data in this data file (it is collected with -EmailSEC, as the desktop shortcut does, from the next run)'};
 if(r.estado!=='ok'&&r.estado!=='previous run')return{ok:false,motivo:r.estado==='skipped'?'the SEC data was not collected (run the script with -EmailSEC)':'the SEC source failed'+(r.erro?': '+String(r.erro):'')};
 /* um preço intradiário (sessão ainda aberta quando a página foi gerada) não é um fecho: não conta como reação */
 const a=ATIVOS.find(x=>x.id===id),H0=HIST[id]||[],H=a&&a.parcial&&H0.length&&a.pts.length&&H0[H0.length-1][0]>=a.pts[a.pts.length-1][0]?H0.slice(0,-1):H0,ts=H.map(p=>p[0]),R=arr(r.resultados),C=[];
 R.forEach(x=>{if(!x||!isoOk(x.sessao))return;const t=Date.parse(x.sessao+'T00:00:00Z'),i=ts.indexOf(t);if(i<1)return;const p0=H[i-1][1];C.push({x,t,r1:(H[i][1]/p0-1)*100,r5:i+4<H.length?(H[i+4][1]/p0-1)*100:null});});
 C.sort((a,b)=>b.t-a.t);const ab=C.map(c=>Math.abs(c.r1)).sort((a,b)=>a-b),n=ab.length,med=n?(n%2?ab[(n-1)/2]:(ab[n/2-1]+ab[n/2])/2):null;
 return{ok:true,C,med,sem:R.length-C.length,ant:r.estado==='previous run',fonte:String(r.fonte||'')};}
const QUANDO={after:'after the close',before:'before the open',during:'during the session',closed:'exchange closed that day'};
function drawEarn(){const box=$('#tbl-earn'),cs=$('#earnCases');if(!box)return;const ids=['AAPL','NVDA','GOOGL'].filter(id=>st.co==='all'||st.co===id),L=ids.map(id=>({id,E:reacoes(id)}));
 const quando=x=>`${fdS(x.entrega)}${x.horaNY?`, ${esc(x.horaNY.slice(11))} New York (${QUANDO[x.quando]||''})`:''}`;
 box.innerHTML='<thead><tr><th scope="col">Company</th><th scope="col">Latest earnings</th><th class="num" scope="col">Reaction session</th><th class="num" scope="col">After 5 sessions</th><th class="num" scope="col">Median move (absolute)</th><th class="num" scope="col">Cases</th></tr></thead><tbody>'+(L.length?L.map(({id,E})=>{
  if(!E.ok||!E.C.length)return`<tr data-co="${id}"><td>${coTag(id)}</td><td colspan="5"><span class="dim">Unavailable: ${esc(E.ok?'no earnings release with a known reaction session in the price history':E.motivo)}</span></td></tr>`;
  const c=E.C[0];return`<tr data-co="${id}"><td>${coTag(id)}${E.ant?`<span class="td-sub old">${esc(E.fonte)}</span>`:''}</td><td>${quando(c.x)}<span class="td-sub">reaction session ${fdS(c.x.sessao)}${c.x.quando==='during'?' (released during the session)':''}</span></td><td class="num ${cls(c.r1)}">${pct(c.r1,2)}</td><td class="num ${cls(c.r5)}">${c.r5!=null?pct(c.r5,2):'—'}</td><td class="num">${nf(E.med,2)}%</td><td class="num">${E.C.length}${E.sem?`<span class="td-sub">${E.sem} without a reaction session</span>`:''}</td></tr>`;}).join(''):'<tr><td colspan="6" class="dim">Choose All or one of the three companies.</td></tr>')+'</tbody>';
 if(cs)cs.innerHTML=L.filter(x=>x.E.ok&&x.E.C.length>1).map(({id,E})=>`<details class="dups"><summary>${ic('layers')}${esc(CO[id].n)}: all ${E.C.length} cases${ic('chevron-down','chev')}</summary><ul>${E.C.map(c=>`<li><span>${quando(c.x)} → ${fdS(c.x.sessao)}</span><span class="${cls(c.r1)}">${pct(c.r1,2)}</span><span class="muted">5 sessions ${c.r5!=null?pct(c.r5,2):'—'}</span></li>`).join('')}</ul></details>`).join('');}
/* marcadores do gráfico de Prices: resultados (sessão de reação, ou data de entrega sem hora conhecida), decisões de juros
   (calendário passado e notícias do histórico com o tema de juros) e notícias "material"; um por dia e por tipo */
function eventosGrafico(R){const ids=R.map(a=>a.id),E=[];
 ['AAPL','NVDA','GOOGL'].filter(id=>ids.includes(id)).forEach(id=>{const r=D.resultadosSec&&typeof D.resultadosSec==='object'?D.resultadosSec[id]:null;if(!r||(r.estado!=='ok'&&r.estado!=='previous run'))return;
  arr(r.resultados).forEach(x=>{if(!x)return;const d=isoOk(x.sessao)?x.sessao:isoOk(x.entrega)?x.entrega:null;if(!d)return;
   E.push({k:'earn',t:Date.parse(d+'T00:00:00Z'),c:CO[id].c,txt:`${CO[id].n}: earnings, ${x.horaNY?`8-K accepted ${x.horaNY} New York`:`8-K filed ${x.entrega}, time unknown`}${isoOk(x.sessao)?`, reaction session ${fdS(x.sessao)}`:''}`});});});
 calendar().filter(e=>e.t<=GEN&&JUROS_EV.test(String(e.ev||''))).forEach(e=>E.push({k:'rate',t:e.t,c:'var(--macro)',txt:String(e.ev)}));
 HN.L.filter(n=>n.temas.includes('Macro and rates')&&JUROS_RE.test(n.titulo)).forEach(n=>E.push({k:'rate',t:sessaoDe('MKT',n.t),c:'var(--macro)',txt:n.titulo}));
 HN.L.filter(n=>n.nivel==='red'&&(st.co==='all'||matchCo(n.empresas,st.co))).forEach(n=>E.push({k:'news',t:sessaoDe(n.empresas.find(c=>c!=='MKT')||'MKT',n.t),c:'var(--red)',txt:n.titulo}));
 const G=new Map(),ROW={earn:0,rate:1,news:2},NOME={earn:'Earnings',rate:'Rate decision',news:'Material news'};
 E.forEach(e=>{const key=e.k+':'+e.t;if(!G.has(key))G.set(key,{k:e.k,t:e.t,c:e.c,L:[]});G.get(key).L.push(e.txt);});
 return[...G.values()].map(g=>({k:g.k,t:g.t,c:g.c,row:ROW[g.k],tip:`${NOME[g.k]} · ${fdS(isoU(g.t))}\n`+g.L.slice(0,3).join('\n')+(g.L.length>3?`\n+${g.L.length-3} more`:'')}));}

/* ---------- Fundamentals ---------- */
/* Fundamentais das três empresas (SEC XBRL, tarefa 13). Cada valor tem a data de publicação (f: entrega mais recente;
   p: a primeira; v0: o primeiro valor, se a entrega mais recente o mudou). O histórico de 10 anos usa, em cada fim de mês,
   só o que já estava publicado (sem look-ahead). Calculado só quando o separador abre (CHARTS.fundamentals). */
const FUND_IDS=['AAPL','NVDA','GOOGL'],FUND_H={};
function fundDados(id){const f=D.fundamentais&&typeof D.fundamentais==='object'?D.fundamentais[id]:null;
 if(!f||typeof f!=='object')return{ok:false,motivo:'no fundamentals in this data file (they are collected with -EmailSEC, as the desktop shortcut does, from the next run)'};
 if(f.estado!=='ok'&&f.estado!=='previous run')return{ok:false,motivo:f.estado==='skipped'?'not collected (run the script with -EmailSEC)':'the SEC source failed'+(f.erro?': '+String(f.erro):'')};
 const T=arr(f.trimestres).filter(q=>q&&isoOk(q.fim)&&q.m&&typeof q.m==='object').map(q=>({fim:q.fim,t:Date.parse(q.fim+'T00:00:00Z'),ano:q.ano?String(q.ano):'',q:+q.q||null,m:q.m})).sort((a,b)=>a.t-b.t);
 if(!T.length)return{ok:false,motivo:'no quarterly figures in the data'};
 return{ok:true,f,T,ant:f.estado==='previous run',faltam:f.faltam&&typeof f.faltam==='object'?f.faltam:{}};}
/* valor de um trimestre conhecido na data t: o da entrega mais recente até t (v se f ≤ t; senão v0, ou v, se p ≤ t) */
function valorEm(x,t){if(!x||typeof x!=='object'||x.v==null||!isFinite(+x.v))return null;const f=Date.parse(x.f+'T00:00:00Z'),p=Date.parse((x.p||x.f)+'T00:00:00Z');
 if(f<=t)return +x.v;if(p<=t)return x.v0!=null&&isFinite(+x.v0)?+x.v0:+x.v;return null;}
/* TTM conhecido em t: os 4 trimestres consecutivos (fins a 80–100 dias) mais recentes já publicados */
function ttmEm(T,k,t){let i=T.length-1;while(i>=0&&valorEm(T[i].m[k],t)==null)i--;if(i<3)return null;let s=0;
 for(let j=i-3;j<=i;j++){const v=valorEm(T[j].m[k],t);if(v==null)return null;if(j>i-3){const g=(T[j].t-T[j-1].t)/DAY;if(g<80||g>100)return null;}s+=v;}return{v:s,fim:T[i].t};}
function acoesEm(T,t){for(let i=T.length-1;i>=0;i--){const v=valorEm(T[i].m.acoes,t);if(v!=null)return{v,fim:T[i].t};}return null;}
/* percentil: % dos valores do histórico iguais ou abaixo do atual */
function percentil(L,x){if(!L.length||x==null||!isFinite(x))return null;return 100*L.filter(v=>v<=x+1e-9*Math.max(1,Math.abs(x))).length/L.length;}
/* P/E e P/FCF no último fecho de cada mês dos últimos 10 anos (histórico longo em USD) e os valores atuais */
function fundHistorico(id){if(FUND_H[id])return FUND_H[id];const F=fundDados(id),H=HIST[id]||[],out={pe:[],pfcf:[],atual:null};if(!F.ok)return FUND_H[id]=out;
 if(H.length){const d1=new Date(H[H.length-1][0]),t0=Date.UTC(d1.getUTCFullYear()-10,d1.getUTCMonth(),d1.getUTCDate());
  for(let i=0;i<H.length;i++){if(H[i][0]<t0)continue;const n=H[i+1];if(n&&isoU(n[0]).slice(0,7)===isoU(H[i][0]).slice(0,7))continue;const t=H[i][0],p=H[i][1];
   const e=ttmEm(F.T,'eps',t);if(e&&e.v>0)out.pe.push([t,p/e.v]);const c=ttmEm(F.T,'fcf',t),s=acoesEm(F.T,t);if(c&&c.v>0&&s)out.pfcf.push([t,p*s.v/c.v]);}}
 const a=ATIVOS.find(x=>x.id===id),pt=a&&a.pts.length?a.pts[a.pts.length-1]:null,e=ttmEm(F.T,'eps',GEN),c=ttmEm(F.T,'fcf',GEN),s=acoesEm(F.T,GEN);
 out.atual={preco:pt?pt[1]:null,tp:pt?pt[0]:null,parcial:!!(a&&a.parcial),eps:e,fcf:c,acoes:s,pe:pt&&e&&e.v>0?pt[1]/e.v:null,pfcf:pt&&c&&c.v>0&&s?pt[1]*s.v/c.v:null};
 return FUND_H[id]=out;}
const FALTA_N={receita:'revenue',lucroBruto:'gross profit',lucroOperacional:'operating income',cfo:'operating cash flow',capex:'capex',eps:'diluted EPS',acoes:'diluted shares'};
function fundBloco(id){const F=fundDados(id),nome=CO[id].n;
 if(!F.ok)return`<div class="fund-co" data-co="${id}"><h3 class="sec-h">${coTag(id)}</h3><p class="empty">${ic('info')}Unavailable: ${esc(F.motivo)}.</p></div>`;
 const T=F.T,U=T[T.length-1],H=fundHistorico(id),A=H.atual,un='<span class="dim">Unavailable</span>',mv=x=>x&&x.v!=null&&isFinite(+x.v)?+x.v:null,dv=x=>x&&x.d?'<sup title="Derived">*</sup>':'';
 const ano=y=>T.find(x=>{const g=(y.t-x.t)/DAY;return g>=350&&g<=380;});
 const linhas=T.slice(-8).reverse().map(q=>{const m=q.m,r=mv(m.receita),ya=ano(q),ry=ya?mv(ya.m.receita):null,gb=mv(m.lucroBruto),op=mv(m.lucroOperacional),fc=mv(m.fcf),ep=mv(m.eps),ac=mv(m.acoes),acy=ya?mv(ya.m.acoes):null;
  const vs=Object.values(m).filter(x=>x&&typeof x==='object'&&x.p),p=vs.length?vs.map(x=>x.p).sort()[0]:null,f=vs.length?vs.map(x=>x.f).sort().pop():null;
  return`<tr><td>${q.ano&&q.q?esc(q.ano)+' Q'+q.q:'Quarter'}<span class="td-sub">ended ${fdt(q.t)}</span></td><td class="num">${r!=null?'$'+nf(r/1e9,2)+' bn'+dv(m.receita):un}</td><td class="num ${cls(r!=null&&ry?r/ry-1:null)}">${r!=null&&ry?pct((r/ry-1)*100):'—'}</td>`+
   `<td class="num">${gb!=null&&r?nf(gb/r*100,1)+'%'+dv(m.lucroBruto):un}</td><td class="num">${op!=null&&r?nf(op/r*100,1)+'%'+dv(m.lucroOperacional):un}</td><td class="num">${fc!=null?'$'+nf(fc/1e9,2)+' bn'+dv(m.fcf):un}</td><td class="num">${ep!=null?'$'+nf(ep,2)+dv(m.eps):un}</td>`+
   `<td class="num">${ac!=null?nf(ac/1e9,2)+' bn':un}${ac!=null&&acy?`<span class="td-sub">${pct((ac/acy-1)*100)} a year${ac<acy?' (buybacks)':ac>acy?' (dilution)':''}</span>`:''}</td><td>${p?fdt(Date.parse(p+'T00:00:00Z')):'—'}${f&&p&&f!==p?`<span class="td-sub">revised ${fdt(Date.parse(f+'T00:00:00Z'))}</span>`:''}</td></tr>`;}).join('');
 const desde=L=>L.length?fdt(L[0][0]):'',perc=(L,x)=>{const pc=percentil(L.map(p=>p[1]),x);return pc==null?'':` · higher than or equal to ${nf(pc,0)}% of the month-end values since ${desde(L)}`;};
 const pTxt=A.preco!=null?`price $${nf(A.preco,2)} (${A.parcial?'live, ':''}${fdt(A.tp)})`:'no current price';
 const peK=A.pe!=null?nf(A.pe,1):A.eps&&A.eps.v<=0?'n/m':'Unavailable',peS=A.pe!=null?`${pTxt} ÷ EPS $${nf(A.eps.v,2)} (4 quarters to ${fdt(A.eps.fim)})${perc(H.pe,A.pe)}`:A.eps&&A.eps.v<=0?'earnings of the last 4 quarters are not positive':A.preco==null?'no current price':'not enough published quarters';
 const pfK=A.pfcf!=null?nf(A.pfcf,1):A.fcf&&A.fcf.v<=0?'n/m':'Unavailable',pfS=A.pfcf!=null?`${pTxt} × ${nf(A.acoes.v/1e9,2)} bn shares ÷ FCF $${nf(A.fcf.v/1e9,1)} bn (4 quarters to ${fdt(A.fcf.fim)})${perc(H.pfcf,A.pfcf)}`:A.fcf&&A.fcf.v<=0?'free cash flow of the last 4 quarters is not positive':!A.acoes?'no diluted share count':'not enough published quarters';
 const falta=Object.keys(F.faltam).filter(k=>F.faltam[k]).map(k=>`${FALTA_N[k]||k} (${esc(F.faltam[k])})`);
 return`<div class="fund-co" data-co="${id}"><h3 class="sec-h">${coTag(id)}<span class="muted">${U.ano&&U.q?esc(U.ano)+' Q'+U.q+' · ':''}latest quarter ended ${fdt(U.t)} · ${F.ant?esc(F.f.fonte):'SEC filings'}</span></h3>`+
  `<div class="kpis">${kpi('chart','P/E, last 4 quarters',peK,peS)}${kpi('wallet','P/FCF, last 4 quarters',pfK,pfS)}</div>`+
  `<div class="tbl-wrap"><table><thead><tr><th scope="col">Quarter</th><th class="num" scope="col">Revenue</th><th class="num" scope="col">Growth, a year</th><th class="num" scope="col">Gross margin</th><th class="num" scope="col">Operating margin</th><th class="num" scope="col">Free cash flow</th><th class="num" scope="col">Diluted EPS</th><th class="num" scope="col">Diluted shares</th><th scope="col">Published</th></tr></thead><tbody>${linhas}</tbody></table></div>`+
  (falta.length?`<p class="note">${ic('info')}<span>Unavailable for ${esc(nome)}: ${falta.join('; ')}.</span></p>`:'')+
  `<div class="fig"><div class="fig-head"><div><p class="fig-title">${esc(nome)}: P/E and P/FCF over 10 years</p><p class="fig-sub">At each month-end, with only the quarters published by then.</p></div></div><div id="ch-fund-${id}"></div></div></div>`;}
function renderFund(){const box=$('#fundOut');if(!box)return;const ids=FUND_IDS.filter(id=>st.co==='all'||st.co===id);
 if(!ids.length){box.innerHTML=`<p class="empty">${ic('info')}Fundamentals cover Apple, NVIDIA and Alphabet: choose All or one of them in the filter.</p>`;return;}
 box.innerHTML=ids.map(fundBloco).join('');
 ids.forEach(id=>{const c=$('#ch-fund-'+id);if(!c)return;const H=fundHistorico(id);if(H.pe.length<2&&H.pfcf.length<2)return empty(c,'Unavailable: not enough published quarters for a history.');
  lineMulti(c,[{n:'P/E',c:CO[id].c,pts:H.pe},{n:'P/FCF',c:'var(--ink)',pts:H.pfcf}],{dec:1,tdec:0,h:240,label:`${CO[id].n}: P/E and P/FCF at month-ends over 10 years`});});}

/* ---------- FX & macro ---------- */
function renderFx(){const s=stats(FX),bl=BCE.length?BCE[BCE.length-1]:null;
 $('#fxKpis').innerHTML=s?`<div class="kpi"><div class="k">${ic('euro')}EUR/USD (latest)</div><div class="v">${nf(s.last,4)}</div><div class="s">dollars per euro · ${FXSRC&&!/^Yahoo/.test(FXSRC)?esc(FXSRC)+', ':''}${fdS(isoU(FX[FX.length-1][0]))}</div></div><div class="kpi"><div class="k">${ic('landmark')}ECB reference rate</div><div class="v">${bl?nf(bl[1],4):'—'}</div><div class="s">${bl?fdt(bl[0]):''}</div></div><div class="kpi"><div class="k">${ic('calendar')}1 month</div><div class="v ${cls(-s.m1)}">${pct(s.m1)}</div><div class="s">${s.m1>0?'dollar losing value':'dollar gaining value'}</div></div><div class="kpi"><div class="k">${ic('chart')}Year to date</div><div class="v ${cls(-s.ytd)}">${pct(s.ytd)}</div><div class="s">change in EUR/USD</div></div>`:`<p class="empty">${ic('info')}No currency data.</p>`;
 lineMulti($('#ch-fx'),[{n:'EUR/USD',c:'var(--ink)',pts:slice(FX)}],{dec:4,tdec:2,area:true,h:260,label:'EUR/USD'});
 const US=ATIVOS.filter(a=>(a.moeda||'').toUpperCase()==='USD'&&(st.co==='all'||st.co===a.id));
 $('#tbl-dec').innerHTML='<thead><tr><th scope="col">Stock</th><th class="num" scope="col">Return in USD</th><th class="num" scope="col">Currency effect</th><th class="num" scope="col">Return in EUR</th></tr></thead><tbody>'+(US.map(a=>{const p=slice(a.pts),f0=p.length?fxAt(p[0][0]):null,f1=p.length?fxAt(p[p.length-1][0]):null;if(p.length<2||!f0||!f1)return`<tr><td>${coTag(a.id)}</td><td colspan="3" class="dim">No data${p.length>=2?' (no EUR/USD rate for these dates)':''}</td></tr>`;const ru=p[p.length-1][1]/p[0][1]-1,fxe=f0/f1-1,re=(1+ru)*(1+fxe)-1;return`<tr><td>${coTag(a.id)}</td><td class="num ${cls(ru)}">${pct(ru*100)}</td><td class="num ${cls(fxe)}">${pct(fxe*100)}</td><td class="num ${cls(re)}">${pct(re*100)}</td></tr>`;}).join('')||'<tr><td colspan="4" class="dim">Choose All or one of the US companies.</td></tr>')+'</tbody>';
 [['GSPC','#ch-gspc',0,'#k-gspc','var(--macro)',{area:true}],['VIX','#ch-vix',1,'#k-vix','var(--orange)',{ref:20,refLabel:'20'}],['TNX','#ch-tnx',2,'#k-tnx','var(--bar)',{}]].forEach(([id,sel,dec,kid,c,ex])=>{const m=MERC.find(x=>x.id===id),ms=m?stats(m.pts):null;
  $(kid).innerHTML=ms?`<span class="mk-v">${nf(ms.last,dec)}${id==='TNX'?'%':''}</span><span class="mk-c ${id==='GSPC'?cls(ms.m1):'muted'}">${id==='GSPC'?pct(ms.m1)+' in 1 month':'latest value'}</span>`:'';
  lineMulti($(sel),[{n:m?m.nome:id,c,pts:m?slice(m.pts):[]}],Object.assign({h:220,dec,label:m?m.nome:id},ex));});
 renderMacro();}
/* ---------- área do euro (BCE): taxa de depósito e inflação HICP; sem alertas novos ---------- */
function macroDados(id){const m=D.macro&&typeof D.macro==='object'?D.macro[id]:null;
 if(!m||typeof m!=='object')return{ok:false,motivo:'no ECB data in this data file (it is collected from the next run)'};
 if(m.estado!=='ok'&&m.estado!=='previous run')return{ok:false,motivo:'the ECB source failed'+(m.erro?': '+String(m.erro):'')};
 const P=toPts(m.pontos);if(!P.length)return{ok:false,motivo:'no values in the data'};return{ok:true,m,P,ant:m.estado==='previous run'};}
function renderMacro(){const t10=Date.UTC(new Date(GEN).getUTCFullYear()-10,new Date(GEN).getUTCMonth(),1),lab=x=>x.ant?`<span class="mk-c muted">${esc(x.m.fonte)}</span>`:'';
 const r=macroDados('BCE_DFR'),kd=$('#k-dfr'),cd=$('#ch-dfr');
 if(!r.ok){if(kd)kd.innerHTML='';if(cd)empty(cd,'Unavailable: '+r.motivo+'.');}
 else{const L=r.P[r.P.length-1];kd.innerHTML=`<span class="mk-v">${nf(L[1],2)}%</span><span class="mk-c muted">since ${fdt(L[0])}</span>${lab(r)}`;
  /* degraus: cada valor vale até à mudança seguinte; o primeiro ponto é o valor em vigor há 10 anos */
  const S=[];let ant=null;r.P.forEach(p=>{if(p[0]<t10){ant=p[1];return;}if(!S.length&&ant!=null)S.push([t10,ant]);if(S.length)S.push([p[0]-DAY,S[S.length-1][1]]);S.push(p);});if(!S.length&&ant!=null)S.push([t10,ant]);S.push([GEN,L[1]]);
  lineMulti(cd,[{n:'Deposit facility rate',c:'var(--macro)',pts:S}],{h:220,dec:2,tdec:1,label:'ECB deposit facility rate, %'});}
 const h=macroDados('HICP_EA'),kh=$('#k-hicp'),ch=$('#ch-hicp');
 if(!h.ok){if(kh)kh.innerHTML='';if(ch)empty(ch,'Unavailable: '+h.motivo+'.');}
 else{const L=h.P[h.P.length-1],d=new Date(L[0]),velho=GEN-L[0]>75*DAY;
  kh.innerHTML=`<span class="mk-v">${nf(L[1],1)}%</span><span class="mk-c ${velho?'old':'muted'}">${MES[d.getUTCMonth()]} ${d.getUTCFullYear()}${velho?' · latest published, not recent':''}</span>${lab(h)}`;
  lineMulti(ch,[{n:'HICP, 12 months',c:'var(--orange)',pts:h.P.filter(p=>p[0]>=t10)}],{h:220,dec:1,tdec:0,ref:2,refLabel:'2% target',label:'Euro area inflation, HICP, %'});}}

/* ---------- calendar ---------- */
const IMPN={'Very high':3,'High':2,'Medium':1};
function renderCal(){const imp=$('#c-imp').value,past=$('#c-past').checked;
 const L=calendar().filter(e=>(st.co==='all'||e.e===st.co||(st.co==='MKT'&&e.e==='MKT')||e.e==='TU')&&(!imp||e.imp===imp)&&(past||diasAte(e.t)>=0));
 $('#tbl-cal').innerHTML='<thead><tr><th scope="col">Date</th><th scope="col">Who</th><th scope="col">Event</th><th scope="col">Importance</th><th scope="col">When</th></tr></thead><tbody>'+(L.length?L.map(e=>{const dd=diasAte(e.t),d=new Date(e.t),lv=IMPN[e.imp]||1;
  return`<tr class="${dd<0?'past':''}"><td><div class="cal-date" title="${fdt(e.t)}"><span class="cal-d">${d.getUTCDate()}</span><span class="cal-m">${MES[d.getUTCMonth()]} ${d.getUTCFullYear()}<br>${WD[d.getUTCDay()]}</span></div></td><td>${coTag(e.e)||esc(e.e)}</td><td><div class="cal-ev">${esc(e.ev)}</div><span class="tag ${e.st==='C'?'t-pos':''}">${ic(e.st==='C'?'check':'clock')}${e.st==='C'?'Confirmed':'Estimated'}${e.auto?' · '+esc(e.auto):''}</span></td><td><span class="imp" data-l="${lv}"><span class="bars" aria-hidden="true">${[1,2,3].map(i=>`<i class="${i<=lv?'on':''}"></i>`).join('')}</span>${esc(e.imp)}</span></td><td><span class="when ${dd===0?'today':dd>0&&dd<=7?'soon':''}">${ic('clock')}${dd===0?'today':dd>0?`in ${dias(dd)}`:`${dias(-dd)} ago`}</span></td></tr>`;}).join(''):'<tr><td colspan="5" class="dim">No events. Update the calendar at the top of the script.</td></tr>')+'</tbody>';}

/* ---------- ETF ---------- */
/* Uma só vista para todos os ETF: o escolhido (etfSel) usa os seus próprios dados (holdings, histórico, simulador).
   Com o SXR8 escolhido, a vista é a de sempre. */
let etfSel='SXR8';
const etfNomeCurto=id=>((ATIVOS.find(a=>a.id===id)||{}).nome||id).replace(/\s*\([A-Z0-9]+\)$/,'');
const etfTemEmpresas=id=>{const w=etfInfo(id);return arr(w.empresas).length>0||(+w.AAPL||0)+(+w.NVDA||0)+(+w.GOOGL||0)>0;};
const simKey=id=>id==='SXR8'?'etf':'etf-'+id;
function renderEtf(){const id=etfSel,w=etfInfo(id),A=+w.AAPL||0,N=+w.NVDA||0,G=+w.GOOGL||0,tot=A+N+G,s=stats((ATIVOS.find(a=>a.id===id)||{pts:[]}).pts),ult=s?` Last price: €${nf(s.last,2)}; ${pct(s.m1)} over a month and ${pct(s.y1)} over a year.`:'';
 $('#f-etf').innerHTML=ETF_IDS.map(x=>`<button type="button" data-etf="${esc(x)}" aria-pressed="${x===id}"><span class="sw-c" style="--c:${CO[x].c}" aria-hidden="true"></span>${esc(etfIdx(x))}<span class="sr"> (${esc(x)})</span></button>`).join('');
 $('#h-etf').textContent=etfNomeCurto(id);
 if(id==='SXR8')$('#etfLead').textContent=`The iShares Core S&P 500 UCITS ETF (Acc), IE00B5BMR087, listed as SXR8 on Xetra in euros, holds about ${nf(tot)}% in these three companies.${ult}`;
 else $('#etfLead').textContent=`The ${w.fundo||etfNomeCurto(id)}${w.isin?', '+w.isin:''}, listed as ${id} on Xetra in euros, tracks the ${w.benchmark||etfIdx(id)}. It is an accumulating ETF${w.domicilio==='IE'?' domiciled in Ireland':''}: dividends are reinvested in its price.${tot>0?` It holds about ${nf(tot)}% in Apple, NVIDIA and Alphabet.`:''}${ult}`;
 $('#fig-etf-w').hidden=!etfTemEmpresas(id);
 $('#etfSub').textContent=`${w.aoVivo?'iShares data':'Reference weights stored in the script (the iShares download failed)'}${w.data?', as of '+w.data:''}${etfIdade(id)>10?` (${etfIdade(id)} days old)`:''}.`;
 $('#etfTotal').innerHTML=`<div class="big">${nf(tot)}%</div><div class="s">in the three companies</div>`;
 const segs=[['Apple',A,'var(--aapl)'],['NVIDIA',N,'var(--nvda)'],['Alphabet',G,'var(--googl)'],['Rest of the ETF',100-tot,'var(--rule-2)']];
 $('#ch-etf').innerHTML='<div class="dist" role="img" aria-label="'+segs.map(x=>`${x[0]} ${nf(x[1],2)}%`).join(', ')+'">'+segs.map(x=>`<div data-tip="${x[0]}: ${nf(x[1],2)}%" style="width:${Math.max(0,x[1])}%;background:${x[2]}"></div>`).join('')+'</div><ul class="etf-legend">'+segs.map(x=>`<li style="--c:${x[2]}"><span class="dot"></span><span class="nm">${x[0]}</span><span class="bar"><i style="width:${Math.max(0,x[1])}%"></i></span><b>${nf(x[1],2)}%</b></li>`).join('')+'</ul>';
}

/* idade, em dias, da data "as of" dos pesos de um ETF (por omissão o SXR8; 0 se desconhecida) */
function etfIdade(id){const w=etfInfo(id||'SXR8'),t=Date.parse((w.dataIso||'')+'T00:00:00Z');return isFinite(t)?Math.max(0,Math.round((GEN-t)/DAY)):0;}

/* ---------- Bitcoin ---------- */
function mm(pts,k){const o=[];let s=0;for(let i=0;i<pts.length;i++){s+=pts[i][1];if(i>=k)s-=pts[i-k][1];if(i>=k-1)o.push([pts[i][0],s/k]);}return o;}
/* correlação dos retornos diários nos dias em que os dois ativos têm preço (a Bitcoin tem todos; as ações só os dias úteis), últimos 3 meses */
function corr(a,b){const mb=new Map(b.map(p=>[p[0],p[1]])),c=a.filter(p=>mb.has(p[0]));if(c.length<20)return null;const t0=c[c.length-1][0]-92*DAY,w=c.filter(p=>p[0]>=t0);if(w.length<20)return null;
 const x=[],y=[];for(let i=1;i<w.length;i++){x.push(Math.log(w[i][1]/w[i-1][1]));y.push(Math.log(mb.get(w[i][0])/mb.get(w[i-1][0])));}
 const mx=x.reduce((s,v)=>s+v,0)/x.length,my=y.reduce((s,v)=>s+v,0)/y.length;let sxy=0,sxx=0,syy=0;x.forEach((v,i)=>{sxy+=(v-mx)*(y[i]-my);sxx+=(v-mx)*(v-mx);syy+=(y[i]-my)*(y[i]-my);});return sxx&&syy?sxy/Math.sqrt(sxx*syy):null;}
const grande=v=>v==null||!isFinite(v)?'—':v>=1e12?nf(v/1e12,2)+' trillion':v>=1e9?nf(v/1e9,1)+' billion':nf(v/1e6,0)+' million';
function renderBtc(){const a=ATIVOS.find(x=>x.id==='BTC'),B=D.bitcoin||{},m=B.mercado,fg=B.sentimento,rd=B.rede,p=a?a.pts:[],s=stats(p),fx=FX.length?FX[FX.length-1][1]:null;
 $('#btcLead').textContent=`Bitcoin trades 24 hours a day, 7 days a week, and is tracked here in euros (BTC/EUR).${s?` Last price: €${nf(s.last,0)}; ${pct(s.m1)} over a month, ${pct(s.y1)} over a year and ${pct(s.dist)} from its 52-week high.`:' No prices in this run: see the Sources section.'}`;
 const k=(i,lab,v,sub,c)=>`<div class="kpi"><div class="k">${ic(i)}${lab}</div><div class="v ${c||''}">${v}</div><div class="s">${sub}</div></div>`,v24=m&&m.var24!=null&&isFinite(m.var24)?+m.var24:(s?s.d1:null),w1=s?s.w1:null,g=MERC.find(x=>x.id==='GSPC'),gs=g?stats(g.pts):null;
 $('#btcKpis').innerHTML=k('bitcoin','Price',s?'€'+nf(s.last,0):(m?'€'+nf(m.eur,0):'—'),m&&m.usd?`≈ $${nf(m.usd,0)}`:(s&&fx?`≈ $${nf(s.last*fx,0)}`:'in euros'))+
  k('clock','24 hours',pct(v24),m?'change over the last 24 h':'since 00:00 UTC',cls(v24))+
  k('calendar','7 days',pct(w1),s?`30 days: ${pct(s.m1)}`:'',cls(w1))+
  k('activity','Volatility',s&&s.vol!=null?nf(s.vol,0)+'%':'—',`annualised, 30 days${gs&&gs.vol!=null?` · S&amp;P 500: ${nf(gs.vol,0)}%`:''}`);
 const sl=slice(p),t0=sl.length?sl[0][0]:0,cut=x=>x.filter(q=>q[0]>=t0),ch=sl.length>1?(sl[sl.length-1][1]/sl[0][1]-1)*100:null;
 $('#k-btc').innerHTML=s?`<span class="mk-v">€${nf(s.last,0)}</span><span class="mk-c ${cls(ch)}">${pct(ch)} over the period</span>`:'';
 lineMulti($('#ch-btc'),[{n:'Bitcoin',c:'var(--btc)',pts:sl},{n:'50-day average',c:'var(--muted)',pts:cut(mm(p,50))},{n:'200-day average',c:'var(--accent)',pts:cut(mm(p,200))}],{dec:0,tdec:0,unit:' €',h:360,label:'Bitcoin price in euros'});
 if(fg&&fg.valor!=null&&isFinite(fg.valor)){const sp=toPts(fg.serie),v=+fg.valor,col={xf:'var(--red)',f:'var(--orange)',n:'var(--yellow)',g:'var(--pos)',xg:'var(--pos)'}[fgClasse(v)],antes=d=>{if(!sp.length)return'—';const t=sp[sp.length-1][0]-d*DAY,q=sp.filter(x=>x[0]<=t);return q.length?q[q.length-1][1]:'—';};
  $('#btcSent').innerHTML=`<div class="big-v"><b>${v}</b><span style="color:${col}">${esc(fg.classe)}</span></div><div class="fng-bar" role="img" aria-label="Index at ${v} out of 100"><i style="left:${Math.max(0,Math.min(100,v))}%"></i></div><div class="fng-lbl"><span>Extreme fear</span><span>Neutral</span><span>Extreme greed</span></div><ul class="kv"><li>A week ago<b>${antes(7)}</b></li><li>A month ago<b>${antes(30)}</b></li></ul>`;
  lineMulti($('#ch-fng'),[{n:'Fear & Greed',c:'var(--btc)',pts:slice(sp)}],{dec:0,h:170,ref:50,refLabel:'neutral',label:'Fear & Greed Index over the period'});}
 else{empty($('#btcSent'),'No sentiment data in this run.');$('#ch-fng').innerHTML='';}
 if(rd){const tH=Date.parse(rd.halvingPrevisto),dd=Math.max(0,Math.round((tH-Date.now())/DAY)),ini=rd.halvingAltura-210000,pr=Math.max(0,Math.min(100,(rd.altura-ini)/210000*100)),bt=x=>String(x);
  $('#btcHalv').innerHTML=`<div class="big-v"><b>${nf(dd,0)}</b><span class="muted">days (forecast)</span></div><div class="prog" role="img" aria-label="${nf(pr,1)}% of the current cycle"><i style="width:${pr}%"></i></div><div class="fng-lbl"><span>Block ${nf(ini,0)}</span><span>${nf(pr,1)}% of the cycle</span><span>${nf(rd.halvingAltura,0)}</span></div><ul class="kv"><li>Expected date<b>${isFinite(tH)?fdt(tH):'—'}</b></li><li>Remaining<b>${nf(rd.blocosFalta,0)} blocks</b></li><li>Block reward<b>${bt(rd.recompensa)} → ${bt(rd.recompensaNova)} BTC</b></li></ul>`;}
 else empty($('#btcHalv'),'No network data in this run.');
 const R=[];if(rd){R.push(['Current block',nf(rd.altura,0)],['Average block time',nf(rd.minBloco,1)+' min'+(rd.minBlocoEpoca?` · ${nf(rd.minBlocoEpoca,2)} min since the last halving (used for the forecast)`:'')],['Hash rate',rd.hashrateEH!=null?nf(rd.hashrateEH,0)+' EH/s':'—']);if(rd.ajusteDif!=null)R.push(['Next difficulty adjustment',pct(rd.ajusteDif)+(rd.ajusteData?' · '+fdt(Date.parse(rd.ajusteData)):'')]);}
 if(m){R.push(['Market cap',m.capEur!=null?'€'+grande(m.capEur):'—'],['24 h volume',m.volEur!=null?'€'+grande(m.volEur):'—']);if(m.dominio!=null)R.push(['Share of the crypto market',nf(m.dominio,1)+'%']);}
 if(R.length)$('#btcRede').innerHTML=`<ul class="kv" style="margin-top:0">${R.map(r=>`<li>${r[0]}<b>${r[1]}</b></li>`).join('')}</ul>`;else empty($('#btcRede'),'No network or market data in this run.');
 const g2=g?[Object.assign({},g,{moeda:g.moeda||'USD'})]:[],cr=ATIVOS.filter(x=>x.id!=='BTC').concat(g2).map(x=>({x,v:corr(p,inCur(x,'EUR'))})).filter(o=>o.v!=null);
 barChart($('#ch-corr'),{cats:cr.map(o=>nmOf(o.x)),series:[{n:'Correlation',c:'var(--btc)',colors:cr.map(o=>colorOf(o.x)),values:cr.map(o=>o.v)}],dec:2,tdec:1,label:'Correlation of Bitcoin with the other assets'});
 const bn=NEWS.filter(n=>n.empresas.includes('BTC')&&n.nivel!=='white').sort((x,y)=>y.score-x.score||(y.t||0)-(x.t||0)).slice(0,6);
 $('#btcNews').innerHTML=bn.length?bn.map(n=>newsItem(n)).join(''):`<p class="empty">${ic('info')}No relevant Bitcoin news in this window.</p>`;
 drawBtcMore();}

/* ---------- dados guardados neste browser (carteira, compras de Bitcoin, simuladores) ---------- */
/* nada sai do computador: fica no localStorage do browser; o backup em ficheiro protege contra limpezas do browser */
/* set: guarda e, se forem dados da carteira, marca a hora da alteração e grava no ficheiro de backup; raw: só guarda */
const store={get(k,d){try{const v=localStorage.getItem('bb.'+k);return v==null?d:JSON.parse(v);}catch(e){return d;}},set(k,v){try{localStorage.setItem('bb.'+k,JSON.stringify(v));if(/^(buys|lots|sales|deleted|targets|policy|notes|fees)$/.test(k)){localStorage.setItem('bb.savedAt',JSON.stringify(new Date().toISOString()));ALTEROU=true;aoMudar();}return true;}catch(e){return false;}},
 raw(k,v){try{localStorage.setItem('bb.'+k,JSON.stringify(v));return true;}catch(e){return false;}}};
const eur=(v,d=0)=>(v==null||!isFinite(v))?'—':(v<0?'−':'')+'€'+nf(Math.abs(v),d);
/* o sinal segue o valor arredondado: −€0,08 com 0 casas decimais aparece como €0 e não como −€0 */
const eurS=(v,d=0)=>{if(v==null||!isFinite(v))return'—';const r=+Number(v).toFixed(d);return(r>0?'+':r<0?'−':'')+'€'+nf(Math.abs(r),d);};
const btcF=q=>(q==null||!isFinite(q))?'—':Number(q).toLocaleString('en-GB',{maximumFractionDigits:8});
const kpi=(i,l,v,s,c)=>`<div class="kpi"><div class="k">${ic(i)}${l}</div><div class="v ${c||''}">${v}</div><div class="s">${s}</div></div>`;
const HIST={};[...new Set(['AAPL','NVDA','GOOGL','SXR8','BTC'].concat(ATIVOS.map(a=>a.id),['FX']))].forEach(k=>HIST[k]=toPts(D.historico&&D.historico[k]));
/* hoje, em Lisboa, como meia-noite UTC (as datas das compras são comparadas dia a dia) */
const hojeL=()=>{const p=tzParts(Date.now(),LIS);return Date.UTC(+p.year,+p.month-1,+p.day);};
function lastEur(id){const a=ATIVOS.find(x=>x.id===id);if(!a)return null;const p=inCur(a,'EUR');return p.length?p[p.length-1][1]:null;}

/* maiores quedas: do máximo anterior até ao fundo e até voltar a esse máximo */
function quedas(pts,min){const out=[];if(!pts.length)return out;let pk=pts[0],tr=null;const fecha=rec=>{if(tr&&(tr[1]/pk[1]-1)*100<=-min)out.push({pk,tr,rec,d:(tr[1]/pk[1]-1)*100});};
 for(const p of pts){if(p[1]>=pk[1]){fecha(p);pk=p;tr=null;}else if(!tr||p[1]<tr[1])tr=p;}fecha(null);return out;}
function tabQuedas(sel,pts,min,n){const Q=quedas(pts,min).sort((a,b)=>a.d-b.d).slice(0,n),dur=ms=>{const m=Math.round(ms/(30.44*DAY));return m<1?'under 1 month':m<24?m+(m===1?' month':' months'):nf(m/12,1)+' years';};
 $(sel).innerHTML='<thead><tr><th scope="col">Peak</th><th scope="col">Bottom</th><th class="num" scope="col">Drop</th><th class="num" scope="col">Time falling</th><th scope="col">Back at the peak</th><th class="num" scope="col">Time to recover</th></tr></thead><tbody>'+
  (Q.length?Q.map(q=>`<tr><td>${fdt(q.pk[0])}${q.pk===pts[0]?'<span class="td-sub">start of the data, not a real peak</span>':''}</td><td>${fdt(q.tr[0])}</td><td class="num neg">${pct(q.d)}</td><td class="num">${dur(q.tr[0]-q.pk[0])}</td><td>${q.rec?fdt(q.rec[0]):'<span class="tag t-warn">Not yet</span>'}</td><td class="num">${q.rec?dur(q.rec[0]-q.tr[0]):'—'}</td></tr>`).join(''):`<tr><td colspan="6" class="dim">No drops of ${min}% or more in this history.</td></tr>`)+'</tbody>';}

/* ---------- ETF: composição e histórico longo ---------- */
let etfLp='10',etfLog=false;   /* escala do gráfico de longo prazo: linear por omissão */
function drawEtfMore(){const id=etfSel,w=etfInfo(id),T10=arr(w.top10),SE=arr(w.setores),own={AAPL:'var(--aapl)',NVDA:'var(--nvda)',GOOGL:'var(--googl)'},emp=etfTemEmpresas(id);
 hbars($('#ch-top10'),T10.map(x=>({l:x.n,v:+x.w,t:nf(+x.w,2)+'%',c:emp&&own[x.t]||'var(--bar)'})),'Top 10 holdings');
 hbars($('#ch-sect'),SE.map(x=>({l:x.s,v:+x.w,t:nf(+x.w,1)+'%',c:'var(--bar)'})),'Sectors');
 const s10=T10.reduce((a,x)=>a+(+x.w||0),0);
 $('#top10Sub').textContent=T10.length?`Together ${nf(s10,1)}% of the ETF, out of ${w.posicoes||(id==='SXR8'?'about 500':'its')} companies.${emp?' Your three companies are highlighted.':''}`:'Unavailable: the iShares download failed in this run.';
 const h=HIST[id]||[],c=CO[id].c,ano=h.length?new Date(h[0][0]).getUTCFullYear():2010;
 $('#etfHistTitle').textContent=`${id} price in euros`;$('#etfDropsTitle').textContent=`Biggest drops since ${ano}`;$('#elp-max').textContent=`Since ${ano}`;
 if(h.length<2){empty($('#ch-etf-hist'),'No long-term history in this run.');$('#k-etf-hist').innerHTML='';$('#tbl-drops-etf').innerHTML='';empty($('#ch-etf-dd'),'No long-term history in this run.');$('#k-etf-dd').innerHTML='';}
 else{const t1=h[h.length-1][0],p=etfLp==='max'?h:h.filter(x=>x[0]>=t1-(+etfLp)*365.25*DAY),ch=(p[p.length-1][1]/p[0][1]-1)*100,yrs=(p[p.length-1][0]-p[0][0])/(365.25*DAY),cagr=yrs>.9?(Math.pow(p[p.length-1][1]/p[0][1],1/yrs)-1)*100:null;
  $('#k-etf-hist').innerHTML=`<span class="mk-v ${cls(ch)}">${pct(ch,0)}</span><span class="mk-c muted">${cagr!=null?pct(cagr)+' per year':''}</span>`;
  lineMulti($('#ch-etf-hist'),[{n:id,c,pts:p}],{area:true,log:etfLog,dec:2,tdec:0,unit:' €',h:340,label:`${id} long-term price in euros${etfLog?', logarithmic scale':''}`});
  tabQuedas('#tbl-drops-etf',h,10,6);
  const U=underwater(h),uc=U[U.length-1][1],um=U.reduce((a,p)=>Math.min(a,p[1]),0);
  $('#k-etf-dd').innerHTML=`<span class="mk-v ${cls(uc)}">${pct(uc)}</span><span class="mk-c muted">now · worst ${pct(um)}</span>`;
  lineMulti($('#ch-etf-dd'),[{n:'Below the previous high',c:'var(--red)',pts:U}],{area:true,base:0,dec:1,tdec:0,unit:'%',h:260,label:`${id}: drop from its highest close so far`});}
 /* retornos rolantes: só com a secção aberta (fechada, fica vazia até ser aberta) */
 if($('#roll-etf')&&$('#roll-etf').open)drawRoll('etf');else if($('#roll-etf-body'))$('#roll-etf-body').innerHTML='';
 /* um simulador por ETF (o do SXR8 é o de sempre, "etf"); só o do ETF escolhido fica visível */
 ETF_IDS.forEach(x=>{const b=$('#sim-'+simKey(x));if(b)b.hidden=x!==id;});
 drawSim(simKey(id));drawStrat('etf');}
function escolheEtf(id){if(!ehEtf(id))return;etfSel=id;renderEtf();drawEtfMore();}

/* ---------- Bitcoin: histórico longo, queda desde o máximo ---------- */
/* halvings marcados no gráfico: os que o script obteve da rede (incluindo os futuros, quando acontecerem); sem a rede, a lista fixa */
const HALV=[['2016-07-09',2],['2020-05-11',3],['2024-04-20',4]],ordinal=n=>n+(n%10===1&&n%100!==11?'st':n%10===2&&n%100!==12?'nd':n%10===3&&n%100!==13?'rd':'th');
const halvings=()=>{const r=D.bitcoin&&D.bitcoin.rede,l=r&&arr(r.halvings).map(pair).filter(x=>/^\d{4}-\d{2}-\d{2}$/.test(x[0]));return(l&&l.length?l:HALV).map(x=>({t:Date.parse(x[0]+'T00:00:00Z'),l:ordinal(+x[1])+' halving'}));};
function drawBtcMore(){const h=HIST.BTC;
 if(h.length<2){empty($('#ch-btc-hist'),'No long-term history in this run.');empty($('#ch-btc-dd'),'No long-term history in this run.');$('#k-btc-dd').innerHTML='';$('#tbl-drops-btc').innerHTML='';}
 else{lineMulti($('#ch-btc-hist'),[{n:'Bitcoin',c:'var(--btc)',pts:h}],{log:true,dec:0,unit:' €',h:380,label:'Bitcoin price in euros since 2014, logarithmic scale',marks:halvings()});
  let mx=0;const dd=h.map(p=>{mx=Math.max(mx,p[1]);return[p[0],(p[1]/mx-1)*100];}),ath=h.reduce((a,p)=>p[1]>a[1]?p:a,h[0]),cur=dd[dd.length-1][1];
  $('#k-btc-dd').innerHTML=`<span class="mk-v ${cur<0?'neg':''}">${pct(cur)}</span><span class="mk-c muted">high of €${nf(ath[1],0)} on ${fdt(ath[0])}</span>`;
  lineMulti($('#ch-btc-dd'),[{n:'Below the all-time high',c:'var(--red)',pts:dd}],{area:true,base:0,dec:1,tdec:0,unit:'%',h:260,label:'Bitcoin drop from its all-time high'});
  tabQuedas('#tbl-drops-btc',h,40,5);}
 drawSim('btc');drawStrat('btc');drawLots();}

/* ---------- underwater e retornos rolantes ---------- */
/* preço ÷ máximo acumulado − 1, em %, ponto a ponto */
function underwater(pts){let mx=0;return pts.map(p=>{mx=Math.max(mx,p[1]);return[p[0],(p[1]/mx-1)*100];});}
/* Retornos de todas as janelas de N anos, uma a começar em cada ponto: 252 sessões por ano nas séries de dias úteis e 365
   dias na Bitcoin (a mesma deteção de stats()). Sem pontos suficientes, null ("Unavailable"). */
function rolar(pts,anos){if(!pts||pts.length<2)return null;const dias=pts.slice(-30).some(p=>{const w=new Date(p[0]).getUTCDay();return w===0||w===6;}),k=(dias?365:252)*anos,c=pts.map(p=>p[1]);if(c.length<=k)return null;
 const r=[];for(let i=k;i<c.length;i++)r.push(c[i]/c[i-k]-1);const s=r.slice().sort((a,b)=>a-b),n=s.length,med=n%2?s[(n-1)/2]:(s[n/2-1]+s[n/2])/2,an=v=>Math.pow(1+v,1/anos)-1;
 return{k,n,anos,dias,r,min:s[0],med,max:s[n-1],neg:r.filter(v=>v<0).length/n,ann:anos>1?{min:an(s[0]),med:an(med),max:an(s[n-1])}:null};}
const ROLLH={etf:1,btc:1};   /* horizonte do histograma em cada separador */
function drawRoll(tab){const body=$(`#roll-${tab}-body`);if(!body)return;const id=tab==='btc'?'BTC':etfSel,h=HIST[id]||[],Y=[1,3,5],R=Y.map(a=>rolar(h,a)),nm=tab==='btc'?'Bitcoin':id,r0=R.find(r=>r),dias=r0?r0.dias:tab==='btc';
 const lab=a=>a===1?'1 year':a+' years',v=(x,y)=>`${pct(x*100)}${y!=null?`<span class="td-sub">${pct(y*100)} a year</span>`:''}`;
 body.innerHTML=`<p class="fig-sub">${esc(nm)}, daily closes in euros${h.length?` from ${fdt(h[0][0])} to ${fdt(h[h.length-1][0])}`:''}.</p>`+
  '<div class="tbl-wrap"><table><thead><tr><th scope="col">Period</th><th class="num" scope="col">Windows</th><th class="num" scope="col">Worst</th><th class="num" scope="col">Median</th><th class="num" scope="col">Best</th><th class="num" scope="col">Negative windows</th></tr></thead><tbody>'+
  R.map((r,i)=>!r?`<tr><td>${lab(Y[i])}</td><td colspan="5" class="dim">Unavailable: the history is shorter than ${lab(Y[i])}.</td></tr>`:
   `<tr><td>${lab(Y[i])}</td><td class="num">${nf(r.n,0)}</td><td class="num ${cls(r.min)}">${v(r.min,r.ann&&r.ann.min)}</td><td class="num ${cls(r.med)}">${v(r.med,r.ann&&r.ann.med)}</td><td class="num ${cls(r.max)}">${v(r.max,r.ann&&r.ann.max)}</td><td class="num">${nf(r.neg*100,1)}%</td></tr>`).join('')+'</tbody></table></div>'+
  `<div class="toolbar"><div class="tb-group"><span class="f-label">${ic('bars')}Histogram</span><div class="seg" role="group" id="roll-${tab}-h">${Y.map((a,i)=>`<button type="button" data-y="${a}" aria-pressed="${ROLLH[tab]===a}"${R[i]?'':' disabled'}>${lab(a)}</button>`).join('')}</div></div></div><div id="roll-${tab}-ch"></div>`+
  `<p class="note">${ic('info')}<span>One window starts on each ${dias?'day (1 year = 365 calendar days)':'trading session (1 year = 252 sessions)'}, so consecutive windows overlap: they are not independent results, and the number of windows is not a number of separate experiences. Price returns in euros${tab==='btc'?'':' (dividends are in the price of these accumulating ETFs)'}. Each bar is the share of windows whose return fell in that range, labelled by its lower bound. Past returns say nothing about future returns.</span></p>`;
 $$(`#roll-${tab}-h button`).forEach(b=>b.addEventListener('click',()=>{ROLLH[tab]=+b.dataset.y;drawRoll(tab);}));
 const j=R[Y.indexOf(ROLLH[tab])]?Y.indexOf(ROLLH[tab]):R.findIndex(r=>r),r=R[j];if(!r){empty($(`#roll-${tab}-ch`),'Not enough history for any window.');return;}
 const vals=r.r.map(x=>x*100),mn=Math.min(...vals),mx=Math.max(...vals),step=[1,2,5,10,20,25,50,100,200,500,1000,2000,5000].find(s=>(Math.ceil(mx/s)-Math.floor(mn/s))<=12)||10000,lo=Math.floor(mn/step)*step,nb=Math.max(1,Math.ceil(mx/step)-Math.floor(mn/step)),cnt=new Array(nb).fill(0);
 vals.forEach(x=>{cnt[Math.min(nb-1,Math.floor((x-lo)/step))]++;});
 barChart($(`#roll-${tab}-ch`),{cats:cnt.map((_,i)=>`${nf(lo+i*step,0)}%`),series:[{n:'Share of windows',c:'var(--bar)',colors:cnt.map((_,i)=>lo+(i+1)*step<=0?'var(--neg)':'var(--pos)'),values:cnt.map(x=>x/r.n*100)}],dec:1,unit:'%',label:`${nm} rolling returns over ${lab(Y[j])}: share of windows in each range of ${nf(step,0)} points`});}

/* ---------- simulador de investimento mensal ---------- */
/* compra no primeiro dia de negociação de cada mês, ao preço de fecho, a partir do mês escolhido (sem comissões nem impostos) */
function simular(pts,amt,from){const r={inv:0,u:0,n:0,serie:[],last:null};if(!pts.length||!(amt>0))return r;const f=Date.parse(from+'-01T00:00:00Z');if(!isFinite(f))return r;let mes='';
 for(const p of pts){if(p[0]<f)continue;const d=new Date(p[0]),m=d.getUTCFullYear()*12+d.getUTCMonth();if(m!==mes){mes=m;r.inv+=amt;r.u+=amt/p[1];r.n++;}r.serie.push([p[0],r.inv,r.u*p[1]]);}
 r.last=pts[pts.length-1][1];return r;}
const SIM={};
function buildSim(key,cfg){const box=$('#sim-'+key);if(!box)return;SIM[key]=cfg;const h=cfg.pts(),minM=h.length?new Date(h[0][0]).toISOString().slice(0,7):'2010-01',maxM=new Date().toISOString().slice(0,7);
 const sv=store.get('sim.'+key,{}),d5=new Date();d5.setUTCFullYear(d5.getUTCFullYear()-5);let from=/^\d{4}-\d{2}$/.test(sv.from||'')?sv.from:d5.toISOString().slice(0,7);if(from<minM)from=minM;
 box.innerHTML=`<div class="fig-head"><div><p class="fig-title">Monthly investment simulator</p><p class="fig-sub">What a fixed amount invested every month in ${esc(cfg.name)} would be worth today, buying on the first trading day of each month at the closing price.</p></div></div>`+
  `<div class="sim-form"><div class="field"><label class="f-label" for="sim-${key}-amt">${ic('euro')}Per month (€)</label><input class="in" type="number" id="sim-${key}-amt" min="1" step="any" value="${+sv.amt>0?+sv.amt:100}"></div><div class="field"><label class="f-label" for="sim-${key}-from">${ic('calendar')}Starting in</label><input class="in" type="month" id="sim-${key}-from" min="${minM}" max="${maxM}" value="${from}"></div></div>`+
  `<div class="kpis sim-kpis" id="sim-${key}-kpis"></div><div id="sim-${key}-ch"></div><p class="note">${ic('info')}${esc(cfg.note)}</p>`;
 ['amt','from'].forEach(f=>$(`#sim-${key}-${f}`).addEventListener('input',()=>drawSim(key)));}
function drawSim(key){const cfg=SIM[key];if(!cfg)return;const amt=+$(`#sim-${key}-amt`).value,from=$(`#sim-${key}-from`).value,h=cfg.pts();
 if(amt>0&&/^\d{4}-\d{2}$/.test(from))store.set('sim.'+key,{amt,from});const r=simular(h,amt,from);
 if(!r.n){$(`#sim-${key}-kpis`).innerHTML=`<p class="empty">${ic('info')}${h.length?'Choose an amount and a start month within the available history.':'No long-term history in this run.'}</p>`;$(`#sim-${key}-ch`).innerHTML='';return;}
 const val=r.u*r.last,g=val-r.inv;
 $(`#sim-${key}-kpis`).innerHTML=kpi('euro','Invested',eur(r.inv),`${r.n} monthly purchases`)+kpi('wallet','Worth today',eur(val),cfg.units(r.u))+kpi('chart','Gain',eurS(g),pct(g/r.inv*100),cls(g))+kpi('tag','Average price paid',eur(r.inv/r.u,cfg.dec),`price now ${eur(r.last,cfg.dec)}`);
 lineMulti($(`#sim-${key}-ch`),[{n:'Value',c:cfg.c,pts:r.serie.map(x=>[x[0],x[2]])},{n:'Amount invested',c:'var(--muted)',pts:r.serie.map(x=>[x[0],x[1]])}],{dec:0,tdec:0,unit:' €',h:300,label:'Simulated value against the amount invested'});}

/* ---------- comparação de estratégias com o mesmo dinheiro e o mesmo período do simulador mensal ---------- */
/* (a) compra mensal fixa: simular(), como no simulador acima. (b) tudo de uma vez: o total investido em (a), todo no primeiro
   dia de compra. (c) "buy the dip": o valor mensal junta-se em caixa (a 0 %) no primeiro dia de negociação de cada mês e é
   todo investido quando o fecho fica X % ou mais abaixo do máximo dos fechos das 52 semanas anteriores (365 dias, sem o
   próprio dia); no fim, a caixa por investir conta pelo valor nominal. Sem comissões nem impostos. */
function estrategias(pts,amt,from,X){const a=simular(pts,amt,from);if(!a.n||!(X>0&&X<100))return null;const f=Date.parse(from+'-01T00:00:00Z'),last=pts[pts.length-1][1];
 let i0=0;while(i0<pts.length&&pts[i0][0]<f)i0++;const uB=a.inv/pts[i0][1],sB=[],sC=[],dq=[];let mes='',cash=0,uC=0,compras=0;
 for(let i=0;i<pts.length;i++){const p=pts[i];while(dq.length&&pts[dq[0]][0]<=p[0]-365*DAY)dq.shift();
  if(i>=i0){const d=new Date(p[0]),m=d.getUTCFullYear()*12+d.getUTCMonth();if(m!==mes){mes=m;cash+=amt;}
   const mx=dq.length?pts[dq[0]][1]:null;if(cash>0&&mx!=null&&p[1]<=mx*(1-X/100)){uC+=cash/p[1];cash=0;compras++;}
   sB.push([p[0],uB*p[1]]);sC.push([p[0],uC*p[1]+cash]);}
  while(dq.length&&pts[dq[dq.length-1]][1]<=p[1])dq.pop();dq.push(i);}
 return{tot:a.inv,n:a.n,t0:pts[i0][0],a:{val:a.u*last,serie:a.serie.map(x=>[x[0],x[2]])},b:{val:uB*last,serie:sB},c:{val:uC*last+cash,cash,compras,serie:sC},X};}
function drawStrat(tab){const box=$('#strat-'+tab);if(!box)return;const key=tab==='btc'?'btc':simKey(etfSel),cfg=SIM[key];if(!cfg)return;
 if(!box.dataset.ok){box.dataset.ok='1';const sv=store.get('strat.'+tab,{}),X0=+sv.dip>0&&+sv.dip<100?+sv.dip:10;
  box.innerHTML=`<div class="fig-head"><div><p class="fig-title">Three ways to invest the same money</p><p class="fig-sub">The same total and the same period as the monthly simulator above, invested in three ways.</p></div></div>`+
   `<div class="sim-form"><div class="field"><label class="f-label" for="strat-${tab}-dip">${ic('trend-down')}Buy-the-dip threshold (% below the 52-week high)</label><input class="in" type="number" id="strat-${tab}-dip" min="1" max="90" step="any" value="${X0}"></div></div>`+
   `<div class="kpis sim-kpis" id="strat-${tab}-kpis"></div><div id="strat-${tab}-ch"></div><p class="note">${ic('info')}<span id="strat-${tab}-note"></span></p>`;
  $(`#strat-${tab}-dip`).addEventListener('input',()=>drawStrat(tab));
  /* acompanha o simulador mensal (montante e mês de início) de cada chave deste separador */
  Object.keys(SIM).filter(k=>tab==='btc'?k==='btc':k!=='btc').forEach(k=>['amt','from'].forEach(f=>{const i=$(`#sim-${k}-${f}`);if(i)i.addEventListener('input',()=>drawStrat(tab));}));}
 const X=+$(`#strat-${tab}-dip`).value,amt=+($(`#sim-${key}-amt`)||{}).value,from=($(`#sim-${key}-from`)||{}).value,h=cfg.pts();if(X>0&&X<100)store.set('strat.'+tab,{dip:X});
 const R=/^\d{4}-\d{2}$/.test(from||'')&&amt>0&&X>0&&X<100?estrategias(h,amt,from,X):null,k=$(`#strat-${tab}-kpis`),ch=$(`#strat-${tab}-ch`),nt=$(`#strat-${tab}-note`);
 if(!R){k.innerHTML=`<p class="empty">${ic('info')}${h.length?'Choose an amount and a start month in the simulator above, and a percentage between 1 and 90.':'No long-term history in this run.'}</p>`;ch.innerHTML='';nt.textContent='';return;}
 const g=v=>`${eurS(v-R.tot)} · ${pct((v/R.tot-1)*100)}`;
 k.innerHTML=kpi('euro','Money put in',eur(R.tot),`${R.n} months from ${fdt(R.t0)}`)+kpi('calendar','Monthly purchases',eur(R.a.val),g(R.a.val),cls(R.a.val-R.tot))+
  kpi('wallet','All at once',eur(R.b.val),g(R.b.val),cls(R.b.val-R.tot))+kpi('trend-down',`Buy the dip (${nf(R.X,R.X%1?1:0)}%)`,eur(R.c.val),`${g(R.c.val)} · ${R.c.compras} purchase${R.c.compras===1?'':'s'}${R.c.cash>0?`, ${eur(R.c.cash)} still in cash`:''}`,cls(R.c.val-R.tot));
 lineMulti(ch,[{n:'Monthly purchases',c:cfg.c,pts:R.a.serie},{n:'All at once',c:'var(--accent)',pts:R.b.serie},{n:`Buy the dip (${nf(R.X,R.X%1?1:0)}%)`,c:'var(--ink)',pts:R.c.serie}],{dec:0,tdec:0,unit:' €',h:300,label:'Value of the three strategies over time'});
 nt.textContent=`${eur(R.tot)} in total: ${eur(amt)} a month from ${fdt(R.t0)}. All at once invests that total on the first day. Buy the dip keeps each month's amount in cash (earning nothing) and invests all of it when the close is ${nf(R.X,R.X%1?1:0)}% or more below the highest close of the previous 52 weeks; cash not invested at the end counts at face value. Fees, spreads and taxes are not included. A comparison of the past, not a forecast or advice.`;}

/* ---------- carteira: compras, vendas e desdobramentos ---------- */
/* Guardado neste browser: bb.buys (compras de ações e ETF {id,a,d,q,p,r,u}), bb.lots (compras de Bitcoin {id,d,q,c}),
   bb.sales (vendas {id,a,d,q,p,r,u}), bb.deleted (ids apagados, para não voltarem ao juntar com o ficheiro),
   bb.savedAt (última alteração), bb.fileSaved / bb.fileWrittenAt (última gravação no ficheiro) e bb.sim.* */
const PF=[['AAPL','Apple','shares'],['NVDA','NVIDIA','shares'],['GOOGL','Alphabet','shares'],['SXR8','iShares Core S&P 500 (SXR8)','ETF units']]
 .concat(ETF_IDS.filter(x=>x!=='SXR8').map(x=>[x,(ATIVOS.find(a=>a.id===x)||{}).nome||x,'ETF units']),[['BTC','Bitcoin','BTC']]);   /* os outros ETF vêm da configuração do script */
/* COMPRAVEIS: ativos do formulário. Compras de ações e ETF ficam em bb.buys; as de Bitcoin em bb.lots (a lista do
   contador dos 365 dias); as vendas de todos em bb.sales */
const COMPRAVEIS=PF.map(x=>x[0]),BUY_IDS=COMPRAVEIS.filter(id=>id!=='BTC'),nomePF=id=>(PF.find(x=>x[0]===id)||[id,id])[1];
const novoId=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,6);
/* LEGADO: true só na primeira abertura com esta versão, se o browser ainda não tinha a lista de apagados (ver juntaBackup) */
let LEGADO=false;
const apagados=()=>{const d=store.get('deleted',{});return d&&typeof d==='object'&&!Array.isArray(d)?d:{};};
/* apaga uma entrada e guarda o id como apagado (assim o ficheiro de backup não a traz de volta) */
function apaga(k,id){const L=arr(store.get(k,[]));if(!L.some(x=>String(x.id)===String(id)))return false;const del=apagados();del[String(id)]=new Date().toISOString();store.raw('deleted',del);
 const N=store.get('notes',null);if(N&&typeof N==='object'&&String(id) in N){delete N[String(id)];store.raw('notes',N);}  /* a nota de uma entrada apagada sai com ela */
 const Fm=store.get('fees',null);if(Fm&&typeof Fm==='object'&&String(id) in Fm){delete Fm[String(id)];store.raw('fees',Fm);}  /* e a comissão */
 return store.set(k,L.filter(x=>String(x.id)!==String(id)));}

/* Desdobramentos (splits): os que o Yahoo indica no histórico, mais esta lista fixa como reserva (se o Yahoo não os der).
   O histórico do Yahoo vem ajustado, por isso uma compra é sempre contada em ações "de hoje". */
const SPLITS_FIXOS={AAPL:[['2014-06-09',7],['2020-08-31',4]],NVDA:[['2021-07-20',4],['2024-06-10',10]],GOOGL:[['2014-04-03',1.998],['2022-07-18',20]]};
const SPLITS={};BUY_IDS.forEach(id=>{const y=arr(D.splits&&D.splits[id]).map(pair).filter(s=>/^\d{4}-\d{2}-\d{2}$/.test(s[0])&&+s[1]>0&&+s[1]!==1).map(s=>[s[0],+s[1]]);
 const perto=(a,b)=>Math.abs(Date.parse(a)-Date.parse(b))<=3*DAY;arr(SPLITS_FIXOS[id]).forEach(s=>{if(!y.some(x=>perto(x[0],s[0])))y.push(s);});SPLITS[id]=y.sort((a,b)=>a[0]<b[0]?-1:1);});
const fatorSplit=(id,iso)=>arr(SPLITS[id]).filter(s=>iso<s[0]).reduce((a,s)=>a*s[1],1);
/* último fecho no dia pedido ou antes dele (fins de semana e feriados usam a sessão anterior) */
function precoEm(pts,iso){const t=Date.parse(iso+'T00:00:00Z');if(!pts||!pts.length||!isFinite(t)||t<pts[0][0])return null;let lo=0,hi=pts.length-1;while(lo<hi){const m=(lo+hi+1)>>1;if(pts[m][0]<=t)lo=m;else hi=m-1;}return pts[lo];}
/* preço de fecho em euros num dia: o SXR8 já está em euros; as ações americanas são convertidas ao EUR/USD desse mesmo dia */
function precoEurEm(id,iso){const a=ATIVOS.find(x=>x.id===id),p=precoEm(HIST[id],iso);if(!a||!p)return null;if((a.moeda||'').toUpperCase()!=='USD')return{t:p[0],eur:p[1],nat:p[1]};
 const f=fxAt(p[0],HIST.FX);return f?{t:p[0],eur:p[1]/f,usd:p[1],fx:f,nat:p[1]}:null;}
/* Cada compra ou venda guarda r = [data, preço de fecho dessa data no histórico usado ao registá-la] e u = data da base
   desse histórico. Se, numa execução futura, o histórico mostrar para essa data um preço k vezes menor, houve um
   desdobramento de k:1 depois do registo: a quantidade passa a q×k e o preço a p/k (o valor investido não muda). */
function ajusteSplit(x){if(!x||x.a==='BTC'||!BUY_IDS.includes(x.a))return{k:1,aviso:''};const S=arr(SPLITS[x.a]),h=HIST[x.a]||[];
 if(Array.isArray(x.r)&&x.r.length===2&&+x.r[1]>0&&h.length){const pt=precoEm(h,x.r[0]);if(pt&&isoU(pt[0])===x.r[0]){const ratio=+x.r[1]/pt[1],dep=S.filter(s=>s[0]>x.r[0]).map(s=>s[1]),cand=[1];
   for(let i=dep.length-1,f=1;i>=0;i--){f*=dep[i];cand.push(f);}const k=cand.find(c=>Math.abs(ratio/c-1)<0.03);if(k!=null)return{k,aviso:''};
   return{k:1,aviso:'The price history for this date changed since this entry was recorded, without a known split: check the quantity.'};}}
 if(/^\d{4}-\d{2}-\d{2}$/.test(x.u||'')){const ate=h.length?isoU(h[h.length-1][0]):isoU(GEN);return{k:S.filter(s=>s[0]>x.u&&s[0]<=ate).reduce((a,s)=>a*s[1],1),aviso:''};}
 return{k:1,aviso:''};}
const efetiva=x=>{const j=ajusteSplit(x);return Object.assign({},x,{q:+x.q*j.k,p:+x.p/j.k,k:j.k,aviso:j.aviso});};
/* completa r e u nas entradas antigas (sem eles), com a data em que estavam guardadas como base. Só calcula r se não
   houver nenhum desdobramento depois dessa base (senão o histórico de hoje já não está na base delas e fica só u) */
function completaBase(L,base){let mudou=false;arr(L).forEach(x=>{if(!x||!BUY_IDS.includes(x.a))return;if(!/^\d{4}-\d{2}-\d{2}$/.test(x.u||'')){x.u=base;mudou=true;}
 if(!Array.isArray(x.r)&&!arr(SPLITS[x.a]).some(s=>s[0]>x.u)){const p=precoEm(HIST[x.a],x.d);if(p){x.r=[isoU(p[0]),p[1]];mudou=true;}}});return mudou;}
const baseDe=v=>{const t=Date.parse(v||'');return isFinite(t)?isoU(t):isoU(GEN);};
/* as compras da versão anterior (só do ETF, em bb.etfLots) passam para a lista geral; as entradas sem base ganham-na.
   Nada disto muda a data da última alteração (bb.savedAt): são os mesmos dados */
function migraCompras(){try{LEGADO=localStorage.getItem('bb.deleted')==null;if(LEGADO)store.raw('deleted',{});}catch(e){}
 try{const old=localStorage.getItem('bb.etfLots');if(old!=null){const o=arr(JSON.parse(old)),b=arr(store.get('buys',[]));
  store.raw('buys',b.concat(o.map(x=>Object.assign({a:'SXR8'},x))));localStorage.removeItem('bb.etfLots');localStorage.removeItem('bb.portfolio');}}catch(e){}
 const base=baseDe(store.get('savedAt',null));['buys','sales'].forEach(k=>{const L=arr(store.get(k,[]));if(completaBase(L,base))store.raw(k,L);});}

/* Posições: compras menos vendas, primeiro a entrar, primeiro a sair (FIFO): cada venda gasta as compras mais antigas
   feitas até à data da venda. Devolve os lotes com a quantidade que resta e as vendas com o custo e o ganho realizado.
   Cada venda leva também a lista dos lotes que gastou (usados: id, data, quantidade e custo unitário, em ações de hoje),
   que é o que a exportação para o IRS usa: é este o único cálculo FIFO da página. */
function carteira(extra){const lotes={};
 arr(store.get('buys',[])).map(efetiva).forEach(x=>{if(!(x.q>0))return;(lotes[x.a]=lotes[x.a]||[]).push({id:x.id,a:x.a,d:x.d,q0:x.q,q:x.q,cu:x.p,k:x.k,aviso:x.aviso,st:'buys'});});
 arr(store.get('lots',[])).forEach(x=>{if(!(+x.q>0))return;(lotes.BTC=lotes.BTC||[]).push({id:x.id,a:'BTC',d:x.d,q0:+x.q,q:+x.q,cu:(+x.c||0)/+x.q,k:1,aviso:'',st:'lots'});});
 Object.values(lotes).forEach(l=>l.sort((a,b)=>a.d<b.d?-1:a.d>b.d?1:String(a.id)<String(b.id)?-1:1));
 const V=arr(store.get('sales',[])).concat(extra?[extra]:[]).map(efetiva).filter(v=>v.q>0).sort((a,b)=>a.d<b.d?-1:a.d>b.d?1:String(a.id)<String(b.id)?-1:1);
 const vendas=V.map(v=>{let falta=v.q,custo=0;const usados=[];for(const x of(lotes[v.a]||[])){if(falta<=1e-12)break;if(x.d>v.d||x.q<=1e-12)continue;const u=Math.min(x.q,falta);x.q=x.q-u<1e-12?0:x.q-u;falta-=u;custo+=u*x.cu;usados.push({id:x.id,d:x.d,q:u,cu:x.cu,k:x.k,aviso:x.aviso,st:x.st});}
  if(falta<1e-9)falta=0;return Object.assign({},v,{custo,falta,usados,real:v.p>0?(v.q-falta)*v.p-custo:null});});
 return{lotes,vendas};}

/* ---------- Bitcoin: contador dos 365 dias ---------- */
function buildLots(){const f=$('#lotForm');if(!f)return;$('#lotDate').max=isoU(hojeL());
 f.addEventListener('submit',e=>{e.preventDefault();const d=$('#lotDate').value,q=+$('#lotQty').value,c=+$('#lotCost').value,msg=$('#lotMsg');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(d)||!(q>0)||!(c>=0)||$('#lotCost').value===''){msg.textContent='Fill in the date, the amount of BTC and the total you paid in euros.';return;}
  if(Date.parse(d+'T00:00:00Z')>hojeL()){msg.textContent='The purchase date cannot be in the future.';return;}
  const L=store.get('lots',[]);L.push({id:novoId(),d,q,c});
  if(!store.set('lots',L)){msg.textContent='Could not save: this browser is blocking local storage.';return;}
  f.reset();msg.textContent='Purchase added.';drawLots();drawBuys();drawPf();});
 $('#tbl-lots').addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-del]');if(!b)return;
  const x=arr(store.get('lots',[])).find(y=>String(y.id)===b.dataset.del);if(!x||!confirm(`Remove the Bitcoin purchase of ${btcF(+x.q)} BTC on ${fdt(Date.parse(x.d+'T00:00:00Z'))}? This cannot be undone.`))return;
  apaga('lots',x.id);$('#lotMsg').textContent='Purchase removed.';drawLots();drawBuys();drawPf();});}
function drawLots(){if(!$('#tbl-lots'))return;const px=lastEur('BTC'),hoje=hojeL(),C=carteira();
 const R=arr(C.lotes.BTC).slice().map(x=>{const t=Date.parse(x.d+'T00:00:00Z'),livre=t+365*DAY,val=px!=null?x.q*px:null,c=x.q*x.cu;return Object.assign({},x,{t,livre,held:Math.floor((hoje-t)/DAY),ok:hoje>=livre,val,c,g:val!=null?val-c:null});});
 const vivos=R.filter(x=>x.q>0),soma=(r,f)=>r.reduce((a,x)=>a+(f(x)||0),0),fr=vivos.filter(x=>x.ok),tx=vivos.filter(x=>!x.ok),nx=tx.slice().sort((a,b)=>a.livre-b.livre)[0];
 $('#lotSum').innerHTML=R.length?kpi('bitcoin','Total held',btcF(soma(vivos,x=>x.q))+' BTC',`worth ${eur(soma(vivos,x=>x.val))} · cost ${eur(soma(vivos,x=>x.c))}`)+
  kpi('check','Tax-free now',btcF(soma(fr,x=>x.q))+' BTC',fr.length?`gain ${eurS(soma(fr,x=>x.g))}`:'none yet','pos')+
  kpi('clock','Still taxable',btcF(soma(tx,x=>x.q))+' BTC',tx.length?`gain ${eurS(soma(tx,x=>x.g))}`:'nothing left to wait for')+
  kpi('calendar','Next to become tax-free',nx?fdt(nx.livre):'—',nx?`in ${dias(Math.ceil((nx.livre-hoje)/DAY))} · ${btcF(nx.q)} BTC`:'all your Bitcoin is tax-free'):'';
 $('#tbl-lots').innerHTML='<thead><tr><th scope="col">Bought on</th><th class="num" scope="col">BTC</th><th class="num" scope="col">Cost</th><th class="num" scope="col">Worth now</th><th class="num" scope="col">Gain</th><th scope="col">365-day status</th><th scope="col"><span class="sr">Remove</span></th></tr></thead><tbody>'+
  (R.length?R.map(x=>{const pr=Math.max(0,Math.min(100,x.held/365*100)),vend=x.q<x.q0-1e-12;return`<tr class="${x.q>0?'':'row-sold'}"><td>${fdt(x.t)}</td><td class="num">${btcF(x.q)}${vend?`<span class="td-sub">of ${btcF(x.q0)} bought</span>`:''}</td><td class="num">${eur(x.c,2)}</td><td class="num">${x.q>0?eur(x.val,2):'—'}</td><td class="num ${x.q>0?cls(x.g):''}">${x.q>0?eurS(x.g,2)+(x.c>0&&x.g!=null?` <small>${pct(x.g/x.c*100)}</small>`:''):'—'}</td>`+
   `<td><div class="lot-st">${x.q<=0?`<span class="tag">${ic('check')}Sold</span>`:x.ok?`<span class="tag t-pos">${ic('check')}Tax-free since ${fdt(x.livre)}</span>`:`<span class="tag t-warn">${ic('clock')}${dias(Math.ceil((x.livre-hoje)/DAY))} to go · tax-free on ${fdt(x.livre)}</span>`}<span class="prog sm" aria-hidden="true"><i style="width:${pr}%"></i></span></div></td>`+
   `<td><button class="btn btn-ic" type="button" data-del="${esc(x.id)}" aria-label="Remove the purchase of ${fdt(x.t)}">${ic('trash')}</button></td></tr>`;}).join(''):'<tr><td colspan="7" class="dim">No purchases yet. Add your first one above.</td></tr>')+'</tbody>';}

/* ---------- carteira: registo de compras e vendas (o preço do dia é preenchido a partir do histórico) ---------- */
function buildBuys(){const f=$('#buyForm');if(!f)return;const sel=$('#buyAsset'),tipo=$('#buyType'),dt=$('#buyDate'),pr=$('#buyPrice'),msg=$('#buyMsg');dt.max=isoU(hojeL());
 sel.innerHTML=COMPRAVEIS.map(id=>`<option value="${id}">${esc(nomePF(id))}</option>`).join('');sel.value='SXR8';
 const venda=()=>tipo.value==='sell';
 const rotulos=()=>{$('#buySubmit').innerHTML=ic(venda()?'minus':'plus')+(venda()?'Add sale':'Add purchase');};
 const lim=()=>{const h=HIST[sel.value]||[];dt.min=h.length?isoU(h[0][0]):'';};
 const preenche=()=>{lim();rotulos();if(!/^\d{4}-\d{2}-\d{2}$/.test(dt.value))return;const id=sel.value,t=Date.parse(dt.value+'T00:00:00Z'),r=precoEurEm(id,dt.value),h=HIST[id]||[];
  if(!r){pr.value='';msg.textContent=h.length?`No price available for this date: enter the price ${venda()?'you sold at':'you paid'}, in euros.`:`No price history in this run: enter the price ${venda()?'you sold at':'you paid'}, in euros.`;return;}
  pr.value=+r.eur.toFixed(2);const a=ATIVOS.find(x=>x.id===id)||{},last=h[h.length-1][0],conv=r.usd!=null?` ($${nf(r.usd,2)} at EUR/USD ${nf(r.fx,4)})`:'',onde=BOLSA_DE[id]==='DE'?'Xetra closing price':id==='BTC'?'BTC/EUR price at the end of the day (UTC)':'Closing price';
  /* o último ponto pode ser um preço intradiário (sessão ainda aberta quando a página foi gerada): não é um fecho */
  const vivo=r.t===last&&a.parcial&&a.hora&&isFinite(Date.parse(a.hora));
  msg.textContent=(vivo?`Price of ${nomePF(id)} at ${hmL(Date.parse(a.hora))} on ${fdt(r.t)}, while the session was still open (not the closing price): ${eur(r.eur,2)}${conv}.`:
   r.t===t?`${onde} of ${nomePF(id)} on ${fdt(r.t)}: ${eur(r.eur,2)}${conv}.`:t>last?`Prices in this page go up to ${fdt(last)}: using that price, ${eur(r.eur,2)}${conv}.`:
   `No trading on ${fdt(t)}: using the last close before it, ${fdt(r.t)}, ${eur(r.eur,2)}${conv}.`)+` Change it if you ${venda()?'sold':'paid'} at a different price.`;
  const fs=fatorSplit(id,dt.value);if(fs>1)msg.textContent+=` ${nomePF(id)} has split its shares since then (×${nf(fs,fs%1?3:0)}), so this price is adjusted: enter the quantity in today's shares (${nf(fs,fs%1?3:0)}× the shares of that day).`;};
 [sel,tipo].forEach(x=>x.addEventListener('change',preenche));dt.addEventListener('change',preenche);dt.addEventListener('input',preenche);lim();rotulos();
 f.addEventListener('submit',e=>{e.preventDefault();const a=sel.value,d=dt.value,q=+$('#buyQty').value,p=+pr.value,v=venda();
  if(!COMPRAVEIS.includes(a)||!/^\d{4}-\d{2}-\d{2}$/.test(d)||!(q>0)||!(p>0)){msg.textContent='Fill in the asset, the date, the quantity and the price per unit.';return;}
  if(Date.parse(d+'T00:00:00Z')>hojeL()){msg.textContent=`The ${v?'sale':'purchase'} date cannot be in the future.`;return;}
  /* base para os desdobramentos: o preço desse dia no histórico desta página e a data da base desse histórico */
  const h=HIST[a]||[],pt=a!=='BTC'?precoEm(h,d):null,base=a==='BTC'?{}:Object.assign({u:h.length?isoU(h[h.length-1][0]):isoU(GEN)},pt?{r:[isoU(pt[0]),pt[1]]}:{});
  /* recusa só o que esta venda estraga: a própria sem unidades, ou uma venda que passe a ter falta (ou mais falta) do que
     tinha. Uma venda antiga já sem compras (dados antigos ou restaurados, de qualquer ativo) não impede as outras (B14) */
  if(v){const f0={};carteira().vendas.forEach(x=>{f0[x.id]=x.falta;});const nova=Object.assign({id:novoId(),a,d,q,p},base),C=carteira(nova),mal=C.vendas.find(x=>x.falta>(f0[x.id]||0)+1e-9);
   if(mal){const tem=arr(C.lotes[a]).filter(x=>x.d<=d).reduce((s,x)=>s+x.q0,0);msg.textContent=mal.id===nova.id?`You did not hold enough ${nomePF(a)} on ${fdt(Date.parse(d+'T00:00:00Z'))} for this sale (purchases up to that day: ${btcF(tem)}, minus earlier sales). Record the purchase first.`:'This sale would leave a later sale without enough units to sell. Check the dates.';return;}
   const L=arr(store.get('sales',[]));L.push(nova);if(!store.set('sales',L)){msg.textContent='Could not save: this browser is blocking local storage.';return;}guardaFeeForm(nova.id);
   msg.textContent=`Sale saved: ${btcF(q)} × ${nomePF(a)} on ${fdt(Date.parse(d+'T00:00:00Z'))} at ${eur(p,2)} each. It uses up your oldest purchases first.`;}
  else{const btc=a==='BTC',k=btc?'lots':'buys',L=arr(store.get(k,[])),nv=btc?{id:novoId(),d,q,c:q*p}:Object.assign({id:novoId(),a,d,q,p},base);L.push(nv);
   if(!store.set(k,L)){msg.textContent='Could not save: this browser is blocking local storage.';return;}guardaFeeForm(nv.id);
   msg.textContent=`Purchase saved: ${btcF(q)} × ${nomePF(a)} on ${fdt(Date.parse(d+'T00:00:00Z'))} at ${eur(p,2)} each.${btc?' It also appears in the 365-day counter of the Bitcoin section.':''}`;}
  dt.value='';pr.value='';$('#buyQty').value=1;$('#buyFee').value='';drawBuys();drawLots();drawPf();});
 $('#tbl-buys').addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-fee]');if(!b)return;const id=String(b.dataset.fee),F=limpaFees(store.get('fees',null));
  const v=prompt('Fee for this entry, in euros (leave it empty to remove it):',F[id]?String(F[id].v):'');if(v==null)return;const s=String(v).trim().replace(',','.');
  if(s===''){delete F[id];store.set('fees',F);msg.textContent='Fee removed.';drawBuys();drawPf();return;}
  if(!(isFinite(+s)&&+s>=0)){msg.textContent='The fee must be a number of 0 or more, in euros.';return;}
  F[id]={v:+s,at:new Date().toISOString()};store.set('fees',F);msg.textContent='Fee saved.';drawBuys();drawPf();});
 $('#tbl-buys').addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-del]');if(!b)return;const k=['lots','sales'].includes(b.dataset.store)?b.dataset.store:'buys';
  const x=arr(store.get(k,[])).find(y=>String(y.id)===b.dataset.del);if(!x)return;const what=k==='sales'?'sale':'purchase';
  if(!confirm(`Remove the ${what} of ${nomePF(k==='lots'?'BTC':x.a)} on ${fdt(Date.parse(x.d+'T00:00:00Z'))}? This cannot be undone.`))return;
  apaga(k,x.id);msg.textContent=k==='sales'?'Sale removed.':'Purchase removed.';drawBuys();drawLots();drawPf();});
 /* nota por entrada: guardada à parte (notes[id] = {t, at}), sem mudar a entrada */
 $('#tbl-buys').addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-note]');if(!b)return;const id=String(b.dataset.note),N=limpaNotas(store.get('notes',null));
  const v=prompt('Note for this entry (leave it empty to remove the note):',N[id]?N[id].t:'');if(v==null)return;
  N[id]={t:String(v).trim().slice(0,1000),at:new Date().toISOString()};store.set('notes',N);msg.textContent=N[id].t?'Note saved.':'Note removed.';drawBuys();});
 drawBuys();}
/* compras (ações e ETF em bb.buys, Bitcoin em bb.lots) e vendas numa só lista; quantidades em ações de hoje */
function drawBuys(){if(!$('#tbl-buys'))return;const C=carteira(),resta={},NT=notasAtuais(),notaH=id=>NT[id]&&NT[id].t?`<span class="td-sub entry-note">${esc(NT[id].t)}</span>`:'',
  notaB=x=>`<button class="btn btn-ic" type="button" data-note="${esc(x.id)}" aria-label="Write a note for the ${x.venda?'sale':'purchase'} of ${esc(nomePF(x.a))} on ${fdt(x.t)}">${ic('book')}</button>`,
  FE=feesAtuais(),feeC=x=>`<td class="num">${FE[x.id]?eur(FE[x.id].v,2):'<span class="dim">—</span>'}<button class="btn btn-ic fee-b" type="button" data-fee="${esc(x.id)}" aria-label="Fee for the ${x.venda?'sale':'purchase'} of ${esc(nomePF(x.a))} on ${fdt(x.t)}">${ic('euro')}</button></td>`;Object.values(C.lotes).forEach(l=>l.forEach(x=>resta[x.id]=x));
 const compras=arr(store.get('buys',[])).map(efetiva).map(x=>Object.assign({st:'buys'},x)).concat(arr(store.get('lots',[])).map(x=>({st:'lots',id:x.id,a:'BTC',d:x.d,q:+x.q,p:x.q>0?(+x.c||0)/x.q:0,k:1,aviso:''})));
 const R=compras.map(x=>{const l=resta[x.id],q=l?l.q:x.q,px=lastEur(x.a),c=q*x.p,val=px!=null?q*px:null;return Object.assign({},x,{t:Date.parse(x.d+'T00:00:00Z'),qr:q,c,val,g:val!=null&&q>0?val-c:null,venda:false});})
  .concat(C.vendas.map(v=>Object.assign({},v,{st:'sales',t:Date.parse(v.d+'T00:00:00Z'),venda:true}))).sort((x,y)=>x.d<y.d?-1:x.d>y.d?1:0);
 const nota=x=>(x.k&&x.k!==1?`<span class="td-sub">adjusted for a ${nf(x.k,x.k%1?3:0)}:1 split</span>`:'')+(x.aviso?`<span class="td-sub old">${esc(x.aviso)}</span>`:'');
 $('#tbl-buys').innerHTML='<thead><tr><th scope="col">Asset</th><th scope="col">Date</th><th class="num" scope="col">Quantity</th><th class="num" scope="col">Price</th><th class="num" scope="col">Cost</th><th class="num" scope="col">Worth now / received</th><th class="num" scope="col">Gain</th><th class="num" scope="col">Fee</th><th scope="col"><span class="sr">Remove</span></th></tr></thead><tbody>'+
  (R.length?R.map(x=>x.venda?
   `<tr><td>${coTag(x.a)} <span class="tag t-sell">Sale</span>${notaH(x.id)}</td><td>${fdt(x.t)}</td><td class="num">−${btcF(x.q)}${nota(x)}${x.falta>0?`<span class="td-sub old">${btcF(x.falta)} more than you held</span>`:''}</td><td class="num">${eur(x.p,2)}</td><td class="num">${eur(x.custo,2)}<span class="td-sub">cost, oldest first</span></td><td class="num">${eur((x.q-x.falta)*x.p,2)}</td><td class="num ${cls(x.real)}">${eurS(x.real,2)}${x.real!=null&&x.custo>0?` <small>${pct(x.real/x.custo*100)}</small>`:''}<span class="td-sub">realised</span></td>${feeC(x)}<td>${notaB(x)}<button class="btn btn-ic" type="button" data-del="${esc(x.id)}" data-store="sales" aria-label="Remove the sale of ${esc(nomePF(x.a))} on ${fdt(x.t)}">${ic('trash')}</button></td></tr>`:
   `<tr class="${x.qr>0?'':'row-sold'}"><td>${coTag(x.a)}${notaH(x.id)}</td><td>${fdt(x.t)}</td><td class="num">${btcF(x.q)}${x.qr<x.q-1e-12?`<span class="td-sub">${x.qr>0?btcF(x.qr)+' left':'all sold'}</span>`:''}${nota(x)}</td><td class="num">${eur(x.p,2)}</td><td class="num">${eur(x.c,2)}</td><td class="num">${x.qr>0?eur(x.val,2):'—'}</td><td class="num ${cls(x.g)}">${eurS(x.g,2)}${x.g!=null&&x.c>0?` <small>${pct(x.g/x.c*100)}</small>`:''}</td>${feeC(x)}<td>${notaB(x)}<button class="btn btn-ic" type="button" data-del="${esc(x.id)}" data-store="${x.st}" aria-label="Remove the purchase of ${esc(nomePF(x.a))} on ${fdt(x.t)}">${ic('trash')}</button></td></tr>`).join(''):
   '<tr><td colspan="9" class="dim">No purchases yet. Choose the asset and the date above to add the first one.</td></tr>')+'</tbody>';
 drawTax();drawDiv8A();}

/* ---------- carteira: posições (só leitura, compras menos vendas) ---------- */
function buildPf(){const t=$('#tbl-pf');if(!t)return;
 t.innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">Quantity</th><th class="num" scope="col">Total invested</th><th class="num" scope="col">Price now</th><th class="num" scope="col">Value</th><th class="num" scope="col">Gain</th><th class="num" scope="col">Weight</th></tr></thead><tbody>'+
  PF.map(([id])=>`<tr><td>${coTag(id)}<div class="pf-src" id="pf-src-${id}"></div></td><td class="num" id="pf-q-${id}"></td><td class="num" id="pf-c-${id}"></td><td class="num" id="pf-p-${id}"></td><td class="num" id="pf-v-${id}"></td><td class="num" id="pf-g-${id}"></td><td class="num" id="pf-w-${id}"></td></tr>`).join('')+
  '</tbody><tfoot><tr><th scope="row">Total</th><td></td><td class="num" id="pf-tc"></td><td></td><td class="num" id="pf-tv"></td><td class="num" id="pf-tg"></td><td class="num" id="pf-tw"></td></tr></tfoot>';}
function pfDados(){const C=carteira();
 return PF.map(([id,n,u])=>{const L=arr(C.lotes[id]),V=C.vendas.filter(v=>v.a===id),q=L.reduce((s,x)=>s+x.q,0),c=L.reduce((s,x)=>s+x.q*x.cu,0),real=V.reduce((s,v)=>s+(v.real||0),0);
  const src=(L.length?`${L.length} purchase${L.length>1?'s':''}`:'No purchases yet')+(V.length?` · ${V.length} sale${V.length>1?'s':''}`:'');
  const p=lastEur(id),v=q>1e-12&&p!=null?q*p:null,g=v!=null&&c>0?v-c:null,a=ATIVOS.find(x=>x.id===id),s=a?stats(inCur(a,'EUR')):null,fr=a?frescura(id,a.pts):null;
  /* velho: o último preço é da execução anterior ou está atrasado em relação à bolsa (o seu "último dia" não é o de hoje) */
  return{id,n,u,q,c,p,v,g,real,vendas:V.length,src,d1:s?s.d1:null,velho:!!a&&(antigo(a)||!!(fr&&fr.velho)),semPreco:q>1e-12&&p==null};});}
function drawPf(){if(!$('#tbl-pf'))return;const R=pfDados(),tv=R.reduce((a,x)=>a+(x.v||0),0),G=R.filter(x=>x.g!=null),tg=G.reduce((a,x)=>a+x.g,0),tc=G.reduce((a,x)=>a+x.c,0),dia=R.reduce((a,x)=>a+(x.v&&x.d1!=null&&!x.velho?x.v*x.d1/(100+x.d1):0),0),diaFora=R.filter(x=>x.v&&x.velho),real=R.reduce((a,x)=>a+x.real,0),nv=R.reduce((a,x)=>a+x.vendas,0),falta=R.filter(x=>x.semPreco);
 drawDiv(R);drawRet(R);drawTgt(R);drawPol();drawStress(R);drawExpo(R);
 R.forEach(x=>{$('#pf-q-'+x.id).textContent=x.q>1e-12?btcF(x.q):'—';$('#pf-c-'+x.id).textContent=x.c>0?eur(x.c,2):'—';$('#pf-p-'+x.id).textContent=eur(x.p,2);$('#pf-v-'+x.id).textContent=x.v!=null?eur(x.v,2):(x.semPreco?'no price':'—');
  const gc=$('#pf-g-'+x.id);gc.className='num '+cls(x.g);gc.innerHTML=x.g!=null?`${eurS(x.g,2)} <small>${pct(x.g/x.c*100)}</small>`:'—';
  $('#pf-w-'+x.id).textContent=x.v&&tv?pct(x.v/tv*100).replace('+',''):'—';$('#pf-src-'+x.id).textContent=x.src;});
 $('#pf-tv').textContent=tv?eur(tv,2):'—';$('#pf-tc').textContent=tc?eur(tc,2):'—';const tgc=$('#pf-tg');tgc.className='num '+(G.length?cls(tg):'');tgc.innerHTML=G.length?`${eurS(tg,2)} <small>${tc?pct(tg/tc*100):''}</small>`:'—';$('#pf-tw').textContent=tv?'100%':'—';
 if(!tv){$('#pfKpis').innerHTML=`<p class="empty">${ic('wallet')}${falta.length?`No current price for ${falta.map(x=>esc(x.n)).join(', ')} in this run (see Sources &amp; method), so the portfolio cannot be valued.`:'Add your purchases below (and Bitcoin purchases in the Bitcoin section) to see your portfolio in euros.'}</p>`;empty($('#ch-pf-alloc'),'Nothing to show yet.');empty($('#ch-pf-exp'),'Nothing to show yet.');return;}
 $('#pfKpis').innerHTML=kpi('wallet','Portfolio value',eur(tv),falta.length?`without ${falta.map(x=>esc(x.n)).join(', ')} (no current price)`:'in euros, at the latest prices')+kpi('euro','Invested',eur(tc),nv?'cost of what you still hold':'total paid')+kpi('chart','Gain',eurS(tg),tc?pct(tg/tc*100):'',cls(tg))+
  (nv?kpi('check','Realised gain',eurS(real),`on ${nv} sale${nv>1?'s':''} (oldest purchases first)`,cls(real)):'')+kpi('activity','Latest daily move',R.some(x=>x.v&&!x.velho)?eurS(dia):'—','last session (Bitcoin: since 00:00 UTC)'+(diaFora.length?` · without ${diaFora.map(x=>esc(x.n)).join(', ')} (price not current)`:''),cls(dia));
 const fat=R.filter(x=>x.v>0).map(x=>[ehEtf(x.id)?'ETF '+x.id:x.n,x.v/tv*100,CO[x.id].c,x.v]);
 $('#ch-pf-alloc').innerHTML='<div class="dist" role="img" aria-label="'+fat.map(x=>`${x[0]} ${nf(x[1],1)}%`).join(', ')+'">'+fat.map(x=>`<div data-tip="${esc(x[0])}: ${nf(x[1],1)}% (${eur(x[3])})" style="width:${x[1]}%;background:${x[2]}"></div>`).join('')+'</div><ul class="etf-legend">'+fat.map(x=>`<li style="--c:${x[2]}"><span class="dot"></span><span class="nm">${esc(x[0])}</span><span class="bar"><i style="width:${x[1]}%"></i></span><b>${nf(x[1],1)}%</b></li>`).join('')+'</ul>';
 /* exposição real: o que tem diretamente + o que tem através de cada ETF (valor do ETF × peso da empresa nesse fundo).
    O resto de cada ETF (fora das três empresas) aparece uma vez, com o nome do índice: nada é contado duas vezes. */
 const vd=id=>(R.find(x=>x.id===id)||{}).v||0,EV=ETF_IDS.map(id=>{const w=etfInfo(id);return{id,w,ev:vd(id),s3:(+w.AAPL||0)+(+w.NVDA||0)+(+w.GOOGL||0)};}).filter(x=>x.ev>0);
 const via=k=>EV.reduce((s,x)=>s+x.ev*(+x.w[k]||0)/100,0),pesoTxt=x=>`${x.w.aoVivo?'the iShares weights':'the reference weights stored in the script (the iShares download failed)'}${x.w.data?' as of '+x.w.data:''}${etfIdade(x.id)>10?` (${etfIdade(x.id)} days old)`:''}`;
 const unico=EV.length>1?null:(EV[0]||{id:'SXR8',w:etfInfo('SXR8')});
 $('#pfExpSub').textContent=unico?`What you hold directly plus what you own through the ETF, using ${pesoTxt(unico)}.`:`What you hold directly plus what you own through the ETFs, using ${EV.map(x=>`for ${x.id} ${pesoTxt(x)}`).join('; ')}.`;
 const E=[['Apple',vd('AAPL')+via('AAPL'),'var(--aapl)'],['NVIDIA',vd('NVDA')+via('NVDA'),'var(--nvda)'],['Alphabet',vd('GOOGL')+via('GOOGL'),'var(--googl)']]
  .concat(EV.map(x=>[x.s3>0?`Rest of the ${etfIdx(x.id)}`:`${etfIdx(x.id)} (${x.id})`,x.ev*(1-x.s3/100),CO[x.id].c]),[['Bitcoin',vd('BTC'),'var(--btc)']]).filter(x=>x[1]>0);
 hbars($('#ch-pf-exp'),E.map(x=>({l:x[0],v:x[1],t:`${eur(x[1])} · ${nf(x[1]/tv*100,1)}%`,c:x[2]})),'Real exposure to each company');}

/* ---------- carteira: rentabilidade ponderada pelo dinheiro (TIR / XIRR) e benchmark pessoal em SXR8 ---------- */
/* Custo de cada movimento (comissões). Ainda não são registadas: quando forem, entram aqui sem mudar rentab() */
const custoFluxo=x=>feeDe(x&&x.id)||0;
/* Fluxos de caixa do registo: compras negativas (q × p de efetiva(), ou seja, já em ações de hoje; Bitcoin: o custo total c),
   vendas positivas (q × p da venda). Datas ISO, à meia-noite UTC (como o resto da página). */
function fluxosCarteira(){const F=[],tt=d=>Date.parse(d+'T00:00:00Z');
 arr(store.get('buys',[])).map(efetiva).forEach(x=>{if(x.q>0&&x.p>0&&isoOk(x.d))F.push({t:tt(x.d),v:-(x.q*x.p+custoFluxo(x)),a:x.a,tipo:'buy'});});
 arr(store.get('lots',[])).forEach(x=>{if(+x.q>0&&+x.c>0&&isoOk(x.d))F.push({t:tt(x.d),v:-(+x.c+custoFluxo(x)),a:'BTC',tipo:'buy'});});
 arr(store.get('sales',[])).map(efetiva).forEach(x=>{if(x.q>0&&x.p>0&&isoOk(x.d)&&COMPRAVEIS.includes(x.a))F.push({t:tt(x.d),v:x.q*x.p-custoFluxo(x),a:x.a,tipo:'sell'});});
 return F.sort((a,b)=>a.t-b.t||a.v-b.v);}
/* TIR de fluxos [[t, valor]] ordenados: a taxa r em que a soma de valor ÷ (1+r)^(dias ÷ base) é zero. base = 365 dá a taxa
   anual; base = dias do período dá a taxa do período. Newton a partir de 10 %; se não convergir, bisseção entre −99,9999 %
   e um limite que cresce até 10^12. Sem troca de sinal nos fluxos, ou sem solução, devolve null ("Unavailable"). */
function xirr(fl,base){if(!fl.length||!fl.some(x=>x[1]<0)||!fl.some(x=>x[1]>0)||!(base>0))return null;const t0=fl[0][0],e=fl.map(x=>(x[0]-t0)/(base*DAY)),escala=fl.reduce((s,x)=>s+Math.abs(x[1]),0);
 const f=r=>fl.reduce((s,x,i)=>s+x[1]*Math.pow(1+r,-e[i]),0),df=r=>fl.reduce((s,x,i)=>s-e[i]*x[1]*Math.pow(1+r,-e[i]-1),0);
 let r=0.1;for(let i=0;i<60;i++){const y=f(r),d=df(r);if(!isFinite(y)||!isFinite(d)||d===0)break;const n=r-y/d;if(!isFinite(n)||n<=-1)break;
  if(Math.abs(n-r)<=1e-12*Math.max(1,Math.abs(n))){if(Math.abs(f(n))<=1e-9*escala)return n;break;}r=n;}
 let lo=-0.999999,hi=1,flo=f(lo),fhi=f(hi);while(isFinite(fhi)&&flo*fhi>0&&hi<1e12){hi*=4;fhi=f(hi);}
 if(!isFinite(flo)||!isFinite(fhi)||flo*fhi>0)return null;
 for(let i=0;i<400&&hi-lo>1e-14*Math.max(1,Math.abs(lo));i++){const m=(lo+hi)/2,fm=f(m);if(!isFinite(fm))return null;if((fm>0)===(flo>0)){lo=m;flo=fm;}else hi=m;}
 return(lo+hi)/2;}
/* Rentabilidade da carteira e "o mesmo dinheiro, nas mesmas datas, em SXR8". Fecho = o valor de Your holdings (pfDados),
   datado no dia (Lisboa) em que a página foi gerada. Com menos de 365 dias desde o primeiro fluxo, a taxa é a do período
   (não anualizada). Benchmark: cada compra (de qualquer ativo, Bitcoin incluída) compra SXR8 pelos mesmos euros ao fecho
   desse dia ou ao último fecho anterior (HIST); cada venda retira os mesmos euros; sem unidades suficientes, vende-as
   todas e assinala-o. Uma data anterior ao histórico do SXR8 deixa o benchmark "Unavailable". */
function rentab(R){R=R||pfDados();const F=fluxosCarteira(),C=carteira(),pl=tzParts(GEN,LIS),tFim=Date.UTC(+pl.year,+pl.month-1,+pl.day);
 const B={estado:'na',r:null,valor:null,unidades:null,curto:false,t:null,inicio:null,tPreco:null};
 const X={estado:'ok',r:null,anual:true,dias:null,t0:null,tFim,valor:null,fl:[],falta:[],motivo:'',nC:F.filter(x=>x.tipo==='buy').length,nV:F.filter(x=>x.tipo==='sell').length,bench:B,dif:null};
 const semP=arr(store.get('buys',[])).concat(arr(store.get('sales',[]))).filter(x=>x&&+x.q>0&&!(+x.p>0)).length+arr(store.get('lots',[])).filter(x=>x&&+x.q>0&&!(+x.c>0)).length;
 if(!F.length&&!semP){X.estado='empty';return X;}
 X.t0=F.length?F[0].t:null;const tEnd=Math.max(tFim,F.length?F[F.length-1].t:tFim);X.tFim=tEnd;X.dias=X.t0!=null?(tEnd-X.t0)/DAY:null;X.anual=X.dias!=null&&X.dias>=365;
 const base=X.anual?365:Math.max(X.dias||0,1);
 if(semP){X.estado='incomplete';X.motivo=`${semP} entr${semP>1?'ies have':'y has'} no price in the register`;}
 else if(C.vendas.some(v=>v.falta>0)){X.estado='incomplete';X.motivo='a sale has no matching purchase in the register';}
 else{X.falta=R.filter(x=>x.semPreco).map(x=>x.n);
  if(X.falta.length)X.estado='noprice';
  else{X.valor=R.reduce((s,x)=>s+(x.q>1e-12&&x.v!=null?x.v:0),0);X.fl=F.map(x=>[x.t,x.v]).concat([[tEnd,X.valor]]);X.r=xirr(X.fl,base);if(X.r==null)X.estado='nosol';}}
 if(X.estado==='ok'||X.estado==='noprice'||X.estado==='nosol'){const H=HIST.SXR8||[],pS=lastEur('SXR8'),aS=ATIVOS.find(x=>x.id==='SXR8');B.inicio=H.length?H[0][0]:null;B.tPreco=aS&&aS.pts.length?aS.pts[aS.pts.length-1][0]:null;B.parcial=!!(aS&&aS.parcial);
  if(!H.length)B.estado='nodata';else if(pS==null)B.estado='noprice';
  else{let u=0;const bf=[];B.estado='ok';
   for(const x of F){const pt=precoEm(H,isoU(x.t));if(!pt){B.estado='early';B.t=x.t;break;}
    if(x.v<0){u+=-x.v/pt[1];bf.push([x.t,x.v]);}else{const q=x.v/pt[1];if(q>u+1e-12){B.curto=true;bf.push([x.t,u*pt[1]]);u=0;}else{u-=q;bf.push([x.t,x.v]);}}}
   if(B.estado==='ok'){B.unidades=u;B.valor=u*pS;bf.push([tEnd,B.valor]);B.r=xirr(bf,base);if(B.r==null)B.estado='nosol';}}}
 if(X.estado==='ok'&&B.estado==='ok')X.dif=X.r-B.r;
 return X;}
function drawRet(R){const box=$('#retKpis'),nt=$('#retNote');if(!box)return;const X=rentab(R),B=X.bench;
 if(X.estado==='empty'){box.innerHTML=`<p class="empty">${ic('scale')}Add your purchases and sales in the register to see the return on your money and compare it with the same money invested in SXR8.</p>`;nt.innerHTML='';return;}
 const taxa=v=>X.anual?`${pct(v*100)} a year`:pct(v*100),per=X.anual?'money-weighted (XIRR), in euros':`for the ${Math.round(X.dias)} days since your first purchase, not annualised`;
 const motivoX={noprice:`no current price for ${X.falta.map(esc).join(', ')}`,incomplete:esc(X.motivo),nosol:'the calculation found no solution for these cash flows'}[X.estado];
 const motivoB={nodata:'no SXR8 price history in this run',noprice:'no current SXR8 price in this run',early:`no SXR8 close on or before ${fdt(B.t)}`,nosol:'the calculation found no solution',na:'it needs the portfolio cash flows'}[B.estado];
 box.innerHTML=kpi('scale','Your return',X.estado==='ok'?taxa(X.r):'Unavailable',X.estado==='ok'?per:motivoX,X.estado==='ok'?cls(X.r):'')+
  kpi('pie','Same money in SXR8',B.estado==='ok'?eur(B.valor):'Unavailable',B.estado==='ok'?`${taxa(B.r)} · ${nf(B.unidades,4)} units`:motivoB)+
  kpi('swap','Difference',X.dif!=null?`${X.dif>0?'+':''}${nf(X.dif*100,1)} pp${X.anual?' a year':''}`:'Unavailable',X.dif!=null?`${eurS(X.valor-B.valor)} in value today`:'needs both returns',X.dif!=null?cls(X.dif):'');
 const held=R.filter(x=>x.q>1e-12&&x.p!=null).map(x=>{const a=ATIVOS.find(y=>y.id===x.id);return a&&a.pts.length?`${esc(x.id)} ${fdt(a.pts[a.pts.length-1][0])}`:esc(x.id);});
 nt.innerHTML=`${ic('info')}<span>Cash flows from your register: ${X.nC} purchase${X.nC===1?'':'s'} and ${X.nV} sale${X.nV===1?'':'s'}${X.t0!=null?`, the first on ${fdt(X.t0)}`:''}. `+
  `Closing value: what you hold now at the latest prices in this run${held.length?` (${held.join(', ')})`:''}, dated ${fdt(X.tFim)}. `+
  `Benchmark: each purchase (any asset, Bitcoin included) buys SXR8 for the same euros at that day's Xetra close, or the last close before it, and each sale takes the same euros out of it; SXR8 closes from the long-term history (Yahoo Finance)${B.tPreco!=null?`, valued at its latest price (${B.parcial?'intraday, ':''}${fdt(B.tPreco)})`:''}.`+
  (X.anual?'':' Not annualised: your first purchase is less than a year old, and a yearly rate over a short period would exaggerate.')+
  (B.curto?' On at least one sale there were not enough SXR8 units in the benchmark: all of them were used, so the benchmark received less than your sale.':'')+
  (B.estado==='early'?` Benchmark Unavailable: SXR8 has no close on or before ${fdt(B.t)}${B.inicio!=null?` (its history starts on ${fdt(B.inicio)})`:''}.`:'')+
  ' Fees are not included. Past returns say nothing about future returns.</span>';}

/* ---------- carteira: alocação-alvo e repartição da próxima contribuição ---------- */
/* Repartição de M: alvos para (total + M); défice de cada ativo = max(0, alvo − atual). Se a soma dos défices for ≥ M,
   M é repartido na proporção dos défices; senão, cobre os défices e reparte o resto pelos pesos-alvo. (Com pesos que
   somam 1, a soma dos défices nunca é menor do que M: no limite é igual, e então cada ativo recebe o seu défice.) */
function reparte(cur,w,M){const ids=Object.keys(w),T=Object.keys(cur).reduce((s,k)=>s+(+cur[k]||0),0),def={},o={};ids.forEach(k=>def[k]=Math.max(0,w[k]*(T+M)-(+cur[k]||0)));
 const S=ids.reduce((s,k)=>s+def[k],0);ids.forEach(k=>o[k]=S>=M?(S>0?M*def[k]/S:0):def[k]+(M-S)*w[k]);return o;}
/* meses de contribuições de M (repartidas como acima, com os preços de hoje) até todos os ativos ficarem dentro da banda,
   sem vender: 0 = já estão; Infinity = mais de 600 meses; null = fora da banda e sem valor mensal */
function mesesBanda(cur,w,M,banda){const c=Object.assign({},cur),dentro=()=>{const T=Object.keys(c).reduce((s,k)=>s+(+c[k]||0),0);return T>0&&Object.keys(w).every(k=>Math.abs((+c[k]||0)/T*100-w[k]*100)<=banda+1e-9);};
 if(dentro())return 0;if(!(M>0))return null;for(let n=1;n<=600;n++){const a=reparte(c,w,M);Object.keys(a).forEach(k=>c[k]=(+c[k]||0)+a[k]);if(dentro())return n;}return Infinity;}
/* valores atuais de Your holdings (pfDados) e alvos gravados. Os ativos sem preço atual ficam de fora (e são nomeados) e os
   alvos dos restantes são reescalados para somar 100 % */
function alocacao(R){R=R||pfDados();const t=limpaAlvos(store.get('targets',null)),inc=PF.map(p=>p[0]).filter(id=>lastEur(id)!=null);
 const excl=R.filter(x=>!inc.includes(x.id)&&(x.q>1e-12||(t&&+t.weights[x.id]>0))).map(x=>x.n),cur={};inc.forEach(id=>{const x=R.find(y=>y.id===id);cur[id]=x&&x.v!=null?x.v:0;});
 const T=inc.reduce((s,id)=>s+cur[id],0),X={estado:t?'ok':'none',alvos:t,T,excl,M:t?t.monthly:0,band:t?t.band:5,meses:null,linhas:[]};let w=null;
 if(t){const s=inc.reduce((a,id)=>a+(+t.weights[id]||0),0);if(s>0){w={};inc.forEach(id=>w[id]=(+t.weights[id]||0)/s);}else X.estado='zero';}
 const al=w&&X.M>0?reparte(cur,w,X.M):null;
 X.linhas=R.map(x=>{const ok=inc.includes(x.id),peso=ok&&T>0?cur[x.id]/T*100:null,alvo=w&&ok?w[x.id]*100:null,desvio=peso!=null&&alvo!=null?peso-alvo:null;
  return{id:x.id,n:x.n,v:ok?cur[x.id]:null,peso,alvo,alvo0:t?(+t.weights[x.id]||0):null,desvio,fora:desvio!=null&&Math.abs(desvio)>X.band+1e-9,aloc:al&&ok?al[x.id]:null};});
 if(w)X.meses=mesesBanda(cur,w,X.M,X.band);
 return X;}
let tgtSujo=false;  /* há alvos escritos na tabela que ainda não foram gravados: não os substituir ao redesenhar */
function somaTgt(){const c=$('#tg-tw');if(!c)return;let s=0,algum=false;PF.forEach(([id])=>{const i=$('#tgt-w-'+id),v=i?i.value.trim():'';if(v!==''){algum=true;s+=+v||0;}});
 c.textContent=algum?String(+s.toFixed(2))+'%':'—';c.className='num tgt-sum'+(algum&&Math.abs(s-100)>0.01?' bad':'');}
function buildTgt(){const t=$('#tbl-tgt'),f=$('#tgtForm');if(!t||!f)return;
 t.innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">Value</th><th class="num" scope="col">Current weight</th><th class="num" scope="col">Target (%)</th><th class="num" scope="col">Deviation</th><th class="num" scope="col">Next contribution</th></tr></thead><tbody>'+
  PF.map(([id,n])=>`<tr><td>${coTag(id)}</td><td class="num" id="tg-v-${id}"></td><td class="num" id="tg-p-${id}"></td><td class="num"><input class="in tgt-in" type="number" min="0" max="100" step="any" id="tgt-w-${id}" form="tgtForm" aria-label="Target weight for ${esc(n)}, in percent"></td><td class="num" id="tg-d-${id}"></td><td class="num" id="tg-a-${id}"></td></tr>`).join('')+
  '</tbody><tfoot><tr><th scope="row">Total</th><td class="num" id="tg-tv"></td><td class="num" id="tg-tp"></td><td class="num tgt-sum" id="tg-tw"></td><td></td><td class="num" id="tg-ta"></td></tr></tfoot>';
 $$('.tgt-in').forEach(i=>i.addEventListener('input',()=>{tgtSujo=true;somaTgt();}));
 f.addEventListener('submit',e=>{e.preventDefault();const msg=$('#tgtMsg'),w={};let s=0,mau=false;
  PF.forEach(([id])=>{const r=$('#tgt-w-'+id).value.trim();if(r==='')return;const v=+r;if(!isFinite(v)||v<0){mau=true;return;}if(v>0)w[id]=v;s+=v;});
  const bR=$('#tgtBand').value.trim(),b=bR===''?5:+bR,mR=$('#tgtMonthly').value.trim(),m=mR===''?0:+mR;
  if(mau){msg.textContent='Each target must be a number, 0 or more. Nothing was saved.';return;}
  if(Math.abs(s-100)>0.01){msg.textContent=`Your targets add up to ${String(+s.toFixed(2))}%: they must add up to 100%. Nothing was saved.`;return;}
  if(!(isFinite(b)&&b>0&&b<=50)){msg.textContent='The tolerance band must be more than 0 and at most 50 percentage points. Nothing was saved.';return;}
  if(!(isFinite(m)&&m>=0)){msg.textContent='The monthly amount must be 0 or more. Nothing was saved.';return;}
  store.set('targets',{weights:w,band:b,monthly:m,at:new Date().toISOString()});tgtSujo=false;msg.textContent='Targets saved.';drawPf();});}
function drawTgt(R){if(!$('#tbl-tgt'))return;const A=alocacao(R),t=A.alvos,o=$('#tgtOut');
 if(!tgtSujo){PF.forEach(([id])=>{const i=$('#tgt-w-'+id);if(i)i.value=t&&t.weights[id]!=null?String(t.weights[id]):'';});$('#tgtBand').value=t?String(t.band):'';$('#tgtMonthly').value=t?String(t.monthly):'';}
 A.linhas.forEach(x=>{$('#tg-v-'+x.id).textContent=x.v!=null?eur(x.v,2):'no price';$('#tg-p-'+x.id).textContent=x.peso!=null?nf(x.peso,1)+'%':'—';
  $('#tg-d-'+x.id).innerHTML=x.desvio!=null?`${x.desvio>0?'+':''}${nf(x.desvio,1)} pp${x.fora?' <span class="tag t-warn">outside the band</span>':''}`:'—';
  $('#tg-a-'+x.id).textContent=x.aloc!=null?eur(x.aloc,2):'—';});
 $('#tg-tv').textContent=A.T?eur(A.T,2):'—';$('#tg-tp').textContent=A.T?'100%':'—';$('#tg-ta').textContent=A.estado==='ok'&&A.M>0?eur(A.M,2):'—';somaTgt();
 const ex=A.excl.length?` ${A.excl.map(esc).join(', ')} ${A.excl.length>1?'have':'has'} no current price in this run: left out, and the other targets are scaled to 100%.`:'';
 if(A.estado==='none'){o.innerHTML=`<p class="empty">${ic('pie')}Set a target for each asset, the band and your monthly amount, then save, to see how far you are from your targets and how to split your next contribution.${ex}</p>`;return;}
 if(A.estado==='zero'){o.innerHTML=`<p class="note">${ic('info')}<span>None of the assets with a target has a current price in this run, so nothing can be split.${ex}</span></p>`;return;}
 const n=A.meses,bd=nf(A.band,1),mt=n===0?'Every asset is already inside the band.':!(A.M>0)?'Set a monthly amount to see the split of your next contribution and how many months it would take to get back inside the band.':
  n===Infinity?`Even after 50 years of contributions of ${eur(A.M,2)}, at least one asset would still be outside the ±${bd} pp band without selling.`:
  `About ${n} month${n===1?'':'s'} of contributions of ${eur(A.M,2)} would bring every asset back inside the ±${bd} pp band without selling, if prices stayed as they are today.`;
 o.innerHTML=`<p class="note">${ic('calendar')}<span>${mt}${ex}</span></p>`;}

/* ---------- carteira: teste de stress com três quedas do S&P 500 (datas de pico e de vale dos fechos de ^GSPC) ---------- */
const EPISODIOS=[{id:'2018Q4',n:'2018 Q4',pk:'2018-09-20',tr:'2018-12-24'},{id:'COVID',n:'COVID',pk:'2020-02-19',tr:'2020-03-23'},{id:'2022',n:'2022',pk:'2022-01-03',tr:'2022-10-12'}];
/* Para cada ativo detido com preço atual: retorno em euros entre o fecho do pico e o do vale (ou o último fecho anterior,
   até 10 dias antes; as ações americanas ao EUR/USD de cada dia), aplicado ao valor atual; e o tempo, depois do vale, até
   o preço em euros voltar ao do pico. Sem histórico nesse período, o ativo fica de fora desse episódio e é nomeado. */
function stress(R){R=R||pfDados();const held=R.filter(x=>x.q>1e-12),com=held.filter(x=>x.v!=null),sem=held.filter(x=>x.v==null).map(x=>x.n),perto=(p,iso)=>p&&Date.parse(iso+'T00:00:00Z')-p.t<=10*DAY;
 const eps=EPISODIOS.map(e=>{const L=com.map(x=>{const a=precoEurEm(x.id,e.pk),b=precoEurEm(x.id,e.tr);if(!perto(a,e.pk)||!perto(b,e.tr)||!(a.eur>0))return{id:x.id,n:x.n,v:x.v,ok:false};
   let rec=null;for(const p of(HIST[x.id]||[])){if(p[0]<=b.t)continue;const c=precoEurEm(x.id,isoU(p[0]));if(c&&c.t===p[0]&&c.eur>=a.eur){rec=p[0];break;}}
   const r=b.eur/a.eur-1;return{id:x.id,n:x.n,v:x.v,ok:true,r,perda:x.v*r,tPk:a.t,tTr:b.t,rec};});
  const ok=L.filter(l=>l.ok),base=ok.reduce((s,l)=>s+l.v,0),total=ok.reduce((s,l)=>s+l.perda,0);
  return Object.assign({},e,{linhas:L,total,base,pct:base?total/base*100:null,excl:L.filter(l=>!l.ok).map(l=>l.n)});});
 return{eps,sem,held:held.length};}
function drawStress(R){const t=$('#tbl-stress'),o=$('#stressOut');if(!t)return;const S=stress(R);
 if(!S.held){o.innerHTML=`<p class="empty">${ic('zap')}Add your purchases to see what past market falls would mean for your holdings today.</p>`;t.innerHTML='';return;}
 const per=e=>`${fdt(Date.parse(e.pk+'T00:00:00Z'))} → ${fdt(Date.parse(e.tr+'T00:00:00Z'))}`,cel=l=>!l.ok?'<span class="dim">no data for this period</span>':
  `<span class="${cls(l.perda)}">${eur(l.perda,0)}</span> <small>${pct(l.r*100)}</small><span class="td-sub">${l.rec!=null?`back in ${durQ(l.rec-l.tTr)}`:'not back yet'}</span>`;
 const linhas=S.eps[0].linhas;
 t.innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">Value now</th>'+S.eps.map(e=>`<th class="num" scope="col">${esc(e.n)}<span class="td-sub">${per(e)}</span></th>`).join('')+'</tr></thead><tbody>'+
  linhas.map((l,i)=>`<tr><td>${coTag(l.id)}</td><td class="num">${eur(l.v,0)}</td>${S.eps.map(e=>`<td class="num">${cel(e.linhas[i])}</td>`).join('')}</tr>`).join('')+
  '</tbody><tfoot><tr><th scope="row">Total</th><td class="num">'+eur(linhas.reduce((s,l)=>s+l.v,0),0)+'</td>'+S.eps.map(e=>`<td class="num"><span class="${cls(e.total)}">${e.base?eur(e.total,0):'—'}</span> <small>${e.pct!=null?pct(e.pct):''}</small>${e.excl.length?`<span class="td-sub">without ${e.excl.map(esc).join(', ')}</span>`:''}</td>`).join('')+'</tr></tfoot>';
 o.innerHTML=S.sem.length?`<p class="note">${ic('info')}<span>${S.sem.map(esc).join(', ')} ${S.sem.length>1?'have':'has'} no current price in this run and ${S.sem.length>1?'are':'is'} left out.</span></p>`:'';}

/* ---------- carteira: política de investimento (textos do utilizador; as regras de queda aparecem nos alertas) ---------- */
let polSujo=false;  /* há texto escrito que ainda não foi gravado: não o substituir ao redesenhar */
function buildPol(){const f=$('#polForm');if(!f)return;POL_K.forEach(k=>{const i=$('#pol-'+k);if(i)i.addEventListener('input',()=>{polSujo=true;});});
 f.addEventListener('submit',e=>{e.preventDefault();const o={};POL_K.forEach(k=>{o[k]=($('#pol-'+k).value||'').trim().slice(0,2000);});const now=Date.now();o.at=new Date(now).toISOString();
  store.set('policy',o);polSujo=false;$('#polMsg').textContent=`Saved on ${fdt(now)} at ${hmL(now)}. Your drop rules now appear next to the drop alerts.`;renderAlerts();});
 drawPol();}
function drawPol(){if(!$('#polForm')||polSujo)return;const P=limpaPolitica(store.get('policy',null));POL_K.forEach(k=>{const i=$('#pol-'+k);if(i)i.value=P?P[k]:'';});}

/* ---------- carteira: "Before you sell" (simulação; nada é gravado) ---------- */
/* Corre carteira() com uma venda hipotética em memória (hoje, ao preço mais recente): os lotes que gastaria (FIFO), a
   mais-valia e uma estimativa de imposto a 28 % sobre o ganho tributável positivo. Na Bitcoin, os lotes detidos 365 dias
   ou mais estão isentos (a regra do contador) e os que faltam 30 dias ou menos são assinalados. */
function simulaVenda(a,q){if(!COMPRAVEIS.includes(a)||!(q>0))return{erro:'Choose an asset and a quantity.'};const p=lastEur(a);if(p==null)return{erro:`No current price for ${nomePF(a)} in this run, so the sale cannot be simulated.`};
 const hj=hojeL(),d=isoU(hj),C=carteira({id:'__sim',a,d,q,p}),v=C.vendas.find(x=>x.id==='__sim');
 if(!v||v.falta>0){const tem=arr(C.lotes[a]).reduce((s,x)=>s+x.q0,0);return{erro:`You do not hold enough ${nomePF(a)} for this sale (you hold ${btcF(Math.max(0,q-(v?v.falta:q)))} after your earlier sales).`};}
 const btc=a==='BTC',L=arr(v.usados).map(m=>{const tA=Date.parse(m.d+'T00:00:00Z'),dias=Math.round((hj-tA)/DAY),custo=m.q*m.cu,valor=m.q*p,isento=btc&&!(hj<tA+ANEXO_J.q94.dias*DAY),faltam=btc&&!isento?Math.round((tA+ANEXO_J.q94.dias*DAY-hj)/DAY):null;
  return{id:String(m.id),d:m.d,q:m.q,cu:m.cu,dias,custo,valor,ganho:valor-custo,isento,faltam,perto:faltam!=null&&faltam<=30};});
 const sum=(l,f)=>l.reduce((s,x)=>s+x[f],0),trib=L.filter(x=>!x.isento),gT=sum(trib,'ganho');
 return{a,q,p,d,btc,lotes:L,valor:sum(L,'valor'),custo:sum(L,'custo'),ganho:sum(L,'ganho'),ganhoTrib:gT,imposto:Math.max(0,gT)*0.28};}
function buildSell(){const f=$('#sellForm');if(!f)return;$('#sellAsset').innerHTML=COMPRAVEIS.map(id=>`<option value="${id}">${esc(nomePF(id))}</option>`).join('');
 f.addEventListener('submit',e=>{e.preventDefault();const S=simulaVenda($('#sellAsset').value,+$('#sellQty').value),o=$('#sellOut'),t=$('#tbl-sell');
  if(S.erro){o.innerHTML=`<p class="note">${ic('triangle')}<span>${esc(S.erro)}</span></p>`;t.innerHTML='';return;}
  const perto=S.lotes.filter(x=>x.perto);
  o.innerHTML=`<div class="kpis">${kpi('wallet','Proceeds',eur(S.valor,2),`${btcF(S.q)} × ${eur(S.p,2)} (latest price)`)+kpi('euro','Cost (oldest first)',eur(S.custo,2),`${S.lotes.length} purchase${S.lotes.length===1?'':'s'} used`)+
   kpi('chart','Gain',eurS(S.ganho,2),S.btc?`taxable part ${eurS(S.ganhoTrib,2)}`:'before fees',cls(S.ganho))+kpi('calculator','Tax estimate (28%)',eur(S.imposto,2),S.btc?'on the taxable part only':'on the positive gain')}</div>`+
   (perto.length?`<div class="callout warn">${ic('clock')}<div>${perto.map(x=>`The purchase of ${fdt(Date.parse(x.d+'T00:00:00Z'))} becomes exempt in ${x.faltam} day${x.faltam===1?'':'s'}`).join('; ')}: selling it before then makes its gain taxable.</div></div>`:'')+
   `<p class="note">${ic('info')}<span>A simplification: 28% is the autonomous rate on the positive gain of this sale alone. It does not offset losses from other sales, does not consider the option to aggregate income (englobamento), and leaves out fees.${S.btc?' Bitcoin held 365 days or more is exempt (the same rule as the 365-day counter).':''} Confirm with the Portal das Finanças or an accountant. Nothing was saved.</span></p>`;
  t.innerHTML='<thead><tr><th scope="col">Purchase</th><th class="num" scope="col">Quantity used</th><th class="num" scope="col">Cost</th><th class="num" scope="col">Value now</th><th class="num" scope="col">Gain</th><th scope="col">Held</th></tr></thead><tbody>'+
   S.lotes.map(x=>`<tr><td>${fdt(Date.parse(x.d+'T00:00:00Z'))}</td><td class="num">${btcF(x.q)}</td><td class="num">${eur(x.custo,2)}</td><td class="num">${eur(x.valor,2)}</td><td class="num ${cls(x.ganho)}">${eurS(x.ganho,2)}</td><td>${x.dias} days${S.btc?(x.isento?' <span class="tag t-pos">exempt</span>':` <span class="tag${x.perto?' t-warn':''}">taxable · exempt in ${x.faltam} days</span>`):''}</td></tr>`).join('')+'</tbody>';});}

/* ---------- carteira: exposição agregada (país, setor, moeda) e concentração por empresa ---------- */
/* setor das ações diretas quando o ficheiro da iShares não as tem no top 10 (setores GICS, com os nomes da iShares) */
const SETOR_ACOES={AAPL:'Information Technology',NVDA:'Information Technology',GOOGL:'Communication'};
const CONC_LIMIAR=10;   /* % do total da carteira acima do qual uma empresa tem um aviso (só informativo) */
const DIM_EXPO=['paises','setores','moedas'];
const agregadosDe=id=>{const a=(etfInfo(id)||{}).agregados;if(!a||typeof a!=='object')return null;const ok=l=>Array.isArray(l)&&l.length&&l.every(x=>x&&typeof x.n==='string'&&isFinite(+x.w))?l:null;
 return{paises:ok(a.paises),setores:ok(a.setores),moedas:ok(a.moedas),fontePaises:a.fontePaises||null};};
function setorAcao(id){for(const e of ETF_IDS){const t=arr((etfInfo(e)||{}).top10).find(x=>x&&x.t===id&&x.s);if(t)return t.s;}return SETOR_ACOES[id]||'Other';}
/* Pondera pelos valores de hoje (pfDados): cada ETF pelos seus agregados (escalados ao valor do fundo), as ações diretas como
   Estados Unidos / USD / o seu setor, a Bitcoin como Crypto. Um fundo sem agregados fica de fora dessa dimensão (sem cobertura). */
function exposicao(R){R=R||pfDados();const H=R.filter(x=>x.q>1e-12&&x.v!=null&&x.v>0),total=H.reduce((s,x)=>s+x.v,0),dims={};
 DIM_EXPO.forEach(dim=>{const m={},soma=(n,v)=>{m[n]=(m[n]||0)+v;};let sem=0;const semN=[];
  H.forEach(x=>{if(x.id==='BTC')soma('Crypto',x.v);else if(!ehEtf(x.id))soma(dim==='paises'?'United States':dim==='moedas'?'USD':setorAcao(x.id),x.v);
   else{const a=agregadosDe(x.id),l=a&&a[dim];if(!l){sem+=x.v;semN.push(x.n);return;}const t=l.reduce((s,y)=>s+(+y.w),0);if(!(t>0)){sem+=x.v;semN.push(x.n);return;}l.forEach(y=>soma(y.n,x.v*(+y.w)/t));}});
  const cob=total-sem,L=Object.keys(m).map(n=>({n,v:m[n],p:cob>0?m[n]/cob*100:null})).sort((a,b)=>b.v-a.v);dims[dim]={L,sem,semN,cob};});
 /* concentração: as três empresas do painel (direto + look-through) e o top 10 de cada fundo, fundo a fundo */
 const vd=id=>(H.find(x=>x.id===id)||{}).v||0,conc=[];
 [['AAPL','Apple'],['NVDA','NVIDIA'],['GOOGL','Alphabet']].forEach(([id,n])=>{let v=vd(id);const via=[];ETF_IDS.forEach(e=>{const w=+(etfInfo(e)||{})[id]||0,ev=vd(e);if(w>0&&ev>0){v+=ev*w/100;via.push(e);}});
  if(v>0)conc.push({n,via:(vd(id)>0?['direct']:[]).concat(via).join(' + '),v,p:v/total*100});});
 ETF_IDS.forEach(e=>{const ev=vd(e),info=etfInfo(e)||{};if(!(ev>0)||!info.aoVivo)return;arr(info.top10).forEach(t=>{if(!t||['AAPL','NVDA','GOOGL'].includes(t.t)||!(+t.w>0))return;const v=ev*(+t.w)/100;conc.push({n:t.n,via:e,v,p:v/total*100});});});
 conc.forEach(c=>c.acima=c.p>CONC_LIMIAR);conc.sort((a,b)=>b.v-a.v);
 return{total,dims,conc,limiar:CONC_LIMIAR,held:H};}
function drawExpo(R){if(!$('#expoOut'))return;const X=exposicao(R);
 if(!X.total){$('#expoOut').innerHTML=`<p class="empty">${ic('globe')}Add your purchases to see where your money really is.</p>`;DIM_EXPO.forEach(d=>empty($('#ch-expo-'+d),'Nothing to show yet.'));$('#expoNote').innerHTML='';$('#concOut').innerHTML='';$('#tbl-conc').innerHTML='';return;}
 const semTxt=DIM_EXPO.map(d=>X.dims[d]).filter(x=>x.sem>0);
 $('#expoOut').innerHTML=semTxt.length?`<p class="note">${ic('triangle')}<span>Not covered: ${nf(semTxt[0].sem/X.total*100,1)}% of the portfolio (${eur(semTxt[0].sem)}), because the breakdown of ${[...new Set(semTxt.flatMap(x=>x.semN))].map(esc).join(', ')} is Unavailable in this run (the iShares download failed or this data file has none). The percentages below are of the rest.</span></p>`:'';
 const cores={'Cash/Other':'var(--muted)',Crypto:'var(--btc)'};
 DIM_EXPO.forEach(d=>{const L=X.dims[d].L,top=L.slice(0,11),resto=L.slice(11),rows=top.concat(resto.length?[{n:`${resto.length} others`,v:resto.reduce((s,x)=>s+x.v,0),p:resto.reduce((s,x)=>s+(x.p||0),0)}]:[]);
  if(!rows.length){empty($('#ch-expo-'+d),'Unavailable in this run.');return;}
  hbars($('#ch-expo-'+d),rows.map(x=>({l:x.n,v:x.p,t:`${nf(x.p,1)}% · ${eur(x.v)}`,c:cores[x.n]||'var(--bar)'})),{paises:'Exposure by country',setores:'Exposure by sector',moedas:'Exposure by underlying currency'}[d]);});
 const idx=X.held.filter(x=>ehEtf(x.id)&&(agregadosDe(x.id)||{}).fontePaises==='index').map(x=>`${esc(x.id)} (${esc((etfInfo(x.id)||{}).paisIndice||'')})`);
 $('#expoNote').innerHTML=`${ic('info')}<span>${X.held.some(x=>x.id==='SXR8')?'SXR8 is quoted in EUR, but its underlying currency exposure is USD. ':''}${idx.length?`The iShares file has no country breakdown for ${idx.join(', ')}: the country of its index is used, which is approximate (the S&amp;P 500, for example, includes a few companies based outside the US). `:''}Cash and derivatives inside the funds count as Cash/Other.</span>`;
 const C=X.conc,ac=C.filter(c=>c.acima);
 $('#concOut').innerHTML=ac.length?`<div class="callout">${ic('info')}<div>${ac.map(c=>`<b>${esc(c.n)}</b> is ${nf(c.p,1)}% of your portfolio (${esc(c.via)})`).join('; ')}: above the ${nf(X.limiar,0)}% threshold. This is information, not an alert.</div></div>`:`<p class="note">${ic('check')}<span>No single company is above ${nf(X.limiar,0)}% of your portfolio.</span></p>`;
 $('#tbl-conc').innerHTML='<thead><tr><th scope="col">Company</th><th scope="col">Through</th><th class="num" scope="col">Value</th><th class="num" scope="col">Share of the portfolio</th></tr></thead><tbody>'+
  (C.length?C.slice(0,15).map(c=>`<tr><td>${esc(c.n)}${c.acima?' <span class="tag t-warn">above '+nf(X.limiar,0)+'%</span>':''}</td><td>${esc(c.via)}</td><td class="num">${eur(c.v)}</td><td class="num">${nf(c.p,1)}%</td></tr>`).join(''):'<tr><td colspan="4" class="dim">No company data for your holdings in this run.</td></tr>')+'</tbody>';}

/* ---------- carteira: dividendos projetados (AAPL, NVDA, GOOGL) ---------- */
/* Ações detidas agora (as de pfDados, ou seja, do mesmo cálculo FIFO da carteira, ajustado aos desdobramentos) × dividendo
   anual por ação em USD (último pagamento × pagamentos por ano, recolhido pelo script). Em euros, à taxa EUR/USD mais
   recente: são pagamentos futuros, por isso não se usa a taxa da data de compra. SXR8 (acumulação) e Bitcoin ficam de fora.
   Sem dados (fonte em falha, dados com mais de 180 dias, resposta inválida) mostra "Unavailable", nunca €0. */
/* as empresas com dividendos recolhidos pelo script; um ETF de acumulação nunca entra (os dividendos ficam no preço) */
const DIV_IDS=[...new Set(['AAPL','NVDA','GOOGL'].concat(D.dividendos&&typeof D.dividendos==='object'&&!Array.isArray(D.dividendos)?Object.keys(D.dividendos):[]))].filter(id=>!(ehEtf(id)&&etfAcum(id))&&id!=='BTC'&&PF.some(p=>p[0]===id)),DIV_MAX_DIAS=180,
 /* arredonda aos cêntimos, metade para longe do zero, sobre o valor decimal (15 algarismos): 1.005 × 100 dá
    100.49999999999999 em vírgula flutuante, e Math.round sozinho daria €1.00 em vez de €1.01 */
 c2=v=>{if(v==null||!isFinite(v))return v;const x=Math.round(+(Math.abs(v)*100).toPrecision(15))/100;return v<0?-x:x;};
function divDados(id){const v=D.dividendos&&D.dividendos[id];if(!v||typeof v!=='object')return{ok:false,motivo:'No dividend data in this run'};
 if(v.estado!=='ok'&&v.estado!=='previous run')return{ok:false,motivo:'Source failed'+(v.erro?': '+v.erro:'')};
 const t=Date.parse(v.obtidoEm||''),a=v.anualPorAcao==null||v.anualPorAcao===''?NaN:+v.anualPorAcao,y=v.rendimentoPct==null?NaN:+v.rendimentoPct,u=pair(v.ultimo);
 if(!isFinite(t)||GEN-t>DIV_MAX_DIAS*DAY)return{ok:false,motivo:'Dividend data is too old'};
 if(!isFinite(a)||a<0)return{ok:false,motivo:v.nota||'No annual dividend in the data'};
 return{ok:true,anual:a,rend:isFinite(y)&&y>=0&&y<=25?y:null,freq:+v.frequencia||null,ult:u.length===2&&isFinite(+u[1])?[String(u[0]),+u[1]]:null,t,antigo:v.estado==='previous run',fonte:String(v.fonte||'')};}
function dividendos(R){R=R||pfDados();const fx=FX.length?FX[FX.length-1]:null;
 const L=DIV_IDS.map(id=>{const p=R.find(x=>x.id===id)||{},q=p.q>1e-12?p.q:0,d=divDados(id),usd=d.ok&&q>0?c2(q*d.anual):null;
  const eu=usd!=null&&fx?c2(q*d.anual/fx[1]):null;return{id,q,d,usd,eur:eu,liq:eu!=null?c2(eu*DIV_LIQ):null,v:p.v!=null?p.v:null};});
 const inc=L.filter(x=>x.usd!=null),falta=L.filter(x=>x.q>0&&!x.d.ok);
 return{L,fx,fxSrc:FXSRC,inc,falta,totUsd:c2(inc.reduce((s,x)=>s+x.usd,0)),totEur:fx?c2(inc.reduce((s,x)=>s+x.eur,0)):null,totLiq:fx?c2(inc.reduce((s,x)=>s+x.liq,0)):null};}
function drawDiv(R){const box=$('#tbl-div');if(!box)return;const X=dividendos(R),usd=v=>v==null||!isFinite(v)?'—':'$'+nf(v,2),algum=X.L.some(x=>x.q>0);
 box.innerHTML='<thead><tr><th scope="col">Asset</th><th class="num" scope="col">Shares held</th><th class="num" scope="col">Dividend per share, a year</th><th class="num" scope="col">Yield</th><th class="num" scope="col">Projected, dollars</th><th class="num" scope="col">Projected, euros</th><th class="num" scope="col">Net (est.), euros</th></tr></thead><tbody>'+
  X.L.map(x=>{const d=x.d,sub=d.ok?`<span class="td-sub${d.antigo?' old':''}">${d.antigo?esc(d.fonte)+' · ':''}${d.ult?`latest $${nf(d.ult[1],4)} (ex-date ${fdS(d.ult[0])})${d.freq?` × ${d.freq} a year`:''}`:''}</span>`:`<span class="td-sub old">${esc(d.motivo)}</span>`;
   return`<tr><td>${coTag(x.id)}</td><td class="num">${x.q>0?btcF(x.q):'—'}</td><td class="num">${d.ok?usd(d.anual):'Unavailable'}${sub}</td><td class="num">${d.ok&&d.rend!=null?nf(d.rend,2)+'%':'—'}</td>`+
    (x.q<=0?'<td class="num dim">No holdings</td><td class="num dim">No holdings</td>':!d.ok?'<td class="num">Unavailable</td><td class="num">Unavailable</td>':`<td class="num">${usd(x.usd)}</td><td class="num">${x.eur!=null?eur(x.eur,2):'FX unavailable'}</td>`)+
    (x.q<=0?'<td class="num dim">No holdings</td>':!d.ok?'<td class="num">Unavailable</td>':`<td class="num">${x.liq!=null?eur(x.liq,2):'FX unavailable'}</td>`)+'</tr>';}).join('')+
  `</tbody><tfoot><tr><th scope="row">Total</th><td></td><td></td><td></td><td class="num" id="div-tusd">${X.inc.length?usd(X.totUsd):'—'}</td><td class="num" id="div-teur">${X.inc.length?(X.totEur!=null?eur(X.totEur,2):'FX unavailable'):'—'}</td><td class="num" id="div-tliq">${X.inc.length?(X.totLiq!=null?eur(X.totLiq,2):'FX unavailable'):'—'}</td></tr></tfoot>`;
 const nn=$('#divNetNote'),Q=ANEXO_J.q8a;if(nn)nn.textContent=`Net (est.) = projected × ${nf(DIV_LIQ,2)}: an estimate of what is left after tax, assuming ${nf(Q.retFonte*100,0)}% withheld in the US (form W-8BEN) and the top-up to ${nf(Q.taxaPT*100,0)}% in Portugal (the special rate, without englobamento). What your broker withholds and converts may differ.`;
 const k=$('#divKpis');if(!k)return;
 if(!algum){k.innerHTML=`<p class="empty">${ic('info')}You hold no Apple, NVIDIA or Alphabet shares, so there are no dividends to project.</p>`;return;}
 const sem=X.falta.length?`without ${X.falta.map(x=>esc(nomePF(x.id))).join(', ')} (data unavailable)`:'';
 if(!X.inc.length){k.innerHTML=`<p class="empty">${ic('triangle')}Dividend data is unavailable for what you hold (see Sources &amp; method), so nothing is projected.</p>`;return;}
 const val=X.inc.every(x=>x.v>0)?X.inc.reduce((s,x)=>s+x.v,0):0,fx=X.fx;
 k.innerHTML=kpi('calendar','Projected dividends',X.totEur!=null?eur(X.totEur,2):'FX unavailable',(X.totEur!=null?'a year, before tax':'the dollar total is in the table')+(sem?' · '+sem:''))+
  kpi('euro','Per month, on average',X.totEur!=null?eur(X.totEur/12,2):'—','usually paid every 3 months')+
  kpi('chart','Dividend yield',X.totEur!=null&&val>0?nf(X.totEur/val*100,2)+'%':'—','on what you hold of the three, at the latest prices')+
  kpi('swap','EUR/USD used',fx?nf(fx[1],4):'unavailable',fx?`latest rate, ${fdt(fx[0])} (${esc(X.fxSrc)})`:'no rate: euro amounts are not shown');}

/* ---------- impostos: ficheiros para preencher o anexo J do IRS ---------- */
/* Ficheiros de preparação: o Portal das Finanças não os importa. Mapeamento oficial num só sítio (Modelo 3, anexo J,
   modelo em vigor a partir de janeiro de 2026, e respetivas instruções da AT). Quadro 9.2A, art.º 10.º, n.º 1, al. b) do CIRS:
   G01 = alienação onerosa de ações/partes sociais; G20 = resgates ou alienação de unidades de participação ou liquidação de
   fundos de investimento. Quadro 9.4A, al. k): criptoativos que não são valores mobiliários detidos menos de 365 dias (sem
   código). País: códigos da Tabela X. Se a AT mudar códigos ou colunas, muda-se aqui (e nas colunas abaixo). */
const ANEXO_J={modelo:'Modelo 3 – Anexo J (model in force from January 2026)',
 q92:{quadro:'9.2A',ativos:{AAPL:{codigo:'G01',pais:'840',admitido:'Sim',isin:'US0378331005'},NVDA:{codigo:'G01',pais:'840',admitido:'Sim',isin:'US67066G1040'},
  GOOGL:{codigo:'G01',pais:'840',admitido:'Sim',isin:'US02079K3059'},SXR8:{codigo:'G20',pais:'372',admitido:'Sim',isin:'IE00B5BMR087'},
  /* os outros ETF UCITS da iShares: a mesma estrutura legal do SXR8 (subfundo de uma sociedade de investimento irlandesa,
     domicílio Irlanda = 372 na Tabela X, admitido à negociação na Xetra), por isso o mesmo tratamento (G20) */
  EUNK:{codigo:'G20',pais:'372',admitido:'Sim',isin:'IE00B4K48X80'},IS3N:{codigo:'G20',pais:'372',admitido:'Sim',isin:'IE00BKM4GZ66'},EUNN:{codigo:'G20',pais:'372',admitido:'Sim',isin:'IE00B4L5YX21'}}},
 q94:{quadro:'9.4A',dias:365,fora:'Not Anexo J: held 365 days or more (exempt; Anexo G1, Quadro 7)'},
 /* Quadro 8A (rendimentos de capitais obtidos no estrangeiro), Tabela V das instruções: E10 = dividendos ou lucros com
    retenção em Portugal, E11 = sem retenção em Portugal. Por omissão E11 (corretora sem retenção em Portugal); com retenção
    em Portugal usa-se E10 e preenchem-se as colunas do imposto retido em Portugal. Retenção nos EUA: 15% com o formulário
    W-8BEN (estimativa); taxa especial em Portugal: 28%. diasPag: datas ex nos últimos dias do ano (o pagamento pode cair no
    ano seguinte, e o ano fiscal segue a data de pagamento) */
 q8a:{quadro:'8A',codigo:'E11',codigoRetPT:'E10',retFonte:0.15,taxaPT:0.28,diasPag:45}};
const DIV_LIQ=1-ANEXO_J.q8a.taxaPT;  /* líquido estimado dos dividendos projetados: 15% nos EUA mais o complemento até 28% em Portugal */
/* EUR/USD de referência de um dia: histórico longo do Yahoo e, sem ele, a série usada no resto da página (1 ano ou BCE);
   nunca fora do período coberto (fxPt). É uma taxa diária: as compras e vendas só guardam a data, não a hora. */
function fxRef(iso){const t=Date.parse(iso+'T00:00:00Z');if(!isFinite(t))return null;let p=fxPt(t,HIST.FX),src='Yahoo Finance (long-term history)';if(!p){p=fxPt(t,FX);src=FXSRC;}return p?{fx:p[1],d:isoU(p[0]),src}:null;}
const ISO_RE=/^\d{4}-\d{2}-\d{2}$/,isoOk=s=>{if(!ISO_RE.test(s||''))return false;const t=Date.parse(s+'T00:00:00Z');return isFinite(t)&&isoU(t)===s;};
/* uma linha por par venda/lote do FIFO da carteira (carteira().vendas[].usados): valores em euros dos preços registados */
function linhaFiscal(v,m){const btc=v.a==='BTC',map=ANEXO_J.q92.ativos[v.a]||{},usd=!btc&&((ATIVOS.find(a=>a.id===v.a)||{}).moeda||'').toUpperCase()==='USD';
 const tA=Date.parse(m.d+'T00:00:00Z'),tV=Date.parse(v.d+'T00:00:00Z'),datas=isoOk(m.d),dias=datas?Math.round((tV-tA)/DAY):null,curto=datas?tV<tA+ANEXO_J.q94.dias*DAY:null;
 const aq=c2(m.q*m.cu),vd=c2(m.q*v.p),fa=usd?fxRef(m.d):null,fv=usd?fxRef(v.d):null,rev=[],obs=[];
 if(!datas)rev.push('BAD_PURCHASE_DATE');if(m.aviso||v.aviso)rev.push('SPLIT_CHECK');if(!(v.p>0))rev.push('NO_SALE_PRICE');if(!btc&&!map.codigo)rev.push('NO_ANEXO_J_MAPPING');
 if(btc&&dias!=null&&Math.abs(dias-ANEXO_J.q94.dias)<=2)rev.push('NEAR_365_DAYS');if(usd&&(!fa||!fv))obs.push('NO_FX_REFERENCE');
 return{id:v.a,ativo:nomePF(v.a),isin:btc?'':map.isin||'',codigo:btc?'':map.codigo||'',pais:btc?'':map.pais||'',admitido:btc?'':map.admitido||'',
  quadro:btc?(curto===false?ANEXO_J.q94.fora:ANEXO_J.q94.quadro):ANEXO_J.q92.quadro,dV:v.d,dA:datas?m.d:String(m.d||''),dias,curto,q:m.q,kA:m.k||1,kV:v.k||1,
  cuA:m.cu,cuV:+v.p,aq,vd,ganho:c2(vd-aq),fa,fv,lote:String(m.id),venda:String(v.id),status:rev.length?'REVIEW':'OK',flags:rev.concat(obs)};}
/* parte de uma venda sem compra que a cubra (dados antigos ou restaurados): não há aquisição para declarar */
function linhaFalta(v){const btc=v.a==='BTC',map=ANEXO_J.q92.ativos[v.a]||{};
 return{id:v.a,ativo:nomePF(v.a),isin:btc?'':map.isin||'',codigo:btc?'':map.codigo||'',pais:btc?'':map.pais||'',admitido:btc?'':map.admitido||'',quadro:btc?ANEXO_J.q94.quadro:ANEXO_J.q92.quadro,
  dV:v.d,dA:'',dias:null,curto:null,q:v.falta,kA:null,kV:v.k||1,cuA:null,cuV:+v.p,aq:null,vd:c2(v.falta*v.p),ganho:null,fa:null,fv:null,lote:'',venda:String(v.id),status:'REVIEW',flags:['OVERSOLD_NO_PURCHASE']};}
/* todas as linhas de um ano (pela data da venda), da carteira atual; não escreve nada. Vendas com data, ativo, quantidade
   ou preço inválidos (de qualquer ano) não entram e são contadas à parte, para o resumo avisar */
function anexoJ(ano){const C=carteira(),acoes=[],cripto=[],mal=arr(store.get('sales',[])).filter(v=>!v||!isoOk(v.d)||!COMPRAVEIS.includes(v.a)||!(+v.q>0)||!(+v.p>=0)).map(v=>v&&v.id);
 C.vendas.forEach(v=>{if(!isoOk(v.d)||!COMPRAVEIS.includes(v.a))return;if(+v.d.slice(0,4)!==ano)return;
  const L=v.a==='BTC'?cripto:acoes,R=arr(v.usados).map(m=>linhaFiscal(v,m)).concat(v.falta>0?[linhaFalta(v)]:[]);despesas(v,R);R.forEach(x=>L.push(x));});
 return{ano,acoes,cripto,mal,csvAcoes:csvLinhas(COLS_ACOES,acoes,ano),csvCripto:csvLinhas(COLS_CRIPTO,cripto,ano)};}
/* Despesas e encargos de cada par venda/compra: comissão da compra × (quantidade usada ÷ quantidade da compra) + comissão da
   venda × (quantidade do par ÷ quantidade da venda). Cada linha é arredondada aos cêntimos e o resíduo vai para a última
   linha da venda (as somas batem certo). Sem nenhuma das duas: em branco (como antes); só uma: preenchida com ela e
   assinalada FEE_PARTIAL; as duas: "Despesas e encargos" sai da lista "Fill in manually" dessa linha. */
function qtdCompra(id){const b=arr(store.get('buys',[])).find(x=>x&&String(x.id)===String(id));if(b)return efetiva(b).q;const l=arr(store.get('lots',[])).find(x=>x&&String(x.id)===String(id));return l?+l.q:null;}
function despesas(v,R){const fS=feeDe(v.id),E=R.map(r=>{const fB=r.lote?feeDe(r.lote):null,qB=r.lote?qtdCompra(r.lote):null,pB=fB!=null&&qB>0?fB*r.q/qB:null,pS=fS!=null&&v.q>0?fS*r.q/v.q:null;
  return{n:(pB!=null)+(pS!=null),e:pB!=null||pS!=null?(pB||0)+(pS||0):null};});
 const com=E.map((x,i)=>i).filter(i=>E[i].e!=null),tot=c2(com.reduce((s,i)=>s+E[i].e,0));let acc=0;
 com.forEach((i,j)=>{const val=j<com.length-1?c2(E[i].e):c2(tot-acc);acc+=val;R[i].desp=val;R[i].despOk=E[i].n===2;if(E[i].n===1)R[i].flags.push('FEE_PARTIAL');});}
/* CSV: separador ";" e vírgula decimal (o formato do Excel em português), sem separador de milhares nem símbolos,
   datas ISO (aaaa-mm-dd), UTF-8 com BOM, linhas CRLF. Números formatados aqui, nunca pelo idioma do browser. */
const CSV_SEP=';',CSV_EOL='\r\n';
const csvTxt=s=>{s=String(s==null?'':s);if(/^[=+\-@\t\r]/.test(s))s="'"+s;return/[;",\r\n]|^\s|\s$/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;};
const csvNum=(v,d)=>{if(v==null||v===''||!isFinite(v))return'';const s=Math.abs(+v).toFixed(d),neg=+v<0&&/[1-9]/.test(s);return(neg?'-':'')+s.replace('.',',');};
function csvLinhas(cols,L,ano){return String.fromCharCode(0xFEFF)+[cols.map(c=>csvTxt(c[0])).join(CSV_SEP)].concat(L.map((r,i)=>cols.map(c=>{const v=c[1](r,i,ano);return c[2]!=null?csvNum(v,c[2]):csvTxt(v);}).join(CSV_SEP))).join(CSV_EOL)+CSV_EOL;}
const ymd=(iso,k)=>ISO_RE.test(iso||'')?+iso.split('-')[k]:null,simNao=b=>b==null?'':b?'Yes':'No',fxu=(f,t)=>f&&t!=null?t*f.fx:null;
/* colunas: [cabeçalho, valor, casas decimais (só nos números)]. As do formulário oficial têm o nome do formulário. */
const COLS_BASE=[['Tax year',(r,i,a)=>a,0],['Row',(r,i)=>i+1,0],['Anexo J table',r=>r.quadro],['Asset',r=>r.ativo],['Ticker',r=>r.id]];
const COLS_DATAS=[['Realização Ano',r=>ymd(r.dV,0),0],['Realização Mês',r=>ymd(r.dV,1),0],['Realização Dia',r=>ymd(r.dV,2),0],['Realização Valor (EUR)',r=>r.vd,2],
 ['Aquisição Ano',r=>ymd(r.dA,0),0],['Aquisição Mês',r=>ymd(r.dA,1),0],['Aquisição Dia',r=>ymd(r.dA,2),0],['Aquisição Valor (EUR)',r=>r.aq,2],
 ['Despesas e encargos (EUR)',r=>r.desp,2],['Imposto pago no estrangeiro (EUR)',()=>''],['País da Contraparte',()=>'']];
const COLS_FIM=[['Sale date',r=>r.dV],['Acquisition date',r=>r.dA],['Days held',r=>r.dias,0],['Held under 365 days',r=>simNao(r.curto)],['Capital gain/loss (EUR)',r=>r.ganho,2],
 ['FIFO lot reference',r=>r.lote],['Sale reference',r=>r.venda],['Status',r=>r.status],['Flags',r=>r.flags.join(' ')]];
const COLS_ACOES=COLS_BASE.concat([['ISIN',r=>r.isin],['Transaction type',r=>r.lote?'Sale matched to a purchase (FIFO)':'Sale without a matching purchase'],['País da fonte',r=>r.pais],['Código',r=>r.codigo]],COLS_DATAS,
 [['Respeita a valores mobiliários admitidos à negociação ou a partes de OIC abertos?',r=>r.admitido]],COLS_FIM,
 [['Quantity (shares of today, split-adjusted)',r=>r.q,8],['Purchase split factor',r=>r.kA,6],['Sale split factor',r=>r.kV,6],['Acquisition unit price (EUR)',r=>r.cuA,6],['Sale unit price (EUR)',r=>r.cuV,6],
  ['Acquisition FX EUR/USD',r=>r.fa&&r.fa.fx,6],['Acquisition FX date',r=>r.fa?r.fa.d:''],['Acquisition unit price (USD, derived)',r=>fxu(r.fa,r.cuA),6],['Acquisition total (USD, derived)',r=>r.fa&&r.aq!=null?c2(r.aq*r.fa.fx):null,2],
  ['Sale FX EUR/USD',r=>r.fv&&r.fv.fx,6],['Sale FX date',r=>r.fv?r.fv.d:''],['Sale unit price (USD, derived)',r=>fxu(r.fv,r.cuV),6],['Sale total (USD, derived)',r=>r.fv?c2(r.vd*r.fv.fx):null,2],
  ['FX source',r=>[r.fa&&r.fa.src,r.fv&&r.fv.src].filter((x,j,A)=>x&&A.indexOf(x)===j).join(' / ')],['Price basis',r=>((ATIVOS.find(a=>a.id===r.id)||{}).moeda||'').toUpperCase()==='EUR'?'EUR as entered (priced in euros)':'EUR as entered; USD columns derived at the EUR/USD of each date'],
  ['Fill in manually',r=>r.despOk?'País da Contraparte; Imposto pago no estrangeiro (if any)':'País da Contraparte; Despesas e encargos; Imposto pago no estrangeiro (if any)']]);
const COLS_CRIPTO=COLS_BASE.concat([['Transaction type',r=>r.lote?'Sale matched to a purchase (FIFO)':'Sale without a matching purchase'],['País da fonte',()=>'']],COLS_DATAS,COLS_FIM,
 [['Quantity (BTC)',r=>r.q,8],['Acquisition unit price (EUR)',r=>r.cuA,2],['Sale unit price (EUR)',r=>r.cuV,2],
  ['Fill in manually',r=>r.curto===false?'':r.despOk?'País da fonte (country of the platform); País da Contraparte; Imposto pago no estrangeiro (if any)':'País da fonte (country of the platform); País da Contraparte; Despesas e encargos; Imposto pago no estrangeiro (if any)']]);
function anosFiscais(){const s=new Set();arr(store.get('sales',[])).forEach(v=>{if(v&&isoOk(v.d))s.add(+v.d.slice(0,4));});return[...s].sort((a,b)=>b-a);}
function drawTax(){const sel=$('#taxYear');if(!sel)return;const A=anosFiscais(),cur=+sel.value,L=A.length?A:[new Date(hojeL()).getUTCFullYear()];
 sel.innerHTML=L.map(y=>`<option value="${y}">${y}</option>`).join('');sel.value=String(L.includes(cur)?cur:L[0]);}
function exportaAnexoJ(ano){const msg=$('#taxMsg');if(!(ano>=2000&&ano<=2100)){msg.textContent='Choose a tax year.';return null;}
 const X=anexoJ(ano),T=X.acoes.concat(X.cripto),rev=T.filter(r=>r.status==='REVIEW').length,semFx=T.filter(r=>r.flags.includes('NO_FX_REFERENCE')).length,
  nv=L=>new Set(L.map(r=>r.venda)).size,c94=X.cripto.filter(r=>r.quadro===ANEXO_J.q94.quadro).length,cFora=X.cripto.length-c94;
 if(!T.length&&!confirm(`There are no sales in ${ano}. Download the two files anyway (headers only)?`)){msg.textContent='Export cancelled.';return X;}
 if(T.length&&!confirm(`Tax year: ${ano}\nStock/ETF disposals: ${nv(X.acoes)} (${X.acoes.length} row${X.acoes.length===1?'':'s'} for Quadro 9.2A, one per purchase used)\nCrypto disposals: ${nv(X.cripto)} (${c94} row${c94===1?'':'s'} under 365 days for Quadro 9.4A${cFora?`, ${cFora} held 365 days or more: not Anexo J`:''})\nReview items: ${rev}${semFx?`\nDollar reference unavailable for ${semFx} row${semFx===1?'':'s'} (euro amounts not affected)`:''}${X.mal.length?`\nSales with an invalid date or asset, not exported: ${X.mal.length}`:''}\n\nThese are preparation files: the Portal das Finanças does not import them. Download them?`)){msg.textContent='Export cancelled.';return X;}
 baixa(`AnexoJ_Stocks_ETFs_${ano}.csv`,X.csvAcoes,'text/csv;charset=utf-8');setTimeout(()=>baixa(`AnexoJ_Crypto_${ano}.csv`,X.csvCripto,'text/csv;charset=utf-8'),400);
 msg.textContent=`Downloaded AnexoJ_Stocks_ETFs_${ano}.csv and AnexoJ_Crypto_${ano}.csv${rev?`: ${rev} row${rev===1?' needs':'s need'} review (Status and Flags columns)`:''}. Nothing in your portfolio was changed.`;return X;}
function buildTax(){const f=$('#taxForm');if(!f)return;f.addEventListener('submit',e=>{e.preventDefault();exportaAnexoJ(+$('#taxYear').value);});drawTax();
 const g=$('#div8aForm');if(g)g.addEventListener('submit',e=>{e.preventDefault();exportaDiv8A(+$('#divYear').value);});drawDiv8A();}

/* ---------- impostos: dividendos (anexo J, Quadro 8A) ---------- */
/* Pagamentos de dividendos recolhidos pelo script ([data ex, dólares por ação], últimos 2 anos), só de dados válidos (estado
   "ok" ou execução anterior); nada é inventado. O Yahoo dá os valores ajustados a desdobramentos, tal como efetiva() dá as
   quantidades em ações de hoje. */
function divPagamentos(id){const v=D.dividendos&&D.dividendos[id];if(!v||typeof v!=='object')return{ok:false,motivo:'no dividend data in this run'};
 if(v.estado!=='ok'&&v.estado!=='previous run')return{ok:false,motivo:'source failed'+(v.erro?': '+v.erro:'')};
 const t=Date.parse(v.obtidoEm||'');if(!isFinite(t))return{ok:false,motivo:'no retrieval date in the data'};
 const P=arr(v.pagamentos).map(pair).filter(p=>p.length===2&&isoOk(String(p[0]))&&+p[1]>0&&isFinite(+p[1])).map(p=>[String(p[0]),+p[1]]).sort((a,b)=>a[0]<b[0]?-1:a[0]>b[0]?1:0);
 return{ok:true,P,de:isoU(t-730*DAY),ate:isoU(t),fonte:String(v.fonte||'Yahoo Finance')};}
/* ações com direito a um pagamento, pelo FIFO de carteira(): os lotes comprados antes da data ex menos o que as vendas
   anteriores à data ex gastaram deles. Uma compra na própria data ex não recebe; uma venda na data ex ainda recebe. */
function comDireito(C,id,ex){let q=0,aviso=false;arr(C.lotes[id]).forEach(x=>{if(x.d<ex){q+=x.q0;if(x.aviso)aviso=true;}});
 C.vendas.forEach(v=>{if(v.a===id&&v.d<ex)arr(v.usados).forEach(m=>{q-=m.q;});});return{q:q>1e-9?q:0,aviso};}
/* uma linha por pagamento com ações detidas, no ano da data ex (a Yahoo não dá a data de pagamento); datas ex futuras não
   entram. Dólares como na fonte; euros só como referência, à taxa da data ex (fxRef: o mesmo fxPt/fxAt do anexo J) */
function anexo8A(ano){const C=carteira(),Q=ANEXO_J.q8a,hoje=isoU(hojeL()),L=[],falta=[],cob=[];
 DIV_IDS.forEach(id=>{const P=divPagamentos(id),tem=arr(C.lotes[id]).length>0;if(!P.ok){if(tem)falta.push({id,motivo:P.motivo});return;}if(tem)cob.push(P);
  P.P.forEach(([ex,dps])=>{if(+ex.slice(0,4)!==ano||ex>hoje)return;const h=comDireito(C,id,ex);if(!(h.q>0))return;
   const map=ANEXO_J.q92.ativos[id]||{},bruto=c2(h.q*dps),ret=map.pais==='840'?c2(bruto*Q.retFonte):null,f=fxRef(ex),rev=[],obs=['BROKER_FX'];
   if(Date.UTC(ano,11,31)-Date.parse(ex+'T00:00:00Z')<Q.diasPag*DAY)rev.push('PAY_DATE_UNKNOWN');if(h.aviso)rev.push('SPLIT_CHECK');if(!map.pais)rev.push('NO_ANEXO_J_MAPPING');if(!f)obs.push('NO_FX_REFERENCE');
   L.push({id,ativo:nomePF(id),isin:map.isin||'',pais:map.pais||'',codigo:Q.codigo,ex,dps,q:h.q,bruto,ret,fx:f,eur:f?c2(bruto/f.fx):null,retEur:f&&ret!=null?c2(ret/f.fx):null,
    fonte:P.fonte,status:rev.length?'REVIEW':'OK',flags:rev.concat(obs)});});});
 L.sort((a,b)=>a.ex<b.ex?-1:a.ex>b.ex?1:a.id<b.id?-1:a.id>b.id?1:0);
 return{ano,L,falta,cob,csv:csvLinhas(COLS_8A,L,ano)};}
const COLS_8A=[['Tax year',(r,i,a)=>a,0],['Row',(r,i)=>i+1,0],['Anexo J table',()=>ANEXO_J.q8a.quadro],['Asset',r=>r.ativo],['Ticker',r=>r.id],['ISIN',r=>r.isin],
 ['Código rendim.',r=>r.codigo],['País da fonte',r=>r.pais],['Ex-dividend date',r=>r.ex],['Dividend per share (USD)',r=>r.dps,6],['Shares entitled (held before the ex-date, shares of today)',r=>r.q,8],
 ['Gross amount (USD)',r=>r.bruto,2],['Tax withheld in the US, estimated (USD)',r=>r.ret,2],['EUR/USD at the ex-date',r=>r.fx&&r.fx.fx,6],['FX date',r=>r.fx?r.fx.d:''],['FX source',r=>r.fx?r.fx.src:''],
 ['Rendimento bruto (EUR, reference)',r=>r.eur,2],['Imposto pago no estrangeiro – No país da fonte (EUR, reference)',r=>r.retEur,2],
 ['Imposto retido em Portugal – NIF da entidade retentora',()=>''],['Imposto retido em Portugal – Retenção na fonte (EUR)',()=>''],
 ['Dividend data source',r=>r.fonte],['Status',r=>r.status],['Flags',r=>r.flags.join(' ')],
 ['Fill in manually',()=>`Rendimento bruto and Imposto pago no estrangeiro in euros, as converted by your broker; if tax was withheld in Portugal: Código ${ANEXO_J.q8a.codigoRetPT} and the Imposto retido em Portugal columns`]];
/* anos com algum pagamento recebido (ações detidas na data ex), do mais recente para o mais antigo */
function anosDiv(){const C=carteira(),hoje=isoU(hojeL()),s=new Set();DIV_IDS.forEach(id=>{const P=divPagamentos(id);if(P.ok)P.P.forEach(([ex])=>{if(ex<=hoje&&comDireito(C,id,ex).q>0)s.add(+ex.slice(0,4));});});return[...s].sort((a,b)=>b-a);}
function drawDiv8A(){const sel=$('#divYear');if(!sel)return;const A=anosDiv(),cur=+sel.value,L=A.length?A:[new Date(hojeL()).getUTCFullYear()];
 sel.innerHTML=L.map(y=>`<option value="${y}">${y}</option>`).join('');sel.value=String(L.includes(cur)?cur:L[0]);}
function exportaDiv8A(ano){const msg=$('#divMsg');if(!(ano>=2000&&ano<=2100)){msg.textContent='Choose a tax year.';return null;}
 const X=anexo8A(ano),n=X.L.length,rev=X.L.filter(r=>r.status==='REVIEW').length,semFx=X.L.filter(r=>r.flags.includes('NO_FX_REFERENCE')).length,
  falta=X.falta.length?`Dividend data unavailable for ${X.falta.map(x=>nomePF(x.id)+' ('+x.motivo+')').join(', ')}: no rows for ${X.falta.length===1?'it':'them'}.`:'',
  cob=X.cob.length?`Payments in the data: ex-dates from ${fdS(X.cob.map(c=>c.de).sort()[0])} to ${fdS(X.cob.map(c=>c.ate).sort().pop())}.`:'';
 if(!n&&!confirm(`There are no dividends for what you held in ${ano}.${falta?'\n'+falta:''}${cob?'\n'+cob:''}\nDownload the file anyway (headers only)?`)){msg.textContent='Export cancelled.';return X;}
 if(n&&!confirm(`Tax year: ${ano} (by ex-dividend date)\nDividend payments: ${n} row${n===1?'':'s'} for Quadro 8A\nReview items: ${rev}${semFx?`\nEUR/USD unavailable for ${semFx} row${semFx===1?'':'s'} (dollar amounts not affected)`:''}${falta?'\n'+falta:''}${cob?'\n'+cob:''}\n\nThis is a preparation file: the Portal das Finanças does not import it. Download it?`)){msg.textContent='Export cancelled.';return X;}
 baixa(`AnexoJ_Dividends_${ano}.csv`,X.csv,'text/csv;charset=utf-8');
 msg.textContent=`Downloaded AnexoJ_Dividends_${ano}.csv${rev?`: ${rev} row${rev===1?' needs':'s need'} review (Status and Flags columns)`:''}.${falta?' '+falta:''}${cob?' '+cob:''} Nothing in your portfolio was changed.`;return X;}

/* ---------- cópia dos dados num ficheiro da pasta do projeto ---------- */
/* "Save to project folder" grava bluechip-board-backup.json na pasta escolhida (o Chrome lembra-se dela) e, enquanto a página
   estiver aberta, as alterações seguintes vão para lá sozinhas. Ao reabrir a página, o Chrome volta a pedir autorização:
   a caixa de estado avisa quando há alterações que ainda não estão no ficheiro. O script lê o ficheiro e inclui-o no site,
   que o junta ao que o browser tem (por id: nada é apagado, exceto o que foi apagado de propósito). Sem acesso direto a
   pastas, o ficheiro vai para as Transferências e o script muda-o para a pasta do projeto na execução seguinte. */
const BK_NOME='bluechip-board-backup.json';let FH=null,bkT=0,AUTO=false;
const bkMsg=t=>{const m=$('#bkMsg');if(m)m.textContent=t;};
const listasLocais=()=>({buys:arr(store.get('buys',[])),lots:arr(store.get('lots',[])),sales:arr(store.get('sales',[]))});
/* versão 5: a mesma da 4, mais "targets" (alocação-alvo), só quando existe */
function dadosBackup(){const L=listasLocais(),alvos=limpaAlvos(store.get('targets',null)),pol=limpaPolitica(store.get('policy',null)),notas=notasAtuais();
 return Object.assign({app:'Bluechip Board',version:5,saved:store.get('savedAt',null)||new Date().toISOString(),exported:new Date().toISOString(),buys:L.buys,lots:L.lots,sales:L.sales,deleted:apagados()},
  alvos?{targets:alvos}:{},pol?{policy:pol}:{},Object.keys(notas).length?{notes:notas}:{},Object.keys(feesAtuais()).length?{fees:feesAtuais()}:{});}
/* comissões por entrada (bb.fees, "fees" no backup): {id: {v (euros, ≥ 0), at}}; as inválidas são ignoradas */
function limpaFees(m){const o={};if(m&&typeof m==='object'&&!Array.isArray(m))Object.keys(m).forEach(id=>{const n=m[id];if(n&&typeof n==='object'&&n.v!==''&&n.v!==null&&typeof n.v!=='boolean'&&isFinite(+n.v)&&+n.v>=0){const ta=Date.parse(n.at||'');o[String(id)]={v:+n.v,at:isFinite(ta)?new Date(ta).toISOString():null};}});return o;}
/* só as comissões de entradas que existem; feeDe(id): a comissão conhecida em euros, ou null */
function feesAtuais(){const ids=new Set(['buys','lots','sales'].flatMap(k=>arr(store.get(k,[])).map(x=>String(x&&x.id)))),F=limpaFees(store.get('fees',null)),o={};Object.keys(F).forEach(id=>{if(ids.has(id))o[id]=F[id];});return o;}
function feeDe(id){if(id==null)return null;const F=feesAtuais();return F[String(id)]?F[String(id)].v:null;}
function guardaFeeForm(id){const i=$('#buyFee');if(!i)return;const s=i.value.trim();if(s===''||!(isFinite(+s)&&+s>=0))return;const F=limpaFees(store.get('fees',null));F[String(id)]={v:+s,at:new Date().toISOString()};store.set('fees',F);}
/* política de investimento (bb.policy, "policy" no backup): textos livres (até 2 000 caracteres cada) e "at". Inválida → null */
const POL_K=['horizon','allocation','monthly','drop20','drop30','sell'];
function limpaPolitica(p){if(!p||typeof p!=='object'||Array.isArray(p))return null;const o={};POL_K.forEach(k=>{o[k]=typeof p[k]==='string'?p[k].slice(0,2000):'';});const ta=Date.parse(p.at||'');o.at=isFinite(ta)?new Date(ta).toISOString():null;return o;}
/* notas por entrada (bb.notes, "notes" no backup): {id: {t, at}}; as inválidas são ignoradas */
function limpaNotas(m){const o={};if(m&&typeof m==='object'&&!Array.isArray(m))Object.keys(m).forEach(id=>{const n=m[id];if(n&&typeof n==='object'&&typeof n.t==='string'){const ta=Date.parse(n.at||'');o[String(id)]={t:n.t.slice(0,1000),at:isFinite(ta)?new Date(ta).toISOString():null};}});return o;}
/* só as notas de entradas que existem (as de entradas apagadas, ou sem entrada, são ignoradas) */
function notasAtuais(){const ids=new Set(['buys','lots','sales'].flatMap(k=>arr(store.get(k,[])).map(x=>String(x&&x.id)))),N=limpaNotas(store.get('notes',null)),o={};Object.keys(N).forEach(id=>{if(ids.has(id))o[id]=N[id];});return o;}
/* alocação-alvo guardada (bb.targets, e "targets" no backup): pesos por ativo (só ativos conhecidos, números ≥ 0, a somar 100),
   banda em pontos percentuais (mais de 0 e até 50; por omissão 5), valor mensal ≥ 0 e "at" (quando foi gravada). Inválida → null */
function limpaAlvos(t){if(!t||typeof t!=='object'||!t.weights||typeof t.weights!=='object'||Array.isArray(t.weights))return null;const w={};let s=0;
 for(const k of Object.keys(t.weights)){const r=t.weights[k],v=+r;if(!COMPRAVEIS.includes(k)||r===''||r===null||typeof r==='boolean'||!isFinite(v)||v<0)return null;if(v>0)w[k]=v;s+=v;}
 if(Math.abs(s-100)>0.01)return null;const b=+t.band,m=+t.monthly,ta=Date.parse(t.at||'');
 return{weights:w,band:isFinite(b)&&b>0&&b<=50?b:5,monthly:isFinite(m)&&m>=0?m:0,at:isFinite(ta)?new Date(ta).toISOString():null};}
/* valida tudo o que vem de um ficheiro (o deste projeto ou um restaurado) antes de o usar */
function limpaBackup(d){const ok=x=>x&&typeof x==='object'&&/^\d{4}-\d{2}-\d{2}$/.test(x.d),
  /* só números (ou texto numérico): +true, +null, +'' e +[5] davam 1, 0, 0 e 5, e um custo em falta passava a €0 */
  num=v=>typeof v==='number'?v:typeof v==='string'&&/^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?\s*$/i.test(v)?+v:NaN,
  /* id estável mesmo nas cópias antigas sem id, para juntar o mesmo ficheiro duas vezes não duplicar nada; duas entradas
     iguais (mesmo ativo, data, quantidade e preço, por exemplo duas compras iguais no mesmo dia) ficam ambas, com #2, #3… */
  rep={},nid=(x,a)=>{if(x.id!=null&&x.id!=='')return String(x.id);const b=['x',a,x.d,x.q,x.p!=null?x.p:x.c].join('_');rep[b]=(rep[b]||0)+1;return rep[b]>1?b+'#'+rep[b]:b;},
  ref=r=>Array.isArray(r)&&r.length===2&&/^\d{4}-\d{2}-\d{2}$/.test(r[0])&&+r[1]>0?[r[0],+r[1]]:null,bas=u=>/^\d{4}-\d{2}-\d{2}$/.test(u||'')?u:null,
  base=(o,x)=>{const r=ref(x.r),u=bas(x.u);if(r)o.r=r;if(u)o.u=u;return o;};
 const buys=arr(d.buys).concat(arr(d.etfLots).map(x=>Object.assign({a:'SXR8'},x))).filter(x=>ok(x)&&BUY_IDS.includes(x.a)&&num(x.q)>0&&num(x.p)>0).map(x=>base({id:nid(x,x.a),a:x.a,d:x.d,q:num(x.q),p:num(x.p)},x));
 const lots=arr(d.lots).filter(x=>ok(x)&&num(x.q)>0&&num(x.c)>=0).map(x=>({id:nid(x,'BTC'),d:x.d,q:num(x.q),c:num(x.c)}));
 const sales=arr(d.sales).filter(x=>ok(x)&&COMPRAVEIS.includes(x.a)&&num(x.q)>0&&num(x.p)>=0).map(x=>base({id:nid(x,x.a),a:x.a,d:x.d,q:num(x.q),p:num(x.p)},x));
 const del={},dd=d.deleted;if(dd&&typeof dd==='object'){if(Array.isArray(dd))dd.forEach(i=>{if(i!=null)del[String(i)]='';});else Object.keys(dd).forEach(i=>{del[i]=String(dd[i]==null?'':dd[i]);});}
 const b=baseDe(d.saved||d.exported);completaBase(buys,b);completaBase(sales,b);return{buys,lots,sales,del,targets:limpaAlvos(d.targets),policy:limpaPolitica(d.policy),notes:limpaNotas(d.notes),fees:limpaFees(d.fees)};}
/* junta um backup ao que o browser tem: as entradas são somadas por id e as apagadas (num lado ou no outro) saem.
   Nada do browser se perde. Devolve quantas entradas vieram do ficheiro e quantas saíram por terem sido apagadas. */
function juntaBackup(d,projeto){const F=limpaBackup(d),L0=listasLocais(),del0=apagados(),del=Object.assign({},F.del,del0),tf=d.saved||d.exported||null;
 /* transição: os dados de versões anteriores não guardavam o que foi apagado. Se este browser tem dados mais recentes do que
    o ficheiro e ainda não tem essa lista, uma entrada do ficheiro que falte aqui foi apagada aqui: não volta */
 const sa0=Date.parse(store.get('savedAt',null)||''),legado=projeto&&LEGADO&&(L0.buys.length+L0.lots.length+L0.sales.length)>0&&isFinite(sa0)&&tf&&sa0>=Date.parse(tf);
 LEGADO=false;
 if(legado){const tem=new Set(['buys','lots','sales'].flatMap(k=>L0[k].map(x=>String(x.id))));['buys','lots','sales'].forEach(k=>F[k].forEach(x=>{if(!tem.has(String(x.id)))del[String(x.id)]=new Date().toISOString();}));}
 const fora=x=>String(x.id) in del;let novos=0,saem=0;const res={};
 ['buys','lots','sales'].forEach(k=>{const m=new Map();L0[k].forEach(x=>m.set(String(x.id),x));F[k].forEach(x=>{if(!m.has(String(x.id))&&!fora(x)){m.set(String(x.id),x);novos++;}});
  res[k]=[...m.values()].filter(x=>{if(fora(x)){if(L0[k].includes(x))saem++;return false;}return true;});});
 const idsF=new Set(['buys','lots','sales'].flatMap(k=>F[k].map(x=>String(x.id))));
 /* o browser tem algo que o ficheiro não tem? (entradas ou apagamentos): então o ficheiro precisa de ser atualizado */
 /* alocação-alvo: fica a gravada mais recentemente ("at"); se for a deste browser, o ficheiro precisa de ser atualizado */
 const tL=limpaAlvos(store.get('targets',null)),aT=x=>x?(Date.parse(x.at||'')||0):-1,alvos=!!F.targets&&aT(F.targets)>aT(tL),alvoLocal=!!tL&&aT(tL)>aT(F.targets);
 if(alvos)store.raw('targets',F.targets);
 /* política: idem; notas: nota a nota, fica a mais recente; as de entradas apagadas saem */
 const pL=limpaPolitica(store.get('policy',null)),pol=!!F.policy&&aT(F.policy)>aT(pL),polLocal=!!pL&&aT(pL)>aT(F.policy);if(pol)store.raw('policy',F.policy);
 const nL=limpaNotas(store.get('notes',null)),nM=Object.assign({},nL),idsR=new Set(['buys','lots','sales'].flatMap(k=>res[k].map(x=>String(x.id))));let notas=0;
 Object.keys(F.notes).forEach(id=>{if(!(id in del)&&(!nM[id]||aT(F.notes[id])>aT(nM[id]))){nM[id]=F.notes[id];notas++;}});Object.keys(nM).forEach(id=>{if(id in del)delete nM[id];});
 const notaLocal=Object.keys(nM).some(id=>idsR.has(id)&&(!F.notes[id]||aT(nM[id])>aT(F.notes[id])));
 if(notas||Object.keys(nM).length!==Object.keys(nL).length)store.raw('notes',nM);
 /* comissões: idem, por entrada */
 const fL=limpaFees(store.get('fees',null)),fM=Object.assign({},fL);let fees=0;
 Object.keys(F.fees).forEach(id=>{if(!(id in del)&&(!fM[id]||aT(F.fees[id])>aT(fM[id]))){fM[id]=F.fees[id];fees++;}});Object.keys(fM).forEach(id=>{if(id in del)delete fM[id];});
 const feeLocal=Object.keys(fM).some(id=>idsR.has(id)&&(!F.fees[id]||aT(fM[id])>aT(F.fees[id])));
 if(fees||Object.keys(fM).length!==Object.keys(fL).length)store.raw('fees',fM);
 const extraLocal=['buys','lots','sales'].some(k=>res[k].some(x=>!idsF.has(String(x.id))))||Object.keys(del0).some(i=>!(i in F.del))||alvoLocal||polLocal||notaLocal||feeLocal;
 if(novos||saem||Object.keys(del).length!==Object.keys(del0).length){['buys','lots','sales'].forEach(k=>store.raw(k,res[k]));store.raw('deleted',del);}
 if(extraLocal||legado){store.raw('savedAt',new Date().toISOString());aoMudar();}
 else if(projeto&&tf&&isFinite(Date.parse(tf))){store.raw('savedAt',tf);store.raw('fileSaved',tf);}  /* browser e ficheiro da pasta iguais */
 return{novos,saem,extraLocal,alvos,pol,notas,fees,n:{buys:res.buys.length,lots:res.lots.length,sales:res.sales.length}};}
/* o "puxador" do ficheiro escolhido fica no IndexedDB deste browser, para não ser preciso escolher a pasta outra vez */
function idb(mode,fn){return new Promise((res,rej)=>{let o;try{o=indexedDB.open('bluechip-board',1);}catch(e){rej(e);return;}o.onupgradeneeded=()=>o.result.createObjectStore('h');o.onerror=()=>rej(o.error);
 o.onsuccess=()=>{const db=o.result;let tx,r;try{tx=db.transaction('h',mode);r=fn(tx.objectStore('h'));}catch(e){db.close();rej(e);return;}  /* sem isto um erro aqui deixava a promessa por resolver */
  tx.oncomplete=()=>{db.close();res(r&&r.result);};tx.onerror=()=>{db.close();rej(tx.error);};tx.onabort=()=>{db.close();rej(tx.error);};};});}
const marcaGravado=sv=>{store.raw('fileSaved',sv);store.raw('fileWrittenAt',new Date().toISOString());};
/* descarrega um ficheiro gerado na página (fica só neste computador) */
function baixa(nome,txt,tipo){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([txt],{type:tipo}));a.download=nome;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1000);}
function descarrega(txt,sv){baixa(BK_NOME,txt,'application/json');
 marcaGravado(sv);estadoBk();bkMsg('This browser cannot save straight into a folder, so the backup went to your Downloads folder. The script moves it into the project folder the next time it runs.');}
async function guardaFicheiro(interativo){const dadosB=dadosBackup(),txt=JSON.stringify(dadosB,null,2);
 if(!window.showSaveFilePicker||!window.indexedDB){if(interativo)descarrega(txt,dadosB.saved);else{AUTO=false;estadoBk();}return;}
 try{if(!FH)FH=(await idb('readonly',s=>s.get('file')).catch(()=>null))||null;
  if(FH){let p=await FH.queryPermission({mode:'readwrite'});if(p!=='granted'&&interativo)p=await FH.requestPermission({mode:'readwrite'});if(p!=='granted'){if(!interativo){AUTO=false;estadoBk();return;}FH=null;}}
  if(!FH){if(!interativo){AUTO=false;estadoBk();return;}
   /* escolhe-se a PASTA e confirma-se logo que é a do projeto (tem o bluechip-board.html): uma pasta errada é recusada já,
      em vez de só se saber na execução seguinte do script. Sem seletor de pastas, o seletor de ficheiro de antes. */
   if(window.showDirectoryPicker){const dir=await window.showDirectoryPicker({id:'bluechip-board',mode:'readwrite',startIn:'documents'});
    let doProjeto=false;try{await dir.getFileHandle('bluechip-board.html');doProjeto=true;}catch(e){}
    if(!doProjeto){bkMsg(`The folder "${dir.name}" has no bluechip-board.html, so it is not the BluechipBoard project folder: nothing was saved. Click Save to project folder again and choose the folder where bluechip-board.html and the script are.`);estadoBk();return;}
    FH=await dir.getFileHandle(BK_NOME,{create:true});}
   else FH=await window.showSaveFilePicker({id:'bluechip-board',startIn:'documents',suggestedName:BK_NOME,types:[{description:'Bluechip Board backup',accept:{'application/json':['.json']}}]});
   await idb('readwrite',s=>s.put(FH,'file')).catch(()=>{});}
  const w=await FH.createWritable();await w.write(txt);await w.close();
  AUTO=true;marcaGravado(dadosB.saved);estadoBk();
  bkMsg(`Saved to ${FH.name} at ${hmL(Date.now())}. It must be in the BluechipBoard folder, next to bluechip-board.html, for the site to load it. While this page is open, every change is saved there automatically.`);
 }catch(e){if(e&&e.name==='AbortError'){bkMsg('Save cancelled.');return;}AUTO=false;estadoBk();if(interativo)descarrega(txt,dadosB.saved);}}
const resumoBk=n=>`${n.buys} stock and ETF purchase${n.buys===1?'':'s'}, ${n.lots} Bitcoin purchase${n.lots===1?'':'s'} and ${n.sales} sale${n.sales===1?'':'s'}`;
function aoMudar(){pedeAuto();clearTimeout(bkT);bkT=setTimeout(()=>guardaFicheiro(false),800);estadoBk();}
/* Depois de reabrir a página, o browser só dá de novo acesso ao ficheiro com um gesto do utilizador. A primeira alteração
   feita aqui (um clique em Add, Save, Remove…) é esse gesto: pede-se logo a autorização, uma vez por página, e com ela as
   gravações automáticas voltam. Sem gesto o pedido é recusado pelo browser e fica para a alteração seguinte. */
let PEDIU=false,ALTEROU=false;
function pedeAuto(){if(AUTO||PEDIU||!FH||typeof FH.requestPermission!=='function')return;let p;try{p=FH.requestPermission({mode:'readwrite'});}catch(e){return;}PEDIU=true;
 Promise.resolve(p).then(r=>{if(r==='granted'){AUTO=true;clearTimeout(bkT);guardaFicheiro(false);}else estadoBk();},()=>{PEDIU=false;});}
/* ao abrir: o ficheiro ligado (se o browser o guardou) e, se a autorização ainda estiver dada (alguns browsers deixam-na
   "em todas as visitas"), as gravações automáticas ficam logo ativas e o que faltar no ficheiro é gravado */
function ligaFicheiro(){if(!window.showSaveFilePicker||!window.indexedDB)return Promise.resolve(null);
 return idb('readonly',s=>s.get('file')).then(h=>{if(!h)return null;if(!FH)FH=h;
  return Promise.resolve(h.queryPermission?h.queryPermission({mode:'readwrite'}):'prompt').then(p=>{if(p==='granted'){AUTO=true;if(pendente())guardaFicheiro(false);else estadoBk();}return h;},()=>h);}).catch(()=>null);}
/* caixa de estado: diz quando o browser tem alterações que ainda não estão no ficheiro de backup */
/* o browser tem dados (entradas, apagamentos, alvos, política ou notas) e alterações mais recentes do que o ficheiro de backup? */
function temDados(){const L=listasLocais();return L.buys.length+L.lots.length+L.sales.length+Object.keys(apagados()).length>0||store.get('targets',null)!=null||store.get('policy',null)!=null||Object.keys(notasAtuais()).length>0;}
function refFicheiro(){const ref=[store.get('fileSaved',null),D.backup&&(D.backup.saved||D.backup.exported)].map(x=>Date.parse(x||'')).filter(isFinite);return ref.length?Math.max(...ref):NaN;}
function pendente(){const sa=Date.parse(store.get('savedAt',null)||''),ft=refFicheiro();return temDados()&&isFinite(sa)&&(!isFinite(ft)||sa>ft+1000);}
function estadoBk(){const el=$('#bkState');if(!el)return;const tem=temDados(),ft=refFicheiro(),pend=pendente();
 el.innerHTML=pend?`<div class="callout warn">${ic('triangle')}<div><b>Some changes are not in the backup file yet.</b> ${AUTO?'Saving them now…':FH&&!PEDIU?'Automatic saving is paused: after the page is reopened, the browser needs your permission again. It asks at your next change on this page; or click <b>Save to project folder</b> now to update bluechip-board-backup.json.':FH?'Automatic saving is off: the browser did not allow access to the backup file on this page. Click <b>Save to project folder</b> to update bluechip-board-backup.json.':'Automatic saving is off on this page. Click <b>Save to project folder</b> to update bluechip-board-backup.json.'}</div></div>`:
  (tem&&isFinite(ft)?`<p class="note">${ic('check')}The backup file has all your data${AUTO?'; while this page is open, changes are saved to it automatically':''}.</p>`:'');}
/* aviso: o ficheiro gravado neste browser não estava na pasta do projeto quando o script correu depois disso */
function verificaPasta(){const fw=Date.parse(store.get('fileWrittenAt',null)||''),fs=Date.parse(store.get('fileSaved',null)||''),b=D.backup,tb=b?Date.parse(b.saved||b.exported||''):NaN;
 if(isFinite(fw)&&fw<GEN-60000&&isFinite(fs)&&(!isFinite(tb)||tb<fs-1000))
  return`The backup you saved on ${fdt(fw)} at ${hmL(fw)} was not in the project folder when the script ran afterwards. Click Save to project folder and choose the BluechipBoard folder (where bluechip-board.html is).`;
 return'';}
/* ao abrir: junta o ficheiro da pasta do projeto (incluído no site pelo script) ao que este browser tem */
function carregaDoFicheiro(){const b=D.backup;if(!b||b.app!=='Bluechip Board')return'';const tf=Date.parse(b.saved||b.exported||''),r=juntaBackup(b,true),quando=isFinite(tf)?` (saved ${fdt(tf)}, ${hmL(tf)})`:'';
 return(r.novos||r.saem?`Merged ${BK_NOME}${quando} with this browser: ${r.novos} entr${r.novos===1?'y':'ies'} added${r.saem?`, ${r.saem} removed because ${r.saem===1?'it was':'they were'} deleted elsewhere`:''}. You now have ${resumoBk(r.n)}.`:
  `Your data in this browser matches ${BK_NOME}${quando}${r.extraLocal?', plus changes made here that are not in the file yet':''}.`)+(r.alvos?' Target allocation taken from the file (it was saved more recently).':'')+(r.pol?' Investment policy taken from the file (it was saved more recently).':'')+(r.notas?` ${r.notas} note${r.notas===1?'':'s'} taken from the file.`:'');}
function buildBackup(){const ex=$('#bkExport'),im=$('#bkImport');if(!ex||!im)return;
 ex.addEventListener('click',()=>guardaFicheiro(true));
 im.addEventListener('change',()=>{const f=im.files&&im.files[0];if(!f)return;
  f.text().then(t=>{const d=JSON.parse(t);if(!d||typeof d!=='object'||d.app!=='Bluechip Board')throw new Error('not a backup');
   if(!confirm(`Merge ${f.name} into this browser? Purchases and sales in the file that are missing here are added, and entries deleted in the file are deleted here too. Nothing else is removed.`))return;
   const r=juntaBackup(d,false);bkMsg(`Backup restored: ${r.novos} entr${r.novos===1?'y':'ies'} added${r.saem?`, ${r.saem} removed`:''}. You now have ${resumoBk(r.n)}.${r.alvos?' Target allocation taken from the file (it was saved more recently).':''}${r.pol?' Investment policy taken from the file (it was saved more recently).':''}${r.notas?` ${r.notas} note${r.notas===1?'':'s'} taken from the file.`:''}`);
   if(r.novos||r.saem||r.alvos||r.pol||r.notas||r.fees){store.raw('savedAt',new Date().toISOString());aoMudar();}drawPf();drawLots();drawBuys();estadoBk();})
  .catch(()=>bkMsg('That file is not a Bluechip Board backup.'));im.value='';});
 const msgs=[];const m1=carregaDoFicheiro();if(m1)msgs.push(m1);const m2=verificaPasta();if(m2)msgs.push(m2);
 if(msgs.length)bkMsg(msgs.join(' '));
 ligaFicheiro().then(h=>{if(h&&!msgs.length)bkMsg(AUTO?`Linked to ${h.name}: changes on this page are saved there automatically.`:`Linked to ${h.name}. The browser asks for permission again after the page is reopened: at your first change here, or click Save to project folder.`);});
 /* sair da página com alterações feitas aqui que ainda não estão no ficheiro: o browser pergunta (os dados ficam no browser,
    mas o ficheiro é a cópia que resiste a limpar o browser) */
 addEventListener('beforeunload',e=>{if(ALTEROU&&pendente()){e.preventDefault();e.returnValue='';}});
 drawPf();drawLots();drawBuys();estadoBk();}

/* ---------- sources ---------- */
const stOf=e=>/^ok \(/.test(e)?['warn','triangle']:/^ok/.test(e)?['ok','check']:/skipped/.test(e)?['skip','circle-minus']:/reference|previous run/.test(e)?['warn','triangle']:['err','x'];
function renderSrc(){const F=arr(D.fontes).slice().sort((a,b)=>(/^ok/.test(a.estado)?1:0)-(/^ok/.test(b.estado)?1:0));const ok=F.filter(f=>/^ok/.test(f.estado)).length,bad=F.length-ok;
 $('#cnt-fontes').textContent=bad?bad+' with errors':'';

 $('#srcLead').textContent=`${ok} of ${F.length} sources responded in this run. If a source keeps failing, its address may have changed: edit the list at the top of the script.`;
 $('#tbl-src').innerHTML='<thead><tr><th scope="col">Source</th><th scope="col">Type</th><th scope="col">Status</th><th class="num" scope="col">Items</th><th class="num" scope="col">Time</th><th scope="col">Detail</th></tr></thead><tbody>'+F.map(f=>{const x=stOf(f.estado||'');return`<tr><td>${/^https?:/.test(f.url||'')?`<a class="src-link" href="${esc(f.url)}" target="_blank" rel="noopener noreferrer">${esc(f.nome)}${ic('external')}<span class="sr"> (opens in a new window)</span></a>`:esc(f.nome)}</td><td><span class="tag">${esc(f.tipo)}</span></td><td><span class="st st-${x[0]}">${ic(x[1])}${esc(f.estado)}</span></td><td class="num">${f.itens||0}</td><td class="num">${nf((f.ms||0)/1000,1)} s</td><td class="detail">${esc(f.erro||'')}</td></tr>`;}).join('')+'</tbody>';
 const M=arr(D.manutencao);$('#manut').innerHTML=M.length?`<div class="callout warn">${ic('refresh')}<div><b>Maintenance:</b> ${M.map(esc).join(' ')}</div></div>`:'';}

/* ---------- horário das bolsas, em hora de Lisboa ---------- */
function tzParts(t,tz){const o={},f=TZF[tz]||(TZF[tz]=new Intl.DateTimeFormat('en-GB',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}));f.formatToParts(new Date(t)).forEach(p=>o[p.type]=p.value);return o;}
function tzOffset(t,tz){const p=tzParts(t,tz);return Date.UTC(+p.year,+p.month-1,+p.day,+p.hour,+p.minute,+p.second)-Math.floor(t/1000)*1000;}
/* hora "de parede" numa bolsa (ex.: 09:30 em Nova Iorque) convertida para um instante real */
function zoned(y,m,d,hhmm,tz){const q=hhmm.split(':').map(Number),g=Date.UTC(y,m-1,d,q[0],q[1]);let u=g-tzOffset(g,tz);u=g-tzOffset(u,tz);return u;}
/* Feriados e sessões curtas calculados pelas regras de cada bolsa, para qualquer ano (assim não deixam de funcionar
   quando acaba a lista do script); a lista do script junta-se a estes e serve para exceções (ex.: luto nacional). */
function pascoa(y){const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),n=h+l-7*m+114;return Date.UTC(y,Math.floor(n/31)-1,n%31+1);}
const isoU=t=>new Date(t).toISOString().slice(0,10),wdU=t=>new Date(t).getUTCDay();
/* n-ésimo dia da semana (0 = domingo) do mês m; n = -1 é o último */
function nesimo(y,m,wd,n){if(n>0){const f=wdU(Date.UTC(y,m-1,1));return Date.UTC(y,m-1,1+(wd-f+7)%7+(n-1)*7);}const u=Date.UTC(y,m,0);return u-((wdU(u)-wd+7)%7)*DAY;}
/* feriado ao sábado passa para sexta e ao domingo para segunda (regra da NYSE) */
const observado=t=>wdU(t)===6?t-DAY:wdU(t)===0?t+DAY:t;
const REGRAS={};
function regras(id,y){const k=id+y;if(REGRAS[k])return REGRAS[k];const P=pascoa(y),fer=[],cur={};
 if(id==='US'){const ano=Date.UTC(y,0,1);if(wdU(ano)!==6)fer.push(observado(ano)); /* Ano Novo ao sábado não fecha a sexta anterior */
  fer.push(nesimo(y,1,1,3),nesimo(y,2,1,3),P-2*DAY,nesimo(y,5,1,-1),observado(Date.UTC(y,5,19)),observado(Date.UTC(y,6,4)),nesimo(y,9,1,1),nesimo(y,11,4,4),observado(Date.UTC(y,11,25)));
  cur[isoU(nesimo(y,11,4,4)+DAY)]='13:00';                                   /* dia a seguir ao Thanksgiving */
  const j4=wdU(Date.UTC(y,6,4));if(j4>=2&&j4<=5)cur[isoU(Date.UTC(y,6,3))]='13:00'; /* véspera do 4 de julho */
  const n24=wdU(Date.UTC(y,11,24));if(n24>=1&&n24<=4)cur[isoU(Date.UTC(y,11,24))]='13:00';} /* véspera de Natal */
 else if(id==='DE'){fer.push(Date.UTC(y,0,1),P-2*DAY,P+DAY,Date.UTC(y,4,1),Date.UTC(y,11,24),Date.UTC(y,11,25),Date.UTC(y,11,26),Date.UTC(y,11,31));
  /* último dia de negociação do ano (antes de 31/12): a Deutsche Börse costuma encurtá-lo até às 14:00 e confirma-o por
     circular; se num ano for diferente, acrescenta a exceção em "curtos" no script */
  let u=Date.UTC(y,11,30);while(wdU(u)===0||wdU(u)===6||[24,25,26].includes(new Date(u).getUTCDate()))u-=DAY;cur[isoU(u)]='14:00';}
 return REGRAS[k]={fer:new Set(fer.map(isoU)),cur};}
const ehFeriado=(b,iso,y)=>arr(b.feriados).includes(iso)||regras(b.id,y).fer.has(iso);
const fechoDe=(b,iso,y)=>(b.curtos&&b.curtos[iso])||regras(b.id,y).cur[iso]||b.fecha;
/* sessão atual ou próxima: salta fins de semana e feriados; usa a hora de fecho mais cedo nos dias de sessão curta */
function sessao(b,now){const p=tzParts(now,b.tz);for(let i=0;i<15;i++){const dt=new Date(Date.UTC(+p.year,+p.month-1,+p.day+i)),y=dt.getUTCFullYear(),m=dt.getUTCMonth()+1,d=dt.getUTCDate(),wd=dt.getUTCDay(),iso=y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');
 if(wd===0||wd===6||ehFeriado(b,iso,y))continue;const cl=fechoDe(b,iso,y),a=zoned(y,m,d,b.abre,b.tz),f=zoned(y,m,d,cl,b.tz);if(now<f)return{a,f,aberta:now>=a,curta:cl!==b.fecha};}return null;}
const hmL=t=>{const p=tzParts(t,LIS);return p.hour+':'+p.minute;};
const diaL=t=>{const p=tzParts(t,LIS),d=new Date(Date.UTC(+p.year,+p.month-1,+p.day));return WD[d.getUTCDay()]+', '+(+p.day)+' '+MES[+p.month-1];};
const mesmoDiaL=(a,b)=>{const x=tzParts(a,LIS),y=tzParts(b,LIS);return x.year===y.year&&x.month===y.month&&x.day===y.day;};
const falta=ms=>{const m=Math.max(1,Math.round(ms/6e4));if(m<60)return m+' min';const h=Math.floor(m/60);if(h<24)return h+' h'+(m%60?' '+(m%60)+' min':'');const d=Math.floor(h/24);return d+' d'+(h%24?' '+(h%24)+' h':'');};
function renderBolsas(){const box=$('#mkts');if(!box)return;const now=Date.now();
 box.innerHTML=arr(D.bolsas).map(b=>{const s=sessao(b,now);if(!s)return'';const dia=mesmoDiaL(s.a,now)?'Today':diaL(s.a),est=s.aberta?'Open':'Closed',quando=s.aberta?'closes in '+falta(s.f-now):'opens in '+falta(s.a-now);
  return`<div class="mkt${s.aberta?' open':''}" title="${esc(b.nome)} (${esc(b.detalhe)}). ${est}: ${quando}. ${dia}, ${hmL(s.a)} to ${hmL(s.f)}, Lisbon time${s.curta?' (shortened session)':''}."><span class="mkt-top"><i class="mkt-dot" aria-hidden="true"></i><b>${esc(b.curto)}</b><span class="mkt-st">${est}</span></span><span class="mkt-when">${dia} · ${hmL(s.a)}–${hmL(s.f)}<small>Lisbon</small></span><span class="sr"> ${quando}</span></div>`;}).join('');}

/* ---------- navigation & filters ---------- */
const TABS=$$('.rail a').map(a=>a.dataset.tab);
/* os gráficos medem o contentor: voltam a ser desenhados quando o separador fica visível ou a janela muda de largura */
function drawTemas(f){const cnt={};f.forEach(n=>n.temas.forEach(t=>cnt[t]=(cnt[t]||0)+1));hbars($('#ch-temas'),Object.entries(cnt).sort((a,b)=>b[1]-a[1]).map(([l,v])=>({l,v,c:'var(--bar)'})));}
function drawNiveis(){const cats=['AAPL','NVDA','GOOGL','BTC','MKT'].concat(ETF_IDS.filter(x=>x!=='SXR8'&&NEWS.some(n=>n.nivel!=='white'&&n.empresas.includes(x))));   /* os outros ETF só quando há notícias */barChart($('#ch-niveis'),{cats:cats.map(c=>CO[c].n),series:['red','orange','yellow'].map(l=>({n:LV[l].n,c:LV[l].c,values:cats.map(c=>NEWS.filter(n=>n.nivel===l&&matchCo(n.empresas,c)).length)})),dec:0,label:'News by asset and level'});}
/* só os gráficos são redesenhados (a lista de notícias mantém o estado: itens abertos, 'Mostrar mais') */
const CHARTS={overview:()=>drawNiveis(),portfolio:()=>drawPf(),news:()=>drawTemas(filteredNews()),prices:()=>renderPrices(),fundamentals:()=>renderFund(),fx:()=>renderFx(),etf:()=>drawEtfMore(),bitcoin:()=>renderBtc()};
let curTab='overview',rsz=0,lastW=0;
function showTab(id,scroll){if(!TABS.includes(id))id='overview';$$('.panel').forEach(p=>p.hidden=p.dataset.panel!==id);curTab=id;if(CHARTS[id])CHARTS[id]();$$('.rail a').forEach(a=>a.dataset.tab===id?a.setAttribute('aria-current','page'):a.removeAttribute('aria-current'));try{history.replaceState(null,'','#'+id);}catch(e){}
 const rail=$('#rail'),cur=$('.rail a[aria-current]');if(cur&&rail.scrollWidth>rail.clientWidth)rail.scrollLeft=cur.offsetLeft-(rail.clientWidth-cur.offsetWidth)/2;
 if(scroll){const y=$('.layout').getBoundingClientRect().top+scrollY-$('#topbar').offsetHeight-8;if(scrollY>y)scrollTo(0,y);$('#main').focus({preventScroll:true});}}
function renderAll(){$$('.chip[data-co]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.co===st.co)));$$('.tk').forEach(t=>t.classList.toggle('dim',st.co!=='all'&&!(t.dataset.tk===st.co||(st.co==='MKT'&&(t.dataset.tk==='SXR8'||t.dataset.tk==='FX')))));
 renderAlerts();renderNews();renderPrices();renderFx();renderCal();renderEtf();drawEtfMore();renderBtc();drawPf();renderSrc();renderStats();if(curTab==='fundamentals')renderFund();}
function setTopbarH(){document.documentElement.style.setProperty('--topbar-h',$('#topbar').offsetHeight+'px');}
function init(){guilloche();renderHero();
 const temas=[...new Set(NEWS.flatMap(n=>n.temas))].sort();$('#f-tema').innerHTML='<option value="">All</option>'+temas.map(t=>`<option>${esc(t)}</option>`).join('');
 if(D.primeiraExecucao)$('#f-novas').disabled=true;
 if(Date.now()-GEN>36*36e5)$('#stale').innerHTML=`<div class="callout warn">${ic('refresh')}<div>This data was generated ${Math.round((Date.now()-GEN)/DAY)} days ago. Run the script again to update it.</div></div>`;
 $('#foot').innerHTML=`<div class="foot-brand"><span class="brand-mark" aria-hidden="true"><svg class="logo" viewBox="0 0 32 32"><use href="#i-logo"/></svg></span>Bluechip Board</div>`+
  `<div class="foot-meta"><span>${ic('clock')}Updated ${fdt(GEN)}, ${hmL(GEN)} Lisbon time</span><span><button type="button" data-goto="sources">Sources &amp; method</button></span></div>`+
  '';
 $$('.rail a').forEach(a=>a.addEventListener('click',e=>{e.preventDefault();showTab(a.dataset.tab,true);}));
 document.addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-goto]');if(b)showTab(b.dataset.goto,true);});
 /* um botão de filtro por ETF além do SXR8 (que fica em "ETF & market"), antes da Bitcoin; nome curto da configuração */
 const btcChip=$('.chip[data-co="BTC"]');ETF_IDS.filter(x=>x!=='SXR8'&&!$(`.chip[data-co="${x}"]`)).forEach(x=>{const b=document.createElement('button');b.className='chip';b.type='button';b.dataset.co=x;b.setAttribute('aria-pressed','false');b.style.setProperty('--c',CO[x].c);b.title=etfInfo(x).fundo||etfNomeCurto(x);
  b.innerHTML=`<span class="ci">${ic(CO[x].i)}</span>${esc(etfInfo(x).chip||etfIdx(x))}<span class="tkr">${esc(x)}</span>`;btcChip.parentNode.insertBefore(b,btcChip);});
 $$('.chip[data-co]').forEach(b=>b.addEventListener('click',()=>{st.co=b.dataset.co;st.lim=S0.lim;if(ehEtf(st.co))etfSel=st.co;else if(st.co==='MKT')etfSel='SXR8';renderAll();}));
 $('#f-etf').addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-etf]');if(b)escolheEtf(b.dataset.etf);});
 $$('#f-lvl button').forEach(b=>b.addEventListener('click',()=>{st.lvl=b.dataset.lvl;st.lim=S0.lim;$$('#f-lvl button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));renderNews();}));
 ['f-tema','f-tier','f-dias','f-ord','f-q','f-novas'].forEach(id=>$('#'+id).addEventListener('input',()=>{st.lim=S0.lim;renderNews();}));
 $('#moreBtn').addEventListener('click',()=>{st.lim+=60;renderNews();});
 /* o período é partilhado pelas Cotações, pelo Câmbio e pela Bitcoin: os botões das duas secções ficam sincronizados */
 $$('[data-per]').forEach(b=>b.addEventListener('click',()=>{st.per=b.dataset.per;$$('[data-per]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.per===st.per)));renderPrices();renderFx();renderBtc();}));
 $('#btcAllNews').addEventListener('click',()=>{st.co='BTC';st.lim=S0.lim;renderAll();showTab('news',true);});
 $$('#f-cur button').forEach(b=>b.addEventListener('click',()=>{st.cur=b.dataset.cur;$$('#f-cur button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));renderPrices();}));
 ['c-imp','c-past'].forEach(id=>$('#'+id).addEventListener('input',renderCal));
 $$('#f-elp button').forEach(b=>b.addEventListener('click',()=>{etfLp=b.dataset.lp;$$('#f-elp button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));drawEtfMore();}));
 $$('#f-elog button').forEach(b=>b.addEventListener('click',()=>{etfLog=b.dataset.log==='1';$$('#f-elog button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));drawEtfMore();}));
 ['etf','btc'].forEach(tab=>{const d=$('#roll-'+tab);if(d)d.addEventListener('toggle',()=>{if(d.open)drawRoll(tab);});});
 buildSim('etf',{name:'the SXR8 ETF',pts:()=>HIST.SXR8,c:'var(--spx)',dec:2,units:u=>nf(u,2)+' ETF units',note:'SXR8 is an accumulating ETF: dividends are reinvested in its price, so they are already included. Fees, spreads and taxes are not. Past returns do not guarantee future ones.'});
 /* o mesmo simulador para cada um dos outros ETF, com o preço do próprio ETF (guardado em bb.sim.etf-<id>) */
 ETF_IDS.filter(x=>x!=='SXR8').forEach(x=>{const d=document.createElement('div');d.className='fig';d.id='sim-'+simKey(x);d.hidden=true;$('#sim-etf-outros').appendChild(d);
  buildSim(simKey(x),{name:`the ${x} ETF`,pts:()=>HIST[x]||[],c:CO[x].c,dec:2,units:u=>nf(u,2)+' ETF units',note:`${x} is an accumulating ETF: dividends are reinvested in its price, so they are already included. Fees, spreads and taxes are not. Past returns do not guarantee future ones.`});});
 buildSim('btc',{name:'Bitcoin',pts:()=>HIST.BTC,c:'var(--btc)',dec:0,units:u=>btcF(+u.toFixed(8))+' BTC',note:'Fees, spreads and taxes are not included. Past returns do not guarantee future ones.'});
 migraCompras();buildLots();buildPf();buildTgt();buildPol();buildBuys();buildSell();buildTax();buildBackup();

 addEventListener('hashchange',()=>showTab(location.hash.slice(1),false));
 addEventListener('resize',setTopbarH);setTopbarH();
 const redraw=()=>{const w=$('#main').clientWidth;if(w&&Math.abs(w-lastW)>2){lastW=w;if(CHARTS[curTab])CHARTS[curTab]();}};
 if(window.ResizeObserver)new ResizeObserver(()=>{clearTimeout(rsz);rsz=setTimeout(redraw,120);}).observe($('#main'));else addEventListener('resize',()=>{clearTimeout(rsz);rsz=setTimeout(redraw,120);});renderAll();renderBolsas();setInterval(renderBolsas,30000);setTopbarH();showTab((location.hash||'#overview').slice(1),false);}
init();
/* gancho para os testes automáticos (Tests\Test-Site.ps1): só existe se a página de teste o definir antes */
if(typeof window.__BB_TEST__==='function')window.__BB_TEST__({stats,inCur,fxAt,fx:()=>({FX,FXSRC}),frescura,ajusteSplit,efetiva,carteira,pfDados,juntaBackup,limpaBackup,precoEurEm,store,alerts,SPLITS,diasAte,regras,sessao,lastEur,ATIVOS,HIST,apaga,migraCompras,estadoBk,verificaPasta,quedas,fxPt,divDados,dividendos,anexoJ,exportaAnexoJ,csvTxt,csvNum,ANEXO_J,fxRef,ETF_IDS,ehEtf,etfInfo,escolheEtf,etfSel:()=>etfSel,PF,DIV_IDS,BOLSA_DE,CO,simKey,simular,rentab,xirr,isoU,pct,stress,EPISODIOS,estrategias,eur,rolar,underwater,nf,exposicao,CONC_LIMIAR,simulaVenda,feeDe,alocacao,reparte,mesesBanda,dadosBackup,limpaAlvos,anexo8A,divPagamentos,expoNoticias,relevancia,NEWS,sessaoDe,grandesMovimentos,reacoes,eventosGrafico,HN,fundDados,valorEm,ttmEm,fundHistorico,percentil,renderFund,macroDados,renderMacro,c2,pendente,guardaFicheiro,auto:()=>AUTO,setFH:h=>{FH=h;PEDIU=false;}});
})();
</script>
</body>
</html>
'@
}

# ============================================================================
# 4. EXECUÇÃO
# ============================================================================

# Sem -EmailSEC, o e-mail da configuração local (nunca nos argumentos da tarefa agendada nem no git)
$emailLocal = Get-EmailSecLocal $PSScriptRoot
if (-not $EmailSEC) {
    $EmailSEC = $emailLocal.email
    if ($emailLocal.aviso) { Write-Warning $emailLocal.aviso }
    elseif (-not $EmailSEC -and -not $AgendarDiariamente) { Write-Host '  ! No SEC e-mail (-SecEmail, bluechip-board.config.json or BLUECHIP_SEC_EMAIL): the SEC sources are skipped.' -ForegroundColor DarkYellow }
}

if ($AgendarDiariamente) {
    $exe = (Get-Process -Id $PID).Path
    # In a Windows command line, backslashes right before a closing quote must be doubled ("C:\Data\" would swallow the quote)
    $aspas = { param($s) '"' + ($s -replace '(\\+)$', '$1$1') + '"' }
    $argumentos = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $(& $aspas $PSCommandPath) -NaoAbrir -Dias $Dias -Pasta $(& $aspas $Pasta)"
    # The SEC e-mail is NOT put in the task's arguments (anyone who can list the tasks or the processes would see it):
    # each scheduled run reads it from bluechip-board.config.json (or BLUECHIP_SEC_EMAIL), like the launcher does.
    if ($EmailSEC -and $EmailSEC -ne $emailLocal.email) { Write-Warning "The scheduled task does not store the e-mail given with -SecEmail: each run reads it from bluechip-board.config.json next to the script (or BLUECHIP_SEC_EMAIL). Put it there, or the scheduled runs will skip the SEC sources." }
    $acao = New-ScheduledTaskAction -Execute $exe -Argument $argumentos
    $gatilho = New-ScheduledTaskTrigger -Daily -At ([datetime]::ParseExact($Hora, 'HH:mm', $Script:Inv))
    # Windows' defaults would skip the run on battery power (and stop it when unplugged): this PC is a laptop
    $definicoes = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
    Register-ScheduledTask -TaskName 'BluechipBoard' -Action $acao -Trigger $gatilho -Settings $definicoes -Description 'Bluechip Board: news and prices for AAPL, NVDA, GOOGL, the iShares ETFs (SXR8, EUNK, IS3N, EUNN) and Bitcoin' -Force | Out-Null
    Write-Host "Task 'BluechipBoard' scheduled every day at $Hora." -ForegroundColor Green
    Write-Host "The latest website will be at: $(Join-Path $Pasta 'bluechip-board.html')" -ForegroundColor Green
    return
}

# Impede duas execuções ao mesmo tempo (ex.: atalho e tarefa agendada), que escreveriam nos mesmos ficheiros
$Script:Trinco = New-Object System.Threading.Mutex($false, 'Local\BluechipBoard')
try { $livre = $Script:Trinco.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $livre = $true }
if (-not $livre) {
    Write-Warning 'Bluechip Board is already running in another window. Wait for it to finish and try again.'
    return
}

New-Item -ItemType Directory -Force -Path $Pasta | Out-Null
$inicio = Get-Date
Write-Host ''
Write-Host 'Bluechip Board: Apple, NVIDIA, Alphabet, iShares ETFs (SXR8, EUNK, IS3N, EUNN) and Bitcoin' -ForegroundColor White
Write-Host "News from the last $Dias days. Output in $Pasta" -ForegroundColor DarkGray

Write-Passo 'Reading news feeds'
$brutas = New-Object System.Collections.Generic.List[object]
foreach ($f in $Feeds) {
    Write-Host "      $($f.Nome)" -ForegroundColor DarkGray
    foreach ($it in (Read-Feed -Feed $f)) { $brutas.Add($it) }
}
if ($EmailSEC) {
    Write-Passo 'Reading SEC filings'
    foreach ($s in $SecEmpresas) {
        $f = @{ Nome = "SEC EDGAR: $($s.Id)"; Url = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=$($s.Cik)&type=&dateb=&owner=include&count=40&output=atom"; Dica = $s.Id }
        foreach ($it in (Read-Feed -Feed $f -UserAgent "BluechipBoard/1.0 $EmailSEC" -Tipo 'sec')) { $brutas.Add($it) }
        Start-Sleep -Milliseconds 400
    }
} else {
    Add-Fonte 'SEC EDGAR (8-K, 10-Q, share issuance)' 'sec' 'https://www.sec.gov/edgar/search/' 'skipped' 0 0 'Run with -EmailSEC "your@email.com" to enable it: the SEC requires a contact in each request.'
}

Write-Passo "Classifying $($brutas.Count) stories"
$vistosPath = Join-Path $Pasta 'vistos.json'
$vistos = @{}
if (Test-Path $vistosPath) {
    try { $obj = Get-Content $vistosPath -Raw -Encoding UTF8 | ConvertFrom-Json; foreach ($p in $obj.PSObject.Properties) { $vistos[$p.Name] = $p.Value } } catch { $vistos = @{} }
}
$primeiraVez = ($vistos.Count -eq 0)
$porChave = @{}
foreach ($b in $brutas) {
    $m = Measure-Noticia $b
    if ($null -eq $m -or -not $m.titulo -or -not $m.chave) { continue }
    if ($porChave.ContainsKey($m.chave)) {
        $ex = $porChave[$m.chave]
        $junta = @(@($m.empresas) + @($ex.empresas) | Select-Object -Unique)
        if ($m.score -gt $ex.score) { $m.empresas = $junta; $porChave[$m.chave] = $m } else { $ex.empresas = $junta }
    } else { $porChave[$m.chave] = $m }
}
$unicas = @($porChave.Values)
Write-Passo 'Grouping duplicate stories'
$agrupadas = @(Join-NoticiasDuplicadas $unicas)
$duplicadas = $unicas.Count - $agrupadas.Count
$agoraIso = $Script:Agora.ToString('o')
foreach ($m in $agrupadas) {
    # um grupo só é "novo" se nenhuma das suas notícias tiver aparecido numa execução anterior
    $chaves = @(@($m.chave) + @($m.outras | ForEach-Object { $_.chave }))
    if (-not @($chaves | Where-Object { $vistos.ContainsKey($_) }).Count) { $m.novo = -not $primeiraVez }
    foreach ($c in $chaves) { if (-not $vistos.ContainsKey($c)) { $vistos[$c] = $agoraIso } }
}
$corte = $Script:Agora.AddDays(-60)
$vistosNovos = @{}
foreach ($k in @($vistos.Keys)) {
    $v = $vistos[$k]
    $d = if ($v -is [datetime]) { [DateTimeOffset]$v } else { ConvertTo-Data "$v" }
    if (-not $d -or $d -ge $corte) { $vistosNovos[$k] = $(if ($v -is [datetime]) { ([DateTimeOffset]$v).ToString('o') } else { "$v" }) }
}
# (vistos.json is written only at the end, after the website: a run that fails halfway must not mark stories as seen)
$noticias = @($agrupadas | Sort-Object -Property @{ Expression = 'score'; Descending = $true }, @{ Expression = 'data'; Descending = $true }, @{ Expression = 'chave'; Descending = $false })
# Histórico de notícias (material e important): lido aqui, escrito só no fim, depois do vistos.json
$histPath = Join-Path $Pasta 'noticias-historico.json'
$histLido = Read-HistoricoNoticias $histPath
if ($histLido.estado -eq 'corrupted') { Write-Warning $histLido.aviso } elseif ($histLido.aviso) { Write-Host "  ! $($histLido.aviso)" -ForegroundColor DarkYellow }
$histNoticias = Merge-HistoricoNoticias $histLido $noticias $Script:Agora

# Data from the previous run (bluechip-board-data.json). Used only when a price source fails, and always labelled
# as such: the site shows how old those prices are, so they never pass for current ones.
$anterior = $null
$anteriorPath = Join-Path $Pasta 'bluechip-board-data.json'
if (Test-Path -LiteralPath $anteriorPath) { try { $anterior = Get-Content -LiteralPath $anteriorPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $anterior = $null } }
function Get-PontosGuardados($Lista) {
    # accepts [date, value] pairs and the {value: [date, value], Count: 2} form that Windows PowerShell 5.1 sometimes writes
    @(foreach ($p in @($Lista)) {
        $q = if ($p -is [array]) { $p } elseif ($null -ne $p -and $p.PSObject.Properties['value']) { @($p.value) } else { $null }
        $v = 0.0
        if ($q -and $q.Count -ge 2 -and "$($q[0])" -match '^\d{4}-\d{2}-\d{2}$' -and [double]::TryParse("$($q[1])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$v) -and $v -gt 0) { , @("$($q[0])", $v) }
    })
}
function Get-SerieAnterior($Pontos, [string]$NomeFonte, $Base) {
    if (-not $anterior) { return $null }
    $pts = @(Get-PontosGuardados $Pontos)
    if ($pts.Count -lt 5) { return $null }
    # the source entry already says "error": add that the previous run's data is shown instead
    $f = @($Script:Fontes | Where-Object { $_.nome -eq $NomeFonte }) | Select-Object -Last 1
    if ($f) { $f.estado = "error (showing previous run's data, up to $($pts[-1][0]))" }
    return [pscustomobject]@{ moeda = $(if ($Base) { "$($Base.moeda)" } else { '' }); pontos = $pts; fonte = "previous run (data up to $($pts[-1][0]))"; splits = @(); parcial = $false; hora = $null }
}
function New-DadosSerie($Ativo, $Serie, [string]$MoedaPadrao) {
    [pscustomobject]@{
        id = $Ativo.Id; nome = $Ativo.Nome; simbolo = $Ativo.Yahoo
        moeda = $(if ($Serie -and $Serie.moeda) { $Serie.moeda } else { $MoedaPadrao })
        pontos = $(if ($Serie) { @($Serie.pontos) } else { @() }); fonte = $(if ($Serie) { $Serie.fonte } else { '' })
        parcial = [bool]($Serie -and $Serie.parcial); hora = $(if ($Serie) { $Serie.hora } else { $null })
        # tipo e bolsa (só para os ETF, de $ETFs): o site trata-os pelo mesmo mecanismo dos outros ativos
        tipo = $(if (@($ETFs | Where-Object { $_.Id -eq $Ativo.Id }).Count) { 'ETF' } else { '' })
        bolsa = "$(@($ETFs | Where-Object { $_.Id -eq $Ativo.Id } | ForEach-Object { $_.Bolsa })[0])"
    }
}

Write-Passo 'Getting 1-year prices'
$ativosDados = @(foreach ($a in $Ativos) {
    $ref = if ($anterior) { @($anterior.ativos | Where-Object { $_.id -eq $a.Id })[0] } else { $null }
    $s = Get-Serie $a -Referencia $(if ($ref) { @(Get-PontosGuardados $ref.pontos) } else { $null })
    if (-not $s -and $anterior) { $s = Get-SerieAnterior (@($anterior.ativos | Where-Object { $_.id -eq $a.Id })[0].pontos) "Prices: $($a.Nome)" (@($anterior.ativos | Where-Object { $_.id -eq $a.Id })[0]) }
    New-DadosSerie $a $s $a.Moeda
})
$mercadoDados = @(foreach ($a in $Mercado) {
    $s = Get-Serie $a
    if (-not $s -and $anterior) { $s = Get-SerieAnterior (@($anterior.mercado | Where-Object { $_.id -eq $a.Id })[0].pontos) "Prices: $($a.Nome)" (@($anterior.mercado | Where-Object { $_.id -eq $a.Id })[0]) }
    New-DadosSerie $a $s ''
})
Write-Passo 'Getting EUR/USD (Yahoo and ECB)'
$fxSerie = Get-Serie @{ Id = 'EURUSD'; Nome = 'EUR/USD'; Yahoo = 'EURUSD=X'; Stooq = 'eurusd'; Moeda = 'USD' }
if (-not $fxSerie -and $anterior -and $anterior.fx) { $fxSerie = Get-SerieAnterior $anterior.fx.yahoo 'Prices: EUR/USD' $null }
$fxBce = Get-TaxaBCE
Write-Passo 'Getting the ETF holdings from iShares'
# um ficheiro por fundo; o do SXR8 continua também em "etf" (o formato anterior dos dados)
$pesosEtfs = [ordered]@{}
foreach ($e in $ETFs) { $pesosEtfs[$e.Id] = Get-PesosETF $e }
$pesos = $pesosEtfs['SXR8']
Write-Passo 'Getting dividends (Yahoo)'
$dividendos = Get-Dividendos
# Se a fonte falhar, usa os dividendos da execução anterior (até 180 dias), sempre marcados como tal e com a data em que
# foram obtidos: nunca passam por atuais. Sem eles, o site mostra "Unavailable".
foreach ($id in @($dividendos.Keys)) {
    if ($dividendos[$id].estado -eq 'ok' -or -not $anterior -or -not $anterior.dividendos) { continue }
    $v = $anterior.dividendos.$id
    if (-not $v -or "$($v.estado)" -notin 'ok', 'previous run') { continue }
    $quando = if ($v.obtidoEm -is [datetime]) { [DateTimeOffset]$v.obtidoEm } else { ConvertTo-Data "$($v.obtidoEm)" }
    $anualV = 0.0; $temAnual = $null -ne $v.anualPorAcao -and [double]::TryParse("$($v.anualPorAcao)", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$anualV) -and $anualV -ge 0
    if (-not $quando -or ($Script:Agora - $quando).TotalDays -gt 180 -or -not $temAnual) { continue }
    $pag = @(Get-PontosGuardados $v.pagamentos)
    $dataPreco = if ($v.precoData -is [datetime]) { ([DateTimeOffset]$v.precoData).ToString('o') } else { "$($v.precoData)" }
    $dividendos[$id] = [pscustomobject]@{
        id = $id; simbolo = "$($v.simbolo)"; moeda = 'USD'; estado = 'previous run'; fonte = "previous run (retrieved $($quando.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv)))"; obtidoEm = $quando.ToString('o')
        anualPorAcao = $anualV; ttmPorAcao = $(if ($null -ne $v.ttmPorAcao) { [double]$v.ttmPorAcao } else { $null }); frequencia = $v.frequencia
        rendimentoPct = $(if ($null -ne $v.rendimentoPct) { [double]$v.rendimentoPct } else { $null })
        ultimo = $(if ($pag.Count) { @($pag[-1][0], $pag[-1][1]) } else { $null }); pagamentos = $pag
        preco = $v.preco; precoData = $dataPreco; nota = "$($v.nota)"; erro = $dividendos[$id].erro
    }
    $f = @($Script:Fontes | Where-Object { $_.nome -like "Dividends: *" -and $_.url -match [regex]::Escape("/chart/$($v.simbolo)?") }) | Select-Object -Last 1
    if ($f) { $f.estado = "error (showing previous run's data, retrieved $($quando.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv)))" }
}
Write-Passo 'Getting Bitcoin indicators (sentiment, dominance, network)'
$bitcoin = Get-DadosBitcoin
# O halving entra no calendário com a data prevista pela rede; sem a rede, uma estimativa a partir do último halving conhecido
# (um evento com a data mal escrita é ignorado com um aviso, em vez de parar a execução inteira)
$eventos = @(foreach ($c in $Calendario) {
    $dt = [datetime]::MinValue
    if (-not [datetime]::TryParseExact("$($c.d)", 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$dt)) { Write-Warning "Calendar event ignored, the date must be yyyy-MM-dd: '$($c.d)' ($($c.ev))"; continue }
    [pscustomobject]@{ d = $c.d; e = $c.e; ev = $c.ev; imp = $c.imp; st = $c.st }
})
if ($bitcoin.rede) { $eventos += [pscustomobject]$bitcoin.rede.evento }
else { $eventos += Get-HalvingAproximado }
Write-Passo 'Getting earnings dates (Nasdaq)'
foreach ($r in @(Get-DatasResultados)) {
    $dAuto = [datetime]::ParseExact($r.d, 'yyyy-MM-dd', $Script:Inv)
    if ($dAuto -lt $Script:Agora.UtcDateTime.Date.AddDays(-1)) { continue }
    # A data manual de resultados mais próxima (até 45 dias) passa a usar a data da Nasdaq; o texto do evento mantém-se.
    # Exceção: uma data manual confirmada (C) não é trocada por uma estimativa (E) da Nasdaq.
    $manual = $eventos | Where-Object { $_.e -eq $r.e -and $_.ev -match 'earnings' -and [math]::Abs(([datetime]::ParseExact($_.d, 'yyyy-MM-dd', $Script:Inv) - $dAuto).TotalDays) -le 45 } | Select-Object -First 1
    if ($manual) {
        if ($manual.st -eq 'C' -and $r.st -ne 'C') { continue }
        $manual.d = $r.d; $manual.st = $r.st; $manual | Add-Member -NotePropertyName auto -NotePropertyValue 'Nasdaq' -Force
    }
    else { $eventos += [pscustomobject]@{ d = $r.d; e = $r.e; ev = 'Quarterly earnings'; imp = 'High'; st = $r.st; auto = 'Nasdaq' } }
}
Write-Passo 'Getting long-term histories (stocks, ETF and EUR/USD since 2010; Bitcoin since 2014)'
# Servem os gráficos de longo prazo, os simuladores e o preço do dia de cada compra registada na carteira.
# Os desdobramentos (splits) vêm com o histórico e servem para manter as compras antigas no número certo de ações.
$historico = [ordered]@{}
$desdobramentos = [ordered]@{}
foreach ($a in @($Ativos) + @(@{ Id = 'FX'; Nome = 'EUR/USD'; Yahoo = 'EURUSD=X'; Stooq = ''; Moeda = 'USD' })) {
    $s = Get-Serie $a -Desde 1262304000   # 1 jan 2010
    if (-not $s -and $anterior -and $anterior.historico) {
        $s = Get-SerieAnterior $anterior.historico.($a.Id) "Price history: $($a.Nome)" $null
        if ($s -and $anterior.splits -and $anterior.splits.($a.Id)) { $s.splits = @(foreach ($x in @($anterior.splits.($a.Id))) { $q = @($x); if ($q.Count -ge 2) { , @("$($q[0])", [double]$q[1]) } }) }
    }
    $historico[$a.Id] = $(if ($s) { @($s.pontos) } else { @() })
    if ($a.Id -ne 'FX') { $desdobramentos[$a.Id] = $(if ($s) { @($s.splits) } else { @() }) }
}

# Datas passadas de resultados (SEC): só com -EmailSEC, com o mesmo User-Agent e 400 ms entre pedidos. As sessões de
# reação usam as datas do histórico longo de cada empresa. Se falhar (ou sem e-mail), fica a execução anterior (até 180
# dias), sempre marcada como tal; sem ela, "Unavailable".
$resultadosSec = [ordered]@{}
$feriadosUS = @(@($Bolsas | Where-Object { $_.id -eq 'US' })[0].feriados)
$secAnt = if ($anterior -and $anterior.resultadosSec) { $anterior.resultadosSec } else { $null }
if ($EmailSEC) {
    Write-Passo 'Getting past earnings dates (SEC)'
    foreach ($s in $SecEmpresas) {
        $sess = [string[]]@(foreach ($p in @($historico[$s.Id])) { $q = @($p); if ($q.Count -ge 2) { "$($q[0])" } })
        $conhecidas = @{}
        if ($secAnt -and $secAnt.($s.Id)) { foreach ($r in @($secAnt.($s.Id).resultados)) { if ($r -and $r.acc -and $r.aceite) { $conhecidas["$($r.acc)"] = (ConvertTo-IsoUtc $r.aceite) } } }
        $resultadosSec[$s.Id] = Get-ResultadosSEC $s "BluechipBoard/1.0 $EmailSEC" $sess $feriadosUS $conhecidas
        Start-Sleep -Milliseconds 400
    }
}
foreach ($s in $SecEmpresas) {
    $x = $resultadosSec[$s.Id]
    if ($x -and $x.estado -eq 'ok') { continue }
    $motivo = if ($x) { $x.erro } else { 'Run with -EmailSEC "your@email.com" to enable it: the SEC requires a contact in each request.' }
    $v = if ($secAnt) { $secAnt.($s.Id) } else { $null }
    $quando = if ($v) { ConvertTo-Data (ConvertTo-IsoUtc $v.obtidoEm) } else { $null }
    if ($v -and "$($v.estado)" -in 'ok', 'previous run' -and $quando -and ($Script:Agora - $quando).TotalDays -le 180) {
        $lista = @(foreach ($r in @($v.resultados)) { if ($r -and $r.acc) {
            [pscustomobject]@{ acc = "$($r.acc)"; entrega = "$($r.entrega)"; aceite = (ConvertTo-IsoUtc $r.aceite); horaNY = $(if ($r.horaNY) { "$($r.horaNY)" } else { $null }); quando = $(if ($r.quando) { "$($r.quando)" } else { $null }); sessao = $(if ($r.sessao) { "$($r.sessao)" } else { $null }); nota = "$($r.nota)" } } })
        $ult = if ($v.ultimoRelatorio) { [pscustomobject]@{ form = "$($v.ultimoRelatorio.form)"; data = "$($v.ultimoRelatorio.data)"; periodo = "$($v.ultimoRelatorio.periodo)"; acc = "$($v.ultimoRelatorio.acc)" } } else { $null }
        $resultadosSec[$s.Id] = [pscustomobject]@{ id = $s.Id; estado = 'previous run'; fonte = "previous run (retrieved $($quando.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv)))"; obtidoEm = $quando.ToString('o'); resultados = $lista; ultimoRelatorio = $ult; erro = $motivo; nota = '' }
    } else {
        $resultadosSec[$s.Id] = [pscustomobject]@{ id = $s.Id; estado = $(if ($x) { 'error' } else { 'skipped' }); fonte = ''; obtidoEm = $Script:Agora.ToString('o'); resultados = @(); ultimoRelatorio = $null; erro = $motivo; nota = '' }
    }
}
# Fundamentais (SEC XBRL companyfacts): só descarrega quando a SEC mostra um 10-Q/10-K mais recente do que o da última
# recolha; se falhar (ou sem e-mail), a execução anterior até 120 dias. EPS e ações ajustados aos splits ($desdobramentos).
$fundamentais = [ordered]@{}
$fundAnt = if ($anterior -and $anterior.fundamentais) { $anterior.fundamentais } else { $null }
if ($EmailSEC) { Write-Passo 'Getting fundamentals (SEC XBRL, only after a new 10-Q/10-K)' }
foreach ($s in $SecEmpresas) {
    $rs = $resultadosSec[$s.Id]
    $rel = if ($rs -and $rs.estado -eq 'ok') { $rs.ultimoRelatorio } else { $null }
    $nFontes = $Script:Fontes.Count
    $fundamentais[$s.Id] = Get-FundamentaisEmpresa $s $(if ($EmailSEC) { "BluechipBoard/1.0 $EmailSEC" } else { '' }) $desdobramentos[$s.Id] $(if ($fundAnt) { $fundAnt.($s.Id) } else { $null }) $rel
    if ($Script:Fontes.Count -gt $nFontes -and $Script:Fontes[-1].estado -ne 'ok (cached)') { Start-Sleep -Milliseconds 400 }
}


# Dados pessoais (carteira e compras), gravados pelo site em bluechip-board-backup.json na pasta do projeto.
# Se o browser não puder gravar em pastas, o ficheiro vai para as Transferências e é trazido para aqui, mas só depois de
# validado e só se for mais recente (pela data gravada dentro dele); o anterior fica guardado em .previous.json.
$bkNome = 'bluechip-board-backup.json'; $bkPath = Join-Path $Pasta $bkNome
function Read-Backup([string]$Caminho) {
    try { $b = Get-Content -LiteralPath $Caminho -Raw -Encoding UTF8 | ConvertFrom-Json; if ($b -and $b.app -eq 'Bluechip Board') { return $b } } catch { }
    return $null
}
function Get-InstanteBackup($B) {
    foreach ($campo in 'saved', 'exported') {
        $v = $B.$campo
        if ($v -is [datetime]) { return [DateTimeOffset]$v }
        $d = ConvertTo-Data "$v"; if ($d) { return $d }
    }
    return [DateTimeOffset]::MinValue
}
$transferencias = $null
try { $transferencias = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path } catch { }
if ($transferencias -and (Test-Path -LiteralPath $transferencias)) {
    $candidatos = @(Get-ChildItem -LiteralPath $transferencias -Filter 'bluechip-board-backup*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($candidatos.Count) {
        $atual = if (Test-Path -LiteralPath $bkPath) { Read-Backup $bkPath } else { $null }
        $tAtual = if ($atual) { Get-InstanteBackup $atual } else { [DateTimeOffset]::MinValue }
        # a backup "saved" in the future (a wrong clock) would win every comparison forever: it is never used to decide
        $futuro = $Script:Agora.AddDays(1)
        foreach ($c in $candidatos) {
            $b = Read-Backup $c.FullName
            if (-not $b) { Write-Warning "$($c.Name) in Downloads is not a valid Bluechip Board backup: it was left there and not used."; continue }
            $tC = Get-InstanteBackup $b
            if ($tC -gt $futuro) { Write-Warning "$($c.Name) in Downloads says it was saved on $($tC.UtcDateTime.ToString('yyyy-MM-dd HH:mm', $Script:Inv)) UTC, in the future: it was left there and not used. Check the PC's clock."; continue }
            if ($tC -le $tAtual) { continue }   # older than (or same as) the project file: left untouched
            if (Test-Path -LiteralPath $bkPath) { Copy-Item -LiteralPath $bkPath -Destination (Join-Path $Pasta 'bluechip-board-backup.previous.json') -Force }
            Move-Item -LiteralPath $c.FullName -Destination $bkPath -Force
            Write-Passo "Backup moved from Downloads to $bkNome (the previous file was kept as bluechip-board-backup.previous.json)"
            break
        }
    }
}
$backup = $null
if (Test-Path -LiteralPath $bkPath) {
    $backup = Read-Backup $bkPath
    if (-not $backup) { Write-Warning "$bkNome is not a valid Bluechip Board backup and was ignored." }
    elseif (($tBk = Get-InstanteBackup $backup) -gt $Script:Agora.AddDays(1)) {
        Write-Warning "$bkNome says it was saved on $($tBk.UtcDateTime.ToString('yyyy-MM-dd HH:mm', $Script:Inv)) UTC, in the future: no backup from Downloads can be newer, so none will be brought in. Check the PC's clock, then save the backup again from the website."
    }
}

# Indicadores da área do euro (ECB Data Portal): se falharem, a execução anterior até 30 dias, marcada como tal
Write-Passo 'Getting euro area indicators (ECB Data Portal)'
$macro = [ordered]@{}
foreach ($s in $SeriesMacro) {
    $m = Get-SerieMacro $s
    $v = if ($m.estado -ne 'ok' -and $anterior -and $anterior.macro) { $anterior.macro.($s.Id) } else { $null }
    $quando = if ($v) { ConvertTo-Data (ConvertTo-IsoUtc $v.obtidoEm) } else { $null }
    if ($v -and "$($v.estado)" -in 'ok', 'previous run' -and $quando -and ($Script:Agora - $quando).TotalDays -le 30) {
        # pares [data, valor] (aceita a forma {value, Count} do Windows PowerShell 5.1); a taxa pode ser negativa
        $pts = @(foreach ($p in @($v.pontos)) {
            $q = if ($p -is [array]) { $p } elseif ($null -ne $p -and $p.PSObject.Properties['value']) { @($p.value) } else { $null }
            $x = 0.0
            if ($q -and $q.Count -ge 2 -and "$($q[0])" -match '^\d{4}-\d{2}-\d{2}$' -and [double]::TryParse("$($q[1])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$x)) { , @("$($q[0])", $x) }
        })
        if ($pts.Count) { $m = [pscustomobject]@{ id = $s.Id; nome = $s.Nome; estado = 'previous run'; fonte = "previous run (retrieved $($quando.UtcDateTime.ToString('yyyy-MM-dd', $Script:Inv)))"; obtidoEm = $quando.ToString('o'); freq = "$($v.freq)"; pontos = $pts; erro = $m.erro } }
    }
    $macro[$s.Id] = $m
}

# Manutenção: dados do script que envelhecem e devem ser revistos à mão (aparecem na consola e em Sources & method)
$manutencao = New-Object System.Collections.Generic.List[string]
try {
    if (($Script:Agora.UtcDateTime - [datetime]::ParseExact($PesosReferencia.Data, 'yyyy-MM-dd', $Script:Inv)).TotalDays -gt 120) {
        $manutencao.Add("The reference ETF weights (`$PesosReferencia, from $($PesosReferencia.Data)) are over 4 months old. They are only used if the iShares download fails: refresh them from the fund page.")
    }
} catch { }
foreach ($e in @($ETFs | Where-Object { $_.Id -ne 'SXR8' -and $_.Referencia })) {
    $x = [datetime]::MinValue
    if ([datetime]::TryParseExact("$($e.Referencia.Data)", 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$x) -and ($Script:Agora.UtcDateTime - $x).TotalDays -gt 120) {
        $manutencao.Add("The reference weights of $($e.Id) (`$ETFs, from $($e.Referencia.Data)) are over 4 months old. They are only used if its iShares download fails: check them in the fund's holdings file.")
    }
}
if (-not @($Calendario | Where-Object { $x = [datetime]::MinValue; [datetime]::TryParseExact("$($_.d)", 'yyyy-MM-dd', $Script:Inv, [Globalization.DateTimeStyles]::None, [ref]$x) -and $x -gt $Script:Agora.UtcDateTime.AddDays(30) }).Count) {
    $manutencao.Add('The manual calendar ($Calendario) has no events more than 30 days ahead: add the next Fed meetings and earnings dates.')
}
foreach ($b in $Bolsas) {
    $ultimoAno = @($b.feriados | ForEach-Object { [int]$_.Substring(0, 4) } | Sort-Object)[-1]
    if ($ultimoAno -and $ultimoAno -lt $Script:Agora.Year) { $manutencao.Add("The exception list of $($b.nome) ends in ${ultimoAno}: regular holidays are computed by rule, but add any extraordinary closure for this year.") }
}
$lembReun = Get-LembreteReunioes $Calendario $Script:Agora
if ($lembReun) { $manutencao.Add($lembReun) }
$lembUA = Get-LembreteUA $Script:UAChrome $Script:UAData $Script:Agora
if ($lembUA) { $manutencao.Add($lembUA) }
$semAlias = @(Get-PosicoesSemAlias $pesosEtfs)
if ($semAlias.Count) { $manutencao.Add("Top 10 holdings without a news alias (`$AliasesPosicoes): $($semAlias -join '; '). Add an alias with word boundaries and exclusions, so their news is linked to the fund.") }
foreach ($m in $manutencao) { Write-Host "  ! $m" -ForegroundColor DarkYellow }

Write-Passo 'Building the website'
$dados = [ordered]@{
    geradoEm = $agoraIso
    dias = $Dias
    versao = '1.4'
    primeiraExecucao = $primeiraVez
    duplicadas = $duplicadas
    ativos = $ativosDados
    mercado = $mercadoDados
    fx = [ordered]@{
        yahoo = $(if ($fxSerie) { @($fxSerie.pontos) } else { @() }); bce = @($fxBce)
        fonte = $(if ($fxSerie) { $fxSerie.fonte } else { '' }); parcial = [bool]($fxSerie -and $fxSerie.parcial)
    }
    etf = $pesos
    etfs = $pesosEtfs
    bitcoin = $bitcoin
    historico = $historico
    splits = $desdobramentos
    dividendos = $dividendos
    noticias = $noticias
    historicoNoticias = [ordered]@{ inicio = $histNoticias.inicio; dias = $HistoricoDias; noticias = @($histNoticias.itens) }
    resultadosSec = $resultadosSec
    fundamentais = $fundamentais
    macro = $macro
    bolsas = $Bolsas
    calendario = $eventos
    manutencao = @($manutencao)
    fontes = $Script:Fontes.ToArray()
}
# "<" só aparece dentro de textos do JSON; trocá-lo pelo escape JSON (barra, "u003c") impede que um título com
# "</script>" ou "<!--" parta a página. O escape é montado a partir de partes para não poder ser convertido de volta
# num "<" ao editar o script (foi isso que tinha anulado esta proteção).
$escapeMenor = '\' + 'u003c'
$json = ($dados | ConvertTo-Json -Depth 10 -Compress).Replace('<', $escapeMenor)
# Os dados pessoais (compras) só entram no site principal, que os carrega sozinho ao abrir. O arquivo e o ficheiro
# de dados ficam sem eles, para não haver 30 cópias da carteira espalhadas pela pasta Archive.
$jsonBackup = if ($backup) { ($backup | ConvertTo-Json -Depth 10 -Compress).Replace('<', $escapeMenor) } else { 'null' }
$jsonSite = '{"backup":' + $jsonBackup + ',' + $json.Substring(1)
$htmlSite = (Get-Plantilla).Replace('__DADOS_JSON__', $jsonSite)
$htmlArquivo = (Get-Plantilla).Replace('__DADOS_JSON__', '{"backup":null,' + $json.Substring(1))
$utf8 = New-Object System.Text.UTF8Encoding($false)
# Escrita segura: grava num ficheiro temporário e só depois o troca pelo verdadeiro, para uma interrupção
# (falta de energia, fecho da janela) nunca deixar um site ou um ficheiro de dados cortado a meio
function Write-Atomico([string]$Caminho, [string]$Texto) {
    $tmp = "$Caminho.tmp"
    [IO.File]::WriteAllText($tmp, $Texto, $utf8)
    try {
        if (Test-Path -LiteralPath $Caminho) { [IO.File]::Replace($tmp, $Caminho, $null) } else { [IO.File]::Move($tmp, $Caminho) }
    } catch {
        Copy-Item -LiteralPath $tmp -Destination $Caminho -Force; Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}
# Cópia datada de cada execução, guardada à parte na pasta Archive (fica só o site mais recente na pasta principal)
$arquivo = Join-Path $Pasta 'Archive'
New-Item -ItemType Directory -Force -Path $arquivo | Out-Null
$ficheiro = Join-Path $arquivo ("bluechip-board-{0}.html" -f (Get-Date).ToString('yyyy-MM-dd_HHmm', $Script:Inv))
$ultimo = Join-Path $Pasta 'bluechip-board.html'
Write-Atomico $ficheiro $htmlArquivo
Write-Atomico $ultimo $htmlSite
Write-Atomico (Join-Path $Pasta 'bluechip-board-data.json') $json
Write-Atomico $vistosPath ($vistosNovos | ConvertTo-Json -Compress)
if ($histNoticias.estado -eq 'corrupted' -and (Test-Path -LiteralPath $histPath)) { Copy-Item -LiteralPath $histPath -Destination "$histPath.bad" -Force }
Write-Atomico $histPath ([ordered]@{ versao = 1; inicio = $histNoticias.inicio; dias = $HistoricoDias; noticias = @($histNoticias.itens) } | ConvertTo-Json -Depth 5 -Compress)
Get-ChildItem -Path $arquivo -Filter 'bluechip-board-2*.html' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue
# e no máximo ~300 MB no total (cada cópia cresce com o histórico de notícias): saem as mais antigas, ficando sempre 5
$copias = @(Get-ChildItem -Path $arquivo -Filter 'bluechip-board-2*.html' | Sort-Object LastWriteTime -Descending)
$total = 0; $nCopia = 0
foreach ($c in $copias) { $total += $c.Length; $nCopia++; if ($nCopia -gt 5 -and $total -gt 300MB) { Remove-Item -LiteralPath $c.FullName -Force -ErrorAction SilentlyContinue } }

$ok = @($Script:Fontes | Where-Object { $_.estado -like 'ok*' }).Count
$conta = { param($n) @($noticias | Where-Object { $_.nivel -eq $n }).Count }
Write-Host ''
Write-Host ("Done in {0:N0} s. Sources that responded: {1} of {2}." -f ((Get-Date) - $inicio).TotalSeconds, $ok, $Script:Fontes.Count) -ForegroundColor Green
Write-Host ("News: {0} in total ({4} duplicates grouped) | {1} material | {2} important | {3} moderate" -f $noticias.Count, (& $conta 'red'), (& $conta 'orange'), (& $conta 'yellow'), $duplicadas) -ForegroundColor Green
Write-Host "Website: $ultimo" -ForegroundColor Green
if (-not $NaoAbrir) { Start-Process $ultimo }
try { $Script:Trinco.ReleaseMutex() } catch { }
