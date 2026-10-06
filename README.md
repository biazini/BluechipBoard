# Bluechip Board

A personal market dashboard for a euro-based investor in Portugal. One PowerShell script collects news, prices and indicators for **Apple (AAPL)**, **NVIDIA (NVDA)**, **Alphabet (GOOGL)**, four accumulating iShares UCITS ETFs listed on Xetra in euros — **Core S&P 500 (SXR8)**, **Core MSCI Europe (EUNK)**, **Core MSCI EM IMI (IS3N)** and **Core MSCI Japan IMI (EUNN)** — and **Bitcoin (BTC, in euros)**. It then builds a single, self-contained HTML website you open in your browser. The Portfolio tab also projects your dividend income (Apple, NVIDIA, Alphabet) and exports your sales as preparation files for the Portuguese IRS return (Anexo J).

Everything runs locally on your PC. The script only *reads* public sources. Your portfolio and purchases are entered in the website and never leave your computer. They are kept in the browser and in `bluechip-board-backup.json`, and the script copies that file into `bluechip-board.html`, so do not share that page if you want to keep your portfolio private.

> Educational, rule-based tool. Not financial advice.

**How to read this document.** Part 1 is for *using* the dashboard. Part 2 explains *how it is built*: diagrams of the whole system and of each run, what depends on what, where each piece of code lives, and how to change it safely. The [glossary](#glossary-portuguese-names-in-the-code) at the end translates the Portuguese names used in the code. The diagrams are plain text, so they display the same in any editor.

---

## Contents

**Part 1 · Using it**

1. [Quick start](#quick-start)
2. [Website sections](#website-sections)
3. [Your data: portfolio, purchases and backup](#your-data-portfolio-purchases-and-backup)

**Part 2 · How it is built**

4. [The system at a glance](#the-system-at-a-glance)
5. [Project files](#project-files)
6. [Requirements and dependencies](#requirements-and-dependencies)
7. [How it works: one run, step by step](#how-it-works)
8. [Script parameters](#script-parameters)
9. [Script configuration](#script-configuration-top-of-bluechip-boardps1)
10. [Data sources](#data-sources)
11. [News classification](#news-classification)
12. [How the website works](#how-the-website-works)
13. [Code map](#code-map)
14. [Data formats](#data-formats)
15. [Website settings you can adjust](#website-settings-you-can-adjust)
16. [Scheduling a daily run](#scheduling-a-daily-run)
17. [Making a change safely](#making-a-change-safely) and [working with git](#working-with-git)
18. [Maintenance checklist](#maintenance-checklist)
19. [Tests](#tests)
20. [Troubleshooting](#troubleshooting)
21. [Glossary: Portuguese names in the code](#glossary-portuguese-names-in-the-code)

---

# Part 1 · Using it

## Quick start

1. Double-click the **Bluechip Board** shortcut on the desktop.
2. A PowerShell window shows the progress (about one minute). When it finishes, the website opens in your default browser.
3. The window closes by itself when everything went well. It stays open, showing the problem, if there was an error or a warning. Press **Enter** to close it.

You can also run it by hand from this folder:

```powershell
pwsh -ExecutionPolicy Bypass -File .\Start-BluechipBoard.ps1
```

`Start-BluechipBoard.ps1` is the launcher used by the shortcut. It calls `Bluechip-Board.ps1` with your SEC e-mail, read from `bluechip-board.config.json` (or the `BLUECHIP_SEC_EMAIL` environment variable), and keeps the window open on problems. You can also call the main script directly with any of the [parameters](#script-parameters) (`powershell` instead of `pwsh` runs it on Windows PowerShell 5.1, with the same result):

```powershell
pwsh -ExecutionPolicy Bypass -File .\Bluechip-Board.ps1 -SecEmail "your@email.com" -Days 3
```

---

## Website sections

```text
┌──────────────────────────────────────────────────────────────────────────────────────────┐
│ HEADER   Bluechip Board │ All · Apple · NVIDIA · Alphabet · ETF & market · Europe · …  │ US · Xetra │
├────────────────────────────────┬─────────────────────────────────────────────────────────┤
│ MARKET SNAPSHOT (left)         │ TABS  Overview · Portfolio · News · Prices ·            │
│  S&P 500 ETF (full width)      │       Fundamentals · Currency & Macroeconomics ·        │
│  Apple · NVIDIA · Alphabet     │       Calendar · ETFs · Bitcoin · Sources & method      │
│  MSCI Europe · EM IMI · Japan  │                                                         │
│  Bitcoin (full width)          │  content of the selected tab                            │
│  EUR/USD (full width)          │                                                         │
├────────────────────────────────┴─────────────────────────────────────────────────────────┤
│ FOOTER   Bluechip Board · Updated <date>, <time> Lisbon time · Sources & method          │
└──────────────────────────────────────────────────────────────────────────────────────────┘
```

**Header** (always visible)

- **Asset filter:** All · Apple · NVIDIA · Alphabet · ETF & market (SXR8 + market and macro news) · Europe (EUNK) · Emerging markets (IS3N) · Japan (EUNN) · Bitcoin. It filters alerts, news, prices and the calendar. The chips of the regional ETFs are built from `$ETFs` (`Chip`). Choosing an ETF chip also selects that fund in the ETFs tab.
- **Exchange hours:** US and Xetra, open or closed, with the next session in Lisbon time and a countdown on hover. They refresh every 30 s.

**Market snapshot** (left sidebar)

- Generation time (Lisbon), news window and EUR/USD.
- One card per asset:
  - price, with a euro equivalent for USD stocks (at the EUR/USD rate of the price's day);
  - the date of the price ("close 2 Oct", "live 15:36" while a session is open, "previous run"), in orange when the source is behind;
  - change over the last session, or "so far today" while it is open (Bitcoin: the CoinGecko 24 h change);
  - a 3-month sparkline;
  - the 1-month change and the distance from the 52-week high;
  - the 52-week range bar.
- An **EUR/USD** card, below Bitcoin. It shows the same series used for the conversions, with its date, a 3-month chart, the last-session change, 1 month, year to date, the 52-week range and the ECB reference rate. The full chart stays in *Currency & Macroeconomics*.

Layout by screen width:

| Width | Sidebar | Cards |
|---|---|---|
| under 1,180 px | above the tabs | a single list (two columns on tablets), S&P 500 first |
| 1,180–1,599 px | 320–400 px wide | a single list, S&P 500 first |
| 1,600–2,399 px | 29 % of the width | a grid: S&P 500 / three stocks / three regional ETFs / Bitcoin / EUR/USD |
| 2,400–2,999 px | 36 % | the same grid; the header filters start exactly where the section column starts |
| 3,000 px and up (4K) | 40 % | the same grid, bigger text |

In the grid, the sidebar fills the height of the screen. Every row of cards grows with the free space: the two rows of three cards take a larger share, the full-width cards a smaller one. The charts grow with their cards, so there is no empty space under the last card and the narrow charts do not become tall and thin. On a screen shorter than the content, the cards keep their minimum size and the sidebar scrolls on its own. The full-width cards are laid out horizontally. The narrow ones put the price under the name and use short labels (no "ETF" prefix, no currency, "From high"); the price symbol shows the currency. The side margins are 32 px from 1,600 px and 44 px from 3,000 px.

The tab bar has ten tabs. Up to 1,560 px, and from 1,561 to 2,399 px, the tabs use a tighter padding and font, so the bar fits from about 1,900 px. On narrower windows it scrolls sideways, and the current tab is scrolled into view.

**Tabs**

| Tab | Contents |
|---|---|
| **Overview** | Summary, KPI tiles (alerts, material and important news, sources OK), rule-based **alerts** (see below), the top stories of the last 72 h and a chart of news by asset and level. |
| **Portfolio** | Your holdings computed from your purchases and sales, **your return compared with the same money in SXR8**, your **target allocation and the split of your next contribution**, your **investment policy**, a **stress test** with three past market falls, your **real exposure by country, sector and underlying currency** with a **concentration** check by company, the projected **dividends** (with an estimated net amount), the register (with a **note** per entry), the **export for taxes (Anexo J)** and the separate **dividend export (Quadro 8A)**, an allocation chart, your real exposure to each company (direct holdings + look-through via each ETF, saying for each fund when its weights are the script's reference values) and the backup. See [Your data](#your-data-portfolio-purchases-and-backup). |
| **News** | All stories with filters (level, theme, source tier, search, period, sort by potential impact, most recent or **Relevance to me**, *only new since the last run*), duplicate groups, a "show more" button and a themes chart. Stories linked to a fund only through one of its largest holdings are tagged *Via top holding*. |
| **Prices** | Relative performance chart (indexed to 100, 1 month to 1 year, in € or $), a returns table (1 session, 1 week, 1 month, 3 months, year to date, 1 year), risk and trend indicators (52-week high, distance, 50/200-session averages, 30-session volatility, max drawdown) and the distance from the 52-week high. Each price shows its date. Bitcoin uses calendar days (7-day week, 365-day year). "1 year", "year to date" and "52-week" figures are only shown when the data covers that period. Markers on the chart for **earnings**, **rate decisions** and **material news** (hover for the headline), a **Big moves explained** table and an **Earnings reactions** table (see [Prices: events, big moves and earnings reactions](#prices-events-big-moves-and-earnings-reactions)). |
| **Fundamentals** | For Apple, NVIDIA and Alphabet, from their SEC filings: the last 8 quarters (revenue and growth over a year, gross and operating margin, free cash flow, diluted EPS, diluted shares and their change, with each quarter's publication date), P/E and P/FCF over the last 4 quarters with their percentile in the company's own 10-year history, and a 10-year chart without look-ahead. See [Fundamentals tab](#fundamentals-tab). Follows the asset filter. |
| **Currency & Macroeconomics** | EUR/USD KPIs and chart, how much of each stock's return came from currency, the S&P 500, VIX and 10-year Treasury yield, and the **euro area**: the ECB deposit facility rate and euro area inflation (HICP), each with its date, a 10-year chart and a short note (see [Euro area indicators](#euro-area-indicators-ecb)). No alerts come from them. |
| **Calendar** | Upcoming events with importance, status (confirmed / estimated / Nasdaq), countdown and an importance filter. Past events are optional. |
| **ETFs** | One view for the four ETFs, with a **Fund** selector (S&P 500 · MSCI Europe · MSCI EM IMI · MSCI Japan IMI; SXR8 by default). For the chosen fund, from its own data: <ul><li>a description (ISIN, Xetra listing, index, accumulating, domicile);</li><li>for SXR8 only, the weights of the three companies (the other funds hold none of them, so this figure is hidden);</li><li>its own top 10 holdings and sectors;</li><li>its price since the start of its history (5 years / 10 years / since 2010, or since 2014 for IS3N, with annualised return);</li><li>a **Linear / Log** scale switch on the long-term chart (linear by default);</li><li>the biggest drops of 10 % or more (depth, time falling, recovery date and time), with an **underwater chart** beside them (how far below its highest close so far the price was each day);</li><li>**rolling returns** over 1, 3 and 5 years (opens on demand, see below);</li><li>a **monthly investment simulator** on the ETF's own price, and the comparison of three ways to invest the same money.</li></ul> |
| **Bitcoin** | Price, 24 h / 7-day change, volatility versus the S&P 500, the BTC/EUR chart with 50/200-day averages, Fear & Greed, the next halving, network and market data, and correlation with your other assets (3 months). Also: the long-term log-scale chart with the halvings marked, the drop from the all-time high (the history starts in Sep 2014, after the late-2013 peak, which the chart notes), the biggest drops of 40 % or more, **rolling returns** over 1, 3 and 5 years (opens on demand), a **monthly investment simulator** and the comparison of three ways to invest the same money, the **365-day tax counter** (see below), the top Bitcoin stories and a note on Portuguese crypto taxation. |
| **Sources & method** | The status, item count, response time and error detail of every source, maintenance reminders (old reference weights, an empty manual calendar, a top 10 holding without a news alias), and how the classification works. |

**Alerts** (Overview)

| Rule | Level |
|---|---|
| No prices at all for an asset | Moderate |
| An asset's latest price is behind its exchange, or comes from the previous run | Important |
| EUR/USD missing (euro values left blank) / from a fallback / behind | Material / Important / Important |
| Daily move ≥ 4 % (Bitcoin ≥ 6 % over 24 h) | Important |
| ≥ 20 % below the 52-week high (Bitcoin ≥ 30 %) | Important |
| ≥ 10 % below the 52-week high (Bitcoin ≥ 15 %) | Moderate |
| Within 1.5 % of the 52-week high | Info |
| Price below the 200-day average | Moderate |
| EUR/USD moved ≥ 2 % in a month | Moderate |
| VIX ≥ 25 / ≥ 20 | Important / Moderate |
| High-importance calendar event within 7 days | Important |
| Material news about an asset in the last 72 h | Material |
| Bitcoin Fear & Greed in "extreme fear" (≤ 24) or "extreme greed" (≥ 76), Alternative.me's own bands | Moderate |

**Context under the drop alerts.** The two "below its 52-week high" alerts (stocks and ETFs from 10 % and 20 %, Bitcoin from 15 % and 30 %) keep their level and text. They gain a line of context underneath:
- **your own rule** for that depth, from your [investment policy](#investment-policy-and-notes): the 30 % rule from a 30 % drop (or the 20 % rule if that one is empty), the 20 % rule from a 20 % drop;
- **the asset's history**, from the long-term prices in its own currency: how many times it fell as much or more from a peak (`quedas()`, the same calculation as the "biggest drops" tables), how many recovered and the median time from the bottom back to the peak, how many have not recovered yet (including the current drop), with the number of cases and the dates of the sample;
- **a caveat:** the sample (since 2010, or 2014 for Bitcoin) is mostly a rising market, so past recoveries may have been quicker than they would be in a long bear market.

With no earlier drop that deep, the line says so plainly.

**Rolling returns** (ETF and Bitcoin tabs, a section you open)

They answer a simple question: how did *every* 1-, 3- and 5-year period in the history turn out?
- A window starts on each trading session (252 sessions = 1 year) or, for Bitcoin, on each calendar day (365 days = 1 year), and runs for that length.
- The table shows, per length:
  - the number of windows;
  - the worst, median and best return (annualised too for 3 and 5 years);
  - the share of windows that ended negative.
- A histogram (choose 1, 3 or 5 years) shows how the returns were spread.
- A length longer than the history shows **Unavailable**; IS3N, for example, has no history before 2014.
- Consecutive windows overlap, so they are not independent results, and the page says so.
- Nothing is calculated until you open the section. In the ETFs tab it follows the fund you choose.

The Bitcoin tab already had its underwater chart (*Drop from the all-time high*, the same formula), so it was not duplicated. The US stocks have no long-term chart in the site (Prices covers up to one year), so the log switch is only on the ETF chart; Bitcoin's long-term chart is always logarithmic.

**Monthly investment simulator** (ETF and Bitcoin tabs)

You choose an amount per month and a start month. It buys on the first trading day of each month at the closing price, and shows:

- invested, value today, gain, units held and average price paid;
- a chart of value against amount invested.

Fees, spreads and taxes are not included. The ETFs are accumulating, so dividends are already in their price. Each ETF has its own simulator with its own saved inputs (`bb.sim.etf` for SXR8; `bb.sim.etf-EUNK`, `bb.sim.etf-IS3N`, `bb.sim.etf-EUNN`).

**Three ways to invest the same money** (ETF and Bitcoin tabs, below the monthly simulator). It uses the same monthly amount and start month as the simulator above, so the total and the period are the same, and compares:

```text
(a) Monthly purchases  every month's amount buys on the first trading day of the month (the simulator above)
(b) All at once        the whole total buys on that first day
(c) Buy the dip        every month's amount waits in cash (earning 0 %); all the cash buys at the first close that is
                       X % or more below the highest close of the previous 52 weeks (X = 10 % by default, adjustable);
                       cash not invested at the end counts at face value
```

It shows each strategy's final value and gain (and, for the dip strategy, the number of purchases and any cash left), plus a chart of the three values over time.
- The X setting is saved in `bb.strat.etf` / `bb.strat.btc`. The monthly simulator's own inputs (`bb.sim.*`) and results do not change.
- In the ETFs tab it follows the fund you choose.
- It is a comparison of the past, without fees, spreads or taxes; not a forecast or advice.

**Stale data warning.** If the page is opened more than 36 hours after it was built, a banner asks you to run the script again.

### Fundamentals tab

`#fundamentals`, between Prices and Currency & Macroeconomics. It uses the [SEC fundamentals](#fundamentals-sec-xbrl) of the data and is drawn only when the tab opens. With the asset filter on one company, only that company is shown. On an ETF, Bitcoin or the market, a note explains that the tab covers the three companies.

**For each company:**
- **The last 8 quarters.**
  - Revenue, and its growth against the quarter a year earlier.
  - Gross and operating margin.
  - Free cash flow (operating cash flow − capex).
  - Diluted EPS.
  - Diluted shares, with their change over a year: *buybacks* when it falls, *dilution* when it rises.
  - When it was **published** at the SEC (and *revised* when a later filing changed it).

  Derived figures are marked `*`: the 4th quarter is the year minus 9 months, and cash-flow quarters come from year-to-date figures.
- **Valuation now.**
  - **P/E** = latest price ÷ diluted EPS of the last 4 quarters.
  - **P/FCF** = latest price × diluted shares ÷ free cash flow of the last 4 quarters.

  Each shows the price date, the quarters used and its **percentile**: the share of the month-end values of the 10-year history at or below today's value. Negative earnings or cash flow show *n/m*.
- **10-year history** of both ratios, at each month-end.
  - **No look-ahead:** each point uses only the quarters already published on that date (by their SEC filing date).
  - A figure revised later counts with its first published value until the revision was filed.
  - Prices are the long-term daily closes in dollars, adjusted for splits like the EPS and the shares.

A note says that valuation gives context and does not predict the short term. A company without data shows **Unavailable** with the reason (not collected, source failed, a metric the company does not report). The other companies are shown normally.
### Prices: events, big moves and earnings reactions

These parts use the data of the [news history](#project-files) and the [SEC past earnings dates](#past-earnings-dates-sec). Without that data, each part shows **Unavailable** with the reason, and the rest of the tab works as before.

**Markers on the chart.** At the bottom of *Relative performance*, one marker per day and kind; hover a marker for its details.

| Row | Marker | From |
|---|---|---|
| bottom | **Earnings** (the company's colour) | The SEC 8-K with item 2.02, on the reaction session. Without a known acceptance time, on the filing date (the tooltip says so). |
| middle | **Rate decision** | Past calendar events such as *Fed decision…*, and stories in the history with the *Macro and rates* theme whose headline is a decision (the same rule as `$DecisaoJuros` in the script). |
| top | **Material news** | Material stories of the history, for the assets in view. |

The tooltip shows up to three headlines (and "+N more"), as text, never as HTML.

**Big moves explained.** It lists the sessions of the chosen period with a daily move of **4 %** or more for stocks and ETFs, or **6 %** for Bitcoin. These are the thresholds of the daily-move alert.
- Moves are measured in the asset's own currency; an intraday price does not count.
- Each move lists the material and important stories of the history for that session and the one before it. They are the asset's own stories, plus Market stories for the US stocks and SXR8.
- A story belongs to a session in the **exchange's time zone**: New York for the US stocks, Xetra for the ETFs, the UTC day for Bitcoin. A story after the close (or after an early close) counts for the next session, and a weekend or holiday story for the next session in the price data.
- The note gives the start of the history ("News history since …"). A move before that start says "Before the news history" instead of showing nothing.
- The 30 most recent moves are shown.

**Earnings reactions.** For Apple, NVIDIA and Alphabet:
- the latest earnings release: its date, the New York acceptance time and the reaction session;
- the move in **dollars** in the reaction session (its close against the close before it), and after **5 sessions** (the close of the 5th session, the reaction session being the 1st);
- the **median absolute move** of the reaction session, and the number of cases;
- all cases, in a list below the table.

The moves use the long-term daily closes, adjusted for splits. Releases without a known acceptance time, or whose reaction session is not yet in the price data, are not counted, and the table says how many.
---

## Your data: portfolio, purchases and backup

### Registering purchases and sales

In **Portfolio → Purchases and sales · stocks, ETF and Bitcoin**:

1. Choose the **asset**, the **type** (purchase or sale) and the **date**.
2. The **price per unit (€)** fills in by itself with that day's closing price, from the long-term history:
   - The ETFs (SXR8, EUNK, IS3N, EUNN): Xetra close in euros.
   - US stocks: the USD close converted at that day's EUR/USD rate (both values are shown).
   - Bitcoin: BTC/EUR at the end of the day (UTC).
   - On weekends or holidays, the last close before that date is used, and the message says so.
   - If the page was built while that day's session was still open, the price is an intraday price, and the message says so.

   Change the price if yours was different.
3. Enter the **quantity** and, if you want, the **fee** you paid (in €, optional). Click **Add purchase** / **Add sale**.
   - Remove any entry with the bin icon (you are asked to confirm).
   - The register has a **Fee** column: the euro button on each row adds, changes or removes that entry's fee.
   - The realised gain shown in the register does not deduct fees. They are used in your return (XIRR) and in the *Despesas e encargos* column of the tax export.

**Before you sell · simulation** (Portfolio). Choose an asset and a quantity and click **Simulate**. Nothing is saved.
- The page runs the same FIFO on an in-memory copy, with a hypothetical sale today at the latest price.
- It shows which purchases it would use, the proceeds, cost and gain, and a **tax estimate at 28 %** on the positive gain.
- For Bitcoin, purchases held 365 days or more are **exempt** and excluded from the estimate. A purchase less than 30 days from becoming exempt gets a warning.
- The estimate is a simplification: no offsetting of losses from other sales, no *englobamento* option, no fees.
- A sale larger than what you hold is refused, as in the register.

**Sales** use up your oldest purchases first (first in, first out). A sale of more than you held on that date is refused. The register shows the cost of each sale and its realised gain. Purchases fully sold are greyed out.

**Your holdings** is read-only. Each asset shows:

- **Quantity:** purchases minus sales.
- **Total invested:** the cost of what you still hold.
- **Value:** at the latest prices.

Above it are the totals: value, invested, gain, realised gain (if you sold) and the latest daily move. An asset with no current price is shown as "no price" instead of a value.

> **Stock splits.** Yahoo's history is split-adjusted, so an old purchase is entered in today's shares. The site warns you when the date is before a split and tells you how many times to multiply.
>
> **Future splits are handled automatically.** Each purchase and sale saves the closing price of its date as it was when you entered it, plus the date of that price data. If a company splits later, Yahoo's history for that date drops by the split ratio. The site detects this, using the split events Yahoo provides (with the list `SPLITS_FIXOS` in the template as a reserve), and shows the entry in the new shares: quantity × ratio, price ÷ ratio, the same amount invested. A note "adjusted for a 10:1 split" appears in the register. If the history changes without a known split, the entry is not adjusted and a warning asks you to check it.

### Your return, compared with the same money in SXR8

**Portfolio → Your return · compared with the same money in SXR8**, just below *Your holdings*, answers two questions:

- How much did your money earn?
- Would the same money, put into the S&P 500 ETF on the same days, have done better or worse?

```text
your register ─► cash flows:  purchase  −(quantity × price you entered)   (Bitcoin: total paid)
                              sale      +(quantity × sale price)
                              today     +(value of what you hold now, as in Your holdings)
        │
        ├─► Your return  = the rate that makes those flows add up to zero (XIRR, money-weighted), in euros
        │
        └─► Same money in SXR8: each purchase buys SXR8 for the same euros, at that day's Xetra close
                                (or the last close before it); each sale takes the same euros out of it
                              ─► SXR8 units × latest SXR8 price = benchmark value ─► its own XIRR
Difference = your return − the benchmark's, in percentage points (and in euros of value today)
```

- **Money-weighted.** The XIRR counts *when* you put money in and took it out, so it is the return of your own decisions, timing included. It is not the price change of any asset.
- **Annualised from one year on.** If your first purchase is less than 365 days old, the figure is the return **for the period**, not annualised, and the page says so. A yearly rate over a few months would exaggerate.
- **Benchmark rules.** Every purchase counts, whatever the asset (Bitcoin included). If a sale is larger than what the benchmark holds at that point, the benchmark's units are all used, it receives less than your sale, and the page says so.
- **Unavailable, never guessed.** Each figure shows "Unavailable" with the reason in these cases:
  - an asset you hold has no current price (it is named);
  - an entry has no price, or a sale has no matching purchase;
  - a purchase is older than the SXR8 history (the date is shown);
  - the calculation finds no solution.
- **Dates and sources.** A note under the figures lists:
  - the number of purchases and sales, and the date of the first one;
  - the date of each price used for today's value;
  - the date of the SXR8 price, marked *intraday* while Xetra is open.
- **Solver.** It uses Newton's method from 10 %, falling back to bisection between −99.9999 % and 10^12. Cash flows are dated at midnight UTC, and a year is 365 days.
- **Fees and taxes.** Fees you record are included: a purchase costs its price plus the fee, and a sale brings in its price minus the fee; the benchmark gets the same cash flows. With no fees recorded, the figures are exactly as before. Taxes are not included. Past returns say nothing about future returns.

### Target allocation and next contribution

**Portfolio → Target allocation · next contribution**, below *Your return*, compares what you hold with the mix you aim for. It also works out how to split the next amount you invest so that you move back towards that mix **without selling**.

1. **Set your targets.** In the table, type a target weight for each asset you want (the weights must add up to 100 %). Above the table, enter the tolerance band (± percentage points, 5 by default) and the amount you plan to invest each month (M). Then click **Save targets**.
   - Targets that do not add up to 100 %, a band outside 0–50 or a negative amount are refused with a message, and nothing is saved.
   - Saved targets go into the backup file with the rest of your data (backup version 5).
2. **Read the table.** Each asset shows its value, its current weight, the target and the deviation in percentage points. An asset further from its target than the band is tagged **outside the band**.
3. **Next contribution.** For M:

```text
target value of each asset = target weight × (what you hold now + M)
gap of each asset          = max(0, target value − current value)
gaps add up to M or more ─► M is shared in proportion to the gaps
gaps add up to less      ─► each gap is covered and the rest is shared by the target weights
```

   With weights that add up to 100 %, the gaps never add up to less than M: at most they equal it, when every asset is below its target, and each asset then gets exactly its gap. The second rule is kept as written in the specification.
4. **Months back inside the band.** The page repeats that split month after month, with today's prices, and says how many monthly contributions it would take to bring every asset back inside the band without selling. If no asset is outside the band, it says so; beyond 50 years, it says that too.

- **Assets without a current price** are left out and named, and the targets of the others are scaled up to 100 %.
- **The fixed note** under the section says that rebalancing by selling realises taxable capital gains, and rebalancing with new contributions does not.
- **It is arithmetic, not advice.** Prices move, so the split is for the next contribution only, at today's prices.

### Real exposure (country, sector, currency) and concentration

**Portfolio → Real exposure · country, sector and currency** shows where your money really is. Each ETF is looked through with the breakdown from its latest iShares file, weighted by what you hold today:
- **Country:** from the file's *Geography/Locations* block (EUNK, IS3N). For SXR8 and EUNN the file has no country breakdown, so the country of the index is used (United States, Japan); the page says it is approximate.
- **Sector and underlying currency** (the currency each holding trades in): from the holdings sheet.

Cash, futures and FX inside a fund count as **Cash/Other**.
- Apple, NVIDIA and Alphabet count as United States and USD. Their sector comes from the iShares file when a fund lists them, otherwise from `SETOR_ACOES`.
- Bitcoin counts as **Crypto** in all three.
- A fund whose breakdown is Unavailable (download failed, reference weights, or an older data file) is left out. The page shows the share of the portfolio not covered, and the percentages are of the rest.
- A fixed note reminds you that **SXR8 is quoted in EUR, but its underlying currency exposure is USD**.

**Concentration by company** only where the company is unambiguous:
- Apple, NVIDIA and Alphabet: directly plus through the ETFs (the existing look-through);
- each fund's top 10, on its own.

Companies are never added up across funds by name.
- A company above `CONC_LIMIAR` (10 % of the portfolio by default) gets an informative note and a tag in the table. It is not an Overview alert.

### Stress test (past market falls)

**Portfolio → Stress test · past market falls** is a clearly labelled **hypothetical scenario**. It asks: what would your holdings today lose if each asset repeated its own move during three falls of the S&P 500?

| Episode | Peak | Trough | S&P 500 (closes) |
|---|---|---|---|
| 2018 Q4 | 20 Sep 2018 | 24 Dec 2018 | −19.8 % |
| COVID | 19 Feb 2020 | 23 Mar 2020 | −33.9 % |
| 2022 | 3 Jan 2022 | 12 Oct 2022 | −25.4 % |

The dates are the peak and trough closes of `^GSPC`, checked against Yahoo's data on 6 Oct 2026; they are in the `EPISODIOS` constant of the template.
- **Each asset's own move, in euros.** For every asset you hold with a current price, the page takes its close on the peak date and on the trough date, or the last close before (Xetra was closed on 24 Dec 2018, so the ETFs use 21 Dec). US stocks are converted at each day's EUR/USD. That return is applied to today's value of the holding.
- **The table** shows, per asset and in total, the hypothetical loss in euros and in %. Under each figure it shows how long the asset took, after the trough, to get back to its starting (peak) value in euros, or "not back yet".
- **Missing data.** An asset whose history does not cover an episode (for example IS3N, listed in 2014, would miss earlier ones) shows "no data for this period" there. It is left out of that episode's total, and the total names it. An asset you hold without a current price is left out and named.

### Investment policy and notes

**Portfolio → Investment policy** keeps your own rules, written while markets are calm. It has six free-text fields: horizon, allocation, monthly amount, what would make you sell, and your rule for a 20 % and for a 30 % drop. Click **Save policy**.
- The two drop rules appear under the matching drop alerts in the Overview (see [Alerts](#website-sections)).
- Nothing in the policy changes a calculation; it is there to be read.

**Notes per entry.** Every purchase and sale in the register has a note button (book icon): write a short note ("why I bought it"), or leave it empty to remove it. The note is shown under the asset.
- Notes are kept beside the entries, not inside them, so the entries and every older backup stay as they were.
- Deleting an entry deletes its note. A note whose entry no longer exists is ignored and is not saved in the backup.

Your policy and notes stay on this computer, like the rest of your data. They go into `bluechip-board-backup.json` and from there into `bluechip-board.html` only, and the page shows them as plain text.

### 365-day tax counter (Bitcoin)

Bitcoin purchases share one list. You can add them in the Portfolio register or in the **Bitcoin → 365-day tax counter** form (date, BTC amount, total paid in €). For each purchase, the counter shows:

- its value and gain;
- days held;
- whether it is already **tax-free** (held 365 days or more) or how many days are left, with a progress bar;
- how much of it is left after sales: sales recorded in the Portfolio use up the oldest purchases first, and a purchase fully sold shows **Sold**.

A summary shows tax-free versus still-taxable BTC and the next purchase to become tax-free. A purchase counts as tax-free from the day it completes 365 days. If you plan to sell close to that day, leave a margin of a day or two, and confirm your situation with the Portal das Finanças or an accountant.

### Dividends (projected income)

**Portfolio → Dividends · projected for the next 12 months** covers Apple, NVIDIA and Alphabet. The ETFs (SXR8, EUNK, IS3N, EUNN) are excluded because they are accumulating, both in the script and in the website (`DIV_IDS` drops any accumulating ETF). Bitcoin pays no dividends.

**Data.** On every run, the script reads Yahoo's dividend events for each company over the last 2 years. Each event has an ex-dividend date and an amount per share in dollars, already adjusted for later splits, like the prices. For each company it stores (`dividendos.<id>` in the embedded JSON):

| Field | Meaning |
|---|---|
| `anualPorAcao` | Annual dividend per share, USD = latest payment × payments per year (the "indicated" rate). The number of payments per year (12, 4, 2 or 1) comes from the usual interval between payments. |
| `ttmPorAcao` | Sum of the payments of the last 365 days, USD (shown for reference). |
| `frequencia`, `ultimo`, `pagamentos` | Payments per year, the latest payment `[ex-date, amount]` and all payments read. |
| `rendimentoPct` | Yield = `anualPorAcao` ÷ the price in the same response, in **percentage points** (0.32 = 0.32 %). It is computed by the script, so it does not depend on whether a provider writes yields as decimals or percentages. |
| `preco`, `precoData` | The USD price used for the yield and its time. |
| `estado`, `fonte`, `obtidoEm`, `erro`, `nota` | `ok`, `previous run` or `error`; source; when the data was retrieved; error text; any note (for example "may have been suspended"). |

**Validation.** The response must be for the right symbol and in USD, with a positive price. Each payment must be positive, below 25 % of the price and not far in the future; other payments are ignored and counted in *Sources & method*. The yield must be between 0 and 25 %. With a single payment the frequency is unknown, and if the latest payment is much older than the usual interval the dividend may have been suspended. In both cases no annual figure is projected.

**Calculation.** Projected annual dividend (USD) = **shares you hold now** × annual dividend per share. "Shares you hold now" is the quantity in *Your holdings*, from the same FIFO calculation (purchases minus sales, split-adjusted, fractional shares included). There is no second holdings calculation. In euros = USD ÷ the **latest EUR/USD rate** of the page, named with its date and source in the "EUR/USD used" tile. The latest rate is used, not the rate on the purchase date, because the payments are still to come. Each row is rounded to cents, and the total is the sum of the rows shown. The tiles also show the monthly average and the yield on the value of what you hold of the three.

**Missing data.** A row shows **Unavailable** (with the reason) when the source failed, the data is invalid or older than 180 days. It shows **No holdings** when you hold none, and **FX unavailable** in the euro column when there is no EUR/USD rate. A missing value is never shown as €0, and an unavailable company is left out of the total, which says so. Amounts are before US withholding tax and Portuguese tax.

**Net (est.).** The last column is an estimate after tax: projected euros × **0.72**. It assumes 15 % withheld in the US (form W-8BEN) plus the top-up to 28 % in Portugal (the special rate, without *englobamento*), so 28 % in all. Each row is rounded to cents and its total is the sum of the rows. It shows **No holdings**, **Unavailable** or **FX unavailable** like the euro column. The existing columns and totals did not change. The rates are in `ANEXO_J.q8a` (`retFonte`, `taxaPT`).

### Export for taxes (IRS, Anexo J)

**Fees in the export.** If you record fees, *Despesas e encargos* of each sale/purchase row is:

```text
purchase fee × (quantity used ÷ quantity bought)  +  sale fee × (quantity of the row ÷ quantity sold)
```

- Each row is rounded to cents, and the remainder goes on the sale's last row, so the totals add up.
- With **no fee** recorded for either side, the cell stays blank, exactly as before (`taxBaseline` checks this byte by byte).
- With **only one** of the two, the known part is filled in and flagged `FEE_PARTIAL`.
- With **both**, "Despesas e encargos" leaves that row's "Fill in manually" list.

**Portfolio → Export for taxes · IRS, Anexo J.** Choose the **tax year** (the years with sales are listed) and click **Export for Taxes (Anexo J)**. A summary shows the tax year, the number of stock/ETF and crypto disposals, how many rows go to each table and how many need review. After you confirm, two files are downloaded:

- `AnexoJ_Stocks_ETFs_<year>.csv`: for **Quadro 9.2A**;
- `AnexoJ_Crypto_<year>.csv`: for **Quadro 9.4A**.

Everything is calculated in the browser. Nothing is sent anywhere, and nothing in your portfolio, `localStorage` or backup is changed.

```text
your register (bb.buys, bb.lots, bb.sales)
        │
        ▼
carteira()  ── the page's only FIFO: matches each sale with the oldest purchases still held
        │        (the same result feeds Your holdings, the realised gain and the 365-day counter)
        ▼
one row per sale/purchase pair ── values = quantity × the euro prices YOU entered (no FX applied)
        │
        ├── stocks and ETFs ──► ANEXO_J mapping (code, country, ISIN) ──► AnexoJ_Stocks_ETFs_<year>.csv (Quadro 9.2A)
        ├── BTC held < 365 days ───────────────────────────────────────► AnexoJ_Crypto_<year>.csv     (Quadro 9.4A)
        └── BTC held ≥ 365 days ──► not exported (exempt; Anexo G1, Quadro 7)
```

> **These are preparation files, not an upload format.** The Portal das Finanças does not import these CSV files. It accepts the whole declaration, filled in on the Portal or as the XML file the Portal itself saves. Copy each row into Anexo J yourself, and check it. This tool is educational and rule-based, not tax advice: confirm with the official instructions or an accountant.

**Official reference used.** The *Modelo 3 – Anexo J* form, "modelo em vigor a partir de janeiro de 2026", and its filling-in instructions, published by the Autoridade Tributária ([form and instructions, PDF](https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/modelos_formularios/irs/Documents/Mod_3_anexo_J.pdf)):

- **Quadro 9.2A** (art. 10, n.º 1, b) CIRS), *alienação onerosa de partes sociais e outros valores mobiliários*. Columns: País da fonte · Código · Realização (Ano, Mês, Dia, Valor) · Aquisição (Ano, Mês, Dia, Valor) · Despesas e encargos · Imposto pago no estrangeiro · País da Contraparte · "Respeita a valores mobiliários admitidos à negociação ou a partes de OIC abertos?" (Sim/Não). Codes (Tabela VII) include **G01** (*alienação onerosa de ações/partes sociais*) and **G20** (*resgates ou alienação de unidades de participação ou liquidação de fundos de investimento*).
- **Quadro 9.4A** (art. 10, n.º 1, k), n.º 19 and n.º 22 CIRS), crypto-assets that are not securities, **held less than 365 days**. Columns: País da fonte · Realização (Ano, Mês, Dia, Valor) · Aquisição (Ano, Mês, Dia, Valor) · Despesas e encargos · Imposto pago no estrangeiro · País da Contraparte. There is no code column.
- Country codes come from **Tabela X**: United States 840, Ireland 372.

**Mapping used** (the `ANEXO_J` object in the template is the single place to change if the AT changes the form):

| Asset | Table | Código | País da fonte | Admitted to trading | ISIN |
|---|---|---|---|---|---|
| AAPL, NVDA, GOOGL | 9.2A | G01 | 840 (issuer in the US) | Sim | US0378331005, US67066G1040, US02079K3059 |
| SXR8 | 9.2A | G20 | 372 (fund domiciled in Ireland) | Sim | IE00B5BMR087 |
| EUNK, IS3N, EUNN | 9.2A | G20 | 372 (funds domiciled in Ireland) | Sim (Xetra) | IE00B4K48X80, IE00BKM4GZ66, IE00B4L5YX21 |
| BTC held < 365 days | 9.4A | (none) | left blank: the country of the platform you used | – | – |
| BTC held ≥ 365 days | not Anexo J | – | – | – | – |

> **Interpretations to confirm.** "País da fonte" as the issuer's country is common practice, but the instructions only say "país da fonte dos rendimentos". G20 for a UCITS ETF follows the usual reading of "unidades de participação", although SXR8 is legally a share of an Irish investment company (some advisers use G01). EUNK, IS3N and EUNN have the same legal form, verified in each fund's Key Facts:
- UCITS sub-funds of Irish iShares investment companies;
- domiciled in Ireland;
- accumulating;
- listed on Xetra.

So they get the same mapping, and the same caveat applies. They are priced in euros, so their rows have no dollar columns. A Bitcoin sale after 365 days or more is exempt and is declared in Anexo G1, Quadro 7, not in Anexo J. If you used a Portuguese platform, crypto goes in Anexo G instead of Anexo J. Check your case.

**FIFO.** The export uses the page's only FIFO calculation (`carteira()`), which also drives *Your holdings*, the realised gain and the 365-day counter. Each sale records the purchases it used up. The export writes **one row per sale/purchase pair**: selling 120 shares from purchases of 100 and 50 gives two rows, 100 + 20. Unsold purchases are not exported, and years are kept apart by the sale date.

**Stock splits.** Quantities and unit prices are the split-adjusted values the page already uses (`efetiva()`), in today's shares. The split factor of the purchase and of the sale is in its own column. Amounts in euros do not change with a split. Each entry is adjusted once, from its own reference price or data date, so a split is never applied twice. Anexo J has no quantity column, so the quantity is for your reference.

**Values and FX.** The prices in the register are the **euro prices you entered** (pre-filled with the close at that day's EUR/USD, which you may have changed). The export uses them as they are and never replaces them with a market close:

- Valor de aquisição = quantity × your purchase price (€).
- Valor de realização = quantity × your sale price (€).
- Gain/loss = realização − aquisição, with each value rounded to cents.

So no exchange rate is applied to the tax values. For US stocks, the dollar columns are a **reference only**: euro amount × the EUR/USD rate of **each transaction's own date** (acquisition and sale are looked up separately). The rate comes from the long-term Yahoo history, or the page's other EUR/USD series, and is never from outside the period those series cover. The rate's date and source are in the file. The register stores dates without a time, so the rates are **daily closing rates**, not the rate at the time of the trade, and they are Yahoo's market rates, not the ECB / Banco de Portugal reference rate. If no rate exists for a date, the dollar columns are left blank and the row gets the `NO_FX_REFERENCE` flag; today's rate is never used instead. If your broker traded in dollars, check how it converted the amounts you entered.

**CSV format.**

- UTF-8 with BOM.
- `;` separator and **decimal comma**, the format Excel uses in Portuguese. Numbers have no thousands separator and no currency symbols, and are formatted by the script, never by the browser's language.
- Euro amounts have 2 decimals, unit prices and rates 6, and quantities 8.
- Dates are ISO (`yyyy-mm-dd`), plus the form's separate Ano / Mês / Dia columns.
- CRLF line ends.
- Text with `;`, quotes or commas is quoted. Text starting with `=`, `+`, `-` or `@` gets a leading `'`, so a spreadsheet cannot run it as a formula.
- The same data always gives the same file.

Columns, stocks/ETF:

1. Tax year · Row · Anexo J table · Asset · Ticker · ISIN · Transaction type.
2. The form's columns, under their Portuguese names: País da fonte · Código · Realização Ano/Mês/Dia/Valor (EUR) · Aquisição Ano/Mês/Dia/Valor (EUR) · Despesas e encargos (EUR) · Imposto pago no estrangeiro (EUR) · País da Contraparte · Respeita a valores mobiliários admitidos à negociação…?
3. Sale date · Acquisition date · Days held · Held under 365 days · Capital gain/loss (EUR) · FIFO lot reference · Sale reference · Status · Flags.
4. Quantity · purchase and sale split factors · unit prices (EUR) · acquisition and sale FX EUR/USD with their dates · unit prices and totals in USD (derived) · FX source · Price basis · Fill in manually.

The crypto file has the same structure for Quadro 9.4A (no Código or ISIN; quantity in BTC).

**What you fill in yourself.** The page does not know these, so they are left blank, never invented:

- País da Contraparte, the residence of the buyer;
- Despesas e encargos, your fees, unless you recorded them (see below);
- Imposto pago no estrangeiro;
- for crypto, País da fonte.

The "Fill in manually" column lists them on every row.

**Status and flags.** `OK` or `REVIEW`. Rows needing review are still exported, so nothing is hidden:

| Flag | Meaning |
|---|---|
| `OVERSOLD_NO_PURCHASE` | Part of a sale has no purchase to match (possible with old or restored data). The row has no acquisition values: record the missing purchase. |
| `SPLIT_CHECK` | The price history changed without a known split, so the quantity may need checking. |
| `NO_SALE_PRICE` | The sale has no price. |
| `BAD_PURCHASE_DATE` | The matched purchase has an invalid date. |
| `NEAR_365_DAYS` | Crypto sold within 2 days of the 365-day mark; the counter's rule (exempt from the day 365 days are completed) applies, but check. |
| `NO_FX_REFERENCE` | Informational: no EUR/USD for a date, so the dollar columns are blank. The euro values are not affected. |
| `FEE_PARTIAL` | Informational: only one of the two fees (purchase or sale) is recorded, so *Despesas e encargos* holds that part only. It stays in "Fill in manually". |
| `NO_ANEXO_J_MAPPING` | A stock or ETF has no line in `ANEXO_J` (for example, an ETF added to the configuration without its tax mapping), so Código and País da fonte are blank. |

Sales with an invalid date, asset or quantity are not exported, and the summary counts them. A year without sales gives header-only files.

### Export dividends (Anexo J, Quadro 8A)

**Portfolio → Export for taxes**, second form. Choose the **Tax year (dividends)** (the years with a dividend on shares you held are listed) and click **Export dividends (Quadro 8A)**. After a summary, one file is downloaded: `AnexoJ_Dividends_<year>.csv`. The stock export above still makes exactly its two files.

```text
dividendos.<id>.pagamentos (script: Yahoo, last 2 years, [ex-date, $ per share])
        │                       carteira()  (the page's only FIFO)
        ▼                           │
one row per payment in the year ◄───┘ shares entitled = purchases before the ex-date
        │                               − what sales before the ex-date used up
        ├── gross $ = shares × $ per share ; US tax withheld (est.) = gross × 15 %
        ├── € reference = $ ÷ EUR/USD of the ex-date (fxRef / fxAt; blank if no rate)
        └── ANEXO_J.q8a: Quadro 8A, Código E11, País da fonte 840 ──► AnexoJ_Dividends_<year>.csv
```

- **Shares entitled.** A purchase on the ex-date does not receive the dividend. A sale on or after the ex-date still does. Quantities are in today's shares, like the dividend amounts from Yahoo (both split-adjusted).
- **Year.** By ex-dividend date: Yahoo gives the ex-date only, while the tax year follows the payment date. Ex-dates after today are not exported.
- **Amounts.** Gross amount in dollars, as in the source. Euros are **only a reference**, at the EUR/USD of the ex-date; the real amount is what your broker converted. The US tax withheld is an estimate at **15 %**, which assumes you filed form W-8BEN.
- **Missing data.** If the dividend source failed or the data is malformed, that company has no rows; nothing is invented. The summary and the message name it, with the reason. They also say which ex-dates the data covers (about 2 years back from the run), so an older year can be incomplete.

**Official reference.** Quadro 8A of the same *Modelo 3 – Anexo J* (verified in the form and its instructions): *rendimentos de capitais* obtained abroad, with the codes of **Tabela V**. **E10** = *dividendos ou lucros com retenção em Portugal*; **E11** = the same *sem retenção em Portugal*. The columns are: Código rendim. · País da fonte · Rendimento bruto · Imposto pago no estrangeiro (No país da fonte) · two columns for the paying agent's country (only for code E23, not used here) · Imposto retido em Portugal (NIF da entidade retentora, Retenção na fonte).

> **Code used: E11.** The file assumes that no tax was withheld in Portugal, which is usual with a foreign broker. If a Portuguese entity withheld tax, use **E10** and fill in the *Imposto retido em Portugal* columns. The code, the rates and the 45 days are in `ANEXO_J.q8a` in the template.

**Columns.** Tax year · Row · Anexo J table · Asset · Ticker · ISIN · Código rendim. · País da fonte · Ex-dividend date · Dividend per share (USD, 6 decimals) · Shares entitled (8 decimals) · Gross amount (USD) · Tax withheld in the US, estimated (USD) · EUR/USD at the ex-date · FX date · FX source · Rendimento bruto (EUR, reference) · Imposto pago no estrangeiro – No país da fonte (EUR, reference) · Imposto retido em Portugal – NIF da entidade retentora · Imposto retido em Portugal – Retenção na fonte (EUR) · Dividend data source · Status · Flags · Fill in manually. The CSV format is the same as above (UTF-8 with BOM, `;`, decimal comma, CRLF, formula protection, deterministic).

| Flag | Meaning |
|---|---|
| `BROKER_FX` | Always: the real euro amounts are the ones your broker converted. |
| `PAY_DATE_UNKNOWN` | The ex-date is in the last 45 days of the year (17 Nov–31 Dec in a non-leap year): the payment may fall in the next year, and the tax year follows the payment date. Status `REVIEW`. |
| `NO_FX_REFERENCE` | Informational: no EUR/USD for the ex-date, so the euro columns are blank. |
| `SPLIT_CHECK` | A purchase used for the count has a price history that changed without a known split. Status `REVIEW`. |
| `NO_ANEXO_J_MAPPING` | The company has no line in `ANEXO_J`. Status `REVIEW`. |

### Where the data lives

Purchases and sales are stored in the browser's `localStorage`:

| Key | Contents |
|---|---|
| `bb.buys` | Stock and ETF purchases `{id, a, d, q, p, r, u}`. `r` is the reference close of the date and `u` the date of that price data, used for future splits. |
| `bb.lots` | Bitcoin purchases `{id, d, q, c}`. |
| `bb.sales` | Sales `{id, a, d, q, p, r, u}`. |
| `bb.deleted` | Ids of deleted entries, so a backup file does not bring them back. |
| `bb.savedAt` | Time of the last change. |
| `bb.fileSaved`, `bb.fileWrittenAt` | Last write to the backup file. |
| `bb.sim.*` | Simulator inputs. |
| `bb.strat.*` | The "buy the dip" percentage of the strategy comparison (`{dip}`), per tab. Not part of the backup, like `bb.sim.*`. |
| `bb.policy` | Investment policy `{horizon, allocation, monthly, drop20, drop30, sell, at}`: free text, up to 2,000 characters each, and the time it was saved. |
| `bb.fees` | Fees per entry `{<entry id>: {v, at}}`: euros (0 or more) and the time they were saved. |
| `bb.notes` | Notes per entry `{<entry id>: {t, at}}`: text up to 1,000 characters and the time it was written. Empty text = note removed. |
| `bb.targets` | Target allocation `{weights: {id: %}, band, monthly, at}`. `weights` holds only the assets with a target above 0, adding up to 100; `band` is in percentage points; `at` is when it was saved. |

Field letters: `a` asset id, `d` date, `q` quantity, `p` unit price in €, `c` total cost in € (Bitcoin), `r` reference close, `u` date of the price data.

Nothing is sent over the internet. The dividend projection, the return and the tax export only read these keys. Since version 5 the backup has optional keys beside the entries: `targets`, `policy`, `notes` and `fees`. The entries themselves did not change, so backups of every version keep loading unchanged.

When two copies are merged (the project file on opening the page, or **Restore backup**), entries are combined as described below. For the target allocation and the investment policy, the copy saved most recently (`at`) wins; for notes and fees, each one keeps its most recent version. If the browser has the newer copy, the page asks for the backup file to be updated. Invalid targets, policy or notes in a file are ignored: an unknown asset, weights not adding up to 100 %, or values that are not numbers or text. Notes of deleted entries are dropped.

### Backup and automatic loading

```text
          ┌───────────────────────────── you add / sell / delete in the website ─────────────────────────────┐
          ▼                                                                                                   │
  browser localStorage (bb.buys, bb.lots, bb.sales, bb.deleted)                                              │
          │  "Save to project folder" (then automatic while the page is open)                                 │
          ▼                                                                                                   │
  bluechip-board-backup.json  ◄── or, if the browser cannot write to folders: downloaded to Downloads,       │
          │                       and the next run moves the newest valid one into the project folder        │
          │  next run of the script                                                                           │
          ▼                                                                                                   │
  embedded in bluechip-board.html ──► page opens ──► MERGED with localStorage, entry by entry ───────────────┘
                                                     (missing entries added, deleted ones removed, nothing lost)
```

1. **Portfolio → Backup → Save to project folder.** The first time, choose this project folder (where `bluechip-board.html` is) and keep the name `bluechip-board-backup.json`. Chrome remembers the folder.
2. **While the page is open,** every later change is saved to that file automatically. When the page is reopened, the browser asks for permission again: a box under the Backup buttons warns you whenever some changes are not in the file yet. Click the button once to save them.
3. **On every run,** the script embeds `bluechip-board-backup.json` into the website (only into `bluechip-board.html`, not the Archive copies or the data file). When the page opens, the file is **merged** with what the browser has, entry by entry:
   - entries missing on either side are added;
   - entries deleted on either side are removed;
   - nothing else is lost.

   Your purchases come back after clearing the browser and on another browser or profile, and changes made in two browsers are combined. If a file you saved is not found in the project folder at the next run, the page says so.
4. **If the browser cannot write to folders,** the backup is **downloaded**. On its next run, the script brings the newest valid `bluechip-board-backup*.json` from your Downloads folder into the project folder:
   - it is checked before anything is replaced, and must be newer by the date saved inside it;
   - the previous file is kept as `bluechip-board-backup.previous.json`;
   - an invalid file is left where it is, with a warning.
5. **Restore backup** merges any backup file by hand, after a confirmation. It also accepts backups from older versions.

---

# Part 2 · How it is built

## The system at a glance

```text
 Desktop shortcut "Bluechip Board.lnk"  (runs pwsh.exe, PowerShell 7)
        │
        ▼
 Start-BluechipBoard.ps1   launcher: passes -SecEmail (from bluechip-board.config.json), closes on success, pauses on errors/warnings
        │
        ▼
 Bluechip-Board.ps1 ◄────── about 70 public sources on the internet (news, prices, holdings, indicators)
   collect                ◄── bluechip-board-data.json      previous run: fallback when a price source fails
   classify               ◄── vistos.json                   stories already seen: the "New" tag
                          ◄── noticias-historico.json       material and important stories of the last 400 days
   build                  ◄── bluechip-board-backup.json    your purchases and sales
        │                  ◄── Downloads\bluechip-board-backup*.json   (only if newer and valid)
        │ writes
        ├──► bluechip-board.html        the website, WITH your backup embedded
        ├──► Archive\bluechip-board-<date>_<time>.html   a copy WITHOUT your backup (30 kept)
        ├──► bluechip-board-data.json   the data alone, WITHOUT your backup
        ├──► vistos.json
        └──► noticias-historico.json    written last
        │ opens
        ▼
 Browser (Chrome / Edge): bluechip-board.html
   reads the embedded JSON ─► draws every tab with JavaScript (no server, works offline)
   your entries ─► localStorage ─► "Save to project folder" ─► bluechip-board-backup.json ─► next run
```

Three ideas explain most of the design:

1. **One script, one page.** `Bluechip-Board.ps1` holds the configuration, the collectors, the classification and the whole website template (HTML, CSS and JavaScript in one here-string). The output is one HTML file with all data inside it.
2. **Nothing is invented.** Every source can fail on its own. A failed price falls back to the previous run's data, clearly labelled with its date; a failed holdings file falls back to reference weights; anything else is shown as unavailable.
3. **Your data stays local.** The script only reads public sources. Your portfolio lives in the browser and in the backup file, and only `bluechip-board.html` carries a copy of it.

---

## Project files

| File / folder | What it is |
|---|---|
| `Bluechip-Board.ps1` | The main script: configuration, data collection, news classification and the website template (HTML, CSS and JavaScript embedded in one here-string). |
| `Start-BluechipBoard.ps1` | Launcher used by the desktop shortcut. Passes `-SecEmail` with the e-mail of `bluechip-board.config.json` (or `BLUECHIP_SEC_EMAIL`), closes on success and pauses on errors or warnings. Without an e-mail it runs with the SEC sources skipped. |
| `bluechip-board.config.json` | **Local, not in git.** Your SEC contact e-mail: `{ "secEmail": "you@example.com" }`. Copy `bluechip-board.config.example.json` to create it. |
| `.gitignore`, `.gitattributes` | What stays out of git (your data and every output), and line endings (LF for the scripts, the byte-exact test CSV files untouched). |
| `.github\workflows\tests.yml` | GitHub Actions: the three test suites on PowerShell 7 and 5.1 for every push to `main` and every pull request. |
| `Bluechip-Board.ico` | Icon of the desktop shortcut (same design as the website logo). |
| `bluechip-board.html` | **The latest website.** Open this file. It includes your backup, so it loads your portfolio by itself. |
| `bluechip-board-data.json` | The data embedded in the latest website, **without** your backup. The next run also uses it as a fallback when a price source fails (see [How it works](#how-it-works)). The site tests build their test page from it. |
| `bluechip-board-backup.json` | Backup of your purchases and sales, saved from the website. The script embeds it into the site (see [backup](#your-data-portfolio-purchases-and-backup)). |
| `bluechip-board-backup.previous.json` | The previous backup, kept when a newer one is brought in from the Downloads folder. |
| `vistos.json` | Stories already seen in previous runs, used for the **New** tag (kept for 60 days). Written only when a run finishes. |
| `noticias-historico.json` | **News history:** every *material* and *important* story of the last **400 days** (title, link, source, date, assets, level, score, themes). It holds public headlines only, no personal data. It is written last, only when a run finishes. If the file is missing, the history starts empty, with a notice. If it cannot be read, it starts empty with a warning, and the old file is kept as `noticias-historico.json.bad`. |
| `Archive\` | A dated copy of every website built (`bluechip-board-YYYY-MM-DD_HHmm.html`), **without** your backup. The 30 most recent are kept. |
| `Tests\` | Automated tests (see [Tests](#tests)). `Tests\fixtures\taxBaseline\` holds the reference Anexo J CSV files for the byte-by-byte check. `Tests\fixtures\bluechip-board-data.sample.json` is a data file of 6 Oct 2026 (public data only), used by the tests when there is no local `bluechip-board-data.json` (in a fresh clone and in GitHub Actions). |
| `README.md` | This document. |
| `AUDIT-REMEDIATION.md` | Record of the technical audit of 4 Oct 2026, of every fix made after it, and a change log of later features. |

**Who reads and writes each file:**

| File | Written by | Read by |
|---|---|---|
| `bluechip-board.html` | the script, every run | you (the browser) |
| `bluechip-board-data.json` | the script, every run | the next run (fallback); `Test-Site.ps1` |
| `bluechip-board-backup.json` | the website ("Save to project folder"); the script when it brings a newer file from Downloads | the script (embeds it) |
| `bluechip-board-backup.previous.json` | the script | you, to recover by hand |
| `vistos.json` | the script, at the end of a run | the next run |
| `noticias-historico.json` | the script, at the very end of a run (after `vistos.json`) | the next run (the history also goes into the data, for the website) |
| `Archive\*.html` | the script | you, to look back |
| `Desktop\Bluechip Board.lnk` | set up once | you |

**Files you should never need to edit by hand:** all of the above except `Bluechip-Board.ps1`. Do not edit `bluechip-board.html` either: it is rebuilt on every run, so change the template in the script instead.

The desktop shortcut is `Desktop\Bluechip Board.lnk`. It runs **PowerShell 7**: `"C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -File "...\Start-BluechipBoard.ps1"`. To go back to Windows PowerShell 5.1, set the shortcut's target to `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`. The arguments stay the same, and both versions give the same result.

---

## Requirements and dependencies

### On your PC

| Dependency | Used for | Without it |
|---|---|---|
| **Windows** | Everything (the script uses Windows-only pieces: the Downloads folder via `Shell.Application`, Task Scheduler, a named mutex). | Does not run. |
| **PowerShell 7** (`pwsh.exe`) or **Windows PowerShell 5.1** (built in) | Running the script. The shortcut uses 7; the tests run on both. In 7 a run takes a few seconds less. PowerShell 7 checks the format of the User-Agent header and would reject the e-mail the SEC requires, so the script skips that check there (`-SkipHeaderValidation`); the header sent is the same. | Use the other version. |
| **.NET** (comes with PowerShell) | Downloads, XML, time zones, file writes. | – |
| **ScheduledTasks** module (built into Windows) | Only `-ScheduleDaily`. | Scheduling fails; normal runs are unaffected. |
| **Google Chrome or Microsoft Edge** | Viewing the website. Saving the backup straight into the project folder uses the File System Access API, which these browsers support. The site tests drive Chrome or Edge headless. | Other browsers show the site, but the backup is downloaded instead of saved, and `Test-Site.ps1` cannot run. |
| **An internet connection** | Collecting data. | Every source fails; the site is still built from the previous run's prices, labelled as such. |
| **UTF-8 with BOM** encoding of the script | Windows PowerShell 5.1 misreads accented characters otherwise. | Broken text, possibly errors. |

No modules to install, no packages and no API keys.

### On the internet

Every source is independent: when one fails, it is listed with its error in *Sources & method* and the rest of the site still works.

| Service | Host | What it gives | If it fails |
|---|---|---|---|
| Google News RSS | `news.google.com` | Most news (EN and PT searches) | Fewer stories |
| Yahoo Finance RSS | `feeds.finance.yahoo.com` | Headlines per symbol | Fewer stories |
| Company newsrooms, Fed, ECB, crypto press | `apple.com`, `nvidianews.nvidia.com`, `blog.google`, `federalreserve.gov`, `ecb.europa.eu`, `coindesk.com`, `theblock.co`, `cointelegraph.com`, `decrypt.co` | Primary-source and crypto news | Fewer stories |
| SEC EDGAR | `sec.gov`, `data.sec.gov` | Company filings and past earnings dates (only with `-SecEmail`) | No filings in the news; past earnings dates from the previous run (≤ 180 days), else Unavailable |
| Yahoo Finance chart API (unofficial) | `query1.finance.yahoo.com` | Prices (1 year and since 2010), EUR/USD, splits, dividends | Kraken (BTC), then the previous run's data; dividends: previous run up to 180 days |
| Kraken | `api.kraken.com` | BTC/EUR fallback | Previous run's data |
| Stooq | `stooq.com` | Stock fallback (currently blocks automated requests) | Previous run's data |
| ECB data | `ecb.europa.eu` | EUR/USD reference rates (90 days) | Yahoo EUR/USD only |
| iShares / BlackRock | `blackrock.com` | Holdings file of each ETF | That fund's reference weights |
| CoinGecko | `api.coingecko.com` | Bitcoin price, 24 h change, market cap, volume, dominance | Those figures are not shown |
| Alternative.me | `api.alternative.me` | Fear & Greed index | Not shown |
| mempool.space | `mempool.space` | Block height, difficulty, hash rate, halving | Halving estimated from the last known one |
| Nasdaq | `api.nasdaq.com` | Next earnings dates | The manual calendar only |
| Google Fonts | `fonts.googleapis.com` | The website's fonts (loaded by the browser) | System fonts |

### Inside the project

```text
Start-BluechipBoard.ps1 ──calls──► Bluechip-Board.ps1
                                      │ 1. CONFIGURATION  ($Ativos, $ETFs, $Feeds, $Calendario, $Bolsas, rules)
                                      │        ▲ used by every function below
                                      │ 2. HELPER FUNCTIONS  (Get-Url, Read-Feed, Measure-Noticia, Get-Serie, …)
                                      │ 3. TEMPLATE  Get-Plantilla  (HTML + CSS + JavaScript; data placeholder __DADOS_JSON__)
                                      │ 4. EXECUTION  (calls 2, fills 3, writes the files)
                                      ▼
Tests\Test-Engine.ps1     loads sections 1–2 without running 4 (cut between the "# 1. CONFIGURA…" and "# 3. MODELO DO SITE" markers)
Tests\Test-Site.ps1       extracts Get-Plantilla + bluechip-board-data.json ─► page ─► site-scenarios.js in headless Chrome
Tests\Test-Resilience.ps1 copies the whole script to %TEMP%, patches a few exact lines (see "Making a change safely"), runs it offline
```

---

## How it works

### One run, step by step

```text
start
  │
  ├─ -ScheduleDaily? ── yes ──► register the task "BluechipBoard" ──► end
  │ no
  ├─ another run in progress? (mutex Local\BluechipBoard) ── yes ──► warning ──► end
  │ no
  ▼
  1. News feeds (32 RSS/Atom feeds) ......................... Read-Feed
  2. SEC filings (3 companies, only with -SecEmail) ......... Read-Feed -Tipo sec
  3. Classify every story, merge identical titles ........... Measure-Noticia
     group duplicates, tag "New" (vistos.json) .............. Join-NoticiasDuplicadas
     add material/important stories to the news history ..... Merge-HistoricoNoticias (400 days)
  4. 1-year prices: 8 assets + S&P 500, VIX, 10-year yield .. Get-Serie ──fail──► previous run's data
  5. EUR/USD (Yahoo) and ECB reference rates ................ Get-Serie, Get-TaxaBCE
  6. ETF holdings, one file per fund ........................ Get-PesosETF ──fail──► reference weights
  7. Dividends (AAPL, NVDA, GOOGL) .......................... Get-Dividendos ──fail──► previous run (≤ 180 days)
  8. Bitcoin indicators ..................................... Get-DadosBitcoin
  9. Calendar: manual events + halving + Nasdaq earnings .... Get-DatasResultados
 10. Long-term histories since 2010 (+ EUR/USD, splits) ..... Get-Serie -Desde ──fail──► previous run's data
     Past earnings dates (SEC, only with -SecEmail) ......... Get-ResultadosSEC ──fail──► previous run (≤ 180 days)
     Fundamentals (SEC XBRL, only after a new 10-Q/10-K) .... Get-FundamentaisEmpresa ──fail──► previous run (≤ 120 days)
     Euro area indicators (ECB Data Portal) ................. Get-SerieMacro ──fail──► previous run (≤ 30 days)
 11. Backup: bring a newer valid file from Downloads, read bluechip-board-backup.json
 12. Maintenance reminders (old reference weights, empty calendar, no Fed/ECB decision more than 60 days ahead, old holiday exceptions, top 10 holdings without a news alias)
 13. Build: all data ─► JSON ─► "<" escaped ─► placed into the template (Get-Plantilla)
 14. Write, each atomically: Archive copy ─► bluechip-board.html ─► bluechip-board-data.json ─► vistos.json ─► noticias-historico.json (last)
 15. Console summary; open the website (unless -NoOpen)
```

If anything throws between steps 1 and 13, nothing is written: the previous website, `vistos.json` and `noticias-historico.json` stay as they were.

### Price fallbacks

```text
Yahoo Finance chart API ──ok──► price dated "close 2 Oct" or "live 15:36"
        │ fail (or wrong symbol / currency / invalid prices)
        ▼
Kraken (BTC only) · Stooq (stocks; currently blocked) ──ok──► "via Kraken"
        │ fail
        ▼
previous run (bluechip-board-data.json) ──found──► "previous run", orange date, alert
        │ none
        ▼
no price: the card says so, the portfolio shows "no price". Nothing is invented.
```

EUR/USD has its own chain, because the euro values of the US stocks depend on it:

```text
Yahoo EURUSD=X, 1 year ─► Yahoo EURUSD=X long-term history ─► ECB reference rates (90 days) ─► none
                                                               (older euro figures left blank)   (euro values of US
                                                                                                  stocks left blank,
                                                                                                  red alert)
```

### Key behaviours

- **One self-contained page.** All data is embedded as JSON inside a `<script id="dados" type="application/json">` tag, and the CSS and JavaScript are inline. The page works offline once built. Only the Google Fonts load from the internet, with system-font fallbacks. Any `<` inside the JSON is escaped as `\u003c`, so a headline can never break the page.
- **Resilient downloads.** Every request is retried once (25 s timeout; 8 s wait after a "too many requests" answer). Responses are decoded with the charset they declare, and XML is parsed without DTDs or external entities. Each source is independent: if one fails, it is listed with its error in *Sources & method*, and the rest of the site still works.
- **Validated data.** A price response is rejected if it is for another symbol or another currency than expected. Zero, negative and missing prices are dropped. ETF weights must add up to about 100 %. Earnings dates must fall between yesterday and about six months ahead. Feed items without a valid date are skipped and counted.
- **Fallbacks.**
  - If Yahoo fails, Bitcoin prices come from **Kraken** (BTC/EUR). Stooq, the stock fallback, currently blocks automated requests.
  - Otherwise the **previous run's prices** are used, marked "previous run" with their date.
  - If the iShares download of a fund fails, that ETF's weights fall back to its own reference values, and the site says so. Each ETF is downloaded and checked separately, so one failing fund does not affect the others.
  - If mempool.space fails, the halving date is estimated from the last known halving.
  - If the dividend data fails, the previous run's dividend data is used if it was retrieved less than 180 days ago, marked "previous run (retrieved …)". Otherwise the Dividends section shows "Unavailable". Dividend values are never invented, and a failure there does not affect anything else.
  - If the 1-year EUR/USD series fails, conversions use the long-term EUR/USD history, and only as a last resort the ECB rates (90 days). With ECB rates, euro figures that need more than 90 days are left blank, and with no rate at all US stock values in euros are left blank. Dollars are never shown as euros.
- **Old data never passes for current.** Every price shows the date it is from ("close 2 Oct", or "live 15:36" for an intraday price), and an alert appears when a series is behind its exchange's last session.
- **Safe writes.** The site and data files are written to a temporary file first and then swapped in (`Write-Atomico`), so an interrupted run never leaves a half-written file. A run that fails halfway leaves the previous site and `vistos.json` untouched.
- **Single instance.** A named mutex (`Local\BluechipBoard`) prevents two runs at once, for example the shortcut and a scheduled task. A second run prints a warning and exits.
- **Dates in the right time zone.** Each daily price is dated in its exchange's time zone (New York, Xetra, London for FX, UTC for Bitcoin). Everything the website shows is in **Lisbon time**, whatever the computer's time zone. Calendar countdowns count Lisbon calendar days.
- **English console.** The script sets the UI culture to `en-US`, so its own messages and Windows/.NET errors appear in English.

---

## Script parameters

Each parameter has an English name. The original Portuguese name still works as an alias.

| Parameter | Alias | Default | Description |
|---|---|---|---|
| `-Days` | `-Dias` | `7` | News window in days (1–30). Older stories are ignored. |
| `-Folder` | `-Pasta` | script folder | Where the website, data, archive and `vistos.json` are written. |
| `-SecEmail` | `-EmailSEC` | *(empty)* | Your e-mail, sent in the User-Agent header to the SEC, as SEC access rules require. Without it, SEC EDGAR is skipped. |
| `-NoOpen` | `-NaoAbrir` | off | Builds the site without opening the browser. |
| `-ScheduleDaily` | `-AgendarDiariamente` | off | Registers the Windows scheduled task `BluechipBoard` and exits (see [scheduling](#scheduling-a-daily-run)). |
| `-Time` | `-Hora` | `08:30` | Time of the daily scheduled run (`HH:mm`). |

Examples:

```powershell
.\Bluechip-Board.ps1 -SecEmail "your@email.com"                      # normal run
.\Bluechip-Board.ps1 -Days 3 -NoOpen                                  # shorter window, don't open the browser
.\Bluechip-Board.ps1 -ScheduleDaily -Time 07:45 -SecEmail "your@email.com"
Get-Help .\Bluechip-Board.ps1 -Full                                   # built-in help
```

---

## Script configuration (top of `Bluechip-Board.ps1`)

Section **1. CONFIGURAÇÃO** at the top of the script holds everything you are expected to edit. Code comments in the script are in Portuguese; names and values are described here and in the [glossary](#glossary-portuguese-names-in-the-code).

### Assets — `$Ativos`, `$Mercado`

```powershell
@{ Id = 'AAPL'; Nome = 'Apple'; Yahoo = 'AAPL'; Stooq = 'aapl.us'; Moeda = 'USD' }
@{ Id = 'BTC';  Nome = 'Bitcoin'; Yahoo = 'BTC-EUR'; Stooq = ''; Kraken = 'XBTEUR'; Moeda = 'EUR' }
```

| Field | Meaning |
|---|---|
| `Id` | Internal id used everywhere in the site (`AAPL`, `NVDA`, `GOOGL`, `SXR8`, `EUNK`, `IS3N`, `EUNN`, `BTC`). The ETF ids are their Xetra tickers. |
| `Nome` | Display name. |
| `Yahoo` | Yahoo Finance symbol (`SXR8.DE`, `BTC-EUR`, …). |
| `Stooq` / `Kraken` | Fallback symbols (empty = no fallback). |
| `Moeda` | Trading currency. USD assets are converted to euros with EUR/USD. |

`$Mercado` holds the market references shown in *Currency & Macroeconomics*: S&P 500 (`^GSPC`), VIX (`^VIX`) and the 10-year Treasury yield (`^TNX`).

`$AtivosDividendos` (`'AAPL', 'NVDA', 'GOOGL'`) lists the assets whose dividends are collected and projected. The ETFs are left out on purpose: they are accumulating, so their dividends are reinvested in their price. Any asset marked `Acumulacao = $true` in `$ETFs` is skipped even if it is added to this list, and the website applies the same rule. Bitcoin pays none.

### Adding an asset

```text
Another iShares ETF (data-driven)                    Another company (several places)
─────────────────────────────────                    ────────────────────────────────
script:  $Ativos  + $ETFs entry                       script:  $Ativos, $EmpresasRe, a feed,
         $EmpresasRe + a news feed                             $SecEmpresas, Get-DatasResultados
template: a colour token (--xxxx) + CO entry          template: CO map, filter chip, PF list,
          ANEXO_J line (tax export)                             BOLSA_DE, alert categories,
                                                                ANEXO_J line, colour token
everything else (card, chip, Prices, ETF view,
simulator, Portfolio, look-through, HIST,             then: run the tests, update this README
exchange) is derived from the data
```

Without a colour token an ETF gets a neutral colour. A sale of an asset without an `ANEXO_J` line is flagged `NO_ANEXO_J_MAPPING` in the tax export.

### ETF holdings and metadata — `$ETFs`, `$UrlPesosETF`, `$PesosReferencia`

Each ETF has one entry in `$ETFs`:

| Field | Meaning |
|---|---|
| `Id` | Same id as in `$Ativos`. |
| `Chip` | Label of its filter chip (SXR8 keeps "ETF & market"). |
| `Fundo`, `Isin` | Fund name and ISIN. |
| `Indice`, `Benchmark` | Short and full name of the index. |
| `MoedaBase` | The fund's base currency. This is not the trading currency, which is `Moeda` in `$Ativos` (EUR for all four on Xetra). |
| `Domicilio`, `Acumulacao`, `Gestora`, `Bolsa`, `Inicio` | Domicile, accumulating flag, fund manager, exchange (`$Bolsas` id) and inception date. |
| `Pagina` | The fund's iShares page. |
| `Url` | Its own holdings file. |
| `Referencia` | Its own fallback weights and date. |
| `Empresas` | The panel companies the fund must contain. |
| `PaisIndice` | Only for single-country indexes whose iShares file has no country breakdown (SXR8: United States; EUNN: Japan). Used as the country of the whole equity part, flagged as approximate. |

The four funds:

| Id | Fund | ISIN | Yahoo | Holdings file (iShares UK portfolioId) |
|---|---|---|---|---|
| `SXR8` | iShares Core S&P 500 UCITS ETF USD (Acc) | IE00B5BMR087 | `SXR8.DE` | 253743 (`$UrlPesosETF`) |
| `EUNK` | iShares Core MSCI Europe UCITS ETF EUR (Acc) | IE00B4K48X80 | `EUNK.DE` | 251861 |
| `IS3N` | iShares Core MSCI EM IMI UCITS ETF USD (Acc) | IE00BKM4GZ66 | `IS3N.DE` | 264659 |
| `EUNN` | iShares Core MSCI Japan IMI UCITS ETF USD (Acc) | IE00B4L5YX21 | `EUNN.DE` | 251867 |

**Listing choice.** All four use their **Xetra listing in euros**: the same exchange, currency and Yahoo pipeline, so no currency conversion and no extra exchange hours. The other listings, checked on Yahoo on 5 Oct 2026:
- in euros on Amsterdam and Milan: IMAE.AS, EMIM.AS, IJPA.AS, SMEA.MI, EIMI.MI;
- in dollars or pence in London: EIMI.L, IJPA.L, SJPA.L.

The Xetra tickers (`EUNK`, `IS3N`, `EUNN`) are also the internal ids, like `SXR8`. `EIMI` was not used, because it is the London dollar listing. The ISIN of each Yahoo listing and of each holdings file was checked against the fund's Key Facts.

- `$UrlPesosETF`: the iShares "Detailed Holdings and Analytics" download for SXR8 (Excel 2003 XML, read with regular expressions). `New-UrlPesosETF <portfolioId>` builds the same download for the other funds; their files have the same sheets and columns (checked on 5 Oct 2026). Each file gives the weights of Apple, NVIDIA and Alphabet (classes A + C combined), the top 10 holdings, the sector breakdown and the number of holdings.
- **Checks per file:**
  - the ISIN in the file's Key Facts must be the fund's;
  - all weights must add up to about 100 %;
  - in SXR8 the three companies must each weigh between 0 and 25 %; in the other funds they may be 0 % (they hold none of them), never negative or above 25 %.
  - Short US ticker aliases ("V" → Visa) are applied only to holdings traded in USD.
- `$PesosReferencia` (SXR8) and each `Referencia` (EUNK, IS3N, EUNN: 0 % for the three companies, from the files of 2 Oct 2026): fallback weights and date, used only if that fund's download fails. Then the top 10 and sectors show "Unavailable". Update them from time to time; a maintenance reminder appears when they are over 4 months old.

### Calendar — `$Calendario`

```powershell
@{ d = '2026-10-28'; e = 'GOOGL'; ev = 'Q3 2026 earnings'; imp = 'High'; st = 'E' }
```

| Field | Values |
|---|---|
| `d` | Date, `yyyy-MM-dd`. An event with a malformed date is skipped with a warning. |
| `e` | Who: `AAPL`, `NVDA`, `GOOGL`, `SXR8`, `EUNK`, `IS3N`, `EUNN`, `BTC`, `MKT` (market / Fed) or `TU` (you, e.g. tax deadlines). The ETFs get no automatic events; add one by hand only for a real date. |
| `ev` | Event text shown in the site (English). |
| `imp` | `Very high`, `High` or `Medium`. *High* and *Very high* events within 7 days raise an alert. |
| `st` | `C` = confirmed, `E` = estimated. |

Automatic additions on every run:

- **Earnings dates (Nasdaq):** the next date for AAPL, NVDA and GOOGL is fetched from Nasdaq.
  - If a manual event of the same company containing "earnings" is within 45 days, its date and status are replaced (the text is kept).
  - A manual date marked confirmed (`C`) is never replaced by a Nasdaq estimate.
  - Otherwise a "Quarterly earnings" event is added.
  - These events show a **Nasdaq** tag.
- **Bitcoin halving:** the next halving block is computed from the current block height. Its date is forecast from the average block time since the last halving, kept between 8 and 12 minutes. Past halvings come from a fixed list; any newer one is looked up on mempool.space.

### Exchange hours — `$Bolsas`

```powershell
@{ id = 'US'; curto = 'US'; nome = 'New York exchanges'; tz = 'America/New_York'; abre = '09:30'; fecha = '16:00'; feriados = @(...); curtos = @{ '2026-11-27' = '13:00' } }
```

Opening and closing times are in each exchange's local time; the site converts them to Lisbon time, including daylight-saving changes. The standard holidays and early closes are **computed by rules for any year**:

- **NYSE:** New Year, MLK Day, Presidents' Day, Good Friday, Memorial Day, Juneteenth, Independence Day, Labor Day, Thanksgiving and Christmas, with weekend observance. Early closes at 13:00 on the day after Thanksgiving, on 3 July and on 24 December when applicable.
- **Xetra:** New Year, Good Friday, Easter Monday, 1 May and 24, 25, 26 and 31 December. The last trading day of the year is shown as closing at 14:00, Deutsche Börse's usual practice. It confirms this each year by circular, so add an exception in `curtos` if a year differs.

Use `feriados` (extra closed days) and `curtos` (date → early closing time) only for exceptions, such as an extraordinary closure.

### News feeds — `$Feeds`, `$SecEmpresas`

```powershell
@{ Nome = 'Google News: NVIDIA (stock)'; Url = (New-GNewsUrl 'Nvidia (stock OR shares OR NVDA)'); Dica = 'NVDA' }
@{ Nome = 'CoinDesk'; Url = 'https://www.coindesk.com/arc/outboundfeeds/rss'; Dica = ''; Exigir = 'BTC' }
```

| Field | Meaning |
|---|---|
| `Nome` | Name shown in *Sources & method* and in the console. |
| `Url` | RSS or Atom URL. `New-GNewsUrl 'query' ['pt']` builds a Google News search limited to the news window (English by default, `pt` for Portugal/Portuguese). `New-YahooRss 'SYMBOL'` builds a Yahoo Finance headline feed. |
| `Dica` | Default asset for stories whose headline has no keyword (e.g. a company's own newsroom). In search feeds (Google News, Yahoo Finance), such stories are only a weak match: tagged *Feed match only*, no company bonus, at most moderate. Empty = keyword-only. |
| `Exigir` | Optional filter: the headline **must** match this asset's keywords, otherwise it is dropped. Used for crypto media, so other coins are not included. |

`$SecEmpresas` lists the SEC CIK numbers for AAPL, NVDA and GOOGL. Insider forms (3, 4, 5 and 144) are ignored. 8-K, 10-Q/10-K and share-issuance filings get extra weight.

### Classification rules

These regular expressions decide how each headline is scored (see [News classification](#news-classification)): `$EmpresasRe` (asset keywords), `$Temas` (themes and weights), `$Severo`, `$DecisaoJuros`, `$Ruido`, `$Positivo` / `$Negativo` (sentiment) and the source tiers `$TierPrimaria`, `$TierReferencia` and `$TierCuidado`. Edit them to tune the classification, and add a case to `Test-Engine.ps1` for every headline you fix.

### Top holdings of the ETFs — `$AliasesPosicoes`

Curated names for the **largest holdings of each ETF**, so news about them can be linked to the fund. The list covers the top 10 of each iShares file of 5 Oct 2026. Apple, NVIDIA and Alphabet are left out because `$EmpresasRe` already covers them. Each fund has a list of `@{ Nome; Tickers; Re }`:

| Key | Meaning |
|---|---|
| `Nome` | The holding's name, shown in the *Via top holding* tag. |
| `Tickers` | The tickers of the holding in the iShares file, used by the maintenance reminder. |
| `Re` | The regular expression, with word boundaries and exclusions for known false positives. Examples: the Amazon rainforest, Nikola Tesla, wet AMD (the eye disease), *ASM International*, *La Roche-Posay*, a *shell company*, *Siemens Energy*, *Samsung Heavy*, *Toyota Tsusho*, *SoftBank Corp*, and lowercase *sap* (only *SAP* counts) and Portuguese *meta* (only *Meta* counts). |

An alias counts only when the headline has **no keyword** from `$EmpresasRe` and comes from a feed **without its own asset** (`Dica` empty), such as Reuters on the three or the Portuguese feeds. Those stories used to be dropped. Feeds with a `Dica` keep their assignment, so stories that were already classified do not change. A match by alias is weak:

- the story goes to the fund (an SXR8 story is also tagged Market, as always);
- it is at most *moderate*;
- it gets no company bonus;
- it is tagged **Via top holding**, with the holding named.

**Maintenance reminder.** When a fund's current iShares file (this run's, not the reference weights) has a top 10 holding with no alias, by ticker or by name, the script lists it, for example *"SXR8 Berkshire Hathaway (BRKB)"*. Add an alias, together with a known false positive in `Test-Engine.ps1`. Without a current file there is no top 10, and no reminder.

---

## Data sources

| Data | Source | Notes |
|---|---|---|
| News | Google News RSS (EN + PT), Yahoo Finance RSS, Apple Newsroom, NVIDIA Newsroom, Google Blog, Federal Reserve, ECB, CoinDesk, The Block, Cointelegraph, Decrypt | Window set by `-Days`. One Google News feed per regional ETF. |
| Company filings | SEC EDGAR (Atom) | Only with `-SecEmail`. 400 ms between requests. |
| Euro area indicators | ECB Data Portal `data-api.ecb.europa.eu/service/data/<flow>/<key>?format=csvdata&detail=dataonly`, no key | Deposit facility rate `FM/B.U2.EUR.4F.KR.DFR.LEV` and HICP annual rate `ICP/M.U2.N.000000.4.ANR`, keys checked on 6 Oct 2026. Fallback: the previous run (up to 30 days old), labelled. See [Euro area indicators](#euro-area-indicators-ecb). |
| Fundamentals | SEC XBRL `data.sec.gov/api/xbrl/companyfacts/CIK##########.json` (3–4 MB per company) | Only with `-SecEmail`, the same User-Agent and 400 ms between requests, and **only when the SEC shows a 10-Q or 10-K newer than the last collection**: otherwise the stored facts are reused ("ok (cached)", no request). See [Fundamentals](#fundamentals-sec-xbrl). Fallback: the previous run's facts (up to 120 days old), labelled. |
| Past earnings dates | SEC: `data.sec.gov/submissions/CIK##########.json` (the filings list) and the EDGAR 8-K Atom feed (acceptance times) | Only with `-SecEmail`, the same User-Agent and 400 ms between requests: two requests per company. See [Past earnings dates](#past-earnings-dates-sec). Fallback: the previous run's data (up to 180 days old), labelled. |
| Prices, 1 year (daily) | Yahoo Finance chart API (unofficial) | Fallbacks: Kraken for BTC, Stooq for stocks (Stooq currently blocks automated requests), then the previous run's data. |
| Long-term history (daily) | Yahoo Finance | AAPL, NVDA, GOOGL, SXR8, EUNK, EUNN and EUR/USD since 2010; IS3N since its listing in Jun 2014; Bitcoin since Sep 2014. Prices are **split-adjusted**; the split events come with them. Fallback: the previous run's data. |
| EUR/USD | Yahoo (`EURUSD=X`) + ECB reference rate (90 days) | See the currency fallbacks above. |
| Dividends (AAPL, NVDA, GOOGL) | Yahoo Finance chart API, dividend events (`range=2y&interval=1mo&events=div`), the same public endpoint as the prices | One request per company, independent of the prices. Yahoo's `quoteSummary` and `quote` endpoints need authentication since 2023, so they are not used. Fallback: the previous run's dividend data (up to 180 days old). |
| ETF holdings | iShares / BlackRock holdings file, one per fund | Each falls back to its own reference weights (`$PesosReferencia` for SXR8, `Referencia` in `$ETFs` for the others). Listed per fund in *Sources & method* as "iShares: ETF holdings (ID)". |
| ETF prices | Yahoo Finance (`SXR8.DE`, `EUNK.DE`, `IS3N.DE`, `EUNN.DE`), Xetra, EUR | Same validation and fallbacks as the other prices. Stooq has no symbol for EUNK, IS3N and EUNN (it blocks automated requests anyway), so they fall back directly to the previous run's data. |
| Bitcoin market | CoinGecko (price, 24 h change, market cap, volume, dominance) | |
| Bitcoin sentiment | Alternative.me Fear & Greed Index (365 days) | |
| Bitcoin network | mempool.space (block height, difficulty adjustment, hash rate) | |
| Earnings dates | Nasdaq (data by Zacks) | |

A full run reads about 70 sources and takes about one minute.

### Euro area indicators (ECB)

`$SeriesMacro` lists two series of the ECB Data Portal, read as CSV with no key:

| Id | Indicator | Series | Plausible range |
|---|---|---|---|
| `BCE_DFR` | Deposit facility rate (%), only the dates it changed | `FM/B.U2.EUR.4F.KR.DFR.LEV` | −2 to 20 |
| `HICP_EA` | Euro area inflation, HICP, annual rate (%), monthly | `ICP/M.U2.N.000000.4.ANR` | −5 to 25 |

**Validation** (`ConvertFrom-CsvSerie`):
- an HTML page (a server error) is rejected, never read as data;
- the header must have `KEY`, `TIME_PERIOD` and `OBS_VALUE`, and every row must be the requested series;
- `.`, `NaN` and empty values are missing, never 0;
- values outside the plausible range are ignored and counted;
- a series without values is an error.

If a series fails, the previous run's data is used when at most 30 days old, labelled. Otherwise the chart says **Unavailable** with the reason.

On 6 Oct 2026 the portal's HICP for the euro area (also as I8 or I9) ended in **December 2025**. The site shows the month next to the value and marks it as not recent after 75 days; it never presents it as current.

**Not shown (pending, see the task log).**
- **US yield curve, high-yield spread, financial conditions.** The FRED series `T10Y3M`, `BAMLH0A0HYM2` and `NFCI` were specified, but FRED (Akamai) resets or ignores this script's requests. It only answers clients that identify as known tools such as curl, and making the script pose as one would mean getting around a bot filter.
- **S&P 500 CAPE.** No public CSV source with current data was found: the GitHub/datahub `s-and-p-500` dataset has PE10 only until September 2023.

The tab says so, with the context the notes were meant to give: an inverted curve preceded US recessions with long and variable lags and has given false alarms, and a high valuation has gone with lower returns over 10 years, not over the next month.

**Calendar.** `$Calendario` has the remaining Fed and ECB decisions of 2026 and those of 2027, from the official pages:
- [federalreserve.gov](https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm) and the [ECB](https://www.ecb.europa.eu/press/calendars/mgcgc/html/index.en.html), read on 6 Oct 2026;
- status `C`, asset `MKT`, on the decision day;
- the Fed decisions are *High* importance, like the existing ones. The ECB decisions are *Medium*, so they add no alert to the Overview (*High* events raise an alert 7 days before).

**Maintenance reminder** (`Get-LembreteReunioes`): when the calendar has no Fed, or no ECB, decision more than 60 days ahead.
### Fundamentals (SEC XBRL)

Only with `-SecEmail`. One `companyfacts` request per company, and only when the SEC shows a 10-Q or 10-K newer than the one of the last collection (from the [past earnings dates](#past-earnings-dates-sec)). The JSON is 3–4 MB and Windows PowerShell 5.1 reads it in about a second (measured on 6 Oct 2026). A first collection adds about 10 s to a run; the next runs reuse the facts.

**Metrics** (`$MetricasSEC`), each with its unit and its us-gaap tags in order of preference:

| Key | Metric | Unit | Tags |
|---|---|---|---|
| `receita` | Revenue | USD | `RevenueFromContractWithCustomerExcludingAssessedTax`, `Revenues`, `SalesRevenueNet`, … |
| `lucroBruto` | Gross profit | USD | `GrossProfit` (Alphabet does not report it: Unavailable) |
| `lucroOperacional` | Operating income | USD | `OperatingIncomeLoss` |
| `cfo` | Operating cash flow | USD | `NetCashProvidedByUsedInOperatingActivities`, … |
| `capex` | Capital expenditure | USD | `PaymentsToAcquirePropertyPlantAndEquipment`, `PaymentsToAcquireProductiveAssets` |
| `eps` | Diluted EPS | USD/shares | `EarningsPerShareDiluted` |
| `acoes` | Diluted average shares | shares | `WeightedAverageNumberOfDilutedSharesOutstanding` (Alphabet only since 2022) |

FCF = operating cash flow − capex.

**Rules.**
- **Tags.** Companies change tags over the years, so each period takes the most preferred tag that has it. The tags actually used are stored (`tags`).
- **Durations and forms.** Only facts of about 90, 180, 270 or 365 days, from 10-Q and 10-K filings, are accepted.
- **Restatements.** For each period, the value of the most recent filing wins (`v`, filed `f`). The first published value and date are also kept (`v0` only if different, `p`), because last year's figures are filed again as comparatives a year later.
- **Quarters.** A quarter is the ~90-day fact. Otherwise it is the difference of two year-to-date figures with the same start, marked `d`:
  - the 4th quarter = the year (10-K) − 9 months;
  - cash flow is only reported year to date.
  - Diluted shares are never derived by subtraction.

  A missing quarter stays empty (Unavailable); nothing is interpolated.
- **Fiscal years** come from the real period dates. Apple ends in September and NVIDIA in late January: NVIDIA's year from 27 Jan 2025 to 25 Jan 2026 is FY2026.
- **Splits.** EPS is divided, and shares multiplied, by the splits **after each value's filing date**. Filings after a split already carry the adjusted figures (NVIDIA restated 5.98 as 0.60 after its 10:1 split), so adjusting by the period would count the split twice.
- **TTM** is the sum of the latest 4 consecutive quarters, all present. For shares, it is the latest quarter.
- **Kept:** the quarters of the last 11 years, about 25 KB per company.
### Past earnings dates (SEC)

Only with `-SecEmail`. For each company (`$SecEmpresas`), the script keeps:

- every **8-K with item 2.02** (*Results of Operations and Financial Condition*, the earnings release) of the last 5 years, from the `recent` list of the submissions JSON;
- the **latest 10-Q or 10-K** (form, filing date and period), for a later task.

```text
data.sec.gov/submissions/CIK##########.json ──► 8-K with item 2.02 (accession number, filing date), latest 10-Q/10-K
        400 ms
EDGAR 8-K Atom feed (40 latest 8-K) ──► acceptance time with its time zone, matched by accession number
        │
        ▼
New York time (US clock changes included) ──► after 16:00: the next session
                                              before 09:30: that same session
                                              09:30–16:00: that same session, with a note
```

- **Why two requests.** The `acceptanceDateTime` field of the submissions JSON is not reliable. For Apple it is several hours off the *Accepted* time on the EDGAR filing page (16:30 New York shows as 00:30Z the next day); for NVIDIA and Alphabet it is right. The Atom feed gives the same time as the filing page, with its offset (for example `2026-07-30T16:30:28-04:00`). Verified on 6 Oct 2026.
- **Missing times.** An 8-K older than the feed (about 4 years) has no verified time, so its time and session are left empty (Unavailable), never guessed. Times already verified in the previous run are reused by accession number.
- **Sessions.** They are the dates of the company's long-term price history. A day the exchange was closed is therefore skipped, even an unplanned one. After the end of the history, the script uses weekdays without the `feriados` of `$Bolsas`.
- **Fallback.** If the SEC fails, or without `-SecEmail`, the previous run's data is used if it is at most 180 days old, labelled "previous run (retrieved …)". Otherwise the data is marked `error` or `skipped`.

---

## News classification

```text
feed item ──► valid date inside the news window? ── no ──► skipped (counted in Sources)
                 │ yes
                 ▼
         asset keywords ($EmpresasRe) ── none ──► feed's Dica? ── none ──► top-holding alias ($AliasesPosicoes)? ── none ──► dropped
                 │ found                              │ yes: assigned to that asset     │ yes: assigned to the fund
                 │                                    │ (from a search feed: weak match, │ (weak match, at most moderate,
                 │                                    │  at most moderate)               │  "Via top holding")
                 ▼                                    ▼                                  ▼
         Exigir filter (crypto feeds): headline must be about that asset ── no ──► dropped
                 ▼
         score = themes (≤ 5) + adverse event (+3) + rate decision (+2) + source tier (+1.5 / +1 / −1.5)
                 + your asset (+1) − click-bait (−3)
                 ▼
         level: ≥ 7 material (red) · ≥ 4 important (orange) · ≥ 2 moderate (yellow) · else noise (white)
                 ▼
         same normalised title ──► kept once;  similar titles (≤ 3 days, same asset, same numbers) ──► one group
                 ▼
         "New" if none of the group's headlines is in vistos.json
```

Each headline goes through these steps:

1. **Assets.** The headline is matched against `$EmpresasRe` (e.g. *iPhone* → Apple, *Blackwell* → NVIDIA, *Gemini* → Alphabet, *spot ETF* / *halving* → Bitcoin, *Fed* / *S&P 500* → Market). Some look-alikes are excluded:
   - *Big Apple* is not Apple.
   - *iShares Bitcoin Trust*, and the iShares MSCI funds, are not SXR8.
   - The mineral *mica* is not the EU's *MiCA* rules.
   - *"…, Nikkei reports"* (the newspaper) is not the Japan ETF.

   **The regional ETFs** (EUNK, IS3N, EUNN) match their own fund (ISIN, Xetra and other tickers, index name) and headlines about the **whole market their index covers**. That is the market that moves the ETF's price:
   - EUNK: *European stocks/shares*, *STOXX 600*;
   - IS3N: *emerging-market stocks*, *MSCI Emerging Markets*;
   - EUNN: *Japanese stocks*, *Nikkei 225*, *Topix*.

   A single company in that market, or a country's economy (GDP, rates), is not enough, so the ETF's news stays about its market. The exception is one of the fund's largest holdings, as a weak match only (see the next paragraphs). Such headlines usually also match Market & macro. A story about SXR8 is always also tagged Market.

   Each regional ETF has its own Google News feed, with that ETF as its `Dica`. Without a keyword, a story from the feed is only a weak match (at most moderate). Like SXR8, the ETFs get no company bonus, so a broad market move is at most *important* unless there is a severe event.

   Without a match, the feed's `Dica` is used; from a search feed that is only a weak match (see `Dica` above). Without a `Dica`, a **top-holding alias** can link the story to its fund: *"Microsoft beats estimates"* goes to SXR8, *"TSMC raises prices"* to IS3N, tagged *Via top holding*, at most moderate and with no company bonus (see [`$AliasesPosicoes`](#top-holdings-of-the-etfs--aliasesposicoes)). Without any of these, the story is dropped.
2. **Score.**
   - Themes: the strongest counts in full and each extra one at half, up to 5 points:
     - Earnings, Regulation and courts, China and exports, Security and fraud: **3** each.
     - AI and CapEx, Capital and shareholders, Management, Macro and rates, ETF flows, Institutional adoption: **2**.
     - Products and technology, Currency, Network and mining: **1**.
     - Analysts: **0.5**.
     - SEC filings add their own theme: 8-K, 10-Q/10-K and share issuance **3**, other forms **1**.
   - Adverse events: **+3**. Examples: a ban, export curbs, a guidance cut, a plunge, a delayed launch, an arrest, smuggling or a hack. Words that need context only count with it: "delays" alone, "to curb inflation" and "X warns Y" do not.
   - A rate decision: **+2**.
   - Primary source **+1.5**, leading press **+1**, "handle with care" sites **−1.5**.
   - A story about one of your assets (AAPL, NVDA, GOOGL or BTC): **+1** (not for weak feed matches).
   - Click-bait ("stocks to buy now", "price prediction"…): **−3**, and classed as noise (unless from a primary source).
3. **Level.** **Material** ≥ 7 · **Important** ≥ 4 · **Moderate** ≥ 2 · **Noise** below that. Weak feed matches and top-holding matches are at most moderate. In the code and data the levels are colours: `red`, `orange`, `yellow`, `white`.
4. **Sentiment.** A count of positive and negative words: positive, negative, mixed or neutral. It is only a hint.
5. **Duplicates.** The same event told by several sources is merged into one story with "Also reported by…". Titles are compared with an IDF-weighted word overlap. Rare words (names, amounts such as "$300M") count more, and synonyms and amounts are normalised. Two stories are grouped only if:
   - they were published at most 3 days apart, and a whole group never spans more than 3 days;
   - they share an asset;
   - they have no conflicting numbers.

   The story with the highest score is kept.
6. **New.** A story (or group) is tagged **New** if none of its headlines appeared in a previous run (`vistos.json`, kept for 60 days). On the very first run nothing is tagged New.

**Relevance to me (website).** *News → Sort by → Relevance to me* orders the stories by what you hold. The default order (*Potential impact*) and the levels do not change. The score and level still come from the script; the relevance only sorts:

```text
relevance = score + 3 × (share of your portfolio exposed to the story's asset)
```

- **Share exposed**, calculated in the browser from *Your holdings* (`carteira()`, the page's only FIFO), at today's values:
  - **Apple, NVIDIA, Alphabet:** what you hold directly, plus what you own through each ETF (the same look-through as *Real exposure to each company*);
  - **an ETF:** its value;
  - **Bitcoin:** its value;
  - **Market:** your stocks and ETFs (everything except Bitcoin).
- **A story with several assets** adds them up, to at most 1. Market counts only when it is the story's only asset, as in the tags.
- While this order is chosen, a note explains the formula and each story shows its relevance next to its points.
- **Without holdings** (or without prices to value them) the option is disabled, with a note saying why.

---

## How the website works

The page has no server and no framework. When it opens:

```text
<script id="dados">  JSON written by the script (+ "backup")
        │ JSON.parse  (a corrupt file shows a message instead of a blank page)
        ▼
D ─► prepared constants:  ATIVOS, MERC (price series) · FX, BCE (EUR/USD) · HIST (long histories)
                          NEWS · ETF_IDS · SPLITS · ANEXO_J · CO (colours) · PF (assets you can buy)
        │
        ├─► init():  merge the embedded backup with localStorage (juntaBackup), migrate old entries,
        │            build the forms, the filter chips and the footer
        ▼
renderAll()  ── runs again on every filter / currency / period change (state in st) ──┐
   renderHero (snapshot cards, cartaoFx)                                               │
   renderAlerts · renderNews · renderPrices · renderFx · renderCal                     │
   renderEtf + drawEtfMore · renderBtc · drawPf (portfolio, dividends, tax) · renderSrc│
        ▲                                                                              │
        └──────────── click on a chip, a tab, a period or a form ◄─────────────────────┘

your entries ─► store.set (localStorage "bb.*") ─► aoMudar ─► save to the backup file (if allowed)
```

- **State.** `st` holds the current filter, period, currency and news level (`S0` is the start value). Tabs are shown and hidden by `showTab`; the current tab is kept in the address (`#overview`, `#portfolio`, `#fundamentals`, …).
- **Prices in euros.** `inCur()` converts a USD series day by day with `fxAt()`, which returns no rate outside the series instead of guessing.
- **Freshness.** `frescura()` compares the last date of each series with its exchange's last completed session (rule-based holidays in `regras()`, `pascoa()`), which gives the "close" / "live" / orange labels and the stale alerts.
- **Portfolio.** `carteira()` is the only FIFO calculation. `efetiva()` applies any split adjustment (`ajusteSplit()`) to an entry before it is used anywhere.
- **Charts.** Hand-written SVG helpers: `lineMulti`, `barChart`, `hbars`, `spark`.
- **Tests.** A hook, `window.__BB_TEST__`, exposes the internal functions to `site-scenarios.js`. It does nothing in normal use.

---

## Code map

`Bluechip-Board.ps1` is about 3,000 lines in four numbered sections. Search for the section titles or function names; line numbers change with every edit.

| Section | What is there |
|---|---|
| Top | Help block, parameters, TLS setting, browser identity (`$Script:UA`), shared state (`$Script:Fontes`, `$Script:Agora`). |
| **1. CONFIGURAÇÃO** | Everything meant to be edited: `$Ativos`, `$Mercado`, `$PesosReferencia`, `$UrlPesosETF`, `$ETFs`, `$AtivosDividendos`, `$Calendario`, `$Bolsas`, `$Feeds`, `$SecEmpresas`, and the classification rules. |
| **2. FUNÇÕES DE APOIO** | The collectors and helpers, below. |
| **3. MODELO DO SITE** | `Get-Plantilla`: the whole website (HTML, then CSS, then JavaScript) with the placeholder `__DADOS_JSON__`. |
| **4. EXECUÇÃO** | The run itself, in the order shown in [How it works](#how-it-works), plus the scheduled-task registration, the backup handling and `Write-Atomico`. |

**PowerShell functions (section 2)**

| Function | Does |
|---|---|
| `Get-Url` | Every download: headers, retry, timeout, charset decoding. |
| `ConvertFrom-Bytes`, `ConvertTo-XmlSeguro`, `Get-Texto`, `ConvertTo-Data` | Decoding, safe XML, text and date parsing. |
| `Add-Fonte` | Records each source's status for *Sources & method*. |
| `Read-Feed` | Reads one RSS/Atom feed (or SEC Atom) into raw stories. |
| `Measure-Noticia` | Classifies one story: assets (keywords, the feed's `Dica`, or a top-holding alias), themes, score, level, sentiment, tier. |
| `Get-PosicoesSemAlias` | Top 10 holdings of the current iShares files with no alias in `$AliasesPosicoes` (maintenance reminder). |
| `Read-HistoricoNoticias`, `ConvertTo-NoticiaHistorico`, `Merge-HistoricoNoticias` | Read the news history (missing or unreadable → empty, with a notice or a warning), keep a story's public fields, merge this run's material and important stories (400 days, one per title). |
| `ConvertFrom-CsvSerie`, `Get-SerieMacro`, `Get-LembreteReunioes` | Read and validate an ECB CSV series, fetch one euro-area indicator, the Fed/ECB calendar reminder. |
| `Get-FundamentaisEmpresa`, `ConvertTo-Fundamentais`, `Get-FatosMetrica`, `Get-TrimestresMetrica`, `Get-FatorSplit` | Fundamentals of one company (cache, request, fallback), the facts converted to quarters and TTM, one metric's facts by period, its quarters (direct or derived), the split factor after a date. |
| `Get-ResultadosSEC`, `Get-SessaoReacao`, `Get-ProximaSessao`, `ConvertTo-IsoUtc` | Past earnings dates from the SEC, the reaction session in New York time, the next session from the price history, ISO UTC text from any date. |
| `Get-PalavrasTitulo`, `Join-NoticiasDuplicadas` | Duplicate grouping. |
| `Get-Serie` | One price series from Yahoo (with Kraken/Stooq fallbacks), validated; also returns splits. |
| `Get-TaxaBCE` | ECB EUR/USD reference rates. |
| `Get-PesosETF`, `Add-InfoETF` | One ETF's holdings file, validated, with its reference fallback. |
| `Get-DatasResultados` | Next earnings dates from Nasdaq. |
| `Get-Dividendo`, `Get-Dividendos` | Dividend events and the annual projection. |
| `Get-DadosBitcoin`, `Get-Halvings`, `Get-HalvingAproximado` | Bitcoin market, sentiment, network and halving. |

**PowerShell functions (section 4)**

| Function | Does |
|---|---|
| `Get-PontosGuardados`, `Get-SerieAnterior` | Read a series back from the previous run's data file. |
| `New-DadosSerie` | Shapes one asset's series for the JSON. |
| `Read-Backup`, `Get-InstanteBackup` | Validate a backup file and read its date. |
| `Write-Atomico` | Safe write: temporary file, then swap. |

**JavaScript in the template (section 3), by job**

| Job | Main names |
|---|---|
| Data and state | `D`, `ATIVOS`, `MERC`, `NEWS`, `FX`, `BCE`, `HIST`, `ETF_IDS`, `st`, `S0`, `store` |
| Numbers | `stats`, `inCur`, `fxAt`, `fxPt`, `slice`, `mm`, `corr`, `quedas` |
| Exchanges and dates | `frescura`, `asOf`, `sessao`, `regras`, `pascoa`, `tzParts`, `zoned`, `renderBolsas` |
| News sorting | `expoNoticias` (share of the portfolio exposed to each asset, from `pfDados()` and `etfInfo`), `relevancia`, `filteredNews(R)`, `newsItem(n, rel)` (with the *Via top holding* tag) |
| Prices: events and reactions | `HN` (news history), `MOV_LIM`, `JUROS_EV` / `JUROS_RE`, `sessaoDe` (session of an instant in the exchange's time zone), `grandesMovimentos`, `drawMoves`, `reacoes`, `drawEarn`, `eventosGrafico`; `lineMulti` option `events` |
| Euro area (Currency & Macroeconomics) | `macroDados`, `renderMacro` (called by `renderFx`) |
| Fundamentals tab | `FUND_IDS`, `fundDados`, `valorEm` (a value as known on a date), `ttmEm`, `acoesEm`, `fundHistorico` (month-end P/E and P/FCF without look-ahead, and today's values), `percentil`, `fundBloco`, `renderFund` (`CHARTS.fundamentals`) |
| Snapshot and tabs | `renderHero`, `cartaoFx`, `alerts`, `renderAlerts`, `renderNews`, `renderPrices`, `renderFx`, `renderCal`, `renderEtf`, `drawEtfMore`, `renderBtc`, `drawBtcMore`, `renderSrc`, `renderAll`, `showTab`, `init` |
| Portfolio | `carteira`, `efetiva`, `ajusteSplit`, `buildBuys`, `drawBuys`, `buildLots`, `drawLots`, `pfDados`, `drawPf` |
| Return and benchmark | `fluxosCarteira`, `custoFluxo` (fees hook, 0 for now), `xirr`, `rentab`, `drawRet` |
| Target allocation | `limpaAlvos`, `alocacao`, `reparte`, `mesesBanda`, `buildTgt`, `drawTgt`, `somaTgt` |
| Sell simulation and fees | `simulaVenda`, `buildSell` (uses `carteira(extra)`), `limpaFees`, `feesAtuais`, `feeDe`, `guardaFeeForm`, `custoFluxo` (fees in the XIRR), `qtdCompra`, `despesas` (Despesas e encargos per row) |
| Policy, notes and drop-alert context | `POL_K`, `limpaPolitica`, `buildPol`, `drawPol`, `limpaNotas`, `notasAtuais` (notes of existing entries only), `contextoQueda`, `durQ`; in `alerts()` the two drop alerts carry `ctx` (data only), drawn by `renderAlerts` |
| Dividends and tax | `divDados`, `dividendos`, `drawDiv` (with `DIV_LIQ`, the Net (est.) column), `ANEXO_J`, `anexoJ`, `linhaFiscal`, `csvLinhas`, `exportaAnexoJ` |
| Dividend export (Quadro 8A) | `ANEXO_J.q8a`, `divPagamentos`, `comDireito` (shares entitled, from `carteira()`), `anexo8A`, `COLS_8A`, `anosDiv`, `drawDiv8A`, `exportaDiv8A` |
| Backup | `dadosBackup`, `limpaBackup`, `juntaBackup`, `migraCompras`, `guardaFicheiro`, `descarrega`, `estadoBk`, `verificaPasta`, `carregaDoFicheiro` |
| Simulators | `simular`, `buildSim`, `drawSim`; strategy comparison `estrategias`, `drawStrat` |
| Rolling returns, underwater, log scale | `rolar` (all windows of N years), `drawRoll` (lazy, on opening the `<details>`), `ROLLH`, `underwater`, `etfLog` (Linear/Log switch `#f-elog`, passed to `lineMulti` as `log`) |
| Stress test | `EPISODIOS`, `stress`, `drawStress` (uses `precoEurEm`, `durQ`) |
| Exposure and concentration | `agregadosDe`, `setorAcao`, `SETOR_ACOES`, `exposicao`, `drawExpo`, `CONC_LIMIAR`; in the script `Get-AgregadosETF` (called by `Get-PesosETF`) |
| Charts | `lineMulti`, `barChart`, `hbars`, `spark`, `legend`, `ticks` |

The CSS is organised by comment headers (`buttons, tags, controls`, `top bar`, `hero / market snapshot`, `layout & navigation`, …, `responsive`).

---

## Data formats

**`bluechip-board-data.json`** (and the JSON inside the page). Data version `1.4` (since 6 Oct 2026; a 1.3 file still works as the previous-run fallback, and the page shows its missing parts as Unavailable).

| Key | Contents |
|---|---|
| `geradoEm`, `dias`, `versao`, `primeiraExecucao`, `duplicadas` | When it was built, news window, format version, first run?, duplicates grouped. |
| `ativos`, `mercado` | One object per asset / market reference: `id`, `nome`, `simbolo`, `moeda`, `pontos` (`[date, close]` pairs), `fonte`, `parcial` (intraday), `hora`, and for ETFs `tipo` and `bolsa`. |
| `fx` | `yahoo` (EUR/USD pairs), `bce` (ECB pairs), `fonte`, `parcial`. |
| `etf`, `etfs` | Holdings: `etf` is SXR8 in the original format; `etfs` has all four by id. Since 1.4 each fund also has `agregados` `{paises, setores, moedas, fontePaises}`. Each is a list of `{n, w}` (name, % of the fund) or `null` when Unavailable, and `fontePaises` is `file` or `index`. They are totals only, never the individual holdings. |
| `bitcoin` | Market, sentiment and network data. |
| `historico`, `splits` | Long-term series by id (plus `FX`), and split events by id. |
| `dividendos` | By company (see [Dividends](#dividends-projected-income)). |
| `noticias` | Stories: `titulo`, `link`, `fonte`, `data`, `empresas`, `temas`, `nivel`, `score`, `sentimento`, `tier`, `novo`, `soFeed`, `viaPosicao` (the top holdings that linked the story to a fund, or empty; missing in older files), `outras` (the grouped duplicates). |
| `historicoNoticias` | Since 1.4 (task 11): `inicio` (when the history starts: the first run that kept it, or the 400-day cut), `dias` (400) and `noticias`, the stories of `noticias-historico.json`: `chave`, `titulo`, `link`, `fonte`, `data` (UTC), `empresas`, `nivel`, `score`, `temas`. Missing in older files. |
| `resultadosSec` | Since 1.4 (task 11), by company: `estado` (`ok`, `previous run`, `error` or `skipped`), `fonte`, `obtidoEm`, `erro`, `nota`, `ultimoRelatorio` `{form, data, periodo, acc}` (latest 10-Q/10-K) and `resultados`, the 8-K with item 2.02: `acc`, `entrega` (filing date), `aceite` (UTC, or `null`), `horaNY`, `quando` (`before`, `during`, `after`), `sessao` (reaction session, or `null`), `nota`. Missing in older files. |
| `macro` | Since 1.4 (task 15), by id (`BCE_DFR`, `HICP_EA`): `estado` (`ok`, `previous run`, `error`), `fonte`, `obtidoEm`, `freq` (`B` change dates, `M` monthly), `pontos` (`[date, value]` pairs), `erro`. Missing in older files. |
| `fundamentais` | Since 1.4 (task 13), by company: `estado` (`ok`, `previous run`, `error`, `skipped`), `fonte`, `obtidoEm`, `relatorio` (the 10-Q/10-K of the collection), `tags`, `faltam` (metrics Unavailable and why), `trimestres` (`fim`, `inicio`, `ano`, `q`, `m` with `receita`, `lucroBruto`, `lucroOperacional`, `cfo`, `capex`, `fcf`, `eps`, `acoes`, each `{v, f, p, v0?, d?}` or `null`) and `ttm`. Missing in older files. |
| `bolsas`, `calendario` | Exchange definitions and calendar events. |
| `manutencao`, `fontes` | Maintenance reminders, and every source's status. |
| `backup` | **Only in `bluechip-board.html`**: your backup, or `null`. |

**`bluechip-board-backup.json`** (version 5): `app` (`"Bluechip Board"`), `version`, `saved`, `exported`, `buys`, `lots`, `sales`, `deleted`, and only when you have them: `targets` (target allocation), `policy` (investment policy), `notes` (notes per entry) and `fees` (fees per entry). Versions 1, 3 and 4 still load, in the script (`Read-Backup`) and in the page (merge and Restore backup).
- The backup goes only into `bluechip-board.html`, never into `Archive\` or `bluechip-board-data.json`, and that includes `targets`, `policy`, `notes` and `fees`.
- PowerShell 7 rewrites ISO dates in the copy it embeds without the milliseconds (for example `2026-10-06T09:00:00Z`). It is the same instant, and the page reads both forms.

**`vistos.json`**: an object of normalised headline → date first seen.

**`noticias-historico.json`**: `{versao: 1, inicio, dias: 400, noticias: [...]}`, with the same stories as `historicoNoticias`. One story per normalised title, keeping the version with the highest score.

---

## Website settings you can adjust

The website has no settings screen. Its tunable values are constants in the JavaScript part of the template inside `Bluechip-Board.ps1` (search for the name):

| Constant / place | Default | What it controls |
|---|---|---|
| `S0` | `{co:'all', per:'1A', cur:'EUR', lvl:'all', lim:60}` | Initial filter, period (`1M`, `3M`, `6M`, `1A`), currency (`EUR`/`USD`), news level and stories per page. |
| `PER` | `1M: 31, 3M: 92, 6M: 183, 1A: 372` days | Length of each chart period. |
| `alerts()` → `L` | stocks `{d:4, o:-20, y:-10}`, Bitcoin `{d:6, o:-30, y:-15}` | Alert thresholds: daily move, and distance from the high for *Important* and *Moderate*. |
| `HALV` | 2016, 2020, 2024 | Halving dates marked on the long-term Bitcoin chart when the script could not read them from the network (normally they come from the script, including future ones). |
| `SPLITS_FIXOS` | Apple 2014 7:1 and 2020 4:1, NVIDIA 2021 4:1 and 2024 10:1, Alphabet 2014 1.998:1 and 2022 20:1 | Reserve list of stock splits, merged with those Yahoo reports. |
| `fgClasse` | 24 / 46 / 54 / 75 | Fear & Greed bands (Alternative.me's), used for colours and alerts. |
| `frescura` | 2 sessions, or 1 after 6 h | When a price counts as behind its exchange. |
| `etfLp` | `'10'` | Default period of the ETF long-term chart (`'5'`, `'10'`, `'max'`), shared by all four ETFs. |
| `etfLog` | `false` | Scale of the ETF long-term chart when the page opens (`false` = linear). |
| `ROLLH` | `{etf:1, btc:1}` | Horizon of the rolling-returns histogram shown first, in years. |
| `etfSel` | `'SXR8'` | ETF shown first in the ETFs tab. |
| `relevancia` | `3` | Weight of your exposure in *Relevance to me* (score + 3 × share exposed). |
| `ANEXO_J.q8a` | `{quadro:'8A', codigo:'E11', codigoRetPT:'E10', retFonte:0.15, taxaPT:0.28, diasPag:45}` | Dividend export: Quadro 8A code (E10 if tax was withheld in Portugal), the estimated US withholding, the Portuguese rate (Net (est.) = 1 − `taxaPT`) and the end-of-year window for `PAY_DATE_UNKNOWN`. |
| `MOV_LIM` | 4 % (stocks and ETFs), 6 % (Bitcoin) | Smallest daily move listed in *Big moves explained*. They are the same thresholds as the daily-move alert in `alerts()`; change both together. |
| `CONC_LIMIAR` | `10` | Share of the portfolio (%) above which a company gets an informative note in *Concentration by company*. |
| `SETOR_ACOES` | Apple, NVIDIA: Information Technology; Alphabet: Communication | Sector of the directly held stocks when no iShares file lists them. |
| `HERO_LUGAR` | SXR8 on top; BTC and the EUR/USD card (`FX`) at the bottom | Which snapshot cards span the full width (the others form the three-column rows). |
| `--side` in the CSS (`@media (min-width:1600px)`, `(min-width:2400px)` and `(min-width:3000px)`) | `clamp(560px,29vw,820px)`; `clamp(860px,36vw,1100px)`; 4K `clamp(1100px,40vw,1700px)` | Width of the snapshot sidebar on wide screens. From 2400 px it is also the width of the brand column in the header, which aligns the filters with the sections. |
| `--gut`, `--max` in the CSS | 32 px / 3200 px from 1600 px; 44 px / 4400 px from 3000 px | Side margins and the maximum page width. |
| `.tickers` `grid-template-rows` (1600 px block) | full-width rows `1fr`, rows of three `1.6fr` | How the free height is shared between the snapshot rows. |
| `--eunk`, `--is3n`, `--eunn` | blue, teal, coral | Colours of the regional ETFs (next to `--spx` for SXR8). |
| `tabQuedas(...)` calls | ETF ≥ 10 % (top 6), Bitcoin ≥ 40 % (top 5) | Minimum drop and number of rows in the "biggest drops" tables. |
| `36*36e5` in `init()` | 36 hours | Age after which the stale-data banner appears. |
| `:root` CSS tokens | navy theme | Colours: `--bg`, `--surface`, …, asset colours `--aapl`, `--nvda`, `--googl`, `--spx`, `--btc`. |

After editing, run the script again to rebuild `bluechip-board.html`.

---

## Scheduling a daily run

```powershell
pwsh -ExecutionPolicy Bypass -File .\Bluechip-Board.ps1 -ScheduleDaily -Time 08:30 -SecEmail "your@email.com"
```

This registers the task **BluechipBoard** in Windows Task Scheduler, with the PowerShell that ran the command. It runs every day at that time with `-NoOpen`, in a hidden window:

- If the PC was off or asleep at that time, it runs later.
- It runs only when there is a network.
- It also runs on battery power (Windows' default would skip laptops on battery).
- It is stopped after 1 hour.

It runs only while you are signed in. Your SEC e-mail is stored in the task's arguments. Open `bluechip-board.html` whenever you want the latest data. To remove the task:

```powershell
Unregister-ScheduledTask -TaskName BluechipBoard -Confirm:$false
```

---

## Making a change safely

```text
edit Bluechip-Board.ps1 (keep UTF-8 with BOM)
        │
        ▼
Tests\Test-Engine.ps1 ──► Tests\Test-Site.ps1 ──► Tests\Test-Resilience.ps1   (on 5.1 and, with -Shell pwsh / pwsh, on 7)
        │ all pass?  no ──► fix the code (never loosen a test that caught a real regression)
        ▼ yes
run the script (shortcut, or -NoOpen) ──► look at the site ──► update README.md (and AUDIT-REMEDIATION.md for a notable change)
```

Pitfalls that have bitten before:

- **Encoding.** Save the script as UTF-8 **with BOM**. Without it, Windows PowerShell 5.1 breaks accented characters.
- **The `<` escape.** The script builds the JSON escape for `<` from two parts (`'\' + 'u003c'`). Typing the escape as one piece in some editors turns it back into `<`, which once silently disabled the protection against headlines that contain `</script>`.
- **Lines the tests depend on.** `Test-Engine.ps1` cuts the script between the `# 1. CONFIGURA…` and `# 3. MODELO DO SITE` comments. `Test-Site.ps1` looks for the function `Get-Plantilla`. `Test-Resilience.ps1` patches these exact texts: `function Get-Url {`, the Downloads lookup `(New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path`, `Write-Passo 'Building the website'`, `# Horário das bolsas`, `foreach ($r in @(Get-DatasResultados)) {` and `-TaskName 'BluechipBoard'`. If you rename any of them, update the test too.
- **Windows PowerShell 5.1 quirks.** Arrays that come out of a pipeline can be written to JSON as `{value, Count}` objects: build `[date, value]` pairs in a plain `foreach`. `Invoke-WebRequest` there does not follow HTTP 308 redirects.
- **PowerShell 7 quirk.** The User-Agent header is validated, so the SEC e-mail needs `-SkipHeaderValidation` (already handled in `Get-Url`).
- **Do not hand-edit the outputs.** `bluechip-board.html` and the data file are rebuilt on every run.

---

## Working with git

**The GitHub repository is the source of truth** for the code and the documentation. This folder is a clone of it. Your data and the outputs never go into git (`.gitignore`); they live only on this PC.

| In git | Only on this PC (ignored) |
|---|---|
| `Bluechip-Board.ps1`, `Start-BluechipBoard.ps1`, `Bluechip-Board.ico` | `bluechip-board-backup.json`, `bluechip-board-backup.previous.json` (your portfolio) |
| `README.md`, `AUDIT-REMEDIATION.md` | `bluechip-board.config.json` (your SEC e-mail) |
| `Tests\` (with the fixtures) | `bluechip-board.html` (embeds your backup), `bluechip-board-data.json`, `noticias-historico.json`, `vistos.json`, `Archive\` |
| `.gitignore`, `.gitattributes`, `.github\`, `bluechip-board.config.example.json` | |

**Set up on a new PC**

```powershell
git clone https://github.com/<your-account>/BluechipBoard.git
cd BluechipBoard
Copy-Item bluechip-board.config.example.json bluechip-board.config.json   # then put your e-mail in it
```

Then point the desktop shortcut at `Start-BluechipBoard.ps1` in the clone. Bring your `bluechip-board-backup.json` from the old PC (or restore it from the website). The first run creates the outputs.

**Making a change**

```text
git switch -c <short-name>        a branch for the change (main stays working)
edit ──► run the 3 test suites on 5.1 and 7 (see Making a change safely)
git add -A ; git commit -m "…"    small commits, with a message that says why
git push -u origin <short-name>   GitHub Actions runs the tests ──► open a pull request ──► merge into main
git switch main ; git pull        this PC back on the shared version
```

- **Never commit your data.** `git status` must never list `bluechip-board-backup.json`, `bluechip-board.config.json` or an output. If it does, fix `.gitignore` first.
- **Line endings.** `.gitattributes` keeps the scripts in LF and the test CSV files byte for byte. Do not change it.
- **Going back.** Every version is in the history: `git log`, then `git revert <commit>` (or `git switch -c look <commit>` to look at an old state). Data file 1.4 and backup version 5 still load in older versions back to task 03 (see [Data formats](#data-formats)).
- **Before using the shortcut on a PC that has not been updated,** run `git pull`. A run changes no file that git tracks.

---

## Maintenance checklist

- **Calendar:** add new manual events to `$Calendario` (Fed meetings, product events, tax deadlines). Earnings dates update themselves through Nasdaq.
- **Exchange holidays:** only exceptions need adding to `$Bolsas` (and the Xetra year-end session if Deutsche Börse announces a time other than 14:00).
- **Stock splits:** nothing to do. They come from Yahoo, and saved purchases are adjusted automatically. `SPLITS_FIXOS` is only a reserve.
- **ETF fallback weights:** refresh `$PesosReferencia` (SXR8) and each `Referencia` in `$ETFs` occasionally. If iShares changes a fund's page, update its `portfolioId` in `$ETFs`; the ISIN check rejects a wrong file.
- **Maintenance reminders:** the script prints them in yellow, and *Sources & method* repeats them, when reference weights are over 4 months old, the manual calendar has no events more than 30 days ahead or no Fed/ECB decision more than 60 days ahead, an exchange's exception list ends before the current year, or a fund's current top 10 has a holding without a news alias (add it to `$AliasesPosicoes`, with a known false positive in `Test-Engine.ps1`).
- **Browser identity:** the Chrome version in `$Script:UA` can be updated now and then.
- **Anexo J, once a year before filing:** compare the new form and instructions on the Portal das Finanças with the `ANEXO_J` object and the CSV columns in the template (codes, table numbers, columns, the 365-day rule, and Quadro 8A with codes E10/E11).
- **A source keeps failing:** check its row in *Sources & method*. The URL may have changed, so edit `$Feeds` or the relevant function.
- **After any edit:** keep the file encoding UTF-8 with BOM and run the [tests](#tests) (see [Making a change safely](#making-a-change-safely)).

---

## Tests

The `Tests` folder has three automated test scripts. None of them changes the project files, your Downloads folder or Task Scheduler, and none needs the internet.

```powershell
pwsh -ExecutionPolicy Bypass -File .\Tests\Test-Engine.ps1                    # PowerShell side, offline
pwsh -ExecutionPolicy Bypass -File .\Tests\Test-Site.ps1                      # website, in headless Chrome or Edge
pwsh -ExecutionPolicy Bypass -File .\Tests\Test-Resilience.ps1 -Shell pwsh    # failures, on a copy in a temp folder
```

Use `powershell` instead of `pwsh` (and leave out `-Shell pwsh`) to test on Windows PowerShell 5.1. Without a local `bluechip-board-data.json` (a fresh clone), the site and resilience tests use `Tests\fixtures\bluechip-board-data.sample.json`. GitHub Actions runs all three suites, on both shells, for every push to `main` and every pull request (`.github\workflows\tests.yml`).

| Script | What it checks |
|---|---|
| `Test-Engine.ps1` (212 checks) | Date parsing and time zones, the **ECB CSV** (a valid file, an HTML error page, `.`/`NaN`/empty values, an empty series, implausible values, another series, an unexpected header, daily dates and negative rates, the source down, a valid response) and the **central-bank calendar** (the 2026–2027 Fed and ECB decisions, dates in order, the 60-day reminder for one or both), the **fundamentals** with canned companyfacts (4th quarter = year − 9 months, a restatement keeping the first published value, cash flow derived from year-to-date figures, FCF, EPS of the 4th quarter and shares never derived, TTM, durations and forms ignored, a missing tag, wrong units, no interpolation, a split adjusted by the filing date, NVIDIA's fiscal year, the cache without a new 10-Q, a newer 10-K, the source down with the 120-day fallback, no `-SecEmail`, another CIK), the **reaction session** in New York time (before, during and after the session, the boundaries, the US clock change, a holiday and a day missing from the price history, after the end of the history), **SEC past earnings** with canned responses (8-K item 2.02 of the last 5 years, acceptance time from the 8-K feed and not from the JSON field, no time when missing, latest 10-Q, the User-Agent, another CIK, lists of different lengths, SEC down, the feed down with times reused), the **news history** (missing and corrupted files, retention, levels kept, one story per title, public fields only, start date, read back on PowerShell 7), charset decoding, safe XML, the classification rules (with known false positives and negatives), the **top-holding aliases** (a story linked to its fund, at most moderate and with no bonus; a keyword or a feed's own asset wins; every alias against a true story and a known false positive; the stories that already matched keep their assets, level and score; the maintenance reminder, none without a current file, and the top 10 of 5 Oct 2026 fully covered), duplicate grouping, Yahoo response validation and split events, the **dividend collector** (annual rate and trailing sum, yield in percentage points, wrong symbol or currency, no dividends, invalid amounts, one payment, suspended payments, implausible yield, source down), the **ETFs** (configuration and metadata of the regional funds, SXR8 rules, wrong listing or currency rejected, holdings files read with test copies of the iShares format: own top 10 and sectors, wrong fund, missing sheet, bad weights, source down, four funds with one failing, accumulating ETFs never in dividends, news matching), the halving estimate, backup reading (versions 1, 3, 4 and 5, with the targets, policy and notes kept when the backup is copied into the site), safe writes and the PowerShell 7 header handling, and the **fund aggregates** (country from the Geography block or the index, sector, currency, Cash/Other, totals that do not add up to about 100 % rejected, reference weights → Unavailable). Works in PowerShell 5.1 and 7. |
| `Test-Site.ps1` (416 checks) | Builds a page from the template and the latest `bluechip-board-data.json`, then runs each scenario in a fresh browser profile. Scenarios: a new user (including the snapshot order and layout, and the EUR/USD card); old and new backups; reopening; merging between browsers; deletions; a future 10:1 split; currency failures; stale and previous-run prices; FIFO sales and the tax counter; the exchange calendar around daylight-saving changes; the statistics against independent calculations; HTML injection from a headline; unreadable data. **Dividends** (`div*`): positive, partial, fractional and zero holdings, the total, EUR conversion, SXR8 excluded, failed, invalid, malformed, stale and previous-run data, no EUR/USD. **Anexo J** (`tax*`): multi-lot FIFO, one purchase / several sales, unsold lots, years, user prices, separate FX per date, mapping, splits before/after/multiple, crypto 365-day boundaries, old backups, missing optional fields, malformed and oversold sales, missing FX, CSV encoding, escaping, decimals, determinism and that exporting changes nothing. **`taxBaseline`**: a fixed test backup (several lots, sales in 2025 and 2026, a 10:1 split, Bitcoin held under and over 365 days) with its own price, EUR/USD and split data. Its four CSV files must be identical, **byte by byte**, to the reference files in `Tests\fixtures\taxBaseline\`, so any change to the tax export is caught. **Return** (`ret*`): XIRR against an independent calculation (purchases only; purchases, a Bitcoin lot and a sale; under one year, not annualised), the SXR8 benchmark by hand (same money, a sale larger than the benchmark, an intraday SXR8 price), an asset without a price, a purchase before the SXR8 history, an empty portfolio, and the solver without a solution. **Target allocation** (`tgt*`): weights and deviations, the band, the split of M against an independent calculation (gaps larger than M, and every asset below target), the months back inside the band against an independent simulation, an asset without a price, targets that do not add up to 100 % or a negative band (nothing saved), saving (backup version 5), merging two backups with different targets (most recent `at` wins, invalid or hostile targets ignored), and version 1, 3 and 4 backups without targets. **Policy, notes and drop-alert context** (`pol*`): the existing drop alerts keep their level and exact text (with and without a policy); the context line with the user's rule for 20 % or 30 %, the history counted with `quedas()` (cases, recovered, median time against independent dates, not yet recovered, sample period, caveat), zero historical cases; saving the policy; notes added, shown, cleaned for missing and deleted entries, entries unchanged; merging (newest policy and each newest note win, invalid ones ignored); version 1, 3 and 4 backups. The `xss` scenario also feeds hostile text into a note and the policy. **ETFs** (`etf*`): <ul><li>cards and filter chips, prices in euros and in dollars, filtered news, alerts and calendar;</li><li>the ETFs view (own holdings, history, drops, simulator; SXR8 view);</li><li>Portfolio purchases and sales through the form, fractional shares, FIFO, removal, allocation and look-through;</li><li>no dividends, the Anexo J mapping and CSV, and backup and restore;</li><li>missing, stale and previous-run prices, the stock thresholds for alerts, holdings fallback and malformed ETF data.</li></ul>Run one scenario with `-Only split`. `-WriteFixtures` writes reference files that are missing and never overwrites an existing one. To accept a deliberate change to the tax files, delete the old reference and record why. **Stress test and strategies** (`stress*`, `strat`):
<ul><li>the episode dates;</li><li>each asset's euro return and loss against independent figures (including the last close before a closed day and a US stock at each day's EUR/USD);</li><li>an episode outside an asset's history (left out and named);</li><li>totals over the assets with data and the time back to the peak value;</li><li>no holdings;</li><li>the three strategies against an independent calculation (same total and period, the dip rule with a brute-force 52-week high);</li><li>the monthly simulator's numbers and saved inputs unchanged;</li><li>an adjustable X.</li></ul> **Rolling returns, underwater and log** (`roll*`): every 1/3/5-year window (252 sessions, or 365 days for Bitcoin) against an independent calculation; worst, median, best, negative share and annualised values; nothing calculated before the section is opened; a short history (5 years Unavailable); the underwater series next to the drops table; the log switch on and off with linear as the default. **Exposure** (`expo*`, `conc`): country, sector and currency weighted by today's values against hand-computed figures, the uncovered share when a fund is Unavailable, stocks only, a version-1.3 data file, the SXR8 currency note and the index flag; concentration with the 10 % threshold, top 10 per fund and no adding up by name. **Sell simulation and fees** (`sell*`, `fees*`):
<ul><li>FIFO over several lots, the gain and the 28 % estimate;</li><li>Bitcoin exempt and taxable lots, and the warning under 30 days;</li><li>oversell refused, and nothing saved;</li><li>fees split per row with the rounding remainder, and `FEE_PARTIAL`;</li><li>"Fill in manually" without Despesas when both fees are known;</li><li>the register's Fee column, with the gain unchanged;</li><li>XIRR with fees;</li><li>orphan fees ignored, editing and merging;</li><li>version 1, 3 and 4 backups.</li></ul> **Dividend export (Quadro 8A)** (`div8a*`):
<ul><li>one row per payment while held, with shares bought before and on the ex-date, a sale after it and a sale on it;</li><li>future ex-dates left out;</li><li>gross in dollars, the 15 % estimate and euros at the ex-date rate;</li><li>E11 / 840, `BROKER_FX`, and `PAY_DATE_UNKNOWN` in December but not on 16 November;</li><li>CSV encoding, decimals and determinism;</li><li>the Net (est.) column and note, with the existing columns and totals unchanged;</li><li>the separate button (one file) and the existing export (still two files);</li><li>no EUR/USD and unavailable dividend data (no rows invented).</li></ul> **News: Relevance to me** (`newsrel*`):
<ul><li>the exposure per asset against an independent calculation (direct, look-through, ETF, Bitcoin, Market);</li><li>the default order unchanged;</li><li>the relevance formula and the new order;</li><li>levels and points unchanged;</li><li>the note;</li><li>the *Via top holding* tag;</li><li>with no holdings, the option disabled with a note.</li></ul> **Prices: events** (`evts*`):
<ul><li>the session of a story in New York time (after the close, during the session, a weekend, the US clock change, the UTC day for Bitcoin);</li><li>the big moves of 4 % or more with the stories of that session and the one before, and "before the news history";</li><li>"News history since";</li><li>the markers per kind and their tooltip;</li><li>HTML in a headline;</li><li>the earnings reaction and the 5-session move against hand-computed figures, the median and the cases;</li><li>a failed and a missing SEC source;</li><li>no news history and no SEC data (Unavailable, the rest of the tab working).</li></ul> **Fundamentals** (`fund*`):
<ul><li>the tab between Prices and Currency, and `#fundamentals` routing;</li><li>no look-ahead (a quarter not yet published, a value revised later);</li><li>month-end history from the first 4 published quarters;</li><li>the percentile;</li><li>today's P/E and P/FCF;</li><li>the quarter table (growth, margins, buybacks, publication date);</li><li>partial data, a failed company and no data at all;</li><li>the note;</li><li>no horizontal overflow;</li><li>the asset filter (one company, an ETF).</li></ul> **Euro area** (`macro*`): the deposit rate with the date it took effect, inflation with its month flagged when not recent and labelled previous run, the notes, no new Overview alert, and no ECB data (Unavailable). |
| `Test-Resilience.ps1` (56 checks) | Runs a patched copy of the script with every download blocked. Covers the ECB indicators (offline: one error per series; the previous run reused if at most 30 days old, negative rates kept), the news history (a new file written at the end with a notice, 400-day retention of an existing file, a corrupted file kept as `.bad` with a warning, never written by a failed run) the fundamentals (none requested without `-SecEmail`; offline, the previous run's facts reused when at most 120 days old, otherwise Unavailable) and the SEC past earnings dates (no request without `-SecEmail`; with it, offline: one error per company, the previous run reused if at most 180 days old, otherwise nothing invented; a failed first run writes no file). Covers a fresh install (including the regional ETFs: no prices invented, each with its own reference weights), the previous run's data used as fallback (including each ETF's prices and history, dividends, and the 180-day limit), the Downloads backup validation, a second instance, a run that fails halfway, the earnings-date merge rules, a malformed calendar date and the scheduled-task settings (dry run). It also checks that a version-5 backup's targets, policy, notes and fees reach the main site only, never the Archive copy or the data file. Offline, every fund's aggregates are Unavailable and the data version is 1.4. Use `-Shell pwsh` for PowerShell 7. |

```text
Test-Engine      script sections 1–2 loaded ─► Get-Url replaced by canned answers ─► checks
Test-Site        Get-Plantilla + data file + a hostile headline ─► test page ─► headless Chrome, one fresh profile per
                 scenario ─► window.__BB_TEST__ hook ─► checks
Test-Resilience  script copied to %TEMP% and patched ─► run offline with every download blocked ─► files and output checked
```

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| The window stays open with "Finished with problems" | Read the error or warning above it. A single failing source does **not** cause this; it only appears in *Sources & method*. |
| "Bluechip Board is already running in another window" | Another run (shortcut or scheduled task) is in progress. Wait for it to finish. |
| "SEC EDGAR … skipped" in Sources | Run with `-SecEmail` (the shortcut already does). |
| SEC sources fail with "The format of value … is invalid" | An old copy of the script on PowerShell 7. The current `Get-Url` handles it; check that the script is up to date. |
| Prices missing for an asset | Yahoo failed and there was no previous run to fall back on. Check *Sources & method*. Bitcoin falls back to Kraken; Stooq (the stock fallback) currently blocks automated requests. |
| A price shows "previous run" or an orange date | The source failed or is behind: the figures for that asset are not current. Run the script again later. |
| Euro values blank for the US stocks | No EUR/USD rate in this run (see the red alert). Run the script again later. |
| Portfolio empty after clearing the browser | Run the script once (it embeds `bluechip-board-backup.json`) and reopen the site, or use **Restore backup**. |
| "Some changes are not in the backup file yet" | Click **Save to project folder**. The browser asks for permission again each time the page is reopened. |
| "The backup you saved … was not in the project folder" | The file was saved in another folder. Save again and choose the BluechipBoard folder. |
| Dividends show "Unavailable" | The dividend source failed and there was no recent previous run's data (see *Sources & method*, rows "Dividends: …"). Run the script again later. |
| The tax CSV opens in a single column, or numbers look wrong | Excel in another language expects `,` as the separator. Use **Data → From Text/CSV**, with delimiter `;` and locale Portuguese (decimal comma). |
| The browser asks to allow multiple downloads | The export saves two files. Allow it for this page. |
| The date picker shows `dd/mm/aaaa` or Portuguese month names | The browser uses the Windows language for native date pickers. Dates are stored correctly. |
| The snapshot cards look too tall or too small | The layout follows the window size (see [Website sections](#website-sections)). Browser zoom changes the width the page sees; `Ctrl+0` resets it. |
| The desktop icon looks outdated | Windows icon cache. It refreshes by itself or after signing out and in. |
| Accented characters broken after editing | The script was saved without a BOM. Save it again as UTF-8 with BOM. |

---

## Glossary: Portuguese names in the code

The code uses Portuguese names and comments. The most common ones:

**Script (PowerShell)**

| Name | Meaning |
|---|---|
| `Ativos` / `Mercado` | Assets / market references (S&P 500, VIX, 10-year yield). |
| `Bolsas`, `feriados`, `curtos`, `abre`, `fecha` | Exchanges, holidays, early closes, opening and closing time. |
| `Calendario` (`d`, `e`, `ev`, `imp`, `st`) | Calendar (date, who, event, importance, status). |
| `Feeds`, `Dica`, `Exigir` | News feeds, default asset (hint), required asset. |
| `EmpresasRe`, `Temas`, `Severo`, `DecisaoJuros`, `Ruido`, `Positivo` / `Negativo` | Asset keywords, themes, adverse events, rate decision, noise (click-bait), sentiment words. |
| `TierPrimaria` / `TierReferencia` / `TierCuidado` | Source tiers: primary, leading press, handle with care. |
| `PesosReferencia`, `Referencia`, `Empresas` | Reference (fallback) weights, the panel companies a fund should hold. |
| `Acumulacao`, `Domicilio`, `Gestora`, `Indice`, `Fundo`, `Inicio` | Accumulating, domicile, fund manager, index, fund, inception. |
| `AtivosDividendos`, `Dividendos`, `anualPorAcao`, `rendimentoPct` | Dividend assets, dividends, annual amount per share, yield in %. |
| `Get-Serie`, `Get-TaxaBCE`, `Get-PesosETF`, `Get-DatasResultados` | Price series, ECB rate, ETF weights, earnings dates. |
| `Measure-Noticia`, `Join-NoticiasDuplicadas` | Score a story, group duplicate stories. |
| `Write-Passo`, `Add-Fonte`, `Write-Atomico` | Progress line, record a source's status, safe write. |
| `Get-Plantilla` | The website template. |
| `noticias`, `empresas`, `nivel`, `outras`, `novo`, `soFeed`, `chave` | Stories, assets of a story, level, grouped duplicates, new, weak feed match, normalised title key. |
| `historicoNoticias`, `HistoricoDias`, `inicio` | News history, how many days it keeps (400), when it starts. |
| `macro`, `SeriesMacro`, `Fluxo`, `Chave`, `faltas`, `invalidos`, `LembreteReunioes` | Euro-area indicators, their configuration, ECB data flow and series key, missing and invalid values, the meetings reminder. |
| `fundamentais`, `MetricasSEC`, `trimestres`, `receita`, `lucroBruto`, `lucroOperacional`, `cfo`, `acoes`, `faltam` | Fundamentals, the metrics configuration, quarters, revenue, gross profit, operating income, operating cash flow, shares, metrics missing. |
| `resultadosSec`, `resultados`, `entrega`, `aceite`, `horaNY`, `quando`, `sessao`, `ultimoRelatorio` | SEC past earnings: the 8-K list, filing date, acceptance time, New York time, before/during/after the session, reaction session, latest 10-Q/10-K. |
| `AliasesPosicoes`, `viaPosicao`, `Get-PosicoesSemAlias` | Top-holding aliases per fund, the holdings that linked a story to a fund, top 10 holdings without an alias. |
| `vistos`, `anterior`, `historico`, `desdobramentos`, `manutencao`, `fontes` | Seen stories, previous run, long-term history, splits, maintenance reminders, sources. |
| `Pasta`, `Dias`, `NaoAbrir`, `AgendarDiariamente`, `Hora`, `Trinco` | Folder, days, don't open, schedule daily, time, lock (mutex). |
| `Agora`, `Inv` | Now (the run's timestamp), invariant culture. |

**Website (JavaScript)**

| Name | Meaning |
|---|---|
| `D`, `st`, `store` | The embedded data, the current filters, the `localStorage` helper. |
| `carteira`, `efetiva`, `ajusteSplit` | Portfolio (FIFO), an entry with any split applied, split detection. |
| `pagamentos`, `comDireito`, `bruto`, `ret`, `liq`, `DIV_LIQ` | Dividend payments, shares entitled to a payment, gross amount, tax withheld, net (estimated), the net factor (0.72). |
| `fundDados`, `valorEm`, `ttmEm`, `fundHistorico`, `percentil` | Fundamentals of a company, a quarter's value as known on a date, the last 4 quarters on a date, the 10-year valuation history, percentile. |
| `sessaoDe`, `grandesMovimentos`, `reacoes`, `eventosGrafico` | Session of a story, big moves, earnings reactions, chart markers (Prices). |
| `expoNoticias`, `relevancia` | Share of the portfolio exposed to each news asset, relevance of a story (Relevance to me). |
| `fluxosCarteira`, `custoFluxo`, `rentab`, `xirr` | Cash flows of the register, cost of a flow (fees), return of the portfolio and of the benchmark, the XIRR solver. |
| `frescura`, `antigo`, `asOf`, `sessao`, `regras`, `pascoa` | Freshness of a series, previous-run data, the date label, exchange session, holiday rules, Easter. |
| `inCur`, `fxAt`, `fxPt` | Convert to a currency, the EUR/USD rate (and its date) on a day. |
| `HIST`, `ATIVOS`, `MERC`, `FX`, `BCE`, `NEWS` | Long-term histories, assets, market references, EUR/USD, ECB rates, stories. |
| `CO`, `PF`, `HERO_LUGAR`, `cartaoFx` | Colours per asset, assets you can buy, which snapshot cards are full width, the EUR/USD card. |
| `etfSel`, `etfInfo`, `ehEtf`, `etfAcum` | Selected ETF, its data, is it an ETF, is it accumulating. |
| `rolar`, `underwater`, `etfLog` | Rolling returns over N years, the underwater series (price ÷ running maximum − 1), the log-scale switch. |
| `agregados`, `paises`, `setores`, `moedas`, `PaisIndice`, `exposicao` | Fund aggregates, countries, sectors, currencies, the index country, the exposure calculation. |
| `EPISODIOS`, `stress`, `estrategias` | The three S&P 500 falls (peak and trough dates), the stress test, the strategy comparison. |
| `simulaVenda`, `feeDe`, `despesas` | Sell simulation, the fee of an entry, the Despesas e encargos of each tax row. |
| `contextoQueda`, `limpaPolitica`, `limpaNotas`, `notasAtuais`, `notas` | Context line under a drop alert, sanitise the policy, sanitise notes, notes of existing entries, notes. |
| `alocacao`, `reparte`, `mesesBanda`, `limpaAlvos`, `alvos` | Target allocation, split of a contribution, months back inside the band, sanitise saved targets, targets. |
| `juntaBackup`, `limpaBackup`, `apagados`, `guardaFicheiro`, `descarrega`, `verificaPasta`, `LEGADO` | Merge a backup, sanitise it, deleted ids, save to the file, download, check the folder, one-time legacy rule. |
| `anexoJ`, `linhaFiscal`, `ANEXO_J`, `anosFiscais` | Tax export, one tax row, tax mapping, tax years. |
| `quedas`, `simular`, `vendas`, `lotes` | Drops, simulate, sales, lots (purchases still held). |
| `nivel` colours | `red` material, `orange` important, `yellow` moderate, `white` noise. |
