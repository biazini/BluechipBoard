# Audit remediation log

Record of the fixes made after the technical audit of Bluechip Board, carried out on 4 Oct 2026. For each finding: what the audit found, why it happened, what changed, how it was verified and its final status.

**Status key:** FIXED · PARTIALLY FIXED (the local handling is done, but an external limit remains) · NOT FIXED (with the reason).

**How to read this file.** The **[follow-up of 7 Oct 2026](#post-remediation-follow-up--7-oct-2026)** (at the end) re-checked every finding below, fixed what could be fixed and records the decisions; the rows of Open items it changed carry a "7 Oct 2026" note. Start with **[Open items](#open-items--6-oct-2026)**: everything still to decide, fix or watch, in one place. The audit sections (Critical to Testing gap) record the state on 4 Oct 2026 and are kept as they were written. Later work is in the dated change logs at the end. The [status review](#status-review--5-oct-2026) says which audit findings still applied on 5 Oct.

**Current test counts (7 Oct 2026, with live prices):** engine 278, site 463, resilience 73, all passing on Windows PowerShell 5.1 and PowerShell 7 (after the remediation: 251, 436, 62; 6 Oct: 212, 416, 56). The table below has the counts at the end of the remediation.

**Test suites at the end of the remediation** (all in `Tests\`, see the README):

| Suite | Checks | Result |
|---|---|---|
| `Test-Engine.ps1` | 45 | Pass on Windows PowerShell 5.1 and PowerShell 7 |
| `Test-Site.ps1` | 84 | Pass in headless Chrome, 21 scenarios |
| `Test-Resilience.ps1` | 29 | Pass on PowerShell 5.1 and 7 |

The suites are referred to below as *engine*, *site* and *resilience*.

**Files changed:** `Bluechip-Board.ps1` and `README.md`. **Files added:** `Tests\*`, `AUDIT-REMEDIATION.md`. **Files not changed:** `Start-BluechipBoard.ps1`, `bluechip-board-backup.json` and `Bluechip-Board.ico`. A copy of the whole project as it was before the fixes was kept outside the project folder during the work.

---

## Open items · 6 Oct 2026

The single list of what is left after the audit (4 Oct) and the feature work of 5–6 Oct (tasks 01–15). The task plan and its progress log (`prompts\`, with `PROGRESSO.md`: the detailed decisions and the final report) were removed from the project on 6 Oct 2026 and sent to the Windows Recycle Bin. Every decision below is also in the change logs further down, so this file is now the reference.

**State on 6 Oct 2026**
- **Tests:** engine 212, site 416, resilience 56, all green on Windows PowerShell 5.1 and PowerShell 7.
- **Real run:** 52–63 s, or about 70 s on a day a new 10-Q/10-K triggers the fundamentals download; 79/79 sources answered.
- **Data:** data file 1.4, backup version 5.
- **Version control:** since 6 Oct 2026 the project is in **git**, with the private GitHub repository `biazini/BluechipBoard` as the source of truth. GitHub Actions runs the three suites on both shells for every push to `main` and every pull request. Personal data, the SEC e-mail (`bluechip-board.config.json`) and every output stay out of git (`.gitignore`). See README → *Working with git*.

### A. Decisions for the owner (nothing is broken; a choice is needed)

| ID | Item | Options / proposal | Where |
|---|---|---|---|
| D1 | **Quadro 8A code** for dividends: the export uses **E11** (no tax withheld in Portugal), the usual case with a foreign broker. | If a Portuguese entity withheld tax, use **E10** and fill in the *Imposto retido em Portugal* columns. | `ANEXO_J.q8a.codigo` (template) |
| D2 | **Anexo J interpretations:** G20 for UCITS ETFs (some advisers use G01); "País da fonte" = the issuer's country. | Confirm with an accountant; review the form and instructions every year before filing. | `ANEXO_J` |
| D3 | **7 Oct 2026:** FRED still does not answer (timeout, honest User-Agent); the Chicago Fed's own NFCI CSV (option c) **does** answer (no key, weekly, 147 KB): verified, not implemented (a new indicator is your call). **FRED series not implemented** (yield curve 10y−3m `T10Y3M`, high-yield spread `BAMLH0A0HYM2`, financial conditions `NFCI`). FRED (Akamai) resets or ignores the script's requests, on 5.1 and 7 and with an honest User-Agent; it answers only clients that identify as known tools such as curl. Posing as one would mean getting around a bot filter, so it was not done. | (a) let the script call Windows' `curl.exe` for FRED: an exception to "every source uses `Get-Url`", and it depends on FRED keeping that filter; (b) the curve from Yahoo `^TNX` − `^IRX`, an existing source; (c) NFCI from the Chicago Fed's own CSV (not verified yet); (d) leave it. The HY spread has no known public alternative. | Currency & Macroeconomics, note `#macroNao` |
| D4 | **S&P 500 CAPE not implemented:** no public CSV with current data. The GitHub/datahub `s-and-p-500` dataset has PE10 only until Sep 2023, then 0.0. | Implement only if a current CSV source appears (Shiller's `.xls` is not usable without modules). | — |
| D5 | **Alert "below the 200-day average"** (`yellow`). It fires for weeks in ordinary corrections and adds to the distance-from-high alerts. | Proposal: make it informational (`white`), or show it only when the asset has no drop alert already. Not changed without your decision. | `alerts()` |
| D6 | **ECB decisions in the calendar are `Medium`**, so they raise no Overview alert (a `High` event alerts 7 days before; task 15 asked for no new alerts). The Fed decisions are `High`, as before. | Change to `High` if you want the 7-day warning for the ECB too. | `$Calendario` |
| D7 | **Top-holding news aliases** apply only to feeds without a `Dica`, so no story that was already classified changes. | Letting them replace a search feed's `Dica` is more precise, but it changes some assignments and needs the old engine test *"one EM company without keyword is only a weak feed match"* to be updated deliberately. | `Measure-Noticia`, `$AliasesPosicoes` |
| D8 | **7 Oct 2026: kept Unavailable** (see the follow-up, D8). **Alphabet gross margin is Unavailable:** Alphabet does not report `GrossProfit`. | Compute it as revenue − `CostOfRevenue` (a derivation, not a tag; not done). | `$MetricasSEC` |
| D9 | **Country exposure of SXR8 and EUNN** uses the index country, flagged as approximate (your choice of 6 Oct). iShares gives no country breakdown for these two funds. | Keep, or read another source. | `Get-AgregadosETF`, `PaisIndice` |
| D10 | ~~Install git~~ **Done on 6 Oct 2026:** private repository `biazini/BluechipBoard`, CI in `.github\workflows\tests.yml`. The history starts with that import; the earlier per-task states exist only in the snapshots in the Recycle Bin (see E). | Optional: protect `main` (pull requests only) once the plan allows it. | — |

### B. Known limitations and approximations (working as designed, but worth knowing)

| ID | Item |
|---|---|
| B1 | **Stress test (05):** the S&P 500 dates of each fall are applied to every asset; it is not each asset's own peak and trough. |
| B2 | **Target allocation (03):** the "gaps smaller than M" branch cannot happen with weights adding up to 100 % (the gaps always add up to M or more). It is kept and tested directly. |
| B3 | **Drop-alert context (04):** "falls as deep or deeper" is read as the current depth from the 52-week high, compared with drops from the running all-time peak. |
| B4 | **Log scale (06):** only on the ETF long-term chart (Bitcoin is always log; the stocks have no long-term chart). |
| B5 | **Fees (08):** the realised gain shown in the register does not deduct fees. They are used in the XIRR and in *Despesas e encargos* of the tax export. The sell simulation excludes fees, loss offsetting and *englobamento* (stated on the page). |
| B6 | **Dividends for Quadro 8A (09):** Yahoo gives ex-dates only, so the tax year is the ex-date's year, and ex-dates in the last 45 days of the year are flagged `PAY_DATE_UNKNOWN`. The US withholding is an estimate (15 %, W-8BEN). Euros are a reference at the ex-date rate (`BROKER_FX`). The dividend data covers about 2 years back from each run. |
| B7 | **7 Oct 2026: resolved.** Every earnings 8-K now has its time: the ones missing from the Atom feed get it from their EDGAR filing header (`ACCEPTANCE-DATETIME`); on the real data 54 of 54 (AAPL 20, NVDA 21, GOOGL 13). **SEC acceptance times (11):** the `acceptanceDateTime` of `data.sec.gov/submissions` is wrong for Apple: several hours late, 8 h in summer and 10 h in winter. The times come from the EDGAR 8-K Atom feed, which lists only the 40 latest 8-K (about 4 years). Older earnings have no time and no reaction session: AAPL 1, NVDA 5, GOOGL 1 on 6 Oct. |
| B8 | **SEC "recent" list:** about 1,000 filings, so for Alphabet (many Form 4) it reaches back only about 3 years (13 earnings releases). Older filings are in the extra files the JSON lists, which are not read. |
| B9 | **Fundamentals (13):** the EPS of a 4th quarter = year − 9 months, an approximation (the share count changes during the year). Diluted shares of a 4th quarter are never derived (empty). Alphabet's diluted shares exist only since 2022. P/FCF uses the latest quarter's diluted shares, not a 4-quarter average. Splits are applied by each value's **filing date** (filings after a split are already adjusted). |
| B10 | **News history (11–12):** it started on 6 Oct 2026, so *Big moves explained* shows "Before the news history" for earlier sessions until it fills up. It keeps material and important stories only. |
| B11 | **Big moves (12):** a stock's moves also list *Market* stories (for the US stocks and SXR8). The thresholds `MOV_LIM` (4 % / 6 %) copy the daily-move alert of `alerts()`, and the rate-decision regex `JUROS_RE` copies `$DecisaoJuros` in the script. **Change both places together.** |
| B12 | **Euro area inflation (15):** on 6 Oct 2026 the ECB Data Portal had HICP data only until **Dec 2025**, for U2, I8 and I9. The page shows the month and flags it as not recent. If it stays like this, check whether the ECB moved the series to another key (for example after Bulgaria joined the euro area in 2026). |
| B13 | **Tab bar:** with ten tabs it fits from about 1,900 px. At 1,600–1,700 px it scrolls sideways, as it did before task 14. |
| B14 | **7 Oct 2026: fixed** (only the sale being added, or a sale it breaks, is refused; scenario `sellother`). **Purchase form:** it still refuses a new sale when *any* existing sale lacks matching purchases (`C.vendas.find(x=>x.falta>0)` in `buildBuys`). It happens only with old or restored data. |
| B15 | **7 Oct 2026: improved** for the US stocks (Nasdaq history as a checked second source); the ETFs still depend on Yahoo. **Live prices** depend on Yahoo only (H5): Stooq blocks automated requests, and a second live source would need a key. |

### C. Things that may break (risks to watch)

| ID | Risk | What to do if it happens |
|---|---|---|
| R1 | **7 Oct 2026: limits added:** news history at most 15,000 stories, Archive at most ~300 MB (and 30 copies). **File size growth.** The news history grows by about 20 stories a day, up to roughly 2–3 MB at 400 days. It is embedded in `bluechip-board.html` and in **each** Archive copy (about 2.5 MB each today). The script keeps 30 copies, so the Archive, trimmed to 3 on 6 Oct, will grow back to 30. The data file is 2.2 MB and the site 2.5 MB. | Lower `$HistoricoDias`, store fewer fields, or keep the history out of the Archive copies. The 30-copy limit is `Select-Object -Skip 30` in the write step. |
| R2 | **ConvertFrom-Json on Windows PowerShell 5.1.** SEC `companyfacts` is 3.2–4.1 MB and parsed fine (about 1.1 s). If it grows a lot, 5.1 could hit its JSON size limit. | Switch to `companyconcept`, one request per tag (the option task 13 allowed). |
| R3 | **PowerShell 7 turns ISO date-time text into `DateTime`** when reading JSON. Everything reread from `bluechip-board-data.json` or `noticias-historico.json` goes through `ConvertTo-IsoUtc`, but date-only text (`yyyy-MM-dd`) stays text. | Any new date field reread from the data must be normalised the same way. |
| R4 | **External formats.** Each has validation and a fallback, but any of them can change: the SEC (submissions JSON, `companyfacts`, the `browse-edgar` 8-K Atom), the ECB Data Portal (series keys, CSV columns), iShares (holdings file layout), Yahoo, Nasdaq, Google News and FRED (already blocked). | Check *Sources & method* after a run; each source shows its error; the README lists every source. |
| R5 | **The site tests use the day's `bluechip-board-data.json`.** Most scenarios fix their own data, but a few read real prices (e.g. 9 rows in the prices table, today's EUR/USD). A real-world change (a source down at the last run) can make such a check fail without a code bug. | Rerun the script, then the tests; if needed, have the scenario fix its own data, as `taxBaseline` does. |
| R6 | **Duplicated rules** (B11): the alert thresholds and the rate-decision regex exist in two places. | Change both together. |
| R7 | **7 Oct 2026: mitigated:** without the Atom feed the times come from the filing headers (at most 12 a run, cached by accession). **`acceptanceDateTime` (B7):** if the SEC fixes it, nothing breaks (the Atom feed is still used). If the Atom endpoint is retired (audit F8), earnings sessions become Unavailable. | Move to another SEC source of acceptance times, such as the filing index pages. |
| R8 | **Headless Chrome / Edge** is needed by `Test-Site.ps1`. A site run takes about 4 min per shell, so a CI run takes about 15 min (Windows minutes count double on private repositories). | Run only the engine suite on pushes if the Actions minutes run short. |
| R9 | **The sample data for the tests** (`Tests\fixtures\bluechip-board-data.sample.json`, from 6 Oct 2026) is used in a fresh clone and in CI. As it ages, checks that compare with today's date could start to fail there but not on this PC, which has a fresh data file. | Refresh the sample from a recent `bluechip-board-data.json` (it holds public data only; check it has no `"backup"` key and no e-mail). |

### D. Dated upkeep (calendar)

| When | Item |
|---|---|
| Every few months | `$PesosReferencia` and each `Referencia` in `$ETFs` (reminder after 4 months); `$Script:UAChrome` / `$Script:UAData` (Chrome 154 on 7 Oct 2026; reminder after about 8 releases). |
| When a fund's top 10 changes | Add the holding to `$AliasesPosicoes`, with a known false positive in `Test-Engine.ps1` (reminder `Get-PosicoesSemAlias`). |
| By **Oct 2027** | Add the 2028 Fed and ECB decisions to `$Calendario`, from the official pages (reminder `Get-LembreteReunioes`, 60 days ahead). Also the companies' earnings dates as they are confirmed. |
| By **Jan 2028** | The exchange exception lists in `$Bolsas` end in 2027 (reminder at the start of the year). |
| Every year, before filing IRS | Compare `ANEXO_J` (Quadros 9.2A, 9.4A and 8A; codes G01, G20, E10, E11; the 365-day rule) with the new form and instructions. |
| Now and then | Empty the Recycle Bin only once you are sure you do not need the old snapshots or `prompts\` (see E); the current code is safe in git. |

### E. How to roll back

**From 6 Oct 2026 on, use git:** `git log` to find the version, then `git revert <commit>` (or `git switch -c look <commit>` to inspect an old state). Your data is not in git, so a rollback never touches it.

**Before the git import,** there is no rollback copy in the project folder. On 6 Oct 2026 the folder snapshots taken during the feature work were moved to the **Windows Recycle Bin**, together with `prompts\`. While the Recycle Bin is not emptied, they can be restored from there, back to `C:\Users\Admin\Documents\Scripts\`.

| Back to… | Snapshot |
|---|---|
| Before all feature work | `BluechipBoard_snapshot_20261006-1020` |
| End of each task | `BluechipBoard_snapshot_20261006-HHMM_tNN` (`_t01` … `_t15`); end of phase 1 = `_t06`, end of phase 2 = `_t10`, final = `_t15` |

From a snapshot, copy back:
- `Bluechip-Board.ps1`;
- `Tests\`;
- `README.md`;
- `AUDIT-REMEDIATION.md`.

Never copy back `bluechip-board-backup.json` (your current data), `bluechip-board.html`, `bluechip-board-data.json` or `Archive\`: they are rebuilt on the next run.

When going back to before task 11, `noticias-historico.json` (and any `.bad`) can be deleted. A newer backup (version 5, with `targets`, `policy`, `notes` or `fees`) still loads in every version from task 03 on. Unknown keys are ignored, but they stay in the file until you save from an older version. Run the three test suites, then the shortcut.

Once the Recycle Bin is emptied, the only way back is to undo changes by hand, using the change logs below. Hence D10.
---

## Status review · 5 Oct 2026

Every FIXED finding was re-checked against the current code and is still fixed: the tests that cover it still pass. The findings that were not fully fixed:

| ID | Status then | Still applies? | Notes today |
|---|---|---|---|
| H4a | PARTIALLY FIXED | Yes | Browsers still ask for folder permission again after each reopen; the status box still warns. |
| H4g | PARTIALLY FIXED | Yes | A backup saved in the wrong folder is still caught only at the next run. |
| H5 | PARTIALLY FIXED | Yes | Live prices still come from one unofficial source (Yahoo). Stooq still blocks automated requests. The three regional ETFs have no Stooq symbol and fall back straight to the previous run's data. |
| L8 | PARTIALLY FIXED (deliberate) | Yes | Code comments stay in Portuguese; the README now has a glossary of the Portuguese names. |
| L9 | PARTIALLY FIXED | Reduced | ETFs are data-driven since 5 Oct (`$ETFs` in the script, `ETF_IDS` in the page). A new **company** still needs edits in several places, listed in the README ("Adding an asset"). |
| F2 | PARTIALLY FIXED | Yes | `$Script:UA` still says Chrome 141; update it now and then. |
| F5 | PARTIALLY FIXED | Yes | Reference weights and the manual calendar are still manual. The reminders also cover each regional ETF's `Referencia`. |
| F7, F9, F10, F11 | PARTIALLY FIXED (external) | Yes | iShares, Nasdaq, news feeds and Chrome's folder rules are outside the project. All 71 sources answered on 5 Oct 2026. |
| F8 | NOT FIXED | Yes | The SEC `browse-edgar` Atom endpoint still works (3 of 3 OK on 5 Oct 2026). Under PowerShell 7 it needed `-SkipHeaderValidation` (see the PowerShell 7 change log). |
| F12 | NOT FIXED (legal) | Yes | Unchanged. The Anexo J export adds its own caveats. |
| F13 | PARTIALLY FIXED | Yes | Still a rule plus manual exceptions. |
| F14 | NOT FIXED | Yes, acceptable | With more feeds, a run now takes 55–60 s for about 1,300 stories. Still fine at the default 7-day window. |
| S2 | NOT FIXED | Yes | Browser design. |
| S3 | NOT FIXED | Yes | The SEC e-mail is still in plain text in the launcher (and in the task arguments, if a task is registered). |
| — | Known limitation (4 Oct change log) | Yes | The purchase form still refuses a new sale when *any* existing sale lacks matching purchases (`C.vendas.find(x=>x.falta>0)` in `buildBuys`). It only happens with old or restored data. |

Verification notes that changed since they were written:

- **H3:** *site/stale* now accepts 2 or 3 missing sessions, because the data can end on an open US session (an intraday point). It still checks that stale data is detected and reported, and that current data is not flagged.
- **H3, M4:** *site/newuser* now checks nine dated cards: the eight assets plus the EUR/USD card.

---

## Critical

### C1 · A future stock split silently corrupted saved purchases — FIXED

- **Finding.** Purchases kept the share count and price of the day they were entered. After a later split, Yahoo's adjusted price made them show a false loss of (1 − 1/ratio): about −90 % for a 10:1 split.
- **Root cause.** `SPLITS` was used only for a warning on the entry form; nothing adjusted saved data. Adding a split to `SPLITS` (the README's advice) did not help.
- **Change.**
  - Each new purchase and sale saves `r` (the adjusted close of its date in the history used when it was entered) and `u` (the date of that history).
  - When shown, `ajusteSplit()` compares `r` with today's history for the same date. A ratio of k (±3 %), equal to a product of the latest split events, means a k:1 split happened since entry: quantity × k, price ÷ k, cost unchanged. Without `r`, the split events after `u` are used.
  - An unexplained change produces a warning and no adjustment.
  - Split events now come from Yahoo (`events=split`), with `SPLITS_FIXOS` as a reserve; the reserve now includes Alphabet's 2014 distribution.
  - Existing entries (local and in backups) get `u`/`r` once, without changing `bb.savedAt`.
- **Files.** `Bluechip-Board.ps1` (`Get-Serie`, the run section and the template: `SPLITS`, `ajusteSplit`, `efetiva`, `completaBase`, `migraCompras`, `buildBuys`, `drawBuys`, `pfDados`).
- **Verification.**
  - *site/split* simulates a future 10:1 NVIDIA split, applied to the whole history as Yahoo would. Detected through `r` and through `u`; an entry made after the split is unchanged; an unexplained change gives a warning; the amount invested is unchanged; the value equals the value before the split exactly.
  - *site/migrate* and *site/v1backup* check the migration of old entries.
  - *engine* checks how split events are read and serialised.
- **Notes.** Purchase quantities are still entered in today's shares, as before; the entry form explains this.

## High

### H1 · With no EUR/USD rate, dollars were shown and valued as euros — FIXED

- **Root cause.** `inCur()` returned the original series when `FX` was empty.
- **Change.**
  - `inCur()` returns no data when it cannot convert, including an unknown currency, or when the last point has no rate (so an older price never passes for the last one).
  - A red alert is raised, the portfolio shows "no price" and the Prices table explains why.
- **Verification.** *site/fxnone*.

### H2 · The ECB fallback held one rate for everything older than 90 days — FIXED

- **Root cause.** `fxAt()` returned the first rate for any earlier date.
- **Change.**
  - Fallback order: the Yahoo 1-year series, then the long Yahoo EUR/USD history, then the ECB rates.
  - `fxAt()` returns `null` outside the series, with a 5-day tolerance for weekends and holidays.
  - `stats()` hides "1 year", "year to date", "52-week" and "max drawdown" when the data does not cover the period.
  - An alert names the fallback in use.
- **Verification.** *site/fxlong*, *site/fxecb*, *site/stats* (year to date when the series starts in the current year).

### H3 · No price date, and old data presented as current — FIXED

- **Change.**
  - `frescura()` compares each series with its exchange's last completed session at build time, using the rule-based holiday and early-close calendar. Bitcoin is checked by UTC days; EUR/USD by London weekdays.
  - Every card and Prices row shows "close D Mon", "live HH:MM" (the session was still open when the data was collected) or "previous run".
  - Data that is behind or from the previous run is shown in orange and raises an alert.
  - `Get-Serie` returns `parcial` (intraday) and `hora` (time of the price).
- **Verification.** *site/stale* (3 missing sessions detected; current data not flagged), *site/prev*, *site/newuser* (5 dated cards).

### H4 · Backup integrity

| ID | Finding | Change | Verification | Status |
|---|---|---|---|---|
| H4a | Automatic saving quietly stopped after a reload | `#bkState` status box: "Some changes are not in the backup file yet… click Save" whenever the browser has changes the file does not; it confirms when the file is up to date. | *site/legacy*, *site/v1backup* | **PARTIALLY FIXED**: the browser asks for permission again after each reopen (a browser security rule); the page now says so, so nothing is lost silently. |
| H4b | A restore could be undone at the next run | Restore now **merges**, and the result is newer than the file, so it is kept and saved. The project file's state (`bb.fileSaved`) only changes for the project file. | *site/restore* | FIXED |
| H4c | No confirmation for restore or delete | `confirm()` before restoring and before deleting any purchase or sale. | Code path; `confirm` is stubbed in *site* | FIXED (there is no undo; the backup file and `.previous.json` act as the recovery path) |
| H4d | Two browsers: the last one to save won | Merge by id (`juntaBackup`): entries missing on either side are added and nothing local is lost. | *site/merge* | FIXED |
| H4e | Deleted purchases came back from the file | Deletion markers (`bb.deleted`, saved in the backup as `deleted`). One-time legacy rule: on the first page load with this version, if the browser holds data newer than the file and has no deletion list yet, file entries missing from the browser are treated as deleted (the old precedence). | *site/tomb*, *site/legacy* | FIXED |
| H4f | A file from Downloads overwrote the backup before it was checked | The file is checked first (valid JSON, `app`) and must be newer by its own `saved` / `exported` date, not by file date. The previous file is kept as `bluechip-board-backup.previous.json`. An invalid file is left in Downloads with a warning; `catch {}` no longer hides errors. | *resilience*: invalid newest file, valid newer file, older-but-recently-written file | FIXED |
| H4g | A backup saved in the wrong folder failed silently | `verificaPasta()`: if the file this browser saved is not the one the script found at the next run, the page tells you to save it in the BluechipBoard folder. The save message says where the file must be. | *site/wrongfolder* | **PARTIALLY FIXED**: browsers do not reveal the chosen folder, so the problem is caught at the next run, not when saving. |

Backup format: version 4 adds `sales`, `deleted` and `r` / `u` on entries. Versions 1 and 3 still load (*site/v1backup*), and old archived pages ignore the new fields.

### H5 · Yahoo was the only price source; the Stooq fallback no longer worked — PARTIALLY FIXED

- **Change.**
  - When a price or history download fails, the previous run's data (`bluechip-board-data.json`) is used. It is labelled "previous run" with its date, the source row says so, and it raises an alert.
  - Stooq's anti-bot page is now reported clearly ("returned a web page instead of CSV").
  - Responses are validated: symbol, currency, matching timestamp and price counts, zero/negative/NaN prices dropped.
- **Verification.** *resilience* (every source blocked, with and without previous data), *engine* (canned Yahoo answers: wrong symbol, wrong currency, invalid prices, Stooq page).
- **Notes.** Live prices still depend on one unofficial source. A second live source would need an account or key, which the project deliberately avoids.

### H6 · Stories assigned to an asset because of the feed they came from — FIXED

- **Change.** In search feeds (Google News, Yahoo Finance), a story with no keyword is only a weak match: tagged "Feed match only", no company bonus, at most moderate. Primary feeds (newsrooms, SEC, Fed, ECB) keep the full assignment.
- **Verification.**
  - *engine* ("China's AI agents…" from the Apple feed → AAPL, weak, moderate at most; the Apple Newsroom story stays a full match).
  - Real run: 67 stories flagged, and none of them is orange or red.

### H7 · Over-scoring and look-alike words — FIXED

- **Change.**
  - Themes now count as: the strongest in full, each extra one at half, up to 5 (was: all added, up to 7).
  - Severity needs context: "delays … launch", trade/US "curbs", "warns of/on", profit warnings. Added: arrests and smuggling.
  - "Curb" counts in the China theme only as trade curbs.
  - "Forecast" and "outlook" count as Earnings only next to earnings words.
  - "Bitcoin treasury" is no longer Macro.
  - Exclusions: "Big Apple" and other companies called Apple; "iShares Bitcoin/Gold…" is not SXR8; MiCA only in its exact spelling.
- **Verification.**
  - *engine*: 16 classification cases covering every audit example.
  - Real runs, before → after (different news days, same 7-day window):

    | Level | Before | After |
    |---|---|---|
    | Material | 32 | 7 |
    | Important | 200 | 156 |
    | Moderate | 380 | 439 |

  - The remaining material stories are a smuggling case, an import ban and US curbs.

### H8 · A Nasdaq estimate replaced a confirmed manual date — FIXED

- **Change.**
  - A manual `C` date is never replaced by a Nasdaq `E`.
  - Nasdaq dates must fall between −2 and +200 days.
  - Malformed manual dates are skipped with a warning instead of stopping the run (found during the work).
- **Verification.** *resilience*: confirmed date kept; estimated date replaced; confirmed Nasdaq date applied; malformed date skipped and the run continues.

## Medium

| ID | Finding | Change | Verification | Status |
|---|---|---|---|---|
| M1 | The tax counter could not record sales | Sales (`bb.sales`) for every asset, matched first in, first out by `carteira()`. Holdings, cost, realised gain and the Bitcoin tax counter use what is left; overselling is refused. The boundary day is explained in the page and the README. | *site/fifo*: FIFO quantities, cost, realised gain, tax counter, oversell | FIXED. Notes: the legal reading of "365 days" cannot be verified from the project, so the page advises a margin. |
| M2 | Bitcoin alert said "over the last day" but used the move since 00:00 UTC | Uses CoinGecko's 24 h change when available, with the correct label otherwise. | Code path | FIXED |
| M3 | Drop-from-high chart wrong for 2014–2017 | The chart subtitle and the README explain that the history starts after the 2013 peak; a drawdown starting at the first data point is labelled "start of the data, not a real peak". | Code path, screenshot | FIXED (the pre-2014 data is not available from the source, so the representation was corrected rather than the data) |
| M4 | Today's intraday price was called a "closing price" | The purchase form names an intraday price as such, with its time; cards show "live HH:MM" and "so far today". | *site/newuser* (dated cards); code path | FIXED |
| M5 | `vistos.json` was written before the outputs | It is written last, safely. | *resilience*: failed run leaves it untouched | FIXED |
| M6 | Portfolio copied into every output; page wording misleading | The backup is embedded only in `bluechip-board.html` (not the Archive or the data file); the Portfolio intro and README explain it. | *resilience* (backup only in the main site) and the real output | FIXED |
| M7 | The exposure chart did not say when reference weights were used | Dynamic subtitle naming the weights' source and date, with their age if over 10 days; same in the ETF tab. | Code path, screenshot | FIXED |
| M8 | Scheduled task skipped on battery, visible window, quoting of a trailing backslash | `-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries`, a 1 h limit, `-WindowStyle Hidden`, backslashes doubled before quotes, English description. | *resilience* dry run (arguments and settings). The task created by the first, faulty attempt at this test was inspected and removed; the test is now isolated: unique task name, stub loaded after the module, clean-up guard. | FIXED (it still runs only while you are signed in, as before; documented) |
| M9 | Duplicate groups could chain beyond 3 days | Each group tracks its earliest and latest date; joins and merges that would span more than 3 days are refused. | *engine* (chain of 3 stories over 4 days); real run: no group over 3 days | FIXED |

## Low / technical debt

| ID | Finding | Change | Status |
|---|---|---|---|
| L1 | Alphabet's 2014 Class C distribution missing from splits | From Yahoo (1998/1000), plus the reserve list. | FIXED |
| L2 | Filter button "Exchange-Traded Funds" also filtered macro news | Renamed "ETF & market". | FIXED |
| L3 | "50-day / 200-day" labels on session counts | "50-session / 200-session", with a note on the counting. | FIXED |
| L4 | Feeds always read as UTF-8; few time-zone abbreviations; undated items bypassed the window | Charset from the header or XML prolog (*engine*); CET/CEST/BST/GMT/CDT/CST/EET/EEST/JST; undated or far-future items skipped and counted in the source row. | FIXED |
| L5 | No guard on `JSON.parse` | A visible message instead of a blank page. | FIXED (*site/corrupt*) |
| L6 | Fear & Greed thresholds differed from the index's bands | Alternative.me's own bands (`fgClasse`) for colours and alerts. | FIXED |
| L7 | Non-atomic writes | `Write-Atomico` (temporary file, then replace). | FIXED (*engine*, *resilience*) |
| L8 | Leftover Portuguese | Scheduled-task description in English. Code comments stay in Portuguese, matching the existing code (the README documents everything in English). | PARTIALLY FIXED (deliberate) |
| L9 | Asset list repeated in many places | `COMPRAVEIS` / `BUY_IDS` now derive from `PF`; the README lists every place to edit when adding an asset. | PARTIALLY FIXED (a single source of truth would need restructuring the template, which was out of scope) |

## Future-failure risks

| ID | Risk | Change | Status |
|---|---|---|---|
| F1 | TLS pinned to 1.2 on `SystemDefault` | TLS 1.2 is added only when the .NET setup lists protocols explicitly without it (PS 5.1 here reports `SystemDefault`). | FIXED (*engine* static check) |
| F2 | Hard-coded Chrome/126 User-Agent | Updated to Chrome 141, with a comment; listed in the maintenance checklist. | PARTIALLY FIXED (still a constant by nature) |
| F3 | Halving marks and fallback date frozen in time | Past halvings: fixed list plus lookup of any new one on mempool.space; fallback estimated from the last known halving (*engine*: before and after 2028). | FIXED |
| F4 | Halving forecast at a fixed 10 min | Average block time since the last halving (9.95 min now), kept within 8–12 min, shown in the site. | FIXED |
| F5 | Manual upkeep | Maintenance reminders in the console and in Sources & method (old reference weights, empty manual calendar, old exception lists). | PARTIALLY FIXED (still manual data) |
| F6 | Manual `SPLITS` | Yahoo split events (see C1). | FIXED |
| F7 | iShares URL and format | Validation: weights add up to 80–120 %, each of the three is plausible; the "as of" date is parsed and its age shown. | PARTIALLY FIXED (external) |
| F8 | SEC `browse-edgar` Atom endpoint | Unchanged: it works, and a migration to the SEC's JSON API would be a rewrite of that source. Documented. | NOT FIXED (external; works today) |
| F9 | Nasdaq endpoint and format | Plausibility window for the date. | PARTIALLY FIXED (external) |
| F10 | News feeds and Google News throttling | Longer wait after 429/503; per-feed errors visible. | PARTIALLY FIXED (external) |
| F11 | Chrome's folder-saving rules | Status box, download fallback, wrong-folder detection. | PARTIALLY FIXED (external) |
| F12 | Portuguese crypto tax rules may change | Text tells you to confirm with the Portal das Finanças or an accountant. | NOT FIXED (legal, outside the code) |
| F13 | Xetra year-end session | Last trading day before 31 Dec closes at 14:00 by rule (Deutsche Börse's usual practice, confirmed each year by circular); override with `curtos`. | PARTIALLY FIXED (*site/cal*) |
| F14 | Duplicate-grouping cost grows with more stories | Measured: 1,655 raw stories; a whole run takes 44–48 s with all downloads. No change. | NOT FIXED (acceptable at the default window; may slow down at `-Days 30`) |

## Security

| ID | Finding | Change | Status |
|---|---|---|---|
| S1 | A file planted in Downloads could replace the backup | Validated before use, previous file kept, only newer by its own date (H4f). A well-formed file can still add entries, now through a merge that keeps everything existing. | FIXED (to the extent possible locally) |
| S2 | `file://` pages share localStorage in Chrome | — | NOT FIXED (browser design; documented) |
| S3 | SEC e-mail in plain text in the launcher and task arguments | — | NOT FIXED (it is your own contact, sent to the SEC by design; documented) |
| S4 | XML parsed with default settings | `ConvertTo-XmlSeguro`: DTDs ignored, no resolver, entity limit. | FIXED (*engine*: entity expansion test) |
| — | Headline injection | Re-tested: a hostile headline goes through the real `<` escaping, and no script runs, no element is injected and `javascript:` links are neutralised. | Still no vulnerability (*site/xss*) |

## Found during the remediation

| Item | Change | Status |
|---|---|---|
| Windows PowerShell 5.1 wrote ECB rates (and the new split and Fear & Greed lists) as `{value, Count}` objects | Pairs rebuilt in plain loops; the output no longer contains that form. | FIXED |
| Calendar countdowns rounded hours and used the PC's time zone | Counted in Lisbon calendar days (`diasAte`). | FIXED (*site/cal*) |
| Chart axis dates used the PC's time zone | UTC date parts (the series are dates). | FIXED |
| A malformed manual calendar date stopped the whole run | Skipped with a warning. | FIXED (*resilience*) |
| The euro equivalent on a card used today's rate instead of the price's day | Same rate as the portfolio (*site/stats*). | FIXED |

## Testing gap from the audit

**FIXED.** The three suites above (158 checks) now cover the engine, the website and the failure modes, on PowerShell 5.1 and 7. They need no internet (except Chrome or Edge for the website) and never touch project data, Downloads or Task Scheduler.

---

## Change log · 4 Oct 2026 · Dividend tracker and Anexo J export

Two features added on top of the remediated version, extending the existing architecture rather than replacing it. The counts in the test table at the top are from the remediation; current counts are below.

### Files

| File | Change |
|---|---|
| `Bluechip-Board.ps1` | Added `$AtivosDividendos`, `Get-Dividendo` and `Get-Dividendos`, the dividend fallback in the run section, and `dividendos` in the data (`versao` 1.2). Template: Dividends section, tax export block, JavaScript for both. Small extensions listed below. |
| `README.md` | Dividend and Anexo J sections, data source, fallback, configuration, tests, maintenance, troubleshooting. |
| `Tests\Test-Engine.ps1`, `Tests\Test-Site.ps1`, `Tests\site-scenarios.js`, `Tests\Test-Resilience.ps1` | New checks (below). |
| Not changed | `Start-BluechipBoard.ps1`, `bluechip-board-backup.json` (hash checked before and after the real runs), `Bluechip-Board.ico`, the scheduled-task code, the backup and merge code, `localStorage` keys. |

### Changes to existing code

Every existing line that changed, and why:

| Where | Change | Effect on existing behaviour |
|---|---|---|
| `fxAt` | Split into `fxPt` (returns the `[date, rate]` point) + `fxAt` (returns the rate, same rules). | None: same result. The export needs the date of the rate it uses. |
| `carteira()` | Each sale also lists the purchases it used (`usados`: id, date, quantity, unit cost, split factor). | None: same lots, cost, realised gain and shortfall. This makes it the single FIFO for holdings, the 365-day counter and the tax export. |
| `drawPf()` / `drawBuys()` / `init()` | Also draw the Dividends section, the tax-year list, and set up the export button. | None on the existing tables. |
| `descarrega()` | Uses the new `baixa()` download helper (same steps). | None: the backup download is identical. |
| `fdt()` | An invalid date shows "—" instead of throwing. | Only for invalid dates: a backup with a date such as `2026-13-45` passes the shape check in `limpaBackup` and used to stop the whole page (found by the new malformed-data test). |
| Test hook | Exposes the new functions. | Tests only. |
| Bitcoin tax note | Says Annex G (Portuguese platform) or Annex J, Quadro 9.4A (foreign), and Annex G1 for exempt sales. | Text only; the previous text named only Annex G. |
| `versao` | `1.1` → `1.2` (one new field). | Nothing reads it; older pages ignore the new field. |

### Decisions and assumptions

- **Dividend source.** Yahoo's `quoteSummary` and `quote` endpoints answer 401 (Unauthorized) without a session cookie and crumb (checked on 4 Oct 2026), so they are not used. The public chart endpoint, the one the prices already use, returns dividend events with `events=div`. It is called once per company with the existing `Get-Url` (same retries, headers, decoding and error reporting), separately from the prices, so a failure affects only the dividends.
- **Annual dividend** = latest payment × payments per year (the "indicated" rate), with the trailing 12-month sum stored too. **Yield** is computed by the script in percentage points, so no provider decimal-or-percent ambiguity can creep in.
- **Fallback** follows the previous-run design used for prices: the previous dividend data is reused for up to 180 days, labelled "previous run (retrieved …)" with its original retrieval date. Older data, invalid values or no data → "Unavailable". Nothing is invented and €0 is never shown for missing data.
- **Holdings** for dividends come from `pfDados()`/`carteira()`: no second calculation.
- **Dividend FX:** the latest EUR/USD of the page, because the payments are in the future; the rate, its date and source are shown.
- **Tax values are the euro prices the user entered.** The register stores euro prices (pre-filled at that day's EUR/USD, editable). The export uses them as they are, never a market close. No exchange rate is applied to the tax amounts, so the "today's rate for history" error cannot happen.
- **Dollar columns** in the stock file are a reference: euro amount × the EUR/USD of each transaction's own date, looked up separately for acquisition and sale through `fxPt` (long-term Yahoo history, then the page's own EUR/USD series). There is no lookup outside the covered period. A missing rate leaves the columns blank with `NO_FX_REFERENCE`; today's rate is never substituted.
- **No timestamps added.** Transactions are date-only, and the FX history is daily closing rates, so intraday precision would be fabricated. The limitation is documented. Adding an optional time field would not change the tax values, which are the euro prices entered, so it was not needed. The backup schema is therefore unchanged.
- **Official form.** *Modelo 3 – Anexo J*, "modelo em vigor a partir de janeiro de 2026", and its instructions (AT): Quadro 9.2A (codes G01, G20; columns including "País da Contraparte" and "admitidos à negociação ou partes de OIC abertos") and Quadro 9.4A (crypto-assets held less than 365 days). Source: https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/modelos_formularios/irs/Documents/Mod_3_anexo_J.pdf. The mapping is in one object (`ANEXO_J`).
- **No official CSV import** for these tables was found. The Portal takes the declaration itself (online or as the XML file it saves), so the files are labelled preparation files everywhere.
- **Interpretations left visible, not hidden:**
  - "País da fonte" is the issuer's country (840 for the US stocks, 372 for SXR8).
  - G20 is used for SXR8, a UCITS ETF that is legally a share of an Irish investment company; some advisers use G01.
  - Bitcoin held 365 days or more is exempt and goes to Anexo G1, not Anexo J.
  - A Portuguese crypto platform means Anexo G.
  - The 365-day boundary is the counter's own rule, with a REVIEW flag within 2 days of it.
  - Fields the app cannot know are left blank: País da Contraparte, expenses, foreign tax, and the crypto "País da fonte".
- **CSV format:** `;` separator, decimal comma, no thousands separator, fixed decimals, ISO dates plus the form's Ano/Mês/Dia, UTF-8 with BOM, CRLF, quoting, and a formula-injection guard. Numbers are formatted by the script, not the browser's language, so the output is deterministic.
- **Privacy:** the export runs entirely in the browser, downloads two local files, and reads `localStorage` without writing to it (a test compares it before and after).

### Tests executed (all passing)

| Suite | Checks | New | Result |
|---|---|---|---|
| `Test-Engine.ps1` | 60 | 15 (dividend collector) | Pass on PowerShell 5.1.26100 and 7.6.6 |
| `Test-Site.ps1` | 155 | 71 in 13 scenarios (`div`, `divnone`, `divfail`, `divbad`, `divstale`, `divfx`, `tax`, `taxsplit`, `taxcrypto`, `taxold`, `taxbad`, `taxfx`, `taxfx0`) | Pass; the 84 existing checks unchanged and passing |
| `Test-Resilience.ps1` | 33 | 4 (dividend source down; previous-run reuse; 180-day limit; source status) | Pass on PowerShell 5.1 and 7; no scheduled task left registered |

Also:

- A real networked run on PowerShell 5.1 in the project: 56 of 57 sources OK, with SEC skipped because the run had no e-mail; the 3 dividend sources OK.
- A real run on PowerShell 7 in a temp folder, with the same result.
- No `{value, Count}` pairs in the data from either shell.
- No new external resources in the page.
- The backup file was unchanged by both runs.
- A visual check of the Portfolio tab with sample data, and sample CSVs inspected.

### Known limitations

- Dividends rely on one unofficial source (Yahoo's chart API); the previous-run fallback covers outages up to 180 days.
- Dividends are a projection from the latest payment, before tax. Announced changes are not known in advance.
- The tax files are preparation aids, not an import format, and are not tax advice. Interpretations above should be confirmed for the user's case.
- Fees and foreign tax are not stored by the register, so they must be added by hand.
- Found, not changed (outside this scope): in the purchase form, a sale is refused with "This sale would leave a later sale without enough units" whenever **any** existing sale already lacks matching purchases (possible only with old or restored data), even for another asset. The check uses the first shortfall in the whole list (`C.vendas.find(x=>x.falta>0)` in `buildBuys`).

---

## Change log · 5 Oct 2026 · Three more iShares ETFs (EUNK, IS3N, EUNN)

Added iShares Core MSCI Europe (IE00B4K48X80), Core MSCI EM IMI (IE00BKM4GZ66) and Core MSCI Japan IMI (IE00B4L5YX21). They go through the same mechanisms as SXR8 wherever possible. The SXR8-specific code was generalised to a list of ETFs, with SXR8 first and unchanged.

### Decisions

- **Listings and ids.** All three use their Xetra listing in euros: `EUNK.DE`, `IS3N.DE`, `EUNN.DE`. That is the same exchange, currency and pipeline as `SXR8.DE`, so there is no FX conversion and no new `$Bolsas` entry.
  - Each listing was checked on Yahoo on 5 Oct 2026: name, `EUR`, `XETRA`, `Europe/Berlin`, history since 2010, or Jun 2014 for IS3N. The candidates in other currencies were rejected, for example `EIMI.L` in USD and `SJPA.L` in GBp.
  - The Xetra tickers are the internal ids, like `SXR8`. `EIMI` was not used: it is the London USD ticker.
- **Holdings.** Each fund has its own official iShares file: the same "Detailed Holdings and Analytics" download with its own `portfolioId` (251861, 264659, 251867).
  - The four files have identical sheets and columns (checked on 5 Oct 2026). The ISIN, base currency, benchmark, accumulating status and Irish domicile were confirmed in each file's Key Facts.
  - None of the three new funds holds AAPL, NVDA or GOOGL, so their reference weights for those are 0%.
- **One configuration list, `$ETFs`.** It holds each fund's metadata, holdings URL, reference weights and expected companies. `$UrlPesosETF` and `$PesosReferencia` stay as they were and are what the SXR8 entry uses. The script writes `etfs` (all four, with metadata) and still writes `etf` (SXR8, the earlier format). The data version is now `1.3`.
- **Holdings parser** (`Get-PesosETF`) now takes one ETF; called with no argument it still reads SXR8. Added:
  - a check that the ISIN in the file is the fund's;
  - per-fund plausibility rules: in SXR8 the three companies must each weigh more than 0 and at most 25%, as before; in the other funds they may be 0;
  - short US ticker aliases applied only to holdings traded in USD.

  Each fund has its own source row ("iShares: ETF holdings (ID)") and its own reference fallback; one failure does not affect the others. The fallback design is unchanged: reference weights, not previous-run holdings.
- **News.** One Google News feed per new ETF. Keywords match the fund itself and the whole market its index covers (for example "European shares", "Nikkei 225", "emerging-market stocks"). A single company in that market does not match, and neither does "Nikkei reports". As with SXR8, there is no company bonus.
  - The SXR8 pattern now excludes "iShares (Core) MSCI …", so stories about the new funds are not tagged SXR8.
  - The existing classification of AAPL, NVDA, GOOGL and BTC is unchanged (engine tests).
- **Website.**
  - The ETF list comes from the data. The `CO` map gained colour tokens (`--eunk`, `--is3n`, `--eunn`).
  - The three filter chips are built from the configuration. "ETF & market" (SXR8 + market) is unchanged.
  - Cards, Prices, alerts (stock thresholds), Portfolio lists (`PF`), `HIST`, `BOLSA_DE` and the Bitcoin correlation now include the new ETFs automatically.
  - The **ETFs** tab is one reusable view with a fund selector; with SXR8 selected it shows the same texts and figures as before. The "three companies" figure is shown only for funds that hold them.
  - Each ETF has its own simulator, built by the same `buildSim` (`bb.sim.etf` for SXR8, unchanged; `bb.sim.etf-<id>` for the others).
  - Look-through now covers every ETF held: the company weights of each fund, and the rest of each fund once under its index name. With SXR8 alone the result and text are identical to before.
- **Dividends.** The new ETFs are excluded both in the script (`Get-Dividendos` skips any accumulating ETF, even if added to `$AtivosDividendos`) and in the website (`DIV_IDS`). AAPL, NVDA and GOOGL are unchanged.
- **Anexo J.** `ANEXO_J` (still the single source of truth) gained EUNK, IS3N and EUNN.
  - Mapping: 9.2A, G20, País da fonte 372 (Ireland), "Sim", with each ISIN. They have the same legal form as SXR8: UCITS sub-funds of Irish iShares investment companies, listed on Xetra.
  - The "Price basis" column now follows the asset's currency instead of the SXR8 id; the result is the same for SXR8.
  - New flag `NO_ANEXO_J_MAPPING` for any future asset without a mapping. The CSV format, columns and other flags are unchanged.
- **FIFO and splits.** No change: `carteira()` and `efetiva()` already work by asset id. The ETFs are not in `SPLITS_FIXOS`, and Yahoo reports no splits for them; an adjustment would still be detected from the reference price.
- **localStorage and backup.** No new keys or fields (apart from the simulator inputs above); the ETFs use `bb.buys` and `bb.sales` with their ids. Old backups load unchanged (`v1backup`, `etfbackup` scenarios).

### Existing lines changed and why

| Where | Change | Effect on existing behaviour |
|---|---|---|
| `Get-PesosETF` | Optional ETF parameter, ISIN check, per-fund rules, USD-only aliases, source name with the id. | SXR8: same weights, top 10, sectors and fallback. The source row is now "iShares: ETF holdings (SXR8)" instead of "iShares: ETF holdings", so the four funds can be told apart. |
| `$EmpresasRe.SXR8` | Excludes "iShares (Core) MSCI". | Only stories about other iShares MSCI funds stop being tagged SXR8. |
| `New-DadosSerie` | Adds `tipo` and `bolsa` (ETFs only). | Additive fields. |
| `Get-Dividendos` | Skips accumulating ETFs. | None for AAPL, NVDA, GOOGL. |
| Template, SXR8-specific expressions | `a.id==='SXR8'` → `ehEtf(a.id)` in card names, alerts, `nmOf`, allocation, "Xetra closing price", price basis; `HIST` keys from the data; `PF` extended from the data; `BOLSA_DE` extended from the data. | SXR8 renders as before. |
| `renderEtf`, `drawEtfMore`, `etfIdade` | Parameterised by the selected ETF (`etfSel`, default SXR8). | SXR8 view identical (test `etfui`, "SXR8 view as before"). |
| `drawPf` look-through | Loops over the ETFs held. | With SXR8 alone: same values and subtitle. |
| `alerts()` news categories, `drawNiveis` | Add the new ETFs (the chart only when they have news). | The existing categories are unchanged. |
| Header and labels | Tab "iShares Core S&P 500" → "ETFs". Title, brand line, hero title and the Xetra "detalhe" mention the ETFs. The dividends note and the Bitcoin and Prices texts are made plural. | Text only. |
| Tests `newuser` | "five price cards" → the exact list of eight; "6 rows" → 9. | The expected values follow the three added assets, and the check is now stricter. |

### Tests executed (all passing)

| Suite | Checks | New | Result |
|---|---|---|---|
| `Test-Engine.ps1` | 92 | 32 (registration and metadata, listing/currency/price validation, holdings parser on test copies of the iShares format, wrong fund, malformed, bad weights, source down, SXR8 rules, four funds with one failing, dividends exclusion, news) | Pass on PowerShell 5.1 and 7 |
| `Test-Site.ps1` | 210 | 55 in 7 scenarios (`etfui`, `etfpf`, `etfbackup`, `etfmissing`, `etfstale`, `etfhold`, `etfbad`) | Pass; the 155 earlier checks pass (2 count checks updated, see above) |
| `Test-Resilience.ps1` | 37 | 4 (new ETFs with every source down, their own reference weights, SXR8 fallback unchanged, each ETF's previous-run prices and history) | Pass on PowerShell 5.1 and 7; no scheduled task registered |

Also:

- Real runs: PowerShell 5.1 in the project (68 of 69 sources, 62 s) and PowerShell 7 in a temp folder (68 of 69, 55 s); SEC was skipped in both because no e-mail was given. All 12 new sources OK.
- No `{value, Count}` pairs in the data from either shell.
- The generated `bluechip-board.html` was opened in headless Chrome with no JavaScript errors.
- A visual check of the ETFs tab (EUNK) and the Portfolio with sample data.
- The backup file was unchanged by the runs.

### Known limitations

- **Runtime:** one run went from about 45 s to 55–62 s, because of three more holdings files (up to 3.7 MB each) and about 180 more stories a week.
- **Market news also counts as market:** broad market headlines for Europe and Japan also match Market & macro, so the "ETF & market" filter shows more stories than before.
- **No Stooq symbols** for the new ETFs: Stooq blocks automated requests anyway. They fall back to the previous run's prices.
- **Reference weights:** the new funds' reference weights cover only the three panel companies (0%). If a holdings download fails, their top 10 and sectors show "Unavailable", as for SXR8.
- **Old pages and new backups:** a page built before this change, if given a backup with the new ETFs, drops those entries on merge, because its asset list does not know them. Always use a page built by the current script.
- **Tax mapping to confirm:** the G20 and País da fonte interpretations recorded for SXR8 apply to the new funds too, so confirm them for your case.

---

## Change log · 5 Oct 2026 · Desktop shortcut on PowerShell 7

- **Shortcut.** At the user's request, `Desktop\Bluechip Board.lnk` now targets `C:\Program Files\PowerShell\7\pwsh.exe`, with the same arguments, folder and icon. The original shortcut was copied aside before the change.
- **Found while checking it (FIXED).** Under PowerShell 7, the three SEC EDGAR sources failed with "The format of value '<e-mail>' is invalid". PowerShell 7 validates the User-Agent header, and the "@" of the e-mail the SEC requires is not valid there; Windows PowerShell 5.1 has no such check.
  - Earlier PowerShell 7 test runs had not shown this, because they ran without `-SecEmail`.
  - Fix in `Get-Url`: on PowerShell 6 and later, `-SkipHeaderValidation` is passed. The header sent is unchanged, and 5.1 is untouched.
- **Verified:**
  - SEC sources OK on 5.1 and 7.
  - The launcher run under PowerShell 7 with the e-mail, on a temp folder without opening the browser: 71 of 71 sources OK, and the e-mail does not appear in the page.
  - Engine tests 93 (one new static check) on both, site 210, resilience 37 on both; no scheduled task registered.

---

## Change log · 5 Oct 2026 · Market snapshot layout, EUR/USD card, footer and documentation

Visual and documentation changes requested by the user. No change to the data, the collectors, `localStorage` or the backup format.

### Changes

- **Snapshot grid (screens of 1,600 px or more).** The left sidebar is a 6-column grid:
  - the S&P 500 ETF across the top;
  - the three stocks in one row;
  - the three regional ETFs in the next row;
  - Bitcoin across the bottom.

  Below 1,600 px the sidebar stays a single list, with SXR8 moved to the top. Placement is set by `HERO_LUGAR`.
- **EUR/USD card** under Bitcoin (`cartaoFx`). It uses the same series and source as the conversions, with its date, change, sparkline, 1 month, year to date, 52-week range and ECB reference rate. The chart in *Currency & Macroeconomics* is unchanged. With no rate, the card says so and shows no price.
- **Proportions and width.**
  - The sidebar fills the screen height, and the free height is shared between the rows: `1fr` for each full-width row, `1.6fr` for each row of three. Charts grow with their card; the SVG is positioned absolutely, so its aspect ratio cannot stretch the card.
  - Sidebar width (`--side`): 29 % of the window from 1,600 px, 36 % from 2,400 px and 40 % from 3,000 px.
  - Side margins (`--gut`): 32 px, or 44 px from 3,000 px. The maximum page width was raised.
  - From 2,400 px the header's brand column has the sidebar's width, so the asset filters ("All", …) line up with the section column.
- **Footer.** Removed "71 sources · Yahoo Finance, ECB, iShares, SEC, CoinGecko"; *Sources & method* still lists every source.
- **Scheduled-task description** now names the iShares ETFs. It is text only: the task name and settings are unchanged.
- **README** reorganised into "Using it" and "How it is built". New parts:
  - system and run diagrams, plus price and EUR/USD fallback charts;
  - dependencies (local and internet) and a code map;
  - data formats, "Making a change safely", with the lines the tests depend on;
  - a glossary of the Portuguese names.

  Stale details were corrected: run time, the alerts table and the layout widths.

### Checks

- At each width from 1,600 to 5,120 px, a headless-Chrome measurement script checked four things:
  - that the filters and the tab bar are not clipped;
  - that there is no horizontal scroll;
  - the card and chart sizes;
  - from 2,400 px, the filter/section alignment (within 4 px: the filter bar's own padding).

  Screenshots were checked at 1,920, 2,560 and 3,840 px.
- **Found during the work:** at 1,920 px the tab bar was already clipped by about 25 px ("Sources & method"); the 29 % sidebar fixes it. Between 1,600 and about 1,900 px, the filter bar and tab bar still scroll sideways, as they did before these changes. The sidebar there is at its 560 px minimum, the same width as before.
- **Tests:** site 214 (4 new checks in *newuser* and *fxnone* for the EUR/USD card and the layout classes), engine 93 on PowerShell 5.1 and 7, resilience 37. *site/stale* was made independent of the time of day (see the status review). The backup file was unchanged by the real runs.

---

## Change log · 6 Oct 2026 · Feature work, phase 0: baseline and Anexo J non-regression test

This work came before any new feature. It followed a task plan kept in `prompts\` (tasks 01–15), with progress in `prompts\PROGRESSO.md`. That folder was removed on 6 Oct 2026 (sent to the Recycle Bin; see [Open items](#open-items--6-oct-2026)).

- **Baseline.** All suites were green before any change: engine 93, site 214, resilience 37, on PowerShell 7 and 5.1.
- **Restore point.** There is no git on this PC, so the whole folder was copied to `..\BluechipBoard_snapshot_20261006-1020`.
- **New test `taxBaseline`** (site). It uses a fixed test backup:
  - several lots;
  - sales in 2025 and 2026;
  - FIFO across two lots;
  - a 10:1 NVDA split;
  - SXR8 and EUNK;
  - Bitcoin held under and over 365 days.

  The scenario brings its own price history, EUR/USD and split data, so the result does not depend on the day's data. The four CSV files (stocks/ETF and crypto, 2025 and 2026) must be identical, byte by byte, to the reference files in `Tests\fixtures\taxBaseline\`, and the export button must download files of the same size.
  - The reference files were generated once, with `Test-Site.ps1 -WriteFixtures`, which never overwrites an existing file. Their values were checked by hand: FIFO amounts, split factor, days held and the 365-day boundary.
  - A one-byte change in a reference file makes the test fail, and points to the byte that differs.
- **Test harness.** `Test-Site.ps1` passes the reference files to the page (as base64) and takes the files a scenario produces back. No existing scenario changed.
- **Tests:** engine 93, site 221 (+7), resilience 37, on PowerShell 7 and 5.1.

---

## Change log · 6 Oct 2026 · F1: your return and a personal benchmark (Portfolio)

A new section in **Portfolio**, just below *Your holdings*: **Your return · compared with the same money in SXR8**. No existing figure, text, table, CSV or format changed. There is no new data, no new `localStorage` key and no network request.

- **Your return (XIRR, money-weighted, in euros).** The cash flows come from the register:
  - each purchase is −(quantity × price), through `efetiva()`, so in today's shares;
  - each Bitcoin purchase is −(total paid);
  - each sale is +(quantity × sale price);
  - a closing flow of today's value of your holdings, from the same `pfDados()` as *Your holdings*, dated on the build day in Lisbon.

  The solver uses Newton's method from 10 % and falls back to bisection. With no solution, or no change of sign in the flows, it shows "Unavailable". Under 365 days the figure is the return for the period, not annualised, with a note.
- **Same money in SXR8.** Each purchase of any asset buys SXR8 for the same euros at that day's close from the long-term history, or the last close before it. Each sale takes the same euros out. If the benchmark has too few units for a sale, it uses all of them and the page says so. A date before the SXR8 history makes the benchmark "Unavailable", and the page names that date. The benchmark is valued at the latest SXR8 price, marked *intraday* while Xetra is open.
- **Unavailable cases:**
  - an asset held without a current price (named);
  - an entry without a price;
  - a sale without a matching purchase;
  - no solution.
- **Fees:** `custoFluxo(x)` returns 0 for now. When fees are recorded (task 09), they enter there without changing `rentab()`.
- **Code:** `fluxosCarteira`, `custoFluxo`, `xirr`, `rentab` and `drawRet` in the template, called from `drawPf`. The only change to existing code is that `drawPf` also calls `drawRet(R)`, and the test hook exposes `rentab`, `xirr`, `isoU` and `pct`.
- **Tests:** site +18 in 5 scenarios (`ret`, `retsales`, `retshort`, `retmissing`, `retempty`). Each XIRR is compared with an independent bisection written in the test, and the benchmark was calculated by hand (8.5 units × €250 = €2,125). `taxBaseline` is still identical byte by byte.
  - **Totals:** engine 93, site 239, resilience 37, on PowerShell 7 and 5.1.
  - **Layout:** no horizontal scroll at 500, 1,366 and 3,840 px.

---

## Change log · 6 Oct 2026 · F2: target allocation and next contribution (Portfolio), backup version 5

A new section in **Portfolio**, below *Your return*: **Target allocation · next contribution**. It is the first change to the backup format since version 4.

- **What it shows.**
  - For each asset: its value, current weight, target, the deviation in percentage points, and an *outside the band* tag when the deviation is larger than the band.
  - The split of the planned monthly amount M.
  - How many months of contributions of M would bring every asset back inside the band without selling, at today's prices (the page simulates it).
  - A fixed note: selling to rebalance realises taxable gains, new contributions do not.
- **Split of M, as specified.** Targets are computed on (holdings + M) and each gap is max(0, target − current). If the gaps add up to M or more, M is shared in proportion to the gaps; otherwise each gap is covered and the rest is shared by the target weights.
  - **Finding:** with weights adding up to 100 %, the gaps always add up to M or more, because the sum of (target − current) is exactly M. The second rule therefore never applies.
  - The edge case is when every asset is below its target: the gaps add up to exactly M and each asset gets its gap. It is tested, and the second rule is tested on its own.
- **Assets without a current price** are left out and named, and the remaining targets are scaled up to 100 %.
- **Inputs and validation.** Targets in the table, plus the band and M; Save. Targets must add up to 100 % (± 0.01), the band must be above 0 and at most 50 (blank = 5), and M must be 0 or more. Otherwise there is an error message and nothing is saved.
- **Backup version 5.** It adds the optional key `targets` `{weights, band, monthly, at}`, also stored in `bb.targets`. Entries are unchanged.
  - Saving targets updates `savedAt` and the backup file, like any other change.
  - When merging, the targets saved most recently (`at`) win. If the browser's copy is newer, the backup file is marked as needing an update.
  - Invalid targets are ignored: an unknown asset, weights that do not add up to 100, or values that are not numbers.
  - Versions 1, 3 and 4 still load in the script and in the page.
  - As before, the backup, `targets` included, goes only into `bluechip-board.html`, never into `Archive\` or `bluechip-board-data.json`.
- **Changes to existing code.**
  - `store.set` also marks changes to `targets`.
  - `dadosBackup` now writes version 5, plus `targets` when they exist.
  - `limpaBackup` also returns `targets`.
  - `juntaBackup` merges the targets and returns `alvos`.
  - The merge and restore messages say when the targets came from the file.
  - `estadoBk` also counts saved targets.
  - `drawPf` and `init` call the new section.
  - Existing behaviour without targets is unchanged.
- **Found during the work.** PowerShell 7 rewrites ISO dates in the backup it embeds without the milliseconds. This already happened to `saved`, and it is the same instant; the page reads both forms. The new engine test compares instants, not text.
- **Tests:**
  - **Totals:** engine 95 (+2: backups v3, v4 and v5 read, and the targets kept in the site copy), site 263 (+24 in 6 `tgt*` scenarios), resilience 38 (+1: targets only in the main site), on PowerShell 7 and 5.1.
  - **Independent checks:** the split and the months are compared with an independent calculation and simulation written in the test.
  - **Unchanged:** `taxBaseline` is still identical byte by byte.

---

## Change log · 6 Oct 2026 · F10: investment policy, notes per entry and context under the drop alerts

- **Investment policy (Portfolio).** Six free-text fields (horizon, allocation, monthly amount, what would make me sell, rule for a 20 % drop, rule for a 30 % drop) and a Save button. They are stored in `bb.policy` and in the optional key `policy` of backup version 5, with `at`. When merging, the most recent `at` wins.
- **Notes per entry.** A note button on every row of the register; the note is shown under the asset. Notes are stored beside the entries, in `bb.notes` and `notes` `{id: {t, at}}`, so the entries are unchanged.
  - When merging, each note keeps its most recent version.
  - Deleting an entry deletes its note.
  - Notes whose entry no longer exists are ignored and are not written to the backup.
- **Context under the drop alerts.** The two existing "below its 52-week high" alerts keep their level and exact text: `alerts()` only attaches data (`ctx`) to them. `renderAlerts` adds a context line:
  - the user's rule for that depth (30 %, or the 20 % rule);
  - the asset's own history: drops from a peak at least as deep as the current one, counted with `quedas()` on the long-term prices; how many recovered, with the median time from the bottom back to the peak (the same measure as the "biggest drops" tables); how many have not recovered yet (including the current one); the dates of the sample;
  - a caveat that the sample is mostly a rising market.

  With no earlier drop that deep, the line says so.
- **Interpretation:** "as deep as the current one" uses the current depth from the 52-week high, rounded down to 0.1 %. It is compared with drops from the running all-time peak in the long-term history.
- **Privacy and safety.** Policy and notes are personal data:
  - they go only into `bluechip-board.html` through the backup, never into `Archive\` or `bluechip-board-data.json`;
  - every user text is escaped where it is shown, and textareas are filled through `.value`;
  - the `xss` scenario now also feeds hostile text into a note and the policy.
- **Changes to existing code.**
  - `alerts()`: one added line after the two drop alerts, which attaches `ctx`. Texts and levels are unchanged, and the tests check the exact texts.
  - `renderAlerts`: draws the context line when `ctx` exists.
  - `apaga`: also removes the entry's note.
  - `drawBuys`: a note line and a note button per row.
  - `buildBuys`: a click handler for the note button.
  - `store.set`: also marks changes to `policy` and `notes`.
  - Backup: `dadosBackup` writes `policy`/`notes`, `limpaBackup` validates them, and `juntaBackup` merges them (and returns `pol` and `notas`).
  - Merge, restore and status messages: also mention them.
  - `drawPf` and `init`: draw the policy.
- **Tests:**
  - **Totals:** engine 96 (+1: v5 with policy and notes read and kept in the site copy), site 284 (+21: 7 `pol*` scenarios and one check in `xss`), resilience 39 (+1: policy and notes only in the main site), on PowerShell 7 and 5.1.
  - **Bugs in my own new tests, fixed during the work:** hostile text containing `</script>` closed the scenarios' script tag; a Bitcoin test series shorter than a year had no 52-week figure; a table was counted before it was redrawn.
  - **Unchanged:** `taxBaseline` is still identical byte by byte.

---

## Change log · 6 Oct 2026 · F11: stress test and strategy comparison

- **Stress test (Portfolio)**, labelled as a hypothetical scenario. The `EPISODIOS` constant holds three S&P 500 falls: 2018 Q4 (20 Sep → 24 Dec 2018), COVID (19 Feb → 23 Mar 2020) and 2022 (3 Jan → 12 Oct 2022).
  - **Dates checked** against the `^GSPC` daily closes from Yahoo (read once for the check, 6 Oct 2026; the script does not fetch it). They are exactly the peak and trough closes: −19.8 %, −33.9 % and −25.4 %. In `HIST`, the stocks, Bitcoin and EUR/USD have a close on all six dates. The Xetra ETFs have none on 24 Dec 2018 (Xetra closed), so they use the last close before it, 21 Dec, as the specification allows.
  - **Per held asset with a current price:** the euro return between the two closes (or the last close before each, at most 10 days earlier; US stocks at each day's EUR/USD), applied to today's value, and the time from the trough until the euro price was back at the peak value.
  - **Excluded and named:** assets without history for an episode, and assets without a current price. Totals and % cover only the assets with data.
- **Three ways to invest the same money (ETF and Bitcoin tabs)**, below the monthly simulator, with the same monthly amount and start month:
  - (a) monthly purchases, `simular()` unchanged;
  - (b) the whole total on the first purchase day;
  - (c) buy the dip: cash at 0 %, all invested at the first close that is X % or more below the highest close of the previous 52 weeks (365 days, the day itself excluded); X = 10 % by default, saved in `bb.strat.*`; cash left at the end counts at face value.

  The section shows the final values, gains, the dip purchases and any cash left, and a three-line chart.
  - The monthly simulator's results and `bb.sim.*` keys are unchanged (tested).
- **Changes to existing code:** `drawEtfMore` and `renderBtc` also call `drawStrat`, `drawPf` also calls `drawStress`, and the test hook exposes `stress`, `EPISODIOS`, `estrategias` and `eur`. No existing function's output changed.
- **Tests:** site 297 (+13 in 3 scenarios: `stress`, `stressnone` and `strat`), against independent calculations written in the tests. The dip rule is checked with a brute-force 52-week high, and there is an episode outside IS3N's history. Engine 96 and resilience 39 are unchanged. All green on PowerShell 7 and 5.1, and `taxBaseline` is still identical byte by byte.

---

## Change log · 6 Oct 2026 · F8: rolling returns, underwater chart, log scale (end of phase 1)

- **Rolling returns (ETFs tab for the chosen fund, and Bitcoin tab)**, in a section that is calculated only when it is opened.
  - Every window of 1, 3 and 5 years, one starting on each point: 252 sessions a year for trading-day series and 365 days for Bitcoin, the same detection as `stats()`.
  - It shows the number of windows; the worst, median and best return (annualised for 3 and 5 years); the share of negative windows; and a `barChart` histogram for 1, 3 or 5 years.
  - It warns that the windows overlap. A length longer than the history shows "Unavailable".
- **Underwater chart (ETFs tab)**, beside the biggest-drops table: price ÷ running maximum − 1, with the current and the worst value. The Bitcoin tab already had the same chart (*Drop from the all-time high*), so it was not duplicated.
- **Log scale (ETF long-term chart):** a Linear/Log switch, linear by default. The chart is unchanged until the switch is used: same `aria-label`, same look. It reuses `lineMulti`'s `log` option, as the Bitcoin chart does.
  - **Not done:** the US stocks have no long-term chart in the site; Prices covers at most one year. The specification mentions a switch "on the long-term charts of the ETFs and the stocks", but creating a new stock chart was not requested, so there is none.
- **Changes to existing code:**
  - `drawEtfMore` passes `log` and draws the underwater chart. It draws the rolling section only when that section is open, and empties it otherwise.
  - The biggest-drops table and its note now sit inside a two-column grid, with the new chart beside them; the order of the sections is unchanged.
  - The toolbar gained the Scale group.
  - `init` adds the switch and the open listeners.
  - The test hook exposes `rolar`, `underwater` and `nf`.
- **Tests:**
  - **Totals:** site 310 (+13 in `roll`, `rollshort` and `rollbtc`), against an independent calculation; engine 96 and resilience 39 unchanged. All green on PowerShell 7 and 5.1, and `taxBaseline` is still identical byte by byte.
- **End of phase 1 (tasks 01–06), real run.** On PowerShell 7, with the SEC e-mail and `-NoOpen`:
  - 71 of 71 sources OK, as in the run before;
  - no warnings or errors;
  - 60 s (earlier runs: 56–62 s);
  - the backup file unchanged, and the e-mail not in the outputs.

---

## Change log · 6 Oct 2026 · F3: real exposure by country, sector and currency; concentration; data version 1.4

- **Checked first, on the real files of the four funds (6 Oct 2026).** The Holdings sheet has Issuer Ticker, Name, Sector, Asset Class, Market Value, Weight (%), Notional Value, Nominal and Market Currency. It has **no Location (country) column**.
  - Country weights are only in the "Exposure Breakdowns" sheet, in a "Geography/Locations" block. EUNK and IS3N have that block; SXR8 and EUNN (single-country indexes) do not.
  - **Decision by the owner:** SXR8 and EUNN use the country of their index (`PaisIndice` in `$ETFs`), and the page flags it as approximate.
  - The test copies of the iShares format now include an Exposure Breakdowns sheet.
- **Script.** `Get-AgregadosETF` (called by `Get-PesosETF`) keeps only totals per fund:
  - country (from the Geography block, or the index country for the equity part), sector and underlying currency (from all holdings rows);
  - non-equity rows (cash, futures, FX, money market) count as Cash/Other;
  - each total must add up to 95–105 %, otherwise it is null (Unavailable); a fund on its reference weights has no aggregates.

  The data version is now **1.4**. Real files: SXR8 United States 99.9 %; EUNK 13 countries; IS3N Taiwan 28.8 %, Korea 21.1 %, China 18.0 %; EUNN Japan 99.1 %. That is 0.5 to 1.4 KB per fund.
- **Page (Portfolio).**
  - **Exposure:** three charts (country, sector, underlying currency) weighted by today's holdings. Apple, NVIDIA and Alphabet count as United States and USD; their sector comes from the iShares file, otherwise `SETOR_ACOES`. Bitcoin counts as Crypto.
  - **Unavailable funds:** left out, with the share of the portfolio not covered.
  - **Notes:** "SXR8 is quoted in EUR, but its underlying currency exposure is USD", and the index-country caveat.
  - **Concentration by company:** only the three panel companies (direct + look-through) and each fund's top 10 on its own, never added up by name. Above `CONC_LIMIAR` (10 %) there is an informative note, not an Overview alert.
- **Changes to existing code:**
  - `Get-PesosETF`: adds `agregados` (null in the fallback).
  - `Add-InfoETF`: adds `paisIndice`.
  - `versao`: changed from 1.3 to 1.4.
  - `drawPf`: also calls `drawExpo`.
  - Test hook: exposes `exposicao` and `CONC_LIMIAR`.
  - The existing ETF tab, look-through and the equity-only `setores` are unchanged.
- **Tests:** engine 101 (+5: aggregates read, totals only, sum rejected, the index country, fallback), site 325 (+15 in `expo`, `expostocks`, `expoold`, `conc`), resilience 40 (+1: offline, aggregates Unavailable, data 1.4). All green on PowerShell 7 and 5.1, and `taxBaseline` is still identical.

---

## Change log · 6 Oct 2026 · F4a: "Before you sell" simulation and fees

- **Before you sell (Portfolio).** You enter an asset and a quantity. The page runs `carteira()` with a hypothetical sale today, at the latest price, in memory only (`carteira(extra)`, the same check the register uses). Nothing is saved; a test compares `localStorage` and the backup before and after. It shows:
  - the purchases used (FIFO), proceeds, cost, gain, and a 28 % tax estimate on the positive taxable gain, with a note that it is a simplification (no loss offsetting, no *englobamento*, no fees);
  - for Bitcoin, exempt lots (held 365 days or more, the counter's rule) and taxable lots, with a warning for lots 30 days or less from becoming exempt.

  A sale larger than the holding is refused.
- **Fees.** A side map `fees {id: {v, at}}` in euros (`bb.fees`, and `fees` in backup version 5); the entries are unchanged.
  - There is an optional "Fee (€)" field in the register form, a **Fee** column, and a button to edit it. The realised gain shown is unchanged.
  - When merging, the most recent `at` wins for each entry. Fees of deleted or missing entries are ignored and dropped.
  - **Anexo J:** *Despesas e encargos* per sale/purchase row = purchase fee × used ÷ bought + sale fee × row ÷ sold. Each row is rounded to cents, with the remainder on the sale's last row.
    - With no fee on either side, the cell stays blank, as before, and `taxBaseline` is still identical byte by byte.
    - With one fee, the row gets `FEE_PARTIAL`.
    - With both fees, Despesas e encargos leaves that row's "Fill in manually".
  - **XIRR (task 02):** `custoFluxo` now returns the entry's fee; without fees, the results are identical.
- **Changes to existing code:**
  - Register (`drawBuys`): a Fee column; the empty row now spans 9 columns.
  - Register form (`buildBuys`): the fee field, the saved entry's id, and a click handler for the fee button.
  - `apaga`: also removes the entry's fee.
  - `store.set`: also marks `fees`.
  - Backup: `dadosBackup`, `limpaBackup` and `juntaBackup` handle `fees`.
  - Anexo J: `anexoJ` builds each sale's rows and calls `despesas`, the *Despesas e encargos* column takes `r.desp`, and the "Fill in manually" columns depend on the row.
  - Test hook: exposes `simulaVenda` and `feeDe`.
- **Tests:** site 343 (+18 in `sellsim`, `sellbtc`, `fees` and `feesold`), resilience 41 (+1: fees only in the main site), engine 101 unchanged. All green on PowerShell 7 and 5.1. One of my new tests used `$()` (a single element) where it needed `querySelectorAll`; that was fixed in the test.

## Change log · 6 Oct 2026 · F4b: dividends for Anexo J (Quadro 8A) and estimated net dividends

- **Verified first** in the official *Modelo 3 – Anexo J* and its instructions:
  - Tabela V codes: **E10**, dividends with tax withheld in Portugal, and **E11**, without;
  - Quadro 8A columns: Código rendim., País da fonte, Rendimento bruto, Imposto pago no estrangeiro (no país da fonte), the paying agent's country (only for E23), and Imposto retido em Portugal (NIF, retenção).
- **Export dividends (Quadro 8A).** A separate form and button (`#div8aForm`, with its own year list) writes `AnexoJ_Dividends_<year>.csv`. There is one row per payment of the dividend companies (`DIV_IDS`) in the year, by ex-date; future ex-dates are left out.
  - **Shares entitled:** taken from `carteira()`. It counts the lots bought before the ex-date, minus what the sales before the ex-date used from them. There is no second FIFO.
  - **Amounts:** gross in USD, and US tax withheld estimated at 15 % (W-8BEN). Euros are only a reference, at the ex-date EUR/USD (`fxRef`, the same `fxPt`/`fxAt` as the Anexo J export). With no rate, the euro columns are blank.
  - **Flags:** `BROKER_FX` on every row; `PAY_DATE_UNKNOWN` (REVIEW) for an ex-date in the last 45 days of the year; `NO_FX_REFERENCE`; `SPLIT_CHECK`; `NO_ANEXO_J_MAPPING`.
  - **CSV:** the same conventions as the existing export, through `csvLinhas`.
  - **Configuration:** `ANEXO_J.q8a`. The code defaults to E11, and the file names E10 for when tax was withheld in Portugal.
- **Dividends table.** A new last column, "Net (est.), euros", equal to the projected euros × 0.72 (`DIV_LIQ = 1 − taxaPT`), with its own total (`#div-tliq`) and a note built from the configured rates. The existing columns, totals and KPIs are unchanged.
- **Changes to existing code:**
  - `dividendos()` returns `liq` and `totLiq` as well;
  - `drawDiv` has the extra column;
  - `ANEXO_J` gained the `q8a` key;
  - `buildTax` also wires the new form;
  - `drawBuys` also refreshes the dividend year list;
  - the test hook exposes `anexo8A` and `divPagamentos`.

  The existing Anexo J export is untouched: the `taxBaseline` CSV files are still identical byte by byte, and its button still downloads exactly two files (tested).
- **Tests:** site 370 (+27 in `div8a`, `div8afx` and `div8anone`); engine 101 and resilience 41 unchanged, because the script and the data format did not change. All green on PowerShell 7 and 5.1.
- **Known limitations:**
  - Yahoo gives ex-dates only, so the year is the ex-date's year, and late-December payments are flagged.
  - The dividend data covers about two years back from the run, which the summary states.
  - Dividends of shares sold and bought back between two runs are counted from the register as it is now.

## Change log · 6 Oct 2026 · F9: news by exposure and the ETFs' top holdings (end of phase 2)

- **Relevance to me (News).** A new option in *Sort by*. The default order (*Potential impact*) and the levels are unchanged; the score and level still come from the script.
  - Relevance = score + 3 × the share of the portfolio exposed to the story's asset, calculated in the browser from `pfDados()` (`carteira()`, the only FIFO).
  - Apple, NVIDIA and Alphabet count directly plus through each ETF (the `etfInfo` look-through of *Real exposure*). An ETF and Bitcoin count by their value. Market counts stocks and ETFs. Several assets add up to at most 1.
  - With no holdings, the option is disabled, with a note. While it is chosen, a note gives the formula and each story shows its relevance.
- **Top holdings (script).** `$AliasesPosicoes` holds 36 curated aliases for the top 10 of each fund on 5 Oct 2026, except Apple, NVIDIA and Alphabet, which are already in `$EmpresasRe`. Each alias has word boundaries and exclusions for a known false positive, with a test for each.
  - An alias is used only when the headline has no `$EmpresasRe` keyword and the feed has no `Dica` (those stories used to be dropped).
  - The story goes to the fund (`viaPosicao` names the holding). It is capped at moderate, gets no company bonus, and the site tags it *Via top holding*.
  - **Decision:** feeds with their own asset keep their assignment, so no story that was already classified changes. A first version also let an alias replace a search feed's `Dica`. That failed the existing test *"one EM company without keyword is only a weak feed match"*, so it was reverted rather than changing the test.
- **Maintenance reminder.** `Get-PosicoesSemAlias` lists the top 10 holdings of the current iShares files that have no alias, by ticker or name. It does nothing with reference weights only. Today's top 10 is fully covered (tested).
- **Changes to existing code:**
  - `Measure-Noticia`: the alias step, plus `viaPosicao` in the bonus and level-cap conditions;
  - the maintenance block: one extra reminder;
  - `filteredNews(R)`, `renderNews` and `newsItem(n, rel)`; the top-news and Bitcoin news lists call `newsItem(n)` explicitly;
  - `NEWS` normalises `viaPosicao`;
  - the test hook exposes `expoNoticias`, `relevancia` and `NEWS`.
- **Tests:**
  - engine 150 (+49: aliases, 36 true/false-positive pairs, levels unchanged with and without aliases, the reminder);
  - site 379 (+9 in `newsrel` and `newsrelempty`);
  - resilience 41 unchanged.

  All green on PowerShell 7 and 5.1. `taxBaseline` is still identical byte by byte.

## Change log · 6 Oct 2026 · F7a: news history and past earnings dates (data collection only)

- **Sources verified first** (6 Oct, one read each, no key, the SEC User-Agent):
  - `data.sec.gov/submissions/CIK##########.json` returns 200 with JSON: parallel lists (`form`, `items`, `accessionNumber`, `filingDate`, `reportDate`, `acceptanceDateTime`) of up to 1,000 recent filings.
  - **Finding:** its `acceptanceDateTime` is wrong for Apple. The EDGAR filing pages say "Accepted 2026-07-30 16:30:28" (New York), but the JSON says `2026-07-31T00:30:28Z`, 8 h later; in winter it is 10 h later. It is right for NVIDIA and Alphabet. The EDGAR 8-K Atom feed gives the page's time with its offset (`2026-07-30T16:30:28-04:00`).
  - **Decision:** the list and the dates come from the JSON, as specified. The acceptance times come from the Atom feed, matched by accession number. Older 8-K missing from the feed get no time and no session, never a guess.
- **Past earnings dates.**
  - `Get-ResultadosSEC` makes two requests per company, only with `-EmailSEC`, with the same User-Agent and 400 ms between requests. It runs after the long-term histories.
  - It keeps the 8-K with item 2.02 of the last 5 years and the latest 10-Q/10-K.
  - `Get-SessaoReacao` converts to New York time with `TimeZoneInfo`, so the US clock changes are handled. After 16:00 the reaction is in the next session; before 09:30, the same session; during the session, the same session with a note. The sessions are the dates of the company's price history, then weekdays without the `$Bolsas` holidays.
  - Fallback: the previous run (up to 180 days old, labelled), and times already verified are reused by accession number.
- **News history.** `noticias-historico.json` keeps the material and important stories (title, link, source, date, assets, level, score, themes) for 400 days, one per title (the highest score wins).
  - It is written atomically, last, after `vistos.json`; the existing order of writes is unchanged.
  - A missing file starts empty with a notice. A corrupted file starts empty with a warning, and the old file is kept as `.bad`.
  - It goes into the data as `historicoNoticias` (with its start date). `resultadosSec` is added too; the data stays at version 1.4.
- **Real run** (6 Oct, PowerShell 7, the SEC e-mail, `-NoOpen`):
  - 60 s and 74/74 sources (+3 new);
  - AAPL 20, NVDA 21 and GOOGL 13 earnings 8-K, with 1, 5 and 1 older ones without a time (before the feed);
  - Apple at 16:30 New York → next session;
  - news history started with 150 stories (93 KB);
  - backup unchanged, e-mail in no output.
- **Changes to existing code:**
  - the `$noticias` step is followed by reading and merging the history;
  - new step after the long-term histories;
  - two keys in `$dados`;
  - one write after `vistos.json`.
- **Tests:** engine 179 (+29), resilience 51 (+10), site 379 unchanged. All green on PowerShell 7 and 5.1. One of my new resilience tests failed at first because its simulated Downloads folder still held a backup from an earlier step; it now uses an empty folder.
- **Known limitations:**
  - The JSON's `recent` list holds about 1,000 filings, so for Alphabet (many Form 4) it reaches back only about 3 years (13 earnings 8-K).
  - The history file grows by about 20 stories a day: roughly 2–3 MB at 400 days, embedded in the site and in each Archive copy.

## Change log · 6 Oct 2026 · F7b: events on the Prices chart, big moves explained, earnings reactions

- **Template only.** It uses the task-11 data (`historicoNoticias`, `resultadosSec`); the script and the data format are unchanged.
- **Markers** (`lineMulti` gained an `events` option). The markers are drawn above the cursor layer as `.ev-mark` circles. Their `data-tip` is shown with `textContent`, so a headline is never read as HTML.
  - Earnings appear on the reaction session, or on the filing date when the time is unknown.
  - Rate decisions come from past calendar events ("Fed decision…") and from history stories with the *Macro and rates* theme and a decision in the headline. The rule mirrors `$DecisaoJuros`.
  - Material history stories are shown for the assets in view.
  - There is one marker per day and kind, with up to 3 headlines.
- **Big moves explained** (`#tbl-moves`):
  - daily moves at or above the existing alert thresholds (4 % stocks and ETFs, 6 % Bitcoin), in the asset's own currency and without the intraday price;
  - the material and important stories of that session and the one before;
  - stories mapped to sessions in the exchange's time zone by `sessaoDe`, with `fechoDe` covering early closes and cached formatters;
  - "News history since …"; moves before the history say so; the 30 most recent are shown.
- **Earnings reactions** (`#tbl-earn`):
  - per company, in USD from the long-term closes: the reaction-session move, the move after 5 sessions, the median absolute move and the number of cases;
  - every case in a list;
  - Unavailable with the reason when SEC data is missing, failed or skipped.
- **Changes to existing code:**
  - `renderPrices` passes `events` and draws the two tables;
  - `#perfSub` mentions the markers;
  - a few CSS rules;
  - the test hook exposes the new functions.
- **Tests:** site 395 (+16 in `evts` and `evtsempty`); engine 179 and resilience 51 unchanged. All green on PowerShell 7 and 5.1. With the real data the history starts today, so the moves listed so far say "Before the news history".

## Change log · 6 Oct 2026 · F5a: fundamentals from SEC XBRL (data collection only)

- **Source measured first** (6 Oct): `companyfacts` is 3.2–4.1 MB per company, uncompressed. Windows PowerShell 5.1 parses 4.1 MB in about 1.1 s. **Decision:** `companyfacts`, one request per company, rather than `companyconcept` per tag.
- **Cache:** a request only when the SEC (task-11 data) shows a 10-Q/10-K newer than the stored collection. Otherwise the stored facts are reused (`ok (cached)`, no request). If the request fails, or without `-EmailSEC`, the previous run's facts are used, up to 120 days old and labelled; otherwise the company is `error` or `skipped`, with no numbers.
- **Metrics:** revenue, gross and operating profit, operating cash flow, capex, FCF, diluted EPS and diluted average shares. Each has a list of tags in order of preference, merged period by period, and the tags used are stored.
  - The rules are the ones specified: 4th quarter = year − 9 months; real fiscal-year dates; the most recent filing wins; splits; validated units and durations; never interpolated.
  - **Additions:** the first published value and date are kept beside the latest (`v0`, `p`), because comparatives are re-filed a year later; without them, the no-look-ahead history of task 14 would lag a year. Cash flow quarters are derived from year-to-date figures. Diluted shares are never derived by subtraction.
  - **Deviation:** splits are applied by each value's **filing date**, not the period end. Filings after a split are already adjusted, so adjusting by the period would count the split twice (verified on NVIDIA's 2024 restatement).
- **Real data:**
  - 43–44 quarters per company (since late 2015), about 25 KB each;
  - Alphabet has no `GrossProfit` (Unavailable) and diluted shares only since 2022.
- **Real runs:**
  - first collection: 70 s, against 60 s before (+17 %), 77/77 sources;
  - the next run, cached: 52 s, with no request.
  - backup unchanged; e-mail in no output.
- **Tests:** engine 197 (+18), resilience 54 (+3), site unchanged. All green on PowerShell 7 and 5.1. Two bugs of mine were caught by the new tests before closing:
  - `[math]::Max(1, …)` picked the Int32 overload;
  - `$A` and `$a` are the same variable in PowerShell.

## Change log · 6 Oct 2026 · F5b: Fundamentals tab

- **New tab:** `#fundamentals`, between Prices and Currency & Macroeconomics (rail, panel, `TABN`, `CHARTS`). It is drawn only when it is open, and redrawn by the asset filter when it is the current tab.
- **For each company** (Apple, NVIDIA, Alphabet; one with the filter on it; a note on an ETF, Bitcoin or the market):
  - **Last 8 quarters:** revenue and growth over a year, gross and operating margin, FCF, diluted EPS, and diluted shares with buybacks or dilution over a year. Each row shows its quarter and publication date ("revised" when a later filing changed it), and derived figures are marked.
  - **P/E and P/FCF** over the last 4 quarters, with their **percentile** in the company's own 10-year month-end history.
  - **10-year chart.** **No look-ahead:** a quarter counts from its first publication date, with its first value until a revision was filed (`valorEm`, `ttmEm`).
- **Note:** valuation gives context and does not predict the short term.
- **Missing data:** Unavailable with the reason, per company and per metric. Alphabet's gross margin, for example, is Unavailable because it does not report gross profit.
- **Layout.** With ten tabs the bar no longer fitted at 1,920 px (it fitted exactly with nine). A rule for 1,561–2,399 px uses the same tighter tab padding as the existing rule under 1,560 px, and the bar fits again at 1,920 px. At 1,600–1,700 px it scrolls, as it did before. Measured at 768, 1,024, 1,280, 1,366, 1,600, 1,920, 2,400, 2,560, 3,000 and 3,840 px: no horizontal page scroll.
- **Tests:** site 411 (+16 in `fund` and `fundempty`); engine 197 and resilience 54 unchanged. All green on PowerShell 7 and 5.1. The old tests had no tab lists to change. One name, `FH`, already existed in the template and was renamed `FUND_H`.

## Change log · 6 Oct 2026 · F6: euro area indicators, central-bank calendar (end of phase 3)

- **Sources verified first** (6 Oct):
  - **FRED** (`fredgraph.csv`): the server resets or ignores this script's requests, on PowerShell 5.1 and 7, with the script's User-Agent and with an honest `BluechipBoard/1.0`. It answers only clients that identify as known tools, such as curl. Posing as one would mean getting around a bot filter, so **the three FRED series were not implemented** (pending decision for the owner).
  - **ECB Data Portal:** confirmed with no key.
    - Deposit facility rate: `FM/B.U2.EUR.4F.KR.DFR.LEV`, change dates only.
    - HICP annual rate: `ICP/M.U2.N.000000.4.ANR`. Its data ends in Dec 2025 on the portal (also for I8 and I9).
  - **CAPE:** no current public CSV. The datahub/GitHub `s-and-p-500` dataset has PE10 only until Sep 2023, so it was **not implemented**, as the task allows.
  - **Calendars:** read from the official Fed and ECB pages.
- **Script:**
  - `$SeriesMacro`, `ConvertFrom-CsvSerie` and `Get-SerieMacro`, validating against an HTML error page, an unexpected header, another series, `.`/`NaN`/empty values (missing, never 0), implausible values and an empty series;
  - fallback to the previous run, up to 30 days old and labelled;
  - the data key `macro`;
  - `$Calendario` gained the remaining ECB decisions of 2026 and the Fed and ECB decisions of 2027 (`C`, `MKT`). The Fed ones are `High`, like the existing ones. The ECB ones are `Medium`, because a `High` event raises an Overview alert 7 days before and the task says there are no new Overview alerts.
  - New maintenance reminder `Get-LembreteReunioes`: no Fed or ECB decision more than 60 days ahead.
- **Currency & Macroeconomics:** a new "Euro area · ECB" section.
  - Each indicator has its value and date, a 10-year chart (the deposit rate as steps; inflation with a 2 % target line) and a short note. Data that is not recent is flagged (HICP: "Dec 2025 · latest published, not recent").
  - A note lists what is not shown and why, with the context the notes were meant to give: thermometers, not clocks; the inverted curve's long lags and false alarms; valuation and 10-year returns.
  - No new alerts.
- **Final real run** (PowerShell 7, SEC e-mail, `-NoOpen`): 63 s, 79/79 sources, no reminders, backup unchanged, e-mail in no output.
- **Tests:** engine 212 (+15), site 416 (+5), resilience 56 (+2). All green on PowerShell 7 and 5.1.

---

## Post-remediation follow-up · 7 Oct 2026

A second, deep pass over the whole project: every finding above re-checked against the code and the tests (the FIXED ones too), the partially fixed ones investigated again, then an independent audit as if this file did not exist. Work on the branch `remediation/deep-audit-2026-10`, in small commits; every bug followed *reproduce → fix → regression test → full suites → real run*. Nothing was removed from the site, and no test was loosened (three checks whose expected text or request count changed with a deliberate change were updated, and say why).

**Baseline before any change** (`main` at `933ee2d`): engine 212, site 416, resilience 56, all green on PowerShell 5.1 and 7; a real run took 59 s (PowerShell 7, 79/79 sources).

### Bugs fixed in this pass

| ID | Problem | Root cause | Fix | Test |
|---|---|---|---|---|
| N1 | On a Windows set to a non-Gregorian calendar (Thai), every price, dividend and Bitcoin date became `2569-…` | `.ToString('yyyy-MM-dd')` uses the current culture | Every date is formatted with the invariant culture | engine: Thai culture; a static check over the whole script |
| N2 | A missing CoinGecko 24 h change showed as "+0.0% over the last 24 hours"; a missing dominance as 0 % | `[double]$null` is 0 in PowerShell, and `isFinite(null)` is true in JavaScript | Missing fields stay `null` in the data; the page tests for `null` | engine; site `btcnull` (fails on the old code) |
| N3 | A `[date, null]` point became a price of 0 (a −100 % move) | `+null` is 0 in `toPts` | Missing values are left out | site `btcnull` |
| N4 | At night the page showed the previous session's prices, flagged as a session behind (seen at 01:17 on 7 Oct: every series ended on 5 Oct) | Yahoo had not yet published the daily bar of 6 Oct, though the same answer had its close in `regularMarketPrice` | When the session has ended, that close is added as the day's point (within 50 % of the previous one), and the source row says so | engine (3 checks); real run: every series on 6 Oct |
| N5 | Kraken and Stooq fallbacks accepted zero, negative or non-numeric closes; an isolated wrong tick was kept | No validation in the fallbacks | The same checks as Yahoo; an isolated point over 50 % away from two agreeing neighbours is dropped and counted; a huge last jump is flagged | engine (5 checks) |
| N6 | Cents rounded on the binary value: €10.005 entered was exported as €10.00 | `Math.round(v*100)/100` on float noise | Rounding on the decimal value (15 digits), half away from zero; the `taxBaseline` files are still identical byte by byte | site `taxround` |
| N7 | A backup with `true`, `null`, `''` or `[5]` as a quantity or price was read as 1, 0, 0 or 5; a Bitcoin purchase with a missing cost became a cost of €0 | `num = v => +v` | Only numbers and numeric text are accepted | site `bkstrict` |
| N8 | Two identical purchases without ids in a version-1 backup became one (data loss) | The derived id was the same for both | A `#2`, `#3`… suffix; merging the same file again adds nothing | site `v1dup` |
| N9 | Stale prices: "rose 5 % in the last session. Look for the cause in today's news" for a move days old; the portfolio's latest daily move added it as today's | The move used the last point without checking its date | The alert gives the date of the move; the daily move leaves out (and names) assets whose price is not current, and shows "—" when none is current | site `stalemove`, `staleall` |
| N10 | An earnings reaction counted an intraday price as the close of the reaction session | The long history's last point can be intraday | That point is left out while the session is open | site `earnpartial` |
| N11 | A release on a day the exchange is closed was "during the session"; a date before the price history mapped to its first session (years later) | No check of the day; the binary search had no lower bound | `quando = closed`, reaction at the next session; dates before the history use weekdays | engine (3 checks) |
| B14 | A new sale was refused when any old sale, of any asset, lacked purchases | `C.vendas.find(x => x.falta > 0)` | Only the new sale, or a sale it breaks, is refused | site `sellother` (fails on the old code) |
| N12 | Duplicate grouping could differ between two runs on the same news, and between PowerShell 5.1 and 7 | Ties in score and date kept the order of a hashtable (random per process on 7) | Ties broken by the title key | engine (input order reversed) |
| N13 | `idb()` could hang forever if IndexedDB threw | The exception was raised inside an event handler | The promise is rejected | site `bkfolder` |
| N14 | A permanent HTTP answer (404, 403…) was retried after 2 s | Every error was retried | No second try for 400/401/403/404/410 | engine |
| N15 | Asset names in the colour tags were not escaped; ETF ids from the data were not checked | Defence in depth (the values come from the script's configuration) | `esc()` in `coTag`; ETF ids must be simple | site (existing `xss`, `etfbad`) |

### Findings re-evaluated

| ID | Before | Now | Detail |
|---|---|---|---|
| H4a | PARTIALLY FIXED | **IMPROVED** | The browser rule stays (a reopened page needs a user gesture before it can write the file again). Now the first change on the page asks for the permission once, and saving resumes; if the browser kept the permission ("Allow on every visit" in newer Chrome), saving resumes on opening; leaving the page with changes made there and not in the file asks first. Scenarios `bkperm`, `bkpermno`. |
| H4g | PARTIALLY FIXED | **FIXED** (for new links) | *Save to project folder* asks for the folder and refuses one without `bluechip-board.html` at once. The next-run check stays as a second line of defence (and covers a file linked before this change). Scenario `bkfolder`. |
| H5 | PARTIALLY FIXED | **IMPROVED** | US stocks: Nasdaq's daily history (no key, the service already used for earnings dates) as a fallback, accepted only when at least 20 dates match the previous run's prices (median within 1 %, each within 3 %): another split basis or symbol is refused; without a previous run it is not used. Plus N4 and N5. The ETFs still have only Yahoo (no keyless Xetra source was found); Stooq still blocks. |
| L9 | PARTIALLY FIXED | **IMPROVED** | A full single source of truth would mean restructuring the template, out of proportion. Instead an engine check fails when an asset is missing from any place a new asset must be added (CO, ANEXO_J, keywords, feed, BOLSA_DE, PF, SEC, Fundamentals, colour token). |
| F2 | PARTIALLY FIXED | **IMPROVED** | `$Script:UAChrome = 154` (the Chrome installed here; it was 141, a year old) and `$Script:UAData`, with a reminder after about 8 releases. All 79 sources answered with the new identity. |
| F5 | PARTIALLY FIXED | IMPROVED | The browser-identity reminder joins the others. Reference weights and the calendar are still data you maintain. |
| F7 | PARTIALLY FIXED | IMPROVED | iShares source rows name a missing column, a missing ISIN and each Unavailable breakdown. |
| F8 | NOT FIXED | **FIXED** | The official EDGAR filing header (`…-index-headers.html`, `ACCEPTANCE-DATETIME` in New York time) gives the acceptance time of any filing, no longer only the Atom feed's 40 latest. Checked on 7 Oct against the submissions JSON for NVIDIA (where that field is right), in summer and winter time: identical. At most 12 a run per company, then cached. Migrating the whole source to the JSON API was not needed. |
| F9, F10, F11 | PARTIALLY FIXED | Unchanged / improved | Nasdaq dates keep their plausibility window; feeds benefit from N14; Chrome's folder rules from H4a and H4g. |
| F12 | NOT FIXED (legal) | Unchanged | No tax rule was changed or added. |
| F13 | PARTIALLY FIXED | Unchanged | Deutsche Börse confirms the year-end session by circular; there is no keyless calendar to read it from. The rule plus `curtos` stays. |
| F14 | NOT FIXED | **FIXED** | Profiling showed that on PowerShell 7.6 every .NET method call costs about 10 µs (AMSI method-invocation logging; about 1 µs on 5.1). Classification and grouping now use hashtables read by index, arrays and operators: identical results on the real data (4,870 titles, 1,334 groups, 24,350 classifications, both shells); grouping 6.9 → 2.3 s, classification about 3× faster. |
| S2 | NOT FIXED | **Kept, with the reasons** | Prefixing the keys per installation would not stop another local page from reading them (same `file://` origin), would break the Archive pages and need a migration of the existing data: risk without a security gain. |
| S3 | NOT FIXED | **FIXED** | The script reads the e-mail itself from `BLUECHIP_SEC_EMAIL` or `bluechip-board.config.json`; the scheduled task never stores it in its arguments (a warning says where it must be), and the launcher no longer passes it on the command line. By design it still goes only in the SEC User-Agent. |
| B7 | Limitation | **FIXED** | See F8: 54 of 54 earnings 8-K have their time on the real data (7 were Unavailable). |
| R1 | Risk | Mitigated | News history at most 15,000 stories (its start moves with it); Archive at most ~300 MB on top of 30 copies (the 5 newest always kept). |
| R2 | Risk | Unchanged | `companyfacts` (3–4 MB) parses fine on 5.1; a size guard is not needed today. |

### Decisions (D1–D9): technical assessment

| ID | Verdict | Why |
|---|---|---|
| D1 Quadro 8A E11/E10 | **KEEP** | A fact of your case (who withheld tax), not a software choice; configurable in `ANEXO_J.q8a`. |
| D2 G20 / G01, País da fonte | **KEEP** | A legal interpretation; the export labels it and leaves the decision visible. |
| D3 FRED | **DEFER** | FRED still blocks (7 Oct). The Chicago Fed's own NFCI CSV answers without a key: a clean option for one of the three series, waiting for your decision to add an indicator. The yield curve could be approximated from Yahoo `^TNX` − `^IRX` (labelled as an approximation); the HY spread has no keyless public source. |
| D4 CAPE | **DEFER** | Still no current, keyless CSV. |
| D5 200-day alert | **KEEP** | A preference about alert noise; nothing is wrong. |
| D6 ECB importance | **KEEP** | A preference (it would add an Overview alert). |
| D7 alias precedence | **KEEP** | It would change existing classifications; needs your decision and a deliberate test update. |
| D8 Alphabet gross margin | **DEFER** | Revenue − `CostOfRevenue` is a sound derivation, but Alphabet itself reports no gross profit, so showing one is an analytical choice. Kept Unavailable until you want it (it would be marked derived). |
| D9 SXR8/EUNN country | **KEEP** | Your choice of 6 Oct; flagged as approximate in the page. |

### Still limited by outside services or rules

- **Yahoo** remains the only price source for the four Xetra ETFs and for EUR/USD's 1-year series (with the long history and the ECB as fallbacks); its chart API is unofficial.
- **Browsers:** a reopened `file://` page needs a user gesture before it can write the backup file again; `file://` pages share `localStorage`.
- **SEC:** the submissions JSON's `acceptanceDateTime` is still wrong for Apple (the times come from the Atom feed and the filing headers); the `recent` list reaches only ~3 years back for Alphabet (B8).
- **iShares, Nasdaq, news feeds, ECB:** each validated, with its own fallback and error row, but outside the project.
- **Tax rules:** the Anexo J files are preparation files; the interpretations (D1, D2, the 365-day reading) are yours or your accountant's to confirm each year.

### New risks found

| ID | Risk | What to do |
|---|---|---|
| R10 | PowerShell 7's per-call cost of .NET methods can come back in any new loop over thousands of items | Follow the note in README → *Making a change safely* (hashtables by index, arrays, operators). |
| R11 | The close taken from the quote (N4) relies on Yahoo's `regularMarketPrice` being the official close once the session has ended | It is checked against the previous close (within 50 %) and named in the source row; the next run replaces it with the daily bar. |
| R12 | The Nasdaq fallback needs a previous run to compare with | On a fresh install with Yahoo down there is still no price for the US stocks (Unavailable, never guessed). |

### Results

| Check | Before | After |
|---|---|---|
| Engine, PowerShell 5.1 / 7 | 212 / 212 | **251 / 251** |
| Site (headless Chrome), run from 5.1 / 7 | 416 / 416 | **436 / 436** |
| Resilience, PowerShell 5.1 / 7 | 56 / 56 | **62 / 62** |
| Real run (PowerShell 7, `-NoOpen`, SEC e-mail from the config) | 59 s, 79/79 sources | 53–59 s, 79/79 sources (the first run also fetched 7 filing headers); the same 1,340 stories and 163 duplicates as the old code on the same news |
| Site in headless Chrome at 1,280–5,120 px | — | no JavaScript error, no horizontal overflow in any tab, every chart drawn; page load ~0.84 s at 3,840 px |
| Portfolio | — | `bluechip-board-backup.json` unchanged (SHA-256 `944F4C24…`) through every run; the copy embedded in the site identical, field by field (buys, lots, sales, deleted, targets, policy, notes, fees) |
| Privacy | — | the SEC e-mail is in no output, log or commit; no personal file is tracked by git |

New regression tests: engine +39, site +20 (12 scenarios), resilience +6.

## Change log · 7 Oct 2026 · Live prices while the page is open

**Request.** When the board is opened from the desktop shortcut, keep the prices up to date minute by minute while the page is open, and stop the background work when the page is closed.

**What was built.**

| Part | Change |
|---|---|
| Script | `-Live` (`-AoVivo`): instead of a run, a process that serves the latest prices on `127.0.0.1:47821` only (`$Vivo` in the configuration). One Yahoo "spark" request a minute brings all 12 symbols (assets, S&P 500, VIX, 10-year yield, EUR/USD), only while a page is open. The quotes are read with the rules of `Get-Serie` (`Get-DiaBolsa`, sharing `$FusosBolsa`): exchange day, partial while the session is open, and a symbol with a wrong currency, an invalid price or a future time left out and named. |
| Launcher | After a run without errors, starts `Bluechip-Board.ps1 -Live` hidden. A live process from an earlier click is asked to quit (`/ping`, then `/quit`) and replaced. |
| Website | `D.vivo.porta` (only in `bluechip-board.html`, never in the Archive copies or the data file). The page asks for `/quotes` every minute (every 5 s while connecting). Each quote is added to its series or replaces that day's point. Then the snapshot, alerts, portfolio, purchase lists and the open price tab are redrawn. A status line in the header shows the state. On `pagehide` it sends `/bye`. |
| Ending | The process ends 20 s after the last page says goodbye (a reload comes back within that time), after 5 minutes without any request (the browser closed without a goodbye), on `/quit` from this computer, or when replaced. A sleep of the PC does not count as silence. |

**Safeguards.**

- **Local only.** The listener is bound to the loopback address, with exclusive use of the port.
- **Request checks.** A request is refused unless:
  - its `Host` is `127.0.0.1` or `localhost` on that port (no DNS rebinding);
  - it carries no `Origin` (a program on this computer) or `Origin: null` (the board opened from the disk). Any other website is refused.
- **`/quit`.** Accepted only without an `Origin`: never from a browser.
- **What the process gives.** It serves public prices only: no personal data, no file access.
- **What the page accepts.** It does not use a quote that is:
  - more than 50 % away from the previous close;
  - in another currency;
  - from an earlier day;
  - more than 6 days after the end of the series.
- **Freshness.** It is counted from the latest live update (`T_PRECOS`), so an asset the updates no longer reach is shown as behind.
- **Bitcoin.** The CoinGecko 24 h change and USD price (from the run) are dropped once Bitcoin is live, because they would no longer match.
- **Nothing saved.** Nothing is written to `localStorage` or to a file. The next run collects everything again.
- **Bitcoin indicator label (found while doing this).** Without CoinGecko's 24-hour change, the Bitcoin section labelled the change since 00:00 UTC as "24 hours". It is now labelled "Today (UTC)", as the snapshot card already did.

**Checked.**

- **Chrome 154.** A `file://` page's `fetch` and `sendBeacon` to `127.0.0.1` arrive with `Origin: null`, without a permission prompt.
- **Real run of `-Live`.** It served the 12 quotes with their exchange days and correct partial flags (Xetra and Bitcoin open, the US closed). It refused another origin, another host and a browser `/quit`, and ended 3.4 s after `/bye` (grace period shortened for the test).

**New tests.**

- **Engine (+27):** the spark reader, the request decisions and static checks.
- **Site (+26):** scenarios `live`, `livestale` and `nolive`, plus one check in `btcnull`.
- **Resilience (+11):** the real process on a test port with a canned Yahoo answer.

All pass on Windows PowerShell 5.1 and PowerShell 7.

**Risks.**

| ID | Risk | What to do |
|---|---|---|
| R13 | The Yahoo spark endpoint is unofficial and could change or require authentication, as `quote` did in 2023 | The page then says *the source is not answering* and keeps the latest prices with their time; the daily run is unaffected (it uses the chart endpoint). |
| R14 | Another program could take port 47821 | The process does not start and the launcher window shows the warning; change `Porta` in `$Vivo`. |
| R15 | A browser could freeze a background tab for more than 5 minutes | The process then ends (as if the page had closed); the page says *Live prices stopped* and the shortcut restarts it. |
