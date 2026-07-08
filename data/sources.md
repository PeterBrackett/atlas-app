# Sources & Methodology

This page explains where Atlas's figures come from and why particular numbers are chosen over others. It's the in-app home for that explanation — update it here rather than in any external document, so the reasoning stays next to the data it explains.

## Segment taxonomy

Every country is broken into the same 19 segments, so figures compare cleanly across countries:

DB Pension (Corp), DC Pension (Corp), DB Pension (Govt), DC Pension (Govt), DB Pension (Union), DC Pension (Union), DB Pension (Healthcare non-profit), DC Pension (Healthcare non-profit), DB Pension (Endowments), DC Pension (Endowments), DB Pension (Tax exempt), DC Pension (Tax exempt), Endowments E&F, Tax exempt E&F, Foundations E&F, Healthcare non-profit E&F, Life insurance, Non-life insurance, SWF.

For UK data, this maps onto the S&P Money Manager Database's own `Category` field (Corporation, Government, Union, Tax Exempt Organization, Endowment, etc.), split further into DB/DC using the DB/DC asset columns in that export.

## How a segment's AUM figure is decided

Each segment's AUM can come from more than one source, and the one actually used isn't always the largest or the most obvious:

- **Bottom-up**: a mechanical sum of the individual institution records collected (currently mainly the S&P Money Manager Database export).
- **Top-down override**: a published, authoritative figure used instead of the bottom-up sum, when the bottom-up number is judged less reliable for that specific segment.
- **Industry aggregate**: a sector-wide published figure used where individual institution records aren't practical to collect (currently insurance).

**Worked example — UK `DB Pension (Govt)`**: the bottom-up sum of S&P's Government-category records is approximately £634bn. The figure actually used, £393.94bn, is a top-down override aligned to official LGPS (Local Government Pension Scheme) statistics, judged more reliable than the raw S&P sum for this segment specifically. This is the kind of decision recorded on this page rather than left unexplained next to the number.

There's no fixed rule like "official sources always win" — each override is a judgement call on which source is more trustworthy for that segment in that country, made case by case as new sources are found. The intent is for this to be genuinely extensible: as new country-specific statistical sources are identified, they get added as another candidate figure for a segment, and whichever is judged best becomes the one shown.

**Reversibility**: wherever a segment has more than one candidate figure, every value that's been sourced for it is kept in that segment's data (not just the one currently shown), under an `alternate_sources` list alongside the live figure. Switching which source is "active" for a segment — for example, reverting `DB Pension (Govt)` back to the raw S&P bottom-up sum — is then just editing that segment's entry in its country JSON file, not re-deriving anything from scratch.

## Opportunity scorecard

Alongside AUM and allocation, every segment carries an opportunity scorecard — the same structure as `FINAL_Country_Matrix.xlsx`, the original research output this whole matrix view is built to reproduce and keep current. Each segment is scored 1–3 on 12 dimensions:

Market opportunity, Outsourced management, Pricing impact, Alignment of investment thinking, Distribution resources required, Regulatory complexity, Client servicing, Local presence required, Languages required, Investor decision-making, Comingled vehicles, Consultant reliant.

**Overall score formula**: `Overall = (3 × Market opportunity) + sum of the other 11 dimensions`. Market opportunity is weighted three times over because it's judged the primary driver of whether a segment is worth pursuing at all — this exactly matches the `=SUM(...)+(3*market_opportunity)` formula found in the original workbook, so Overall scores are directly comparable to the historic matrix.

**Scoring sources**: Regulatory complexity uses TMF Group's country rankings as a reference where available, falling back to analyst judgement where TMF doesn't cover a market. Every other dimension is pure analyst judgement — there's no external tracker for things like local presence or consultant reliance.

**Overall only shows once all 12 dimensions are scored.** A segment missing even one dimension shows Overall as "—", not a partial sum — a partial total would look like a real, comparable score when it isn't. This is deliberate: it's what makes blank cells visible and worth going back to fill in, rather than a gap quietly defaulting to zero. The per-country matrix page shows a "Scored X/12" count per segment, and blank dimension cells are highlighted, so finishing a segment's scorecard is a visible, trackable task rather than something that has to be cross-checked against the workbook by hand.

**Cross-country view**: the global overview page shows every built country's segments against their Overall score in one table — the same shape as the workbook's own `Overall` sheet — so segments can be compared for attractiveness across markets at a glance, with the per-country matrix one click away for the full AUM/allocation/dimension-level detail.

## Asset allocation

Where shown, allocation figures are pulled from the S&P Money Manager Database's allocation export, which reports at the level of an individual institution's individual plan (a firm can sponsor more than one distinct scheme, each with its own totals). Four dimensions are tracked per allocation line: broad asset type (Equities, Fixed Income, Cash & Short-Term, Real Estate, Alternatives), a finer asset style (e.g. Gilts, Corporate Bonds, Private Equity, Infrastructure), investment region, and vehicle type (e.g. Mutual Fund, Unit Trust, Pooled Separate Account).

## Coverage by entity type

- **Pensions, endowments, charities, foundations, government/LGPS, union funds**: S&P Money Manager Database (UK plan-sponsor export plus its companion allocation export).
- **Insurance (Life / Non-life)**: not covered by the S&P exports used so far — S&P's plan-sponsor data is built around pension/endowment sponsorship, not balance-sheet institutional investing. Insurance figures currently come from industry aggregates (e.g. OECD Global Insurance Market Trends, ABI, Bank of England releases) rather than individual-firm records. A firm-level source exists in principle — UK insurers' annual Solvency and Financial Condition Reports (SFCRs) — but hasn't been ingested yet.
- **Sovereign wealth funds**: not applicable to the UK (no domestic SWF). For future countries, the Sovereign Wealth Fund Institute and Global SWF are the identified sources.

## Known open caveats

- **Currency/units**: S&P export figures are confirmed to be in thousands, but the currency itself is inferred from context (magnitudes consistent with GBP for UK institutions) rather than stated explicitly by the source data — worth treating as provisional until independently confirmed.
- **Coverage is not exhaustive**: the S&P database captures a large share of UK institutions but not all of them, particularly smaller schemes. Figures should be read as "tracked" totals, not definitive market totals, unless a segment is specifically noted as reconciled to an official/top-down figure.
- **Manager relationships**: data on which asset managers a given institution uses isn't currently captured from any source in use.
