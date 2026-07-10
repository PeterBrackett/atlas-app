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

**Worked example — UK `DB Pension (Govt)`**: a live MMD read of the Government-category "Total DB Assets" figure is approximately $689.6bn (confirmed USD — see the currency caveat below; roughly £514.65bn at the 2026-07-10 spot rate). The figure actually used, £393.94bn, is a top-down override aligned to official LGPS (Local Government Pension Scheme) statistics, judged more reliable than the raw S&P sum for this segment specifically — even once both figures are compared in the same currency, the LGPS override remains meaningfully lower than the bottom-up estimate. This is the kind of decision recorded on this page rather than left unexplained next to the number.

There's no fixed rule like "official sources always win" — each override is a judgement call on which source is more trustworthy for that segment in that country, made case by case as new sources are found. The intent is for this to be genuinely extensible: as new country-specific statistical sources are identified, they get added as another candidate figure for a segment, and whichever is judged best becomes the one shown.

**Every segment carries its full set of sources, not just the one shown.** Each segment's `sources` array holds every candidate figure that's been captured for it — bottom-up, top-down override, industry aggregate, whatever applies — each with its own `basis`, `aum_bn`, `source` citation, optional `note`, and an `as_of` date marking when that particular figure was captured or published. Exactly one entry is marked `active`; that's the one shown as the segment's headline AUM. Segments with only one known source still get a one-entry `sources` array, so the shape is consistent everywhere and every figure on the site carries an explicit as-of date — which is what would have caught the DB Pension (Govt) bottom-up figure quietly going a month and roughly £55bn out of date, the incident that prompted this section.

On the country page, whenever a segment has more than one source, all of them render as visible rows underneath it — not hidden behind a click — with the active one badged and every figure's as-of date shown alongside it, so staleness is visible at a glance rather than discovered by accident.

**Audit trail.** Every segment also carries a `source_history` array logging every event that's changed its data: the segment's initial build, any later refresh of a source figure, and any switch of which source is active. Each entry records the date, the event type (`created` / `refresh` / `switch`), who made the change, and a plain-language description of what changed and why. This is visible on the country page as an expandable "Audit trail" line under a segment's sources, so it's always possible to see not just what's live now but what was live before, when it changed, and the reasoning — full transparency on which figure is prevailing and why, without losing the history of how it got there.

**Reversibility**: switching which source is "active" for a segment — for example, reverting `DB Pension (Govt)` back to the raw S&P bottom-up sum — is a click on the "Apply" control next to a segment's sources (visible once signed in), which flips the `active` flag, updates the segment's headline `aum_bn`/`basis`, and appends a `switch` entry to `source_history`. Nothing is re-derived from scratch and nothing is lost — the previously active figure stays in the `sources` list, just no longer marked active.

## Opportunity scorecard

Alongside AUM and allocation, every segment carries an opportunity scorecard — the same structure as `FINAL_Country_Matrix.xlsx`, the original research output this whole matrix view is built to reproduce and keep current. Each segment is scored 1–3 on 12 dimensions:

Market opportunity, Outsourced management, Pricing impact, Alignment of investment thinking, Distribution resources required, Regulatory complexity, Client servicing, Local presence required, Languages required, Investor decision-making, Comingled vehicles, Consultant reliant.

**Overall score formula**: `Overall = (3 × Market opportunity) + sum of the other 11 dimensions`. Market opportunity is weighted three times over because it's judged the primary driver of whether a segment is worth pursuing at all — this exactly matches the `=SUM(...)+(3*market_opportunity)` formula found in the original workbook, so Overall scores are directly comparable to the historic matrix.

**Scoring sources**: Regulatory complexity uses TMF Group's country rankings as a reference where available, falling back to analyst judgement where TMF doesn't cover a market. Every other dimension is pure analyst judgement — there's no external tracker for things like local presence or consultant reliance.

**Overall only shows once all 12 dimensions are scored.** A segment missing even one dimension shows Overall as "—", not a partial sum — a partial total would look like a real, comparable score when it isn't. This is deliberate: it's what makes blank cells visible and worth going back to fill in, rather than a gap quietly defaulting to zero. The per-country matrix page shows a "Scored X/12" count per segment, and blank dimension cells are highlighted, so finishing a segment's scorecard is a visible, trackable task rather than something that has to be cross-checked against the workbook by hand.

**Cross-country view**: the global overview page shows every built country's segments against their Overall score in one table — the same shape as the workbook's own `Overall` sheet — so segments can be compared for attractiveness across markets at a glance, with the per-country matrix one click away for the full AUM/allocation/dimension-level detail.

## Asset allocation

The matrix's allocation row is dynamic, not fixed to equities: a dropdown above the matrix picks which broad asset type it shows (Equities, Fixed Income, Cash & Short-Term, Real Estate, Alternatives, or a catch-all Other/Unclassified bucket for asset types too messy to bucket cleanly), and a second dropdown narrows to a specific style within that type where the underlying data supports it (e.g. Gilts or Corporate Bonds within Fixed Income, Private Equity or Infrastructure within Alternatives). This exists so a fixed-income-focused manager can see fixed income figures instead of having to mentally translate from an equity-only view — the point raised that prompted building it.

**How the numbers are derived.** Equities is the originally validated per-segment figure, unchanged. Every other asset type (and the style breakdown within it) is estimated: it comes from the real S&P Money Manager Database allocation export — which reports at the level of an individual institution's individual plan, tagging each allocation line with a broad asset type, a finer asset style, investment region, and vehicle type — but scaled to fill whatever's left of the segment's already-validated AUM once Equities is subtracted out. This keeps the one number already checked against the workbook exactly as it was, while still giving a real-data-informed view of everything else, rather than either leaving those asset types blank or guessing at them from nothing.

**Splitting DB from DC.** The allocation export's `Type of Plan` field usually says directly whether a line belongs to a Defined Benefit or Defined Contribution plan. Where it doesn't (mainly a large "Investment Pool Funds" bucket, plus a smaller "Hybrid DB/DC Plan" bucket), the line is split between DB and DC in proportion to that specific institution's own overall DB/DC asset split from its plan-sponsor record — not a country-wide average — so a DC-heavy firm's pooled assets still weight towards DC.

**Excluded from every calculation**: Bank of England and British Business Bank (S&P's `Reserve Banks` category). Their records reflect central bank balance sheet assets, not staff pension schemes — Bank of England alone shows roughly £1.09 trillion, which would swamp every real segment it touched if included. A single ambiguous for-profit health-sector record is also excluded rather than guessed at.

**Materiality threshold**: within a given segment and asset type, a style only gets its own line if it accounts for at least 5% of that type's value for that segment; smaller ones are folded into "Other/Unspecified" for that segment specifically. This is why picking a specific style in the dropdown can show a real figure for one segment and nothing for another — it doesn't mean the style is absent there, only that it wasn't broken out separately.

## Coverage by entity type

- **Pensions, endowments, charities, foundations, government/LGPS, union funds**: S&P Money Manager Database (UK plan-sponsor export plus its companion allocation export).
- **Insurance (Life / Non-life)**: not covered by the S&P exports used so far — S&P's plan-sponsor data is built around pension/endowment sponsorship, not balance-sheet institutional investing. Insurance figures currently come from industry aggregates (e.g. OECD Global Insurance Market Trends, ABI, Bank of England releases) rather than individual-firm records. A firm-level source exists in principle — UK insurers' annual Solvency and Financial Condition Reports (SFCRs) — but hasn't been ingested yet.
- **Sovereign wealth funds**: not applicable to the UK (no domestic SWF). For future countries, the Sovereign Wealth Fund Institute and Global SWF are the identified sources.

## Known open caveats

- **Currency/units**: confirmed 2026-07-10 — S&P stated directly that the Money Manager Database defaults to USD platform-wide for exported asset values (a local-currency view exists only in the Firm Profile screen, not in exports or asset-related queries). Every country's bottom-up S&P figures across Atlas are USD, not each country's local currency as originally assumed when the site was built — this was corrected across all country data files on 2026-07-10. The one exception is genuinely-local-currency top-down overrides sourced independently of S&P (currently only UK `DB Pension (Govt)`'s LGPS figure, which is GBP) — these are marked with their own `currency` field on the source entry where a segment has more than one source, precisely so they aren't compared to a USD figure as if they were the same currency.
- **Coverage is not exhaustive**: the S&P database captures a large share of UK institutions but not all of them, particularly smaller schemes. Figures should be read as "tracked" totals, not definitive market totals, unless a segment is specifically noted as reconciled to an official/top-down figure.
- **Manager relationships**: data on which asset managers a given institution uses isn't currently captured from any source in use.
- **Non-equity allocation figures are estimated, not entity-verified**: they're a real-data-informed scaling exercise (see "Asset allocation" above), not a re-audited bottom-up sum in their own right — treat the shape (relative proportions) as more reliable than any single absolute figure to two decimal places.
- **Bottom-up source figures can go stale between refreshes**: a segment's bottom-up `sources` entry reflects MMD as of its `as_of` date, not a live query — it only updates when someone re-checks it (manually, or via a future automated refresh against the raw institutions file). Always check a source's `as_of` date before treating it as current; this is exactly what caught UK `DB Pension (Govt)`'s bottom-up figure being about a month and £55bn out of date.
- **Raw institution-level data exists for UK and Ireland**, via a paginated per-country institutions subtable (`getInstitutions.js`/`institutions.html`) built from the S&P exports — but segment-level AUM/allocation figures in each country's JSON are not yet wired to recompute automatically from that subtable when it's refreshed. Refreshing a segment total today means re-deriving it (as was done for `DB Pension (Govt)` above), not something that happens automatically.
