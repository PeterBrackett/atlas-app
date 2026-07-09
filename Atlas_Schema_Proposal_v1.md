# Atlas — Data Schema Proposal (v1, pilot: United Kingdom)

Status: draft for review. No ingestion has started. Country codes use ISO 3166-1 alpha-2 throughout so the model generalises beyond the UK without changes.

Storage model assumed: one raw-records file and one derived-stats file per country on SharePoint (JSON preferred for nesting, CSV export alongside for spreadsheet use), versioned by extraction date rather than overwritten in place.

---

## 1. Institutional Asset Owner Record (raw data, one row/object per institution)

### Identity

| Field | Notes |
|---|---|
| `entity_id` | Atlas-generated stable ID (not from S&P — their IDs may not persist across exports). Format: `{country_code}-{sequence}`, e.g. `GB-0001`. |
| `entity_name` | Legal/registered name. |
| `former_names` | Array — schemes rename or merge often (esp. UK LGPS pools, DB consolidations). |
| `entity_type` | Enum: `Pension – DB`, `Pension – DC`, `Pension – Hybrid`, `Insurance Company`, `Sovereign Wealth Fund`, `Charity/Foundation`, `Endowment`, `Public Reserve Fund`, `Corporate Treasury`, `Other`. |
| `sub_type` | Free text within type, e.g. `Local Government Pension Scheme`, `Master Trust`, `Industry-wide`, `Corporate DB`. |
| `country_code` / `country_name` | ISO 3166-1. |
| `region_state` | For federated countries later (blank for UK, useful for US/Canada/Australia). |
| `domicile_regulator` | e.g. TPR, FCA, PRA. |
| `regulatory_id` | Pension Scheme Registry number / FCA firm reference / charity number — whatever the domicile uses. |
| `sponsor_parent` | Sponsoring employer, parent group, or founding body. |
| `year_established` | |
| `status` | `Active` / `Closed to new members` / `In wind-up` / `Merged` (if merged, link to successor `entity_id`). |
| `website` | |

### Financials

| Field | Notes |
|---|---|
| `total_aum` | Numeric. |
| `aum_currency` | ISO 4217. |
| `aum_as_of_date` | Critical — S&P records lag; without this field the country stats are meaningless. |
| `aum_source` | `S&P export` / `Claude in Chrome` / `Annual report` / `Manual` / `Regulatory filing (SFCR)` / `Industry aggregate (OECD/ABI/BoE)` / `Third-party fund tracker (SWFI/Global SWF)` — see Section 4 for the insurer/SWF-specific sources. |
| `member_count` | Beneficiaries/policyholders, where applicable. |
| `funding_ratio` | DB pensions only. |
| `net_cashflow` | Optional — contributions minus outflows. |

### Plans & asset allocation (nested, not flat columns) — **revised after seeing the allocation export**

The real S&P allocation data (see Section 3) comes in at the **plan level, not the entity level** — a single institution can sponsor more than one distinct plan (e.g. a legacy closed DB scheme plus a separate money-purchase plan), each with its own asset total and its own allocation split. The original flat `allocation` array (one per entity) undersold this — it's now a `plans` array, each element carrying its own allocation breakdown:

```
plans: [
  {
    "plan_type": "Closed or Frozen Defined Benefit Plan",
    "plan_assets": 6691810,
    "as_of_date": "2026-03-31",
    "allocation": [
      { "asset_type": "Fixed Income", "asset_style": "Gilts", "investment_region": "United Kingdom",
        "vehicle_type": "Pooled Separate Account", "value": 970312, "pct_of_plan": 15 },
      { "asset_type": "Fixed Income", "asset_style": "Unspecified", "investment_region": "Unspecified",
        "vehicle_type": "Unspecified", "value": 5720159, "pct_of_plan": 85 }
    ]
  },
  {
    "plan_type": "Money Purchase Plan",
    "plan_assets": 1258,
    "as_of_date": "2026-03-31",
    "allocation": []
  }
]
```

Four dimensions per allocation line, matching what the export actually carries: `asset_type` (broad class — fixed reference list: `Equities`, `Fixed Income`, `Cash & Short-Term`, `Real Estate`, `Alternatives`, `Other/Unclassified`), `asset_style` (finer sub-class, e.g. `Gilts`, `Corporate Bonds`, `Private Equity`, `Infrastructure`, `Hedge Funds` — free text, not enumerable in advance), `investment_region` (geography, mostly `Unspecified` in practice), `vehicle_type` (e.g. `Mutual Fund`, `Unit Trust`, `Pooled Separate Account`). Entity-level `total_aum` is the sum of `plan_assets` across all plans, which reconciles exactly against the plan-sponsor export's `Total PS Assets` in the sample checked (see Section 3).

Caveat: the export has no explicit plan-instance ID — where a firm has two plans of the *same* `plan_type` label (e.g. two separate closed DB schemes), the only way to tell them apart is that they carry different `plan_assets` totals. Ingestion needs to group allocation rows by (`Pid#`, `plan_type`, `plan_assets`) as a composite key. Flagging as a fragility: if two distinct plans at the same firm ever happen to share an identical assets figure, they'd wrongly merge.

### Governance & relationships

| Field | Notes |
|---|---|
| `investment_consultant` | |
| `fiduciary_manager` | If applicable. |
| `custodian` | |
| `actuary` | DB pensions. |
| `asset_managers_used` | Array of `{manager_name, mandate/asset_class, value/pct}` — this is the field S&P Money Manager Database is traditionally strongest on, since it's built around manager–client relationships. |
| `key_contact` | `{name, title, email, phone}` — optional, useful for direct research, not for publication. |

### Provenance (important given two ingestion paths)

| Field | Notes |
|---|---|
| `record_created_date` | |
| `record_last_updated_date` | |
| `data_sources` | Array — a single record may combine S&P export fields with Claude-in-Chrome-filled gaps; track per-field or at least per-record. |
| `data_quality_flag` | `Confirmed` / `Estimated` / `Stale (>12mo)`. |
| `notes` | Analyst commentary. |
| `source_links` | SharePoint links to the underlying export file or extraction screenshot. |

---

## 2. Country-Level Derived Statistics (generated from the raw records above)

| Field | Notes |
|---|---|
| `country_code` / `country_name` | |
| `reporting_period` | The as-of date this rollup represents. |
| `institutions_tracked` | Count of records in scope. |
| `total_aum_tracked` | Sum, normalised to one reporting currency (GBP for UK; suggest also holding a USD-converted figure for cross-country comparison later). |
| `aum_by_segment` | Replaces the originally-drafted `aum_by_type` — see 2b/2c: `[{segment, reported_aum, bottom_up_aum, count, pct_of_country_total}]`, using the confirmed 19-segment taxonomy and the source-registry override mechanism. |
| `top_10_by_aum` | `[{rank, entity_id, entity_name, entity_type, aum, pct_of_country_total}]` |
| `aggregate_asset_allocation` | `[{asset_class, aum_weighted_pct, total_value}]` — weighted average across tracked institutions. Needs a decision: weight by AUM across *all* tracked institutions, or only the top N? (see open question below) |
| `concentration_metrics` | Optional: top-10 share of total AUM, HHI. |
| `coverage_note` | e.g. "94 of an estimated ~130 UK DB schemes >£1bn AUM captured; smaller schemes and most DC master trusts out of scope for v1." — necessary caveat since S&P coverage won't be exhaustive. |
| `generated_date` | |
| `methodology_version` | So historical stats snapshots stay comparable even as the calc logic evolves. |

### 2b. Segment taxonomy — confirmed as the Atlas standard

Confirmed: `FINAL_Country_Matrix.xlsx` is the target end output, so its 19-segment taxonomy replaces the coarser `entity_type` enum from Section 1 as the primary cut for country-level stats (`entity_type`/`sub_type` stay on the raw institution record — Section 1 — as the underlying classification that segments get built from). `aum_by_type` becomes `aum_by_segment`, using:

`DB Pension (Corp)`, `DC Pension (Corp)`, `DB Pension (Govt)`, `DC Pension (Govt)`, `DB Pension (Union)`, `DC Pension (Union)`, `DB Pension (Healthcare non-profit)`, `DC Pension (Healthcare non-profit)`, `DB Pension (Endowments)`, `DC Pension (Endowments)`, `DB Pension (Tax exempt)`, `DC Pension (Tax exempt)`, `Endowments E&F`, `Tax exempt E&F`, `Foundations E&F`, `Healthcare non-profit E&F`, `Life insurance`, `Non-life insurance`, `SWF`.

This maps cleanly onto data already in the schema: it's the S&P `Category` field (Section 3's mapping table) crossed with DB/DC split from `Total DB Assets`/`Total DC Assets`, plus the two insurance segments and SWF sourced per Section 4.

### 2c. Source overrides — designed as an extensible registry, not a single override field

Confirmed: government/official sources should override the bottom-up S&P aggregation (matches the UK `DB Pension (Govt)` finding — $393.94bn in your matrix vs. $634.45bn summing the raw S&P Government category, landing on the LGPS official figure instead). But you've also said you want the portal itself to let you add new sources later that could override whatever's currently active — e.g. a country-specific pension statistics body you haven't sourced yet. A single `override_source` field can't do that; it needs to be a small registry per (country, segment, metric), so a new source can be added at any time without discarding what's already there:

```
segment_aum_sources: [
  {
    "source_id": "GB-DBGOVT-01",
    "source_name": "S&P bottom-up aggregate",
    "source_type": "Bottom-up",
    "value": 634453886,
    "as_of_date": "2026-06-01",
    "precedence": 3,
    "active": false
  },
  {
    "source_id": "GB-DBGOVT-02",
    "source_name": "DLUHC LGPS annual return, FY2023/24",
    "source_type": "Government",
    "value": 393940000,
    "as_of_date": "2024-03-31",
    "precedence": 1,
    "active": true
  }
]
reported_aum: 393940000   // value of whichever source has active: true
```

**Precedence is a judgment call on data veracity, not a fixed rule.** Government sources will usually win, but not automatically — `precedence` is set case by case based on how much the specific source is trusted for that specific segment/country, and `active`/`precedence` can be changed by hand at any time as new sources get added or an existing one turns out to be less reliable than assumed. So there's no hardcoded "government > aggregate > bottom-up" ranking baked into the system — it's a manually-set ordering per (country, segment, metric), with the registry just holding whatever's been decided plus the reasoning (`override_rationale`) for why. Same registry shape applies to the `asset_managers_used`-style gaps and to Section 4's insurer/SWF sourcing — it's one mechanism, not a special case for pensions.

**Pro-ration rule, since you flagged this specifically**: when `reported_aum` differs from `bottom_up_aum`, every percentage/allocation figure that was computed against the bottom-up total (segment share of country AUM, `aggregate_asset_allocation` weights, the matrix's "Equity allocation ($bn)"-style rows) must be rebased against `reported_aum`, not left keyed to the smaller/larger bottom-up number. Mechanically: keep the *shape* of the bottom-up allocation breakdown (e.g. "20% of bottom-up institutions' assets are in equities") but scale the absolute values so they sum to `reported_aum` instead of `bottom_up_aum`. A `prorated: true` flag on any derived figure marks that its dollar value no longer reconciles to a literal sum of underlying institution records — necessary so no one mistakes a pro-rated figure for a bottom-up one later.

This also confirms the `plans`/`allocation` nested structure from Section 1 was the right call: it lets any "opportunity allocation" figure (e.g. the `Equity allocation ($bn)` row in your matrix) be computed on demand by filtering `asset_type` at query time, rather than needing a fixed allocation metric baked into the schema.

---

## 3. S&P Money Manager Database export → schema mapping

Confirmed against four real samples, all from the same underlying database (`Pid#`/`Firm Name` values match across all four, and firm-level totals reconcile exactly). The first two were pulled manually; the second two were pulled via a live Claude-in-Chrome MMD WebAccess session on 2026-07-08, re-querying UK with Report Output set to Manager and Provider respectively (Company and Asset Allocation were also re-pulled in that same session as a freshness check — row counts match the original samples almost exactly, see below):

- `fed68457-results.xls` / re-pulled as `results (2).xls` — UK plan sponsors, 1,606 / 1,606 records, one row per institution. Covered in the original mapping below.
- `results (1).xls` — UK plan **allocation detail**, 9,832 / 9,832 rows, grain is one row per (institution, plan, asset-type line). This is the export that fills the allocation gap flagged below — updated mapping follows. (Re-pull landed as `results (1).xls` too, overwriting nothing — both counts identical.)
- `results (3).xls` — UK **manager relationships**, 6,436 rows, grain is one row per (institution, plan, named manager). New in this round — fills the `asset_managers_used` gap. See File 3 below.
- `results (4).xls` — UK **provider relationships**, 2,757 rows, grain is one row per (institution, named provider). New in this round — fills the `custodian`/`actuary`/`investment_consultant` gaps. See File 4 below.

All four files share the same quirk: the `.xls` extension is misleading — they're HTML tables (Excel's "web page" export format), not binary/OOXML workbooks. Fine for `pandas.read_html` or Excel itself, but any ingestion script needs to parse them as HTML.

### File 2 — allocation export (`results (1).xls`)

**Actual column headers (16, in order):** `Firm Name`, `Category`, `Address`, `City`, `State`, `Phone#`, `Total PS Assets`, `Type of Plan`, `Plan Assets`, `Investment Region`, `Pooled Investments`, `Asset Type`, `Asset Style`, `Dollar Amount of Asset Type in Plan (in 000's)`, `% of Asset type in Plan`, `Pid#`

Key findings:

- **Grain confirmed as (firm, plan, asset-type line)** — 1,605 unique firms across 9,832 rows, matching the `Pid#` set from the plan-sponsor file exactly. `Type of Plan` breaks a firm into its distinct schemes (`Defined Benefit Plan`, `Closed or Frozen Defined Benefit Plan`, `Money Purchase Plan`, `Investment Pool Funds`, `Hybrid DB/DC Plan`, `Endowment Fund`, and 7 others) — this is the detail that `Total DB Assets`/`Total DC Assets` in File 1 was only summarising.
- **Units confirmed**: the column header itself states `(in 000's)` — this resolves the units half of the earlier open question. Currency is still not explicit (the column is generically labelled "Dollar Amount", almost certainly a template artefact rather than a literal indication of USD — magnitudes only make sense as GBP thousands for UK entities, e.g. Bank of England's plan assets consistent with the ~£1.09bn figure from File 1). Still needs your confirmation, just narrowed.
- **Allocation coverage**: 1,545 of 1,605 firms (96%) have at least one populated allocation line; 231 rows (mostly smaller DC/money-purchase plans) carry a plan total with no breakdown at all — consistent with smaller plans being reported at a coarser level.
- **Core `Asset Type` taxonomy**: five clean buckets dominate — `Fixed Income`, `Cash & Short-Term`, `Equities`, `Alternatives`, `Real Estate` — plus a long tail (~215 more values, many clearly free-text/data-entry noise: manager names, plan names, or one-off labels like "Homebuy & Equity Loans"). Needs a cleaning pass mapping the long tail down to the fixed reference list plus `Other/Unclassified`, exactly as the original schema draft anticipated.
- **`Asset Style` is a second, finer dimension**, not a duplicate of `Asset Type` — e.g. `Gilts`, `Corporate Bonds`, `Private Equity`, `Infrastructure`, `Hedge Funds`, `Liability-Driven`. Mostly `Unspecified` (5,730 of 9,832 rows) but populated often enough to be worth keeping as its own field rather than collapsing into `Asset Type`.
- **`Investment Region`** is mostly `Unspecified` (8,191/9,832) but shows real values (`United Kingdom`, `Global`, `Emerging Markets`, etc.) often enough to keep as its own dimension rather than folding into the asset class label as originally sketched (e.g. `Equities – UK`).
- **Percentages reconcile per plan-instance**: `% of Asset type in Plan` sums to ~100 for a given (`Pid#`, `Type of Plan`, `Plan Assets`) group in the large majority of cases (median exactly 100), with some noise up to 200 in a small number of groups — worth a data-quality check during ingestion rather than trusting blindly.
- **`Total PS Assets` reconciles exactly** against the sum of `Plan Assets` across a firm's plans (checked against AkzoNobel UK Ltd.: 6,691,810 + 2,933,685 + 1,258 + 4,785 = 9,631,538, matching File 1's `Total PS Assets` for the same firm to the pound). Good sign the two exports are internally consistent and safe to join on `Pid#`.
- **`Category` distribution matches File 1 proportionally** (`Corporation`, `Tax Exempt Organization`, `Government`, `Endowment`, etc. in the same rank order) — same entity-type mapping table below applies to both files.
- New identity field available here that File 1 lacked: **`Address`** (street-level, in addition to `City`).

### File 3 — manager export (`results (3).xls`)

**Actual column headers (16, in order):** `Firm Name`, `Category`, `Total PS Assets`, `Plan Name`, `Type of Plan`, `Plan Assets`, `Management Co. Name`, `Dollar Amount Managed (in 000's)`, `Managed Investment Region`, `Managed Pooled Investments`, `Managed Asset Type`, `Managed Asset Style`, `Year hired`, `Total DB Assets`, `Total E/F Assets`, `Total DC Assets`

Key findings:

- **Grain confirmed as (firm, plan, named manager)** — 6,436 rows (incl. one blank trailer); 5,199 have a populated `Management Co. Name`, the rest are plan rows on file with no named manager attached (same "known plan, no detail" pattern as File 2's 231 unallocated allocation rows).
- **No `Pid#` column** — the only identifier in this export is `Firm Name`, unlike every other MMD export sampled so far. Cross-checked: all 1,604 distinct firm names in this file match exactly against File 1's `Firm Name` set (zero orphans), so a name join is safe *except* for File 1's one duplicate name (Bank of England, 2 rows with different `Pid#`s) — flag duplicate firm names for manual `Pid#` assignment rather than assuming a 1:1 join.
- **Coverage**: 675 of 1,604 UK firms (42%) have at least one named manager relationship on file — a meaningfully higher hit rate than expected for what the original schema draft flagged as a total gap.
- **Grain double-checked against two multi-plan firms** (AkzoNobel UK Ltd., Associated British Foods plc) — manager sets genuinely differ plan-to-plan, confirming this isn't a repeated cross-join artifact. One large multi-plan firm (Church of England) does show a near-identical ~19-name manager panel recurring across several of its plans — plausible for a shared manager panel serving related schemes under one sponsor, not treated as a defect, but worth a sense-check if the pattern recurs at scale during full ingestion.
- **`Dollar Amount Managed` populated on only 1,795 of 5,199 named-manager rows (35%)** — most relationships are disclosed by manager name only, with no value. Treat the value as optional per manager entry, same as File 2's own patchiness on dollar amounts.
- **`Managed Asset Type` mirrors File 2's core taxonomy** (`Equities`, `Alternatives`, `Fixed Income`, `Real Estate`, `Cash & Short-Term` dominate) — same cleanup/mapping rules apply.
- **`Managed Investment Region` mostly `Unspecified`**, same pattern as File 2.
- **`Year hired` populated on only 216 of 5,199 rows (4%)** — sparse, keep optional.
- **`Total PS/DB/DC/E&F Assets` repeat File 1's totals verbatim on every row** — redundant but useful as a self-contained sanity check without a join.

**Proposed `asset_managers_used` population rule**: group by (`entity_id` via `Firm Name`→`Pid#`, `Management Co. Name`), collapsing the plan-level repeats to one array entry per distinct manager per firm. The schema's existing shape — `{manager_name, mandate/asset_class, value/pct}` — accommodates this with a straight collapse, taking whichever row for that manager carries a non-null `Dollar Amount Managed` / `Managed Asset Type` / `Managed Investment Region` / `Year hired` where more than one plan-level row exists for the same manager.

### File 4 — provider export (`results (4).xls`)

**Actual column headers (14, in order):** `Firm Name`, `Category`, `Address`, `City`, `State`, `Phone#`, `Total PS Assets`, `Total DB Assets`, `Total DC Assets`, `Total E/F Assets`, `Provider Name`, `Year Provider was hired`, `Provider Type`, `Pid#`

Key findings:

- **Grain confirmed as (firm, named provider relationship)** — 2,757 rows (incl. one blank trailer); 1,075 have a populated `Provider Name`. Has `Pid#`, unlike File 3 — joins cleanly against File 1 on `Pid#` (all 1,605 UK `Pid#`s present, zero orphans).
- **Coverage**: 391 of 1,605 UK firms (24%) have at least one named provider on file.
- **`Provider Type` breakdown** across the 1,075 named rows: `Custodian` 349, `Actuary` 303, `Consultant` 276, `Third Party Administrator` 78, `Global Custodian` 28, `Master Trustee` 19, `OCIO` 11, `Recordkeeper` 9, plus 2 rows labelled `AUD` (likely a data-entry artifact, not a real provider type).
- **No `Plan Name` field** — unlike File 3, a provider relationship here can't be attributed to a specific plan when a firm has several. A genuine loss of granularity relative to the manager export, not an ingestion choice.
- **Moderate duplication**: the 1,075 named rows collapse to 893 distinct (`Pid#`, `Provider Name`, `Provider Type`) combinations — a ~1.2x duplication factor, most visible on firms with several plans, where the same provider recurs once per plan (e.g. Lloyds Banking Group plc shows "Willis Towers Watson – Actuary" five times within its 24-row block). Dedupe on (`Pid#`, `Provider Name`, `Provider Type`) before loading.
- **`Year Provider was hired` populated on only 47 of 1,075 rows (4%)** — sparse, same pattern as File 3's `Year hired`.
- New identity field available here that File 1/File 3 lack: street-level **`Address`** (same as File 2).
- **`Total PS/DB/DC/E&F Assets` repeat File 1's totals verbatim**, same redundant-but-useful pattern as File 3.

**Proposed schema mapping**: `Provider Type` is the natural discriminator for three of the existing single-value governance fields — `custodian` (`Provider Type` = `Custodian` or `Global Custodian`), `actuary` (`Provider Type` = `Actuary`), `investment_consultant` (`Provider Type` = `Consultant`). `Master Trustee`, `OCIO`, `Third Party Administrator`, and `Recordkeeper` have no matching field in the current schema — `OCIO` in particular overlaps conceptually with the still-unpopulated `fiduciary_manager` field, which is a reasonable home for it, but flagging for sign-off rather than deciding unilaterally. Separately: `custodian`/`actuary`/`investment_consultant` are currently scalar fields in the schema, but a small number of firms have more than one provider of the same type after dedup — these fields likely need to become arrays, or the schema needs an explicit "primary provider" rule.

### File 1 — plan-sponsor export (`fed68457-results.xls`)

**Actual column headers (13, in order):**

`Firm Name`, `Category`, `City`, `State`, `Country`, `Phone#`, `Total PS Assets`, `Total E/F Assets`, `Pension Assets`, `Other Assets`, `Total DC Assets`, `Total DB Assets`, `Pid#`

**Answers to the five open questions:**

1. **Format/headers**: HTML table masquerading as `.xls`, 13 columns as listed above. One trailing blank row and one fully-blank row present — both need stripping on ingest.
2. **Asset allocation**: Not included in *this* file — only aggregate AUM split by plan type (DB / DC / Pension total / Endowment & Foundation total / "Other"). The `Other Assets` column exists in the header but is empty for every UK row in this sample. **Update: a second export (`results (1).xls`, see above) turns out to carry full allocation detail** — it's a separate report from the same database, not a field missing from this one. Original answer stands for File 1 specifically; the broader open question is now resolved at the database level.
3. **Unit of record**: The asset owner (plan sponsor), one row per institution — not the manager. This appears to be S&P MMD's "plan sponsor" module specifically. There is no manager-relationship data (no manager names, mandates, or values) anywhere in this export — the `asset_managers_used` field in the draft schema cannot be populated from this file at all.
4. **Persistent unique ID**: Yes — `Pid#`, populated for 1,605 of 1,606 rows (the one gap is the blank trailer row), no duplicates. This is a reliable key to use as (or map to) `entity_id`, better than matching on `Firm Name` alone (one duplicate name already appears in this sample).
5. **Snapshot vs. historical**: Snapshot only. There is no date field anywhere in the file — not in the headers, not in a title row, not in a footer. Each export is a single point-in-time pull with no self-declared as-of date; Atlas has to stamp `aum_as_of_date` itself at extraction time (or ask you to note the date S&P shows on-screen when downloading, if it shows one there).

**Category → entity_type mapping** (S&P's `Category` field, 9 values observed):

| S&P `Category` | Count | Maps to draft `entity_type` | Notes |
|---|---|---|---|
| `Corporation` | 774 | `Pension – DB` / `Pension – DC` / `Pension – Hybrid` | Split by which of `Total DB Assets`/`Total DC Assets` is populated, not by `Category` itself — this is the sponsoring employer, not the scheme name. |
| `Tax Exempt Organization` | 541 | `Charity/Foundation` | Largest non-pension bucket. |
| `Endowment` | 152 | `Endowment` | |
| `Government` | 111 | `Pension – DB` (public), `sub_type: Local Government Pension Scheme` | Sampled rows are almost all LGPS funds and pools (e.g. Avon Pension Fund, Brunel Pension Partnership). |
| `Health Service Organization – nonprofit` | 10 | `Charity/Foundation` | NHS trusts, hospices. |
| `Union` | 8 | `Pension – DB`, `sub_type: Industry-wide` | Union-sponsored pension funds. |
| `Private Foundation` | 6 | `Charity/Foundation` | |
| `Reserve Banks` | 2 | *No clean fit* | Bank of England and British Business Bank — their own staff pension schemes, not really a distinct institutional-investor type. Suggest folding into `Pension – DB` with `sub_type: Public Institution`. |
| `Health Service Organization – for profit` | 1 | `Corporate Treasury` or `Charity/Foundation` | Only one row (HCA International) — needs a judgment call. |

**Gaps confirmed across Files 1 and 2 — nothing here populates**: `regulatory_id`, `domicile_regulator`, `sponsor_parent` (beyond firm name itself), `year_established`, `status`, `website`, `funding_ratio`, `member_count`, `net_cashflow`, and everything under Governance & relationships (`investment_consultant`, `fiduciary_manager`, `custodian`, `actuary`, `asset_managers_used`, `key_contact`). **Update**: the allocation sub-table is no longer a gap — File 2 covers it — and manager relationships and three of the six governance fields are no longer gaps either: File 3 populates `asset_managers_used` (42% firm coverage) and File 4 populates `custodian`, `actuary`, and `investment_consultant` (24% firm coverage) — see Files 3 and 4 above. Still genuinely unsourced by any of the four files: `regulatory_id`, `domicile_regulator`, `sponsor_parent`, `year_established`, `status`, `website`, `funding_ratio`, `member_count`, `net_cashflow`, `fiduciary_manager` (though File 4's `OCIO` provider type is a plausible partial source, see File 4 above), and `key_contact` (would need S&P's separate `Contacts` report output, not pulled in this round).

**Coverage gap — no insurers, no SWF**: Neither export contains a single `Insurance Company` or `Sovereign Wealth Fund` record. That tracks with these being *plan sponsor* reports — insurers and a SWF (moot for the UK, which has none domestically) wouldn't appear as "sponsors" of a pension/endowment plan. If insurers are meant to be in scope for the UK database, they'll need a different S&P module (if one exists in your subscription) or a different source entirely (e.g. ABI/PRA regulatory returns, annual reports) — flagging rather than assuming, per your original ask.

---

## 4. Insurance & sovereign wealth fund data — alternative sources

Agreed: S&P's plan-sponsor exports are the wrong tool for these two entity types (they're built around pension/endowment sponsorship, not balance-sheet institutional investing). Researched alternatives below — split by what's usable now (UK insurers, pilot-relevant) versus later (SWFs, not relevant to the UK pilot since the UK has no domestic SWF, but worth scoping now so the schema doesn't need rework when non-UK countries come on).

### UK insurance companies

**Identity/universe layer** — the Bank of England/PRA publishes a free, public **"List of PRA-regulated insurers"** as a CSV, updated monthly, covering every UK-authorised insurer (including EEA/Gibraltar branches). This is the natural `entity_id`/`entity_name`/`regulatory_id` source for `entity_type: Insurance Company` — same role S&P's `Pid#` plays for pensions.

**AUM + asset allocation layer** — every UK insurer/group must publish an annual **Solvency and Financial Condition Report (SFCR)** under Solvency UK (the post-Brexit successor to Solvency II). Each SFCR discloses total assets, own funds, technical provisions, and a genuine asset-class breakdown (by CIC — Complementary Identification Code — categories: government bonds, corporate bonds, equities, CIUs/collective investment undertakings, property, etc.). This is a real per-entity primary source, filed annually, and structurally similar enough to the `plans`/`allocation` shape already in Section 1 that it should slot in without a schema change — an insurer's SFCR balance sheet plays the same role as a pension's DB/DC plan.

Two ways to get SFCR data in: pull each insurer's SFCR directly from their own website (free, but PDF-based — Claude in Chrome territory, one filing per firm per year) or use a commercial aggregator like **Solvency II Wire Data**, which indexes 24,000+ SFCRs/QRTs across the UK and Europe in a queryable form — effectively the "S&P" of this space, at a cost, if manual PDF pulls across however many UK insurers matter turns out to be too slow.

**Country-level benchmark layer** — **OECD Global Insurance Market Trends** (annual, covers 67 jurisdictions including the UK) publishes aggregate, country-level asset allocation and total-assets figures — e.g. UK life insurers' average bond/equity/CIS split. This is aggregate, not per-entity, so it can't populate individual institution records, but it's a useful sense-check for the derived country-level stats (Section 2) once real UK insurer records are loaded — if Atlas's aggregate allocation is wildly off the OECD figure, that's a signal something's wrong with coverage or classification. The **Bank of England's own quarterly/annual insurance aggregate data** and the **ABI's industry data packages** (including a UK top-20 insurer list by gross written premium) serve the same benchmarking role, UK-specific.

**Practical open question for you**: the UK has several hundred PRA-authorised insurers, most small. Pulling an SFCR per firm is real per-firm effort (unlike the one-shot S&P CSV export). Worth deciding now whether v1 targets the largest N insurers by assets (mirroring how the pension side will likely apply a coverage threshold) rather than the full authorised universe.

### Sovereign wealth funds (not UK-relevant now, but scoping for later)

Not applicable to the UK pilot — no domestic UK SWF exists — but since the schema needs to generalise, worth naming sources now: the **Sovereign Wealth Fund Institute (SWFI)** publishes fund-level rankings and profiles by country and AUM; **Global SWF** is a data platform tracking 400+ SWFs *and* public pension funds together (closer to Atlas's own combined institutional-investor remit than most sources), including a "top 100" portfolio view; **IFSWF** offers governance/policy disclosures for member funds but is thinner on AUM/allocation detail. None of these are S&P, and none are needed for UK — flagging for whenever the second country is scoped.

### Schema implication

`aum_source` (Financials) and `data_sources` (Provenance) need a couple more enum values beyond `S&P export` / `Claude in Chrome` / `Annual report` / `Manual`: something like `Regulatory filing (SFCR)` and `Industry/aggregate benchmark (OECD/ABI/BoE)` for cases where a figure is a sense-check rather than the entity's own record, and `Third-party fund tracker (SWFI/Global SWF)` for when SWFs come into scope. No structural change needed beyond the enum — the `plans`/`allocation` shape already accommodates a single-balance-sheet insurer as a degenerate one-plan case.

---

## 5. The opportunity scorecard layer — confirmed in scope, built

Confirmed and now built: the scorecard belongs in Atlas, mixing quantitative (AUM, equity allocation — already on the segment record per Section 2b) with qualitative (the 1–3 judgment scores) dimensions, matching how `FINAL_Country_Matrix.xlsx` lays each (country, segment) cell out.

**Implementation simplified from the original per-engagement design below.** Rather than a separate `scorecard` object keyed by `engagement_id` (multiple snapshots per segment, one per client project), the built version stores a single live `scorecard` object directly on each segment record in the country JSON — the same place `aum_bn`/`basis`/`sources`/`source_history` already live. There's one current scorecard per (country, segment), not one per engagement. This was a deliberate scope cut: a running, editable-in-place scorecard is what actually nudges completion (missing dimensions show as gaps to fill in, not a form to duplicate per project), and per-engagement history can be layered on later if a real need for it shows up — nothing about the shape below blocks that.

Actual shape, per segment:

```
"scorecard": {
  "market_opportunity": 1,
  "outsourced_management": 1,
  "pricing_impact": 1,
  "alignment_of_investment_thinking": 1,
  "distribution_resources_required": 1,
  "regulatory_complexity": 1,
  "client_servicing": 2,
  "local_presence_required": 1,
  "languages_required": 1,
  "investor_decision_making": 2,
  "comingled_vehicles": 1,
  "consultant_reliant": 3,
  "scored_date": "2026-06-01"
}
```

Any dimension key can simply be absent — that's how "not yet scored" is represented, rather than a null or a placeholder score. `Overall` is not stored; it's computed on the fly as `(3 × market_opportunity) + sum(the other 11 dimensions)`, matching the workbook's own formula, and deliberately returns "not computed" (not a partial sum) if any of the 12 keys is missing. A country-level `scorecard_note` field carries the provenance text (e.g. "regulatory_complexity uses TMF Group rankings where available; all else is analyst judgment") once per country rather than repeating it on every segment.

Resolved: **Regulatory complexity uses TMF Group rankings as its go-to source when available** — external data preferred, falling back to analyst judgment for a country TMF doesn't cover. **All other 11 dimensions are pure analyst judgment**, with no external tracker to fall back on.

Resolved (was open): **new segments/countries start with no scorecard fields at all — nothing pre-populates from a prior country or engagement.** A blank segment is the intended default; filling it in is the explicit workflow the matrix UI is built around (a "Scored X/12" indicator and highlighted blank cells per segment, editable inline once signed in).

---

## 6. Dynamic multi-asset-class allocation — built

Confirmed and built: the matrix's single "Equity allocation" row is now a dynamic row driven by two dropdowns (broad asset type, then an optional finer style within it), so a segment can be viewed by whichever asset class a given mandate actually cares about, not just equities.

**Why this couldn't just be a mechanical re-aggregation.** Attempting to reconstruct segment membership directly from the raw S&P files (Category field + DB/DC asset columns) does not reconcile against the already-validated segment AUM figures — Corporation-category institutions summed to roughly £1,864bn expected vs. as little as ~£1,380bn from a naive Category+DB/DC split, with Government and Tax Exempt Organization off by tens of billions in the other direction. Some of this traces to specific entities S&P categorises in a way that doesn't match how they were originally treated (Pension Protection Fund and NEST both show up under `Corporation` despite being public statutory bodies), but the full gap couldn't be safely reverse-engineered — meaning real manual reclassification judgment went into the original matrix numbers that isn't recoverable purely from the raw columns.

**Resolved approach**: rather than risk quietly presenting allocation figures that contradict AUM numbers already checked and used with clients, each segment's `allocation` array anchors its Equities entry to the already-validated figure unchanged, and derives every other asset type (Fixed Income, Cash & Short-Term, Real Estate, Alternatives, Other/Unclassified) plus the finer style breakdown within each by taking the real *shape* (relative proportions) from the S&P allocation export for that segment's DB/DC/E&F institutions, then scaling those proportions to exactly fill the remainder of the segment's known AUM once Equities is subtracted. Every segment's allocation entries sum exactly to its `aum_bn`. DB vs DC vs Endowment/Foundation splitting uses the allocation export's own `Type of Plan` field directly where it's unambiguous, falling back to each specific institution's own DB/DC asset ratio (from its plan-sponsor record) for the ambiguous plan types — mainly a large "Investment Pool Funds" bucket.

**Excluded throughout**: `Reserve Banks` (Bank of England, British Business Bank) and one ambiguous for-profit health-sector record. Bank of England's S&P record shows ~£1.09 trillion — clearly its own balance sheet/reserve assets, not a staff pension scheme — and would have distorted any segment it touched by an order of magnitude larger than the segment itself.

**Shape**, per segment:

```
"allocation": [
  { "asset_type": "Equities", "value_bn": 205.196711, "pct_of_segment": 11.0,
    "styles": [ { "asset_style": "Other/Unspecified", "value_bn": 205.196711 } ] },
  { "asset_type": "Fixed Income", "value_bn": 1023.199847, "pct_of_segment": 54.9,
    "styles": [
      { "asset_style": "Bonds", "value_bn": 396.514944 },
      { "asset_style": "Government Bonds", "value_bn": 109.443483 },
      { "asset_style": "Other/Unspecified", "value_bn": 134.353605 }
    ] }
]
```

A style only gets its own line if it's at least 5% of that asset type's value for that segment; smaller ones fold into `Other/Unspecified`. A segment-level `allocation_incomplete: true` flag marks the two insurance segments, which have no S&P allocation data at all (S&P's plan-sponsor exports don't cover insurers) and so only carry an Equities entry.

**Built**: `atlas-site/data/uk_institutions_2026-07-08.json` — the first real raw per-country institutions subtable, per Section 1's shape, ingested 2026-07-08 from the four MMD exports (Files 1–4 above). **1,603 records**, `entity_id` format `GB-####`. Per Peter's instruction, the two `Reserve Banks`-category records (Bank of England's £1.09tn central-bank balance sheet and British Business Bank's £5.6bn) are excluded from this dataset entirely — not just from aggregate stats, as `uk.json`'s segment allocation already did, but dropped from the raw institutions file itself, matching the 1,605 → 1,603 record count now used everywhere (this file, `global.json`'s GB row). One side benefit: a second `Bank of England`-named record (Category `Corporation`, a legitimate £4.4bn staff pension scheme, distinct from the excluded £1.09tn Reserve Banks record) was previously flagged as ambiguous for the Manager-export join since two same-named rows existed — with the Reserve Banks record now removed, the name is unique and that ambiguity note is gone; 0 named managers were actually on file for it, though, so the join now resolves cleanly but has no data to show. Headline stats on the 1,603: 674 firms (42%) carry at least one `asset_managers_used` entry, 227 have a `custodian`, 239 an `actuary`, 191 an `investment_consultant`, 11 a `fiduciary_manager` (mapped from `OCIO` provider type). Spot-check reconciled exactly against the doc's own AkzoNobel UK Ltd. example (£9,631,538k total, matching the sum of its four plans).

`global.json`'s GB row has also been corrected: it previously showed 2,437 institutions / £3.45tn, traceable to an earlier, differently-scoped S&P pull from schema-design time rather than the 1,605/1,603-firm baseline this doc is anchored to. It now reads **1,603 institutions / £2,531,575,923k (£2.53tn)**, matching this file's bottom-up sum with Reserve Banks excluded — the same convention `uk.json`'s segment stats already use, per Peter's confirmation.

**Now wired into the site**: `atlas-site/institutions.html` (linked from `country.html` via a "Browse individual institutions →" link, shown only when a country has an institutions file) is a searchable/sortable/paginated table over all 1,603 records — filter by entity type or by which relationship fields are populated (manager/custodian/actuary/consultant), click a row to expand its plans, allocation detail, and governance relationships. It reads `data/{code}_institutions_latest.json`, a stable-filename pointer copy of the current dated extraction (`uk_institutions_2026-07-08.json` today) — re-pulls should still land as a new dated file per the storage model above, with `_latest.json` re-copied to point at whichever is current.

**Paginated API added (2026-07-08), needed before this scales past small/mid-size countries.** A single flat JSON file works fine at the UK's size (1,603 records, ~7MB), but projecting the same per-record density onto the largest country in `global.json` — the US's placeholder 281,767 institutions — lands at roughly **1.2GB for one file**. That breaks three things independently: no browser can hold that in memory to render a table, it exceeds Azure Static Web Apps' own app-size ceiling (250MB free / 500MB standard), and it exceeds GitHub's 100MB single-file limit if this repo is ever pushed there. `atlas-api/src/functions/getInstitutions.js` (route `institutions/{country}`) addresses this by doing search/filter/sort/pagination **server-side** and returning only the requested page (default 50 records, capped at 200) — `institutions.html` now tries this endpoint first and only falls back to downloading the whole static file client-side if the API isn't reachable (e.g. running the site off a bare local `http.server` with no Functions backend attached, which is how this was tested).

This is a real fix for the file-size and browser-memory problem, but only a partial fix for a truly large country: the Function still loads and JSON-parses the *entire* per-country file into memory on each cache miss (5-minute in-memory cache per warm instance, keyed by country) before filtering — trivial for the UK's 7MB, but a ~1.2GB file would risk Consumption-plan memory/timeout limits on the Function itself even though the browser never sees it. Flagged directly in `getInstitutions.js`'s comments: a US-scale ingestion will need the SharePoint-JSON-per-country model replaced with a real indexed store (Cosmos DB, Azure SQL, or Table Storage) that the Function queries directly, rather than loading a full file per request. Not built — worth scoping specifically when a country that large is actually ingested, not before.

**Deployment note, not yet done**: like the rest of the backend, `getInstitutions.js` reads from the `AtlasData` SharePoint folder via Graph API (same pattern as `getData.js`/`getNotes.js`), not from the local `atlas-site/data/` folder used for local testing. For the live API to actually serve UK institutions data, `uk_institutions_latest.json` needs uploading to that SharePoint folder, and the new Function needs deploying alongside the existing ones — neither has happened yet as part of this session's work.

**Partial progress on the refresh mechanism**: a Claude-in-Chrome session can now drive the MMD WebAccess Quick Search end-to-end — set Report Output, set Country, run the search, and trigger "Export to Excel" — confirmed working for all four report types (Company, Asset Allocation, Manager, Provider) on 2026-07-08. The remaining manual step: MMD's export button opens a popup that immediately fires a native browser download, which lands in Peter's own Downloads folder rather than anywhere Claude's session can reach — Claude cannot currently retrieve the exported file itself, so Peter still has to move each file into the Atlas folder by hand before Claude can parse it. Full automation would need either a way to redirect/intercept that download, or a subscription tier with real API access.

One real lesson from this round: the first attempt at all four Ireland export clicks silently failed — a JS timing bug (measuring the "Go" button's position via `getBoundingClientRect()` immediately after `scrollIntoView()`, before the browser had finished the resulting layout reflow) meant every click landed ~50px off target. Nothing downloaded, but nothing errored either, so it looked like it had worked until Peter noticed the file timestamps predated the session. Fixed by separating the scroll and the position-read into two distinct steps (screenshot after scrolling, rather than trusting a same-call rect), and by explicitly confirming a transient popup tab appears after every export click rather than assuming success from the click alone. Worth remembering for any future country pull using this same browser-automation approach.

**Second country validated (2026-07-08): Ireland.** Same four-report MMD pull, same ingestion script (parameterised for country code/name and `IE-####` entity IDs), same Reserve Banks exclusion (Central Bank of Ireland's €91.96bn Reserve Banks-category record removed, mirroring the Bank of England treatment) — landed at 43 institutions (44 Company rows minus the one Reserve Bank), €62.5bn total AUM. Coverage is thinner than the UK's: only 5 firms (12%) have a named manager, 3 a custodian, 2 an actuary, 1 an investment consultant — real numbers from a real pull, not a partial/broken run, just a smaller market with less MMD relationship data on file. Currency assumption switched to EUR for this country (still unconfirmed from the source data itself, same caveat as GBP for the UK — S&P's export never states a currency explicitly).

Saved as `atlas-site/data/ie_institutions_2026-07-08.json` / `ie_institutions_latest.json`, same shape and same access pattern as the UK — `institutions.html?c=ie` should work immediately via the static-file fallback. The live `/api/institutions/ie` path needs the same SharePoint upload + nothing-else-required deploy step already documented above for the UK (the Function code itself is already country-agnostic, no changes needed there).

`global.json`'s Ireland row has been corrected from the old schema-design-era placeholder (77 institutions / €298.6bn — provenance unknown, same category of stale estimate as the UK's original 2,437/£3.45tn figure) to the real sourced numbers (43 / €62.5bn). Status is deliberately left as `coming_soon` rather than `built` — that status is tied to having a segment/scorecard derivation (the Section 2 treatment `uk.json` has), which does not exist for Ireland. Raw institution-level browsing works; the country-level matrix page does not, and building it would mean the same kind of judgment-heavy segment classification and scorecard work that was done by hand for the UK — not attempted here.

## Open questions before build

Resolved:

- ~~Asset allocation has no source~~ — resolved. `results (1).xls` is the source; see Section 3.
- ~~Units unclear~~ — narrowed. Confirmed as thousands (explicit in the column header); currency symbol still unconfirmed (see below).
- ~~Insurers/SWFs aren't in S&P~~ — resolved as a sourcing question. See Section 4: PRA register + SFCR filings for UK insurers; SWFI/Global SWF for future non-UK SWFs. Still open is the practical scoping question below.

Still open, carried over or sharpened:

- **Currency is still not explicit anywhere.** Units are now confirmed as thousands, but no file states the currency — the allocation export literally labels its value column "Dollar Amount," which is almost certainly a generic S&P template artefact rather than literal USD, given the magnitudes only make sense as GBP for UK entities. Please confirm before any figures get loaded.
- **Allocation weighting methodology**: with real allocation data now in hand, should `aggregate_asset_allocation` weight by AUM across all tracked institutions or only the top 10/20, and should it weight by plan or by firm (a firm's plans can have different allocations)?
- **Coverage threshold — now applies to insurers too**: is there a minimum AUM cutoff for a UK institution to be worth tracking? This matters more now that insurers are in scope, since pulling an SFCR per firm is real manual/Claude-in-Chrome effort per insurer (unlike the one-shot S&P CSV) — worth deciding whether v1 targets the largest N UK insurers by assets rather than the full PRA-authorised universe (several hundred firms, most small).
- **Category → entity_type mapping** (`Corporation`, `Reserve Banks`, `Health Service Organization – for profit`) involves judgment calls — flagging for sign-off rather than deciding unilaterally, especially the `Corporation` split.

New, raised by the Manager and Provider exports (pulled 2026-07-08):

- **`custodian`/`actuary`/`investment_consultant` are scalar fields but File 4 shows some firms with multiple providers of the same type after dedup** — need to decide arrays vs. a "primary provider" rule before ingestion.
- **`OCIO` (11 rows in File 4) is the closest match to the still-unpopulated `fiduciary_manager` field** — proposing that mapping, flagging for sign-off rather than assuming.
- **File 3 has no `Pid#`**, only `Firm Name` — safe to join except for File 1's one duplicate name (Bank of England). Needs a manual fix for that one firm rather than a purely mechanical join.
- **File 4 has no `Plan Name`**, so provider relationships can't be attributed to a specific plan the way File 3's manager relationships can — an inherent granularity limit of that report, not fixable by re-pulling.

Resolved by this round:

- ~~Manager relationships (`asset_managers_used`) have no source in either S&P file~~ — resolved. `results (3).xls` (MMD's `Manager` report output) is the source, pulled via a live Claude-in-Chrome MMD session on 2026-07-08 — 42% of UK firms now have at least one named manager. See File 3, Section 3.
- ~~`custodian`/`actuary`/`investment_consultant` have no source~~ — resolved. `results (4).xls` (MMD's `Provider` report output), same session — 24% of UK firms now have at least one named provider, discriminated by `Provider Type`. See File 4, Section 3.

New, raised by the allocation export:

- ~~`Asset Type` has a long noisy tail (~215 non-core values)~~ — resolved for now: non-core values default to `Other/Unclassified` rather than being reviewed one-by-one, since several (`Loans`, `Unspecified`, `Other Investments`) turned out to be dominated by a single outlier entity (Bank of England) once that was excluded. See Section 6.
- **No explicit plan-instance ID.** Where a firm has two plans sharing the same `Type of Plan` label, they're only distinguishable by having different `Plan Assets` values — a fragile composite key. Worth flagging now rather than discovering it mid-ingestion.
- **A small number of plan-instances have allocation percentages that don't sum to ~100%** (up to 200 in the sample) — likely duplicate or overlapping rows. Needs a data-quality rule (e.g. flag and exclude from stats rather than silently including).

Resolved by this round:

- ~~Is the 19-segment taxonomy the Atlas standard?~~ — yes, confirmed. See 2b.
- ~~How should top-down overrides be sourced/tracked?~~ — resolved as a design: an extensible source registry (2c) so new sources can be added and re-ranked at any time, not a single hardcoded override field. Percentages/allocations pro-rate to the active source's total, per your instruction.
- ~~Does the scorecard belong in Atlas?~~ — yes, confirmed as a mixed quantitative/qualitative structure. See Section 5.
- ~~Is the Nordics/Germany benchmark table a one-off or a pattern?~~ — confirmed as a pattern: it should be "one of many" possible sources per segment, feeding the same 2c registry mechanism rather than being special-cased. Its exact origin (OECD Pension Markets in Focus is my best guess) still needs naming once that country is scoped — not urgent for the UK pilot.
- ~~Precedence defaults for the source registry~~ — resolved: no fixed global rule. Precedence is set case by case based on the assessed veracity of each specific source for that specific segment/country, not an automatic "government always wins." See 2c.
- ~~Which qualitative scorecard dimensions are externally sourced vs pure judgment?~~ — resolved: Regulatory complexity uses TMF Group rankings as a go-to external source when available, falling back to judgment where TMF doesn't cover a country. Local presence, Languages, and the rest are pure analyst judgment with no external tracker. See Section 5.

Resolved by the scorecard build:

- ~~What identifies an "engagement" in the `engagement_id` field?~~ — superseded. The built scorecard dropped the per-engagement design in favor of one live scorecard per (country, segment) directly on the segment record, so there's no `engagement_id` to define. See Section 5.
- ~~Should judgment dimensions pre-populate from a prior engagement's score?~~ — resolved: no. New segments/countries start with no scorecard fields at all; the matrix UI's "Scored X/12" indicator and highlighted blank cells are the intended mechanism for filling them in over time, not a pre-populated starting point. See Section 5.

Resolved by the allocation build, one new item raised:

- ~~Category → entity_type mapping involves judgment calls needing sign-off~~ — partially superseded: rather than fully resolving the Corporation/Government/Tax Exempt classification (which would need entity-by-entity sign-off), the allocation feature works around it by anchoring Equities to the already-validated figure and only using raw-file data to shape the *other* asset types. AUM itself hasn't been recomputed from raw records and doesn't depend on getting this classification exactly right. See Section 6.
- **New**: no raw per-country institutions subtable exists yet — today's figures came from a one-off derivation script run against the exported files, not a live, editable set of institution records that AUM/allocation recompute from automatically. Peter confirmed interest in this ("subtables per country that cascade up") and in automating the S&P refresh itself via a Claude-in-Chrome-driven export (no API exists at his subscription tier) rather than manual download/upload. Both remain to be built.
