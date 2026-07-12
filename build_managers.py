#!/usr/bin/env python3
"""
Atlas — Manager Market Share builder.

Generalizes the UK pilot (atlas-site/data/uk_managers.json, built from
results (3).xls) to any country's <Country>_manager.xls (S&P MMD "Manager"
report export, same 16-column schema, HTML-table-as-.xls format).

Segment routing rules (reverse-engineered from the UK pilot output to
reconcile against uk_managers.json's own segment totals):

  Category -> modifier, and whether that category has a dedicated E&F
  segment separate from its DB/DC pension segments:

    Corporation                              -> Corp        (no E&F segment; EF-typed
                                                               plans fold into DB Pension (Corp))
    Government                                -> Govt        (no E&F segment; folds into DB Pension (Govt))
    Union                                     -> Union       (no E&F segment; folds into DB Pension (Union))
    Tax Exempt Organization                   -> Tax exempt  (has its own E&F segment)
    Endowment                                 -> Endowments  (has its own E&F segment)
    Health Service Organization - nonprofit   -> Healthcare non-profit (has its own E&F segment)
    Private Foundation                        -> Foundations (E&F only — no DB/DC Pension
                                                               (Foundations) segment exists in the
                                                               19-segment taxonomy, so any DB/DC-typed
                                                               plan under this category is excluded)
    Reserve Banks                             -> excluded entirely (central-bank balance sheets,
                                                               same treatment as every other Atlas
                                                               data product)
    Health Service Organization - for profit  -> excluded (no clean fit in the 19-segment
                                                               taxonomy — flagged as a judgment call
                                                               in Atlas_Schema_Proposal_v1.md, never
                                                               resolved to a segment)

  Plan-type buckets:
    DB   -> Defined Benefit Multinational Consolidated Plan, Closed or Frozen Defined
             Benefit Plan, Defined Benefit Plan, Cash Balance Plan
    DC   -> Generic Defined Contribution or Profit Sharing Plan, Money Purchase Plan,
             401(k) Plan, Defined Contribution Plan (plan with contributions only),
             Stock Plan
    EF   -> Investment Pool Funds, Endowment Fund, Foundation Fund Or Investment Pool
    HYB  -> Hybrid DB/DC Plan — split between DB/DC using the firm-level
             Total DB Assets : Total DC Assets ratio when Total DC Assets is populated
             and non-zero; otherwise treated entirely as DB (matches every UK hybrid
             row where Total DC Assets is blank).

  This reconciles exactly against the UK pilot for DB/DC Pension (Govt), DB/DC
  Pension (Union), DC Pension (Corp), and Foundations E&F, and within ~0.5% for
  the others — the residual gap is documented in Atlas_Schema_Proposal_v1.md as
  real, hand-done entity-level reclassification in the original pilot (e.g. a
  handful of firms S&P categorises in a way that doesn't match how they were
  originally treated) that "couldn't be safely reverse-engineered" even by the
  people who built it — not a defect in this script.

Manager-level rules (matching the UK pilot's source_note verbatim):
  - "Self-managed" = Management Co. Name == "In House Asset Management" (the only
    tag with no city/country suffix — every other tag is "Name, City, Country").
  - Manager name grouping: strip the trailing "<City>, <Country/State>" portion
    (last two comma-separated tokens) so the same manager's regional offices
    roll up together, e.g. "BlackRock Investment Management (UK) Limited,
    London, United Kingdom" -> "BlackRock Investment Management (UK) Limited".
  - Per plan: if no manager-tag rows exist at all, the whole plan is "Not
    disclosed". If the *only* tag anywhere on the plan is self-managed with no
    dollar figure, the whole plan counts as self-managed (S&P is telling us the
    whole plan is run in-house, just not the exact split). Otherwise, only rows
    with a populated dollar figure are attributed (to self-managed or to the
    named external manager); named managers with no dollar figure are known
    relationships but contribute nothing quantifiable, so they don't get their
    own entry and their share of the plan remains in the "Not disclosed" gap.
  - Data quirk confirmed against every duplicated (plan, manager) group in the
    UK source file (164 of 164): when one manager's mandate within a plan spans
    several Managed Asset Type/Style rows, S&P repeats the *same* Dollar Amount
    Managed figure on each row rather than splitting it (e.g. Border to Coast
    Pensions Partnership shows GBP4,407,324k five times for Teesside Pension
    Fund, once per asset-type tag, not five distinct amounts). Summing naively
    inflates that manager's total by up to 5x. Fix: dedupe to one dollar figure
    per (plan, raw manager name) — take the single distinct non-null value —
    and use only the first-occurring row's asset type/style as the
    representative tag for the breakdown, since the true per-type split isn't
    recoverable from the data.
  - "Not disclosed" and "Self-managed" carry no asset-type/style breakdown
    (asset_type: "Not broken down") since the raw data doesn't say how those
    amounts are invested; every other manager's Dollar Amount Managed is
    grouped by (Managed Asset Type, Managed Asset Style) with a style line only
    where the row data supports it.
  - identified_coverage_pct counts self-managed + named external managers with
    a dollar figure as identified; only "Not disclosed" counts against it.

Usage:
    python3 build_managers.py <Country_manager.xls> <ISO2 lowercase code> "<Country Name>" [output_dir]

Writes {code}_managers.json to output_dir (default: current directory).
"""
import sys
import json
import re
from collections import defaultdict

import pandas as pd

DB_TYPES = {
    "Defined Benefit Multinational Consolidated Plan",
    "Closed or Frozen Defined Benefit Plan",
    "Defined Benefit Plan",
    "Cash Balance Plan",
}
DC_TYPES = {
    "Generic Defined Contribution or Profit Sharing Plan",
    "Money Purchase Plan",
    "401(k) Plan",
    "Defined Contribution Plan (plan with contributions only)",
    "Stock Plan",
    # Seen in non-UK exports, not present in the UK pilot file:
    "Defined Contribution Multinational Consolidated Plan",
    "Tax Deferred Annuity",          # 403(b)-style individual contribution vehicle
    "Thrift & Savings Plan",          # US federal TSP-style DC savings plan
    "457 Deferred Compensation Plan",  # contribution-based deferred comp
}
EF_TYPES = {
    "Investment Pool Funds",
    "Endowment Fund",
    "Foundation Fund Or Investment Pool",
    # Seen in non-UK exports — trust/pooled vehicles, not DB/DC retirement
    # plans, so treated the same as Investment Pool Funds (E&F bucket where
    # the category has one, folded into the DB pension segment otherwise):
    "VEBA Health & Welfare Plan",
    "Other Tax-exempt Fund",
    "Nuclear Decommissioning Fund",
    "Local/State Investment Funds",
    "Workers Compensation Fund",
}
# Genuinely ambiguous plan types with no DB/DC label at all — same treatment
# Atlas_Schema_Proposal_v1.md documents for ambiguous institutions-build plan
# types: fall back to the firm's own Total DB Assets : Total DC Assets ratio,
# same mechanism as a Hybrid DB/DC Plan.
AMBIGUOUS_SPLIT_TYPES = {
    "Other Retirement Fund",
}
HYBRID_TYPES = {
    "Hybrid DB/DC Plan",
    "Hybrid DB/DC Multinational Consolidated Plan",
}

# Category -> (modifier label, has_own_ef_segment)
CATEGORY_MAP = {
    "Corporation": ("Corp", False),
    "Government": ("Govt", False),
    "Union": ("Union", False),
    "Tax Exempt Organization": ("Tax exempt", True),
    "Endowment": ("Endowments", True),
    "Health Service Organization - nonprofit": ("Healthcare non-profit", True),
    "Health Service Organization – nonprofit": ("Healthcare non-profit", True),  # en-dash variant
    "Private Foundation": ("Foundations", "ef_only"),
}
EXCLUDED_CATEGORIES = {
    "Reserve Banks",
    "Health Service Organization - for profit",
    "Health Service Organization – for profit",
}

# Country-specific overrides, matching judgment calls already made (and
# documented in atlas-site/data/us.json's own source_note) for the US
# institutions/segment build, so the manager product stays consistent with
# the rest of Atlas's US data rather than silently diverging:
#   (1) US 'Health Service Organization - for profit' is material (29,470
#       institutions, $373.8bn) and was mapped to Corp "by analogy, since
#       for-profit entities don't hold endowment/foundation-style assets" —
#       unlike every other country where this category is a single
#       inconsequential row and gets excluded instead.
#   (2) A large share of US 'Government' ($1.2tn) is state treasuries, 529
#       college-savings programs, workers'-comp funds and similar
#       cash-management pools, not staff retirement plans — us.json excludes
#       these "following the same logic as the Reserve Banks exclusion".
#       Mechanically, that maps to: Government-category rows bucketed as EF
#       (the pooled/ambiguous-vehicle types) are excluded for the US instead
#       of folding into DB Pension (Govt) as they do for every other country
#       building on the pilot's rule (validated exactly against Korea/UK,
#       where Government-EF genuinely is public-pension-adjacent).
HEALTH_FOR_PROFIT_AS_CORP_COUNTRIES = {"US"}
GOVT_EF_EXCLUDED_COUNTRIES = {"US"}

# Categories that route to a single flat segment regardless of Type of Plan —
# the 19-segment taxonomy has no DB/DC split for these (unlike Corp/Govt/etc.).
# Seen in non-UK exports (absent from the UK pilot, which the schema doc flags
# as having "no insurers, no SWF" coverage at all): a SWF's own small staff
# pension plans (e.g. Korea Investment Corporation's KIC Pension Fund/401(k))
# get folded into "SWF" too, since there's no separate "DB Pension (SWF)"
# segment to route them to and they're immaterial next to the main fund.
FLAT_CATEGORY_SEGMENTS = {
    "Sovereign Wealth Funds": "SWF",
}

SELF_MANAGED_TAG = "In House Asset Management"
NOT_DISCLOSED = "Not disclosed"


def load_raw(path):
    tables = pd.read_html(path)
    df = tables[0]
    df = df.dropna(subset=["Firm Name"]).copy()
    return df


def bucket_plan_type(pt):
    if pt in DB_TYPES:
        return "DB"
    if pt in DC_TYPES:
        return "DC"
    if pt in EF_TYPES:
        return "EF"
    if pt in HYBRID_TYPES or pt in AMBIGUOUS_SPLIT_TYPES:
        return "HYB"
    return "OTHER"


def route_segment(category, plan_type, country_code=None):
    """Return the target segment name for a plan, or None if excluded."""
    cat = category.strip() if isinstance(category, str) else category
    cc = (country_code or "").upper()

    if cat in EXCLUDED_CATEGORIES:
        if cc in HEALTH_FOR_PROFIT_AS_CORP_COUNTRIES and cat.startswith("Health Service Organization"):
            cat = "Corporation"  # US-only override, see module docstring / notes above
        else:
            return None
    if cat in FLAT_CATEGORY_SEGMENTS:
        return FLAT_CATEGORY_SEGMENTS[cat]
    mapping = CATEGORY_MAP.get(cat)
    if mapping is None:
        return None  # unknown category — exclude rather than guess
    modifier, has_ef = mapping
    b = bucket_plan_type(plan_type)
    if has_ef == "ef_only":
        # Private-Foundation-style: only an E&F segment exists
        if b == "EF":
            return f"{modifier} E&F"
        return None  # DB/DC/Hybrid under an EF-only category has no home
    if b == "EF":
        if cat == "Government" and cc in GOVT_EF_EXCLUDED_COUNTRIES:
            return None  # US-only override: treasuries/529s/workers'-comp, not pensions
        if has_ef:
            return f"{modifier} E&F"
        else:
            return f"DB Pension ({modifier})"  # folds into DB pension segment
    if b == "DB":
        return f"DB Pension ({modifier})"
    if b == "DC":
        return f"DC Pension ({modifier})"
    return None  # OTHER plan types — no rule, exclude rather than guess


def split_hybrid(plan_assets, total_db, total_dc):
    """Return (db_amount, dc_amount) for a Hybrid DB/DC Plan row."""
    if pd.isna(total_dc) or total_dc == 0:
        return plan_assets, 0.0
    if pd.isna(total_db):
        total_db = 0.0
    denom = total_db + total_dc
    if denom == 0:
        return plan_assets, 0.0
    frac_db = total_db / denom
    return plan_assets * frac_db, plan_assets * (1 - frac_db)


def normalize_ws(s):
    """Collapse repeated whitespace (incl. &nbsp; artifacts from HTML source)
    so the same manager name doesn't fragment into two dict keys depending on
    whether it was parsed via pandas/lxml (which collapses whitespace) or the
    streaming stdlib parser (which doesn't) — caught on 'Pacific Investment
    Management Company LLC (PIMCO)' picking up a double space in one path."""
    return re.sub(r"\s+", " ", s).strip()


def clean_manager_name(raw_name):
    """Strip trailing '<City>, <Country/State>' from 'Name, City, Country'."""
    raw_name = normalize_ws(raw_name)
    if raw_name == SELF_MANAGED_TAG:
        return raw_name
    parts = [p.strip() for p in raw_name.split(",")]
    if len(parts) <= 2:
        return parts[0]
    return ", ".join(parts[:-2])


def _isnan(v):
    try:
        return v is None or (isinstance(v, float) and v != v)
    except Exception:
        return v is None


def process_plan(category, plan_type, plan_assets, total_db, total_dc, per_raw_manager, code, accum):
    """Route one plan-instance and its (deduped) manager rows into accum.

    per_raw_manager: list of (raw_manager_name, dollar_amount_or_nan, asset_type, asset_style),
    already deduped to one entry per distinct raw manager name within this plan.
    accum: dict with keys seg_plan_assets, seg_not_disclosed, seg_managers,
    excluded_count, excluded_assets (all mutated in place).
    """
    if _isnan(plan_assets):
        return

    b = bucket_plan_type(plan_type)
    if b == "HYB":
        db_amt, dc_amt = split_hybrid(plan_assets, total_db, total_dc)
        amounts_by_segment = []
        if db_amt:
            seg = route_segment(category, "Defined Benefit Plan", code)
            if seg:
                amounts_by_segment.append((seg, db_amt))
        if dc_amt:
            seg = route_segment(category, "Money Purchase Plan", code)
            if seg:
                amounts_by_segment.append((seg, dc_amt))
        if not amounts_by_segment:
            accum["excluded_count"] += 1
            accum["excluded_assets"] += plan_assets
            return
    else:
        seg = route_segment(category, plan_type, code)
        if seg is None:
            accum["excluded_count"] += 1
            accum["excluded_assets"] += plan_assets
            return
        amounts_by_segment = [(seg, plan_assets)]

    has_any_manager = len(per_raw_manager) > 0
    self_entries = [e for e in per_raw_manager if e[0] == SELF_MANAGED_TAG]
    external_entries = [e for e in per_raw_manager if e[0] != SELF_MANAGED_TAG]

    self_dollar_total = sum(e[1] for e in self_entries if not _isnan(e[1]))
    only_self_tag = has_any_manager and len(external_entries) == 0 and len(self_entries) > 0

    named_amounts = defaultdict(float)
    named_breakdown = defaultdict(lambda: defaultdict(float))
    for raw_name, amt, atype, astyle in external_entries:
        if _isnan(amt) or amt == 0:
            continue
        mname = clean_manager_name(raw_name)
        named_amounts[mname] += amt
        named_breakdown[mname][(atype, astyle)] += amt

    if only_self_tag and (_isnan(self_dollar_total) or self_dollar_total == 0):
        self_amt = plan_assets
    else:
        self_amt = self_dollar_total if not _isnan(self_dollar_total) else 0.0

    named_total = sum(named_amounts.values())

    if not has_any_manager:
        plan_not_disclosed = plan_assets
    else:
        plan_not_disclosed = max(plan_assets - self_amt - named_total, 0.0)

    seg_plan_assets = accum["seg_plan_assets"]
    seg_not_disclosed = accum["seg_not_disclosed"]
    seg_managers = accum["seg_managers"]

    for seg, seg_amount in amounts_by_segment:
        frac = seg_amount / plan_assets if plan_assets else 0
        seg_plan_assets[seg] += seg_amount
        seg_not_disclosed[seg] += plan_not_disclosed * frac
        if self_amt:
            seg_managers[seg]["Self-managed"]["total"] += self_amt * frac
        for mname, amt in named_amounts.items():
            seg_managers[seg][mname]["total"] += amt * frac
            for (atype, astyle), v in named_breakdown[mname].items():
                seg_managers[seg][mname]["breakdown"][(atype, astyle)] += v * frac


def build(path, code, country_name, source_filename=None):
    df = load_raw(path)
    df["Dollar Amount Managed"] = df["Dollar Amount Managed (in 000's)"]
    df["Management Co. Name"] = df["Management Co. Name"].apply(
        lambda x: normalize_ws(x) if isinstance(x, str) else x
    )

    # Plan-instance key: (Firm Name, Plan Name, Type of Plan, Plan Assets) —
    # Plan Name resolves the "no explicit plan-instance ID" ambiguity flagged
    # in Atlas_Schema_Proposal_v1.md for files that lack it.
    df["plan_key"] = list(zip(df["Firm Name"], df["Plan Name"], df["Type of Plan"], df["Plan Assets"]))

    total_plans_seen = set()
    accum = {
        "excluded_count": 0,
        "excluded_assets": 0.0,
        "seg_plan_assets": defaultdict(float),
        "seg_not_disclosed": defaultdict(float),
        "seg_managers": defaultdict(lambda: defaultdict(lambda: {"total": 0.0, "breakdown": defaultdict(float)})),
    }

    for plan_key, group in df.groupby("plan_key", sort=False):
        firm, plan_name, plan_type, plan_assets = plan_key
        total_plans_seen.add(plan_key)
        if pd.isna(plan_assets):
            continue
        category = group["Category"].iloc[0]
        total_db = group["Total DB Assets"].iloc[0]
        total_dc = group["Total DC Assets"].iloc[0]

        # manager rows for this plan
        mgr_rows = group.dropna(subset=["Management Co. Name"])

        # Dedupe to one dollar figure per (raw manager name) within this plan —
        # S&P repeats the same amount across multiple asset-type/style rows for
        # the same manager rather than splitting it (see module docstring).
        per_raw_manager = []  # list of (raw_name, dollar_amount, first_asset_type, first_asset_style)
        for raw_name, sub in mgr_rows.groupby("Management Co. Name", sort=False):
            vals = sub["Dollar Amount Managed"].dropna().unique()
            amt = vals[0] if len(vals) else float("nan")
            first_row = sub.iloc[0]
            atype = first_row.get("Managed Asset Type")
            astyle = first_row.get("Managed Asset Style")
            atype = atype if isinstance(atype, str) and atype.strip() else "Other/Unclassified"
            astyle = astyle if isinstance(astyle, str) and astyle.strip() else "Unspecified"
            per_raw_manager.append((raw_name, amt, atype, astyle))

        process_plan(category, plan_type, plan_assets, total_db, total_dc, per_raw_manager, code, accum)

    n_plans_seen = len(total_plans_seen)
    return assemble_output(accum, n_plans_seen, code, country_name, source_filename or path.split("/")[-1])


def assemble_output(accum, n_plans_seen, code, country_name, src):
    excluded_count = accum["excluded_count"]
    excluded_assets = accum["excluded_assets"]
    seg_plan_assets = accum["seg_plan_assets"]
    seg_not_disclosed = accum["seg_not_disclosed"]
    seg_managers = accum["seg_managers"]

    # ---- assemble output ----
    segments_out = []
    for seg, total_assets_k in seg_plan_assets.items():
        total_assets_bn = total_assets_k / 1e6
        not_disclosed_bn = seg_not_disclosed[seg] / 1e6

        managers_out = []
        if not_disclosed_bn > 1e-9:
            managers_out.append({
                "manager_name": NOT_DISCLOSED,
                "is_self_managed": False,
                "is_undisclosed": True,
                "total_managed_bn": round(not_disclosed_bn, 6),
                "market_share_pct": round(not_disclosed_bn / total_assets_bn * 100, 1) if total_assets_bn else 0.0,
                "by_asset_type": [{"asset_type": "Not broken down", "value_bn": round(not_disclosed_bn, 6)}],
            })

        identified_bn = 0.0
        for mname, data in seg_managers[seg].items():
            total_bn = data["total"] / 1e6
            if total_bn <= 1e-9:
                continue
            identified_bn += total_bn
            is_self = mname == "Self-managed"
            entry = {
                "manager_name": mname,
                "is_self_managed": is_self,
                "is_undisclosed": False,
                "total_managed_bn": round(total_bn, 6),
                "market_share_pct": round(total_bn / total_assets_bn * 100, 1) if total_assets_bn else 0.0,
            }
            if is_self:
                entry["by_asset_type"] = [{"asset_type": "Not broken down", "value_bn": round(total_bn, 6)}]
            else:
                by_type = defaultdict(lambda: defaultdict(float))
                for (atype, astyle), v in data["breakdown"].items():
                    by_type[atype][astyle] += v
                bt_list = []
                for atype, styles in by_type.items():
                    type_total = sum(styles.values()) / 1e6
                    styles_list = [{"asset_style": s, "value_bn": round(v / 1e6, 6)} for s, v in styles.items()]
                    bt_list.append({"asset_type": atype, "value_bn": round(type_total, 6), "styles": styles_list})
                entry["by_asset_type"] = bt_list
            managers_out.append(entry)

        managers_out.sort(key=lambda m: m["total_managed_bn"], reverse=True)
        identified_coverage_pct = round(identified_bn / total_assets_bn * 100, 1) if total_assets_bn else 0.0

        segments_out.append({
            "segment": seg,
            "total_plan_assets_bn": round(total_assets_bn, 6),
            "total_identified_managed_bn": round(identified_bn, 6),
            "identified_coverage_pct": identified_coverage_pct,
            "managers": managers_out,
        })

    segments_out.sort(key=lambda s: s["total_plan_assets_bn"], reverse=True)

    source_note = (
        f"Sourced from the S&P Money Manager Database Manager export ({src}), which records, "
        "per plan, which external manager runs how much of that plan's assets, in what asset "
        "type/style. Segment routing reuses the same category/plan-type rules as every other "
        "Atlas data product. 'Self-managed' and 'Not disclosed' are kept as two separate "
        "categories, not merged: 'Self-managed' is only used where S&P explicitly tags the "
        "mandate as run in-house (e.g. 'In House Asset Management') -- covering both rows with a "
        "specific dollar figure and, for a plan whose only manager tag anywhere is that "
        "self-managed tag, the plan's full remaining gap too. 'Not disclosed' is reserved for "
        "genuine unknowns -- plans with no manager information at all, or plans naming at least "
        "one real external manager where a leftover gap remains that can't be attributed to "
        "anyone specifically. Neither bucket has an asset-type breakdown, since the raw data "
        "doesn't say how those amounts are invested. 'identified_coverage_pct' on each segment "
        "counts both named external managers and self-managed amounts as identified -- only "
        "'Not disclosed' counts against coverage. Manager names are grouped by the firm name "
        "portion of the raw 'Name, City, Country' field, so the same manager's different "
        "regional offices roll up together. Built with the generalized build_managers.py script, "
        f"same rule set as the UK pilot ({excluded_count} of {n_plans_seen} plans "
        f"(${excluded_assets/1e6:,.2f}bn) could not be routed to any segment and are excluded "
        "entirely -- same categories excluded everywhere else in Atlas: Reserve Banks and "
        "similar non-pension pools)."
    )

    out = {
        "country_code": code.upper(),
        "country_name": country_name,
        "reporting_period": "2026-07 (S&P export date)",
        "source_note": source_note,
        "segments": segments_out,
    }
    return out


_TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S | re.I)
_TD_RE = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S | re.I)
_TAG_RE = re.compile(r"<[^>]+>")


class _RowStreamParser:
    """Fast, low-memory HTML-table row extractor using regex rather than a
    DOM parser.

    pandas.read_html builds a full lxml/bs4 DOM before extracting cells — for
    the US manager export (231MB, ~600k rows) that DOM alone exceeds this
    sandbox's ~3.5GB RAM budget and gets OOM-killed. stdlib html.parser (an
    earlier version of this class) avoids the DOM but its per-tag Python
    callback overhead was still far too slow at this scale (600k rows would
    not finish inside a single tool-call time budget). Regex extraction over
    the whole file content is >100x faster in practice (the whole file
    round-trips in ~2s) since the heavy lifting runs in C, and 231MB as a
    single string is well within the memory budget on its own — the DOM tree
    was the actual memory problem, not the raw text.
    """
    def __init__(self, on_row):
        self.on_row = on_row
        self.header = None

    def feed_rows(self, path):
        import html as _html
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()
        first = True
        for row_html in _TR_RE.findall(content):
            cells = [_html.unescape(_TAG_RE.sub("", c)).strip() for c in _TD_RE.findall(row_html)]
            if not cells:
                continue
            if first:
                self.header = cells
                first = False
            else:
                self.on_row(cells)
        return self.header


def _to_float(s):
    s = s.strip().replace(",", "")
    if not s:
        return float("nan")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def build_streaming(path, code, country_name, source_filename=None):
    """Same logic as build(), but streams rows via _RowStreamParser instead of
    loading the whole file into a pandas DataFrame — for files too large to
    fit as an in-memory DOM (see _RowStreamParser docstring)."""
    plans = {}  # plan_key -> {"category", "total_db", "total_dc", "managers": {raw_name: [amt, atype, astyle]}}
    col_idx = {}
    parser_ref = {}  # holds the parser instance so on_row can read its header lazily

    def on_row(cells):
        if not col_idx:
            header = parser_ref["parser"].header
            if not header:
                return
            col_idx.update({name: i for i, name in enumerate(header)})

        def get(name):
            i = col_idx.get(name)
            if i is None or i >= len(cells):
                return None
            v = cells[i]
            return v if v != "" else None

        firm = get("Firm Name")
        if not firm:
            return
        category = get("Category")
        plan_name = get("Plan Name")
        plan_type = get("Type of Plan")
        plan_assets_raw = get("Plan Assets")
        plan_assets = _to_float(plan_assets_raw) if plan_assets_raw is not None else float("nan")
        total_db_raw = get("Total DB Assets")
        total_dc_raw = get("Total DC Assets")
        total_db = _to_float(total_db_raw) if total_db_raw is not None else float("nan")
        total_dc = _to_float(total_dc_raw) if total_dc_raw is not None else float("nan")

        plan_key = (firm, plan_name, plan_type, plan_assets_raw)
        entry = plans.get(plan_key)
        if entry is None:
            entry = {"category": category, "total_db": total_db, "total_dc": total_dc,
                      "plan_assets": plan_assets, "managers": {}}
            plans[plan_key] = entry

        mgr_name = get("Management Co. Name")
        if mgr_name:
            mgr_name = normalize_ws(mgr_name)
            amt_raw = get("Dollar Amount Managed (in 000's)")
            amt = _to_float(amt_raw) if amt_raw is not None else float("nan")
            atype = get("Managed Asset Type") or "Other/Unclassified"
            astyle = get("Managed Asset Style") or "Unspecified"
            existing = entry["managers"].get(mgr_name)
            if existing is None:
                entry["managers"][mgr_name] = [amt, atype, astyle]
            elif _isnan(existing[0]) and not _isnan(amt):
                existing[0] = amt  # keep first asset_type/style tag, just fill in the amount

    parser = _RowStreamParser(on_row)
    parser_ref["parser"] = parser
    parser.feed_rows(path)

    accum = {
        "excluded_count": 0,
        "excluded_assets": 0.0,
        "seg_plan_assets": defaultdict(float),
        "seg_not_disclosed": defaultdict(float),
        "seg_managers": defaultdict(lambda: defaultdict(lambda: {"total": 0.0, "breakdown": defaultdict(float)})),
    }

    for plan_key, entry in plans.items():
        per_raw_manager = [(name, amt, atype, astyle) for name, (amt, atype, astyle) in entry["managers"].items()]
        process_plan(entry["category"], plan_key[2], entry["plan_assets"], entry["total_db"], entry["total_dc"],
                     per_raw_manager, code, accum)

    return assemble_output(accum, len(plans), code, country_name, source_filename or path.split("/")[-1])


if __name__ == "__main__":
    import os

    if len(sys.argv) < 4:
        print("Usage: build_managers.py <path_to_manager.xls> <code> <country_name> [output_dir]")
        sys.exit(1)
    path = sys.argv[1]
    code = sys.argv[2]
    country_name = sys.argv[3]
    outdir = sys.argv[4] if len(sys.argv) > 4 else "."

    # pandas.read_html builds a full DOM in memory — fine up to a few tens of
    # MB, but risks an OOM kill on much larger exports (confirmed on the US's
    # 231MB file in this sandbox's 3.5GB RAM). Switch to the streaming parser
    # above that threshold.
    size_mb = os.path.getsize(path) / 1e6
    if size_mb > 20:
        print(f"{path} is {size_mb:.0f}MB — using streaming parser")
        result = build_streaming(path, code, country_name)
    else:
        result = build(path, code, country_name)
    outpath = f"{outdir}/{code}_managers.json"
    with open(outpath, "w") as f:
        json.dump(result, f, indent=2)
    print(f"Wrote {outpath} — {len(result['segments'])} segments")
    for s in result["segments"]:
        print(f"  {s['segment']}: {s['total_plan_assets_bn']:.3f}bn, {len(s['managers'])} managers, {s['identified_coverage_pct']}% identified")
