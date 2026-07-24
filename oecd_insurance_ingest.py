"""
Switches Life insurance / Non-life insurance segments over to OECD's "Asset
allocation of insurance companies" dataset (OECD.DAF.CM:DSD_INS@DF_ASSET_ALLOC,
data-explorer.oecd.org), replacing the placeholder "Industry aggregate" source
each segment currently has.

Data below was pulled live via the OECD SDMX API on 2026-07-14 (see chat for
the exact API calls) -- REF_AREA x {LIFE, NLIFE}, MEASURE = INV (total) plus
the instrument breakdown, UNIT_MEASURE=USD, OWNERSHIP=UND_T (all undertakings),
INSURER_TYPE=DIR (direct insurer, excludes reinsurance), latest year with data
per country (mostly 2024, a few 2022/2023 where a country hasn't reported yet).
Values from OECD are in USD millions; divided by 1000 below to match Atlas's
$bn convention.

Mapping from OECD's instrument codes to Atlas's 6 allocation buckets:
  Equities        = SHARE (direct listed + unlisted equity)
  Fixed Income    = BOND + LOAN
  Cash & Short-Term = CASH
  Real Estate     = REST (land & buildings)
  Alternatives    = PEQUF + HEDGE + STRPRDT
  Other/Unclassified = CIS (mutual funds) + UNITLINK (assets backing
    unit-linked contracts) + INV_O (other) -- both CIS and UNITLINK are
    "wrapper" categories OECD doesn't look through to underlying holdings,
    so per Peter's instruction (2026-07-14) they fold into Other/Unclassified
    rather than being guessed at. This dominates Other/Unclassified for
    unit-linked-heavy markets (UK life is ~84% Other for exactly this reason)
    -- a real feature of the data, not a mapping bug.

Countries with no OECD coverage for this dataset (New Zealand, Saudi Arabia,
UAE, Qatar, Oman, Botswana) are left untouched -- their existing placeholder
stays active.
"""
import json
from pathlib import Path
from datetime import date

DATA_DIR = Path(__file__).parent / "atlas-site" / "data"
TODAY = "2026-07-14"
EDITOR = "peter.brackett@institutionaladviser.co.uk"

SOURCE_LABEL = "Industry aggregate (OECD Global Insurance Market Trends)"
SOURCE_NAME = ("OECD 'Asset allocation of insurance companies' dataset "
               "(OECD.DAF.CM:DSD_INS@DF_ASSET_ALLOC), data-explorer.oecd.org")
SOURCE_NOTE = ("All undertakings, direct business (excludes reinsurance). "
               "Mutual funds/CIS and assets backing unit-linked contracts have "
               "no look-through breakdown in the source and are folded into "
               "Other/Unclassified rather than estimated.")

# {ISO3: Atlas slug}
ISO3_TO_SLUG = {
    "GBR": "uk", "USA": "us", "DEU": "de", "NLD": "nl", "NOR": "no", "CHE": "ch",
    "DNK": "dk", "SWE": "se", "ZAF": "za", "IRL": "ie", "FIN": "fi", "ISL": "is",
    "CAN": "ca", "FRA": "fr", "ITA": "it", "ESP": "es", "AUS": "au", "AUT": "at",
    "JPN": "jp", "HKG": "hk", "SGP": "sg", "KOR": "kr", "ARG": "ar", "PER": "pe",
    "CHL": "cl", "BRA": "br", "MYS": "my", "BEL": "be",
}

# {ISO3: {LIFE|NLIFE: [year, total_musd, equities_musd, fixed_income_musd,
#                       cash_musd, real_estate_musd, alternatives_musd, other_musd]}}
# All figures USD millions, straight from the OECD API (see docstring).
OECD = {
    "AUS": {"LIFE": ["2024", 53422.33, 0, 28574.72, 4546.71, 0, 0, 20300.9],
            "NLIFE": ["2024", 52049.42, 0, 36987.56, 4684.69, 0, 0, 10377.17]},
    "SGP": {"LIFE": ["2022", 123373.95, 10121.83, 66577.63, 4406.93, 1598.38, 1714.23, 38954.95],
            "NLIFE": ["2022", 12189.76, 209.91, 6902.99, 4588.18, 209.81, 22.47, 256.4]},
    "ZAF": {"LIFE": ["2023", 219685.79, 28979.09, 39684.29, 12017.52, 311.56, 3633.15, 135060.18],
            "NLIFE": ["2023", 15855.25, 1046.88, 4791.68, 3186.34, 83.92, 30.68, 6715.74]},
    "JPN": {"LIFE": ["2024", 2723918.11, 242340.51, 1960831.05, 56091.86, 42015.01, 0, 422639.68],
            "NLIFE": ["2024", 184374.56, 40230.63, 44633.27, 10515.71, 5286.06, 0, 83708.89]},
    "GBR": {"LIFE": ["2024", 1579823.06, 7882.31, 235677.56, 5322.92, 2306.09, 3703.88, 1324930.3],
            "NLIFE": ["2024", 15761.69, 1577.51, 10581.8, 1024.65, 106.83, 133.9, 2337]},
    "PER": {"LIFE": ["2024", 2247.64, 67.76, 1753.02, 110.91, 146.83, 0.35, 168.77],
            "NLIFE": ["2024", 635.57, 166.95, 196.94, 70.93, 19.63, 0, 181.12]},
    "KOR": {"LIFE": ["2023", 669047.67, 34596.85, 384091.51, 14187.24, 8920.7, 0, 227251.37],
            "NLIFE": ["2023", 238596.26, 11462.28, 148778.77, 7196.8, 3873.11, 0, 67285.3]},
    "BRA": {"LIFE": ["2023", 198959.8, 35.17, 8951.49, 190.62, 25.24, 1088.21, 188669.07],
            "NLIFE": ["2023", 3576.47, 0.5, 2097.81, 59.49, 11.71, 52.35, 1354.61]},
    "NOR": {"LIFE": ["2024", 218315.8, 16866.85, 78471.95, 389.86, 13759.75, 0, 108827.39],
            "NLIFE": ["2024", 19587.38, 1582.38, 11619.88, 1057.55, 1084.42, 0, 4243.15]},
    "CHE": {"LIFE": ["2024", 306916.35, 11924.1, 176397.43, 8678.91, 42654.35, 4321.75, 62939.81],
            "NLIFE": ["2024", 177951.98, 8230.8, 73150.03, 10060.85, 8002.88, 3839.35, 74668.07]},
    "ITA": {"LIFE": ["2024", 298484.66, 3113.35, 124172.85, 2259.27, 197.06, 6901.89, 161840.24],
            "NLIFE": ["2024", 19938.31, 1818.8, 14969.72, 916.86, 450.12, 711.89, 1070.92]},
    "NLD": {"LIFE": ["2024", 360376.83, 21635.59, 186001.3, 8598.53, 3675.14, 3235.31, 137230.96],
            "NLIFE": ["2024", 65598.23, 10875.64, 32690.18, 4412.9, 122, 386.02, 17111.49]},
    "ESP": {"LIFE": ["2024", 19833.39, 708.83, 11488.77, 569.71, 256.86, 1904.96, 4904.26],
            "NLIFE": ["2024", 26981.67, 4182.55, 15071.63, 2733.1, 1768.92, 560.41, 2665.06]},
    "HKG": {"LIFE": ["2024", 395957.93, 36160.18, 222631.14, 13896.4, 0, 0, 123270.21],
            "NLIFE": ["2024", 20562.12, 607.68, 8515.12, 6885.54, 0, 0, 4553.78]},
    "USA": {"LIFE": ["2024", 5713261, 200652, 4576738, 176363, 21864, 35386, 702258],
            "NLIFE": ["2024", 2912947, 643026, 1715851, 355455, 17988, -3, 180630]},
    "ARG": {"LIFE": ["2024", 2592.28, 6.93, 1688.77, 11.52, 40.69, 0, 844.37],
            "NLIFE": ["2024", 3549.2, 171.26, 1964.64, 93.75, 150.57, 0, 1168.98]},
    "MYS": {"LIFE": ["2023", 75800.13, 6076.12, 39112.78, 2389.04, 904.17, 0, 27318.02],
            "NLIFE": ["2023", 10777.04, 140.38, 3261.38, 2309.49, 104.45, 0, 4961.34]},
    "SWE": {"LIFE": ["2023", 661178, 172573.59, 125155.95, 15345.26, 5473.63, 0, 342629.57],
            "NLIFE": ["2023", 83302.36, 23986.02, 30408.6, 3333.53, 2585.94, 0, 22988.27]},
    "IRL": {"LIFE": ["2024", 353342.19, 3345.04, 16909.93, 3384.32, 32.79, 121.6, 329548.51],
            "NLIFE": ["2024", 32599.9, 372.62, 22260.25, 5140.41, 117.07, 405.73, 4303.82]},
    "FIN": {"LIFE": ["2024", 75427.25, 1240.62, 9745.37, 2308.29, 931.18, 0, 61201.79],
            "NLIFE": ["2024", 15211.76, 2668.37, 4627.23, 862.43, 827.55, 0, 6226.18]},
    "BEL": {"LIFE": ["2024", 11844.8, 1044.79, 7751.1, 280.82, 205.55, 15.81, 2546.73],
            "NLIFE": ["2024", 20630.96, 1247.67, 13915.41, 1945.94, 94.29, 145.18, 3282.47]},
    "DNK": {"LIFE": ["2024", 243288.85, 19718.13, 57936.24, 4031.44, 902.11, 16760.18, 143940.75],
            "NLIFE": ["2024", 21386.86, 822.99, 11496.44, 934.72, 104.7, 44.98, 7983.03]},
    "ISL": {"LIFE": ["2024", 140.54, 1.13, 76.82, 7.63, 0, 0.6, 54.36],
            "NLIFE": ["2024", 1419.1, 496.98, 681.16, 39.09, 0, 33.74, 168.13]},
    "AUT": {"LIFE": ["2024", 8995.49, 101.51, 2304.99, 100.66, 14.81, 47.95, 6425.57],
            "NLIFE": ["2024", 17647.2, 11054.42, 1961.1, 1752.4, 714.43, 22.37, 2142.48]},
    "CHL": {"LIFE": ["2024", 67690.86, 439.61, 46271.6, 1219.41, 5222.46, 0, 14537.78],
            "NLIFE": ["2024", 2805.1, 27.92, 2043.17, 334.77, 87.35, 0, 311.89]},
    "CAN": {"LIFE": ["2022", 341256.34, 30221.28, 197595.38, 10949.06, 0, 0, 102490.62],
            "NLIFE": ["2022", 122987.58, 15862.68, 84535.03, 6533.12, 0, 0, 16056.75]},
    "FRA": {"LIFE": ["2024", 862944.69, 55118.46, 382133.61, 14484.26, 13732.32, 31573.5, 365902.54],
            "NLIFE": ["2024", 281083.8, 84994.25, 126017.69, 14987.34, 6122.35, 6732.76, 42229.41]},
    "DEU": {"LIFE": ["2024", 1220392.42, 124180.41, 426508.14, 11409.68, 14809.92, 31946.67, 611537.6],
            "NLIFE": ["2024", 683920.42, 117501.55, 289332.47, 14602.49, 10002.82, 20752.19, 231728.9]},
}

BUCKET_NAMES = ["Equities", "Fixed Income", "Cash & Short-Term", "Real Estate", "Alternatives", "Other/Unclassified"]

SEGMENT_TO_INS_TYPE = {"Life insurance": "LIFE", "Non-life insurance": "NLIFE"}


def build_allocation(total_bn, bucket_values_bn):
    allocation = []
    for name, value_bn in zip(BUCKET_NAMES, bucket_values_bn):
        if round(value_bn, 6) == 0:
            continue
        pct = round(value_bn / total_bn * 100, 1) if total_bn else 0
        allocation.append({
            "asset_type": name,
            "value_bn": round(value_bn, 3),
            "pct_of_segment": pct,
            "styles": [{"asset_style": "Unspecified", "value_bn": round(value_bn, 3)}]
        })
    return allocation


def process_country(iso3, slug):
    row = OECD[iso3]
    path = DATA_DIR / f"{slug}.json"
    if not path.exists():
        print(f"  SKIP {slug}: no local data file at {path}")
        return
    data = json.loads(path.read_text())
    data.setdefault("segments", [])
    changed = False

    existing_by_name = {s.get("segment"): s for s in data["segments"]}

    for ins_type, segment_name in [("LIFE", "Life insurance"), ("NLIFE", "Non-life insurance")]:
        if ins_type not in row:
            continue
        seg = existing_by_name.get(segment_name)
        is_new = seg is None
        if is_new:
            seg = {
                "segment": segment_name,
                "aum_bn": 0,
                "basis": "",
                # No "scorecard" key here -- as of 2026-07-24, scores live in
                # their own {code}_scores.json file, not inline on segments,
                # precisely so ingestion scripts like this one can never wipe
                # them out. See getScores.js for the full rationale.
                "allocation": [],
                "sources": [],
                "source_history": []
            }
            data["segments"].append(seg)
            existing_by_name[segment_name] = seg

        # Idempotency guard: if this segment's active source is already this
        # exact OECD source, skip -- prevents duplicate sources[]/source_history
        # entries if the script is ever re-run (e.g. GBR was already switched
        # over in an earlier pass).
        active_source = next((s for s in seg.get("sources", []) if s.get("active")), None)
        if active_source and active_source.get("source") == SOURCE_NAME:
            print(f"  {slug} / {segment_name}: already on OECD source, skipping")
            continue

        year, total_m, eq_m, fi_m, cash_m, re_m, alt_m, other_m = row[ins_type]
        total_bn = total_m / 1000
        buckets_bn = [v / 1000 for v in (eq_m, fi_m, cash_m, re_m, alt_m, other_m)]

        for s in seg.get("sources", []):
            s["active"] = False
        new_source = {
            "basis": SOURCE_LABEL,
            "aum_bn": round(total_bn, 3),
            "as_of": f"{year}-12",
            "source": SOURCE_NAME,
            "note": SOURCE_NOTE,
            "active": True
        }
        seg.setdefault("sources", []).append(new_source)

        seg["aum_bn"] = round(total_bn, 3)
        seg["basis"] = SOURCE_LABEL
        seg["allocation"] = build_allocation(total_bn, buckets_bn)
        seg.pop("allocation_incomplete", None)

        if is_new:
            history_detail = (f"Segment created from {SOURCE_LABEL} ({year} data). "
                               f"Scorecard not yet scored -- dimensions to be filled in manually.")
            history_event = "created"
        else:
            history_detail = (f"Switched active source to {SOURCE_LABEL} ({year} data). "
                               f"Total investments and full 6-bucket allocation now sourced "
                               f"directly from OECD (previously a placeholder aggregate with "
                               f"only a partial Equities estimate). Superseded placeholder "
                               f"source retained in sources[] for history.")
            history_event = "source_changed"

        seg.setdefault("source_history", []).append({
            "date": TODAY,
            "event": history_event,
            "by": EDITOR,
            "detail": history_detail
        })
        changed = True
        tag = "NEW" if is_new else "updated"
        print(f"  {slug} / {seg['segment']} [{tag}]: aum_bn -> {seg['aum_bn']} ({year}), "
              f"{len(seg['allocation'])} allocation buckets")

    if changed:
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n")


def main():
    print(f"Processing {len(OECD)} countries with OECD insurance coverage...")
    for iso3, slug in ISO3_TO_SLUG.items():
        process_country(iso3, slug)
    skipped = ["nz", "sa", "ae", "qa", "om", "bw"]
    print(f"\nNo OECD coverage, left untouched: {', '.join(skipped)}")


if __name__ == "__main__":
    main()
