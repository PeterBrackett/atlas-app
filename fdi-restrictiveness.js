// OECD FDI Regulatory Restrictiveness Index (FDIRRI) -- an external
// reference used to auto-score the "Local presence required" scorecard
// dimension. The index runs 0 (fully open to foreign investment) to 1
// (fully closed), based on statutory restrictions: foreign equity limits,
// discriminatory screening/approval, restrictions on key foreign personnel,
// and other operational restrictions.
//
// Two design notes worth flagging:
//
// 1. The sector-specific score for fund/trust management (OECD sector code
//    K643_663, the closest match to "asset management") turned out to have
//    almost no variance across Atlas's countries -- everywhere from Italy to
//    Saudi Arabia scores between 0 and 0.0015, i.e. fund management itself
//    is close to universally liberalised, so that number can't discriminate
//    between markets. What's used here instead is the ECONOMY-WIDE "Total -
//    All activities" score (OECD sector code _T), which does have real
//    spread (UK 0.0001 fully open vs Saudi Arabia 0.33 clearly restrictive).
//    This is a general-market-openness proxy, not something asset-management
//    specific -- it captures the screening/approval/foreign-personnel regime
//    a country applies broadly, which still bears on how easily a foreign
//    manager can set up a local entity, but it isn't a pure read on fund
//    management access. Treat it as a structural signal to combine with
//    judgment, not a literal "can managers operate here" answer.
//
// 2. Direction is genuinely inverted from every other auto-scorer on this
//    site (including TMF's Regulatory complexity, which LOOKS inverted
//    because rank 1 is the worst jurisdiction, but is actually a normal
//    increasing scale once you know that). FDIRRI's own scale has 0 = best
//    at the bottom, so low raw values score HIGH here. See
//    applyLocalPresenceThresholds() in overview.html for the reversed
//    comparison logic.
//
// Source: OECD FDI Regulatory Restrictiveness Index 2024 update (statutory
// measures in force as of end-December of the stated year), sector "Total -
// All activities", pulled via the OECD Data Explorer SDMX API
// (OECD.DAF.INV:DSD_FDIRRI_SCORES@DF_FDIRRI_SCORES) 2026-07-14. OECD
// republishes this annually -- when that happens, update FDIRRI_YEAR and
// re-pull each covered country's score.
const FDIRRI_YEAR = 2023;

// FDIRRI "Total - All activities" scores (0 = fully open, 1 = fully closed)
// for the Atlas countries the index actually covers. Only Atlas's covered
// countries are listed here (unlike the TMF reference file, which keeps the
// full jurisdiction table) -- FDIRRI covers 100+ economies and Atlas only
// needs 30 of them.
const FDIRRI_SCORES = {
  'GB': 0.0001,
  'US': 0.051,
  'DE': 0.0088,
  'NL': 0.0091,
  'NO': 0.0135,
  'CH': 0.1118,
  'DK': 0.0138,
  'SE': 0.0075,
  'ZA': 0.0341,
  'IE': 0.0074,
  'FI': 0.0065,
  'IS': 0.154,
  'CA': 0.1501,
  'FR': 0.0224,
  'IT': 0.0711,
  'ES': 0.0104,
  'AU': 0.1941,
  'AT': 0.0555,
  'JP': 0.0484,
  'SG': 0.0652,
  'NZ': 0.1436,
  'KR': 0.1048,
  'SA': 0.331,
  'BE': 0.0101,
  'AR': 0.1058,
  'PE': 0.0088,
  'CL': 0.0263,
  'BW': 0.1482,
  'BR': 0.0601,
  'MY': 0.3219
};

// Countries FDIRRI doesn't cover at all -- Hong Kong, UAE, Qatar and Oman
// are genuinely absent from the dataset (not a query issue; confirmed via
// direct API check 2026-07-14). Skipped, not guessed at, same treatment as
// Iceland's old TMF/EF-EPI gaps.
const FDIRRI_UNCOVERED_CODES = ['HK', 'AE', 'QA', 'OM'];

function fdirriScoreForCode(code) {
  const score = FDIRRI_SCORES[(code || '').toUpperCase()];
  return typeof score === 'number' ? score : null;
}
