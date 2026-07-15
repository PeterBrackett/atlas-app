// Individual investment-consultant headcount by country, sourced from the
// IC Research Institute's global database ("consultants by country.xlsx",
// pulled by Peter via his IC Research Global Advisory Board access -- this
// is NOT a published/public index like TMF, EF EPI, FDIRRI or Mercer
// elsewhere on this site, so there's no re-pull URL or API to note; it's a
// one-off manual export). Drives the "Consultant reliant" scorecard
// dimension.
//
// Two design notes worth flagging:
//
// 1. Raw headcount is not itself the metric. It's dominated by where the
//    major global consulting firms (Mercer, Aon, WTW, Cambridge Associates,
//    etc.) are headquartered -- the US (10,831) and UK (4,462) swamp every
//    other country by an order of magnitude, which mostly reflects "this is
//    where the big firms have their offices," not "local asset owners here
//    are unusually reliant on consultants." consultantDensityForCode()
//    below normalizes headcount against a country's total AUM (summed
//    across all its segments, passed in by the caller) to get consultants
//    per $1tn AUM instead -- a much more usable cross-country comparison,
//    though the UK in particular likely still carries some residual
//    "global HQ" effect rather than pure local intermediation intensity.
//    Treat it as a structural signal, not a precise reliance measurement.
//
// 2. Not every Atlas country is covered. Only the 28 countries the IC
//    Research export actually listed are in CONSULTANT_HEADCOUNT --
//    Argentina, Chile, Iceland, Oman, Peru and Qatar have no entry and are
//    genuinely absent from the source file (not a zero), so they're
//    skipped, not guessed at, same treatment as the other sources' gaps.
const CONSULTANT_HEADCOUNT_DATE = '2026-07';

// Raw individual-consultant headcount per country (ISO code), as listed in
// "consultants by country.xlsx". Non-Atlas countries in that file (e.g.
// Bahrain, Cyprus, India, Luxembourg, Mexico...) are omitted here -- only
// Atlas's 34 tracked countries are relevant.
const CONSULTANT_HEADCOUNT = {
  AE: 38,
  AT: 17,
  AU: 466,
  BE: 24,
  BR: 7,
  BW: 6,
  CA: 829,
  CH: 375,
  DE: 422,
  DK: 68,
  ES: 54,
  FI: 11,
  FR: 104,
  GB: 4462,
  HK: 122,
  IE: 338,
  IT: 42,
  JP: 102,
  KR: 10,
  MY: 5,
  NL: 493,
  NO: 67,
  NZ: 34,
  SA: 7,
  SE: 55,
  SG: 105,
  US: 10831,
  ZA: 219
};

// Countries the IC Research export has no entry for at all.
const CONSULTANT_HEADCOUNT_UNCOVERED_CODES = ['AR', 'CL', 'IS', 'OM', 'PE', 'QA'];

// Consultants per $1tn ($1,000bn) of a country's total AUM, given that
// total (summed across every segment Atlas tracks for the country -- this
// is a country-level metric applied identically to every segment, same as
// TMF/FDIRRI/Mercer). Returns null if there's no headcount entry for this
// code, or totalAumBn isn't a usable positive number.
function consultantDensityForCode(code, totalAumBn) {
  if (!code) return null;
  const count = CONSULTANT_HEADCOUNT[code.toUpperCase()];
  if (typeof count !== 'number') return null;
  if (typeof totalAumBn !== 'number' || totalAumBn <= 0) return null;
  return (count / totalAumBn) * 1000;
}
