// EF English Proficiency Index (EF EPI) -- an external reference used to
// auto-score the "Languages required" scorecard dimension (question:
// "Is English an acceptable language for marketing and client service?").
// Higher EF EPI score = higher English proficiency among the local
// population = more likely English alone is acceptable = better score.
// No direction inversion needed here (unlike TMF rank for Regulatory
// complexity), and this is a COUNTRY-LEVEL dimension, same as Regulatory
// complexity -- one value applies to every segment in a country.
//
// Source: EF EPI 2025 edition (ef.com/wwen/epi/, based on 2.2m test takers
// across 123 countries/regions), scores read directly off each country's
// EF EPI profile page 2026-07-10. EF publishes a new edition annually --
// when that happens, update EF_EPI_YEAR and re-check each covered country's
// score.
//
// IMPORTANT: EF EPI does not rank countries where English is already the
// native/primary language of business -- the US, UK, Ireland, Canada,
// Australia and New Zealand simply aren't in EF's country list, because
// the index measures proficiency among *non-native* speakers. Singapore is
// also absent from EF's list (likely too high a proportion of near-native
// speakers to be a meaningful test pool), even though English is one of its
// official languages and the dominant language of government, courts and
// business there. Rather than leaving these countries "uncovered" like
// Iceland (a genuine data gap), they get an automatic top score with a
// clear label explaining why -- English is either the native language or
// the de facto business language, so no local-language capability is
// required by definition.
const EF_EPI_YEAR = 2025;

// EF EPI 2025 scores (0-800 scale) for the Atlas countries EF actually
// ranks. Only Atlas's covered countries are listed here (unlike the TMF
// reference file, which keeps the full jurisdiction table) -- EF EPI covers
// 123 countries and Atlas only needs 14 of them.
const EF_EPI_SCORES = {
  'Netherlands': 624,
  'Austria': 616,
  'Germany': 615,
  'Norway': 613,
  'Denmark': 611,
  'Sweden': 609,
  'Finland': 603,
  'South Africa': 602,
  'Switzerland': 564,
  'Spain': 540,
  'France': 539,
  'Hong Kong': 538,
  'Italy': 513,
  'Japan': 446
};

// Atlas's global.json "code" field -> EF EPI country name, for the 14
// countries EF EPI actually scores.
const EF_CODE_TO_NAME = {
  'DE': 'Germany',
  'NL': 'Netherlands',
  'NO': 'Norway',
  'CH': 'Switzerland',
  'DK': 'Denmark',
  'SE': 'Sweden',
  'ZA': 'South Africa',
  'FI': 'Finland',
  'FR': 'France',
  'IT': 'Italy',
  'ES': 'Spain',
  'AT': 'Austria',
  'JP': 'Japan',
  'HK': 'Hong Kong'
};

// Countries EF EPI doesn't rank because English is the native/primary
// language -- automatic top score, not a data gap.
const NATIVE_ENGLISH_CODES = ['GB', 'US', 'IE', 'CA', 'AU', 'NZ'];

// Countries EF EPI doesn't rank, but where English is the de facto language
// of government, courts and business despite not being the majority native
// tongue. Treated the same as native-English for this dimension. This is a
// judgment call, not an EF EPI finding -- flagged separately so Peter can
// override it per-segment if he disagrees.
const ENGLISH_AS_BUSINESS_LANGUAGE_CODES = ['SG'];

function epiScoreForCode(code) {
  const name = EF_CODE_TO_NAME[(code || '').toUpperCase()];
  if (!name) return null;
  const score = EF_EPI_SCORES[name];
  return typeof score === 'number' ? score : null;
}

// Returns one of: 'native', 'business-english', 'epi-scored', 'uncovered'.
// 'uncovered' is a genuine gap (currently only Iceland) -- same treatment
// as TMF's Iceland gap for Regulatory complexity: skipped, not guessed at.
function languageCoverageStatus(code) {
  const up = (code || '').toUpperCase();
  if (NATIVE_ENGLISH_CODES.includes(up)) return 'native';
  if (ENGLISH_AS_BUSINESS_LANGUAGE_CODES.includes(up)) return 'business-english';
  if (epiScoreForCode(up) !== null) return 'epi-scored';
  return 'uncovered';
}
