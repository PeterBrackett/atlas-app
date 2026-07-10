// TMF Group Global Business Complexity Index (GBCI) rankings -- an external
// reference used to auto-score the "Regulatory complexity" scorecard
// dimension. 1 = most complex jurisdiction to do business in, 79 = least
// complex (2025 edition covers 79 jurisdictions). Source: "TMF Global
// Business Complexity Index: Rankings 2025" PDF, supplied by Peter and
// transcribed 2026-07-10 from its "GBCI rankings 2025" table.
//
// TMF publishes a new edition annually (the same PDF's "GBCI rankings"
// table each year is the thing to re-transcribe). When that happens: update
// TMF_GBCI_YEAR and TMF_GBCI_RANKINGS below, leave TMF_CODE_TO_NAME alone
// unless TMF renames a jurisdiction. Full 79-jurisdiction table is kept
// (not just Atlas's current 21 mapped countries) so a newly-added Atlas
// country already has a rank available without re-reading the PDF.
const TMF_GBCI_YEAR = 2025;

const TMF_GBCI_RANKINGS = {
  'Greece': 1, 'France': 2, 'Mexico': 3, 'Turkey': 4, 'Colombia': 5,
  'Brazil': 6, 'Italy': 7, 'Bolivia': 8, 'Kazakhstan': 9, "China's Mainland": 10,
  'Argentina': 11, 'Paraguay': 12, 'Peru': 13, 'Indonesia': 14, 'Poland': 15,
  'Belgium': 16, 'Spain': 17, 'India': 18, 'Croatia': 19, 'Chile': 20,
  'Portugal': 21, 'Venezuela, Bolivarian Republic of': 22, 'South Korea': 23,
  'Romania': 24, 'Malaysia': 25, 'Philippines': 26, 'Uruguay': 27, 'Russia': 28,
  'Ukraine': 29, 'Ecuador': 30, 'Slovakia': 31, 'Hungary': 32, 'Slovenia': 33,
  'Germany': 34, 'Serbia': 35, 'Austria': 36, 'Egypt': 37, 'Saudi Arabia': 38,
  'United Arab Emirates': 39, 'Bulgaria': 40, 'Sweden': 41, 'Panama': 42,
  'Japan': 43, 'Qatar': 44, 'Guatemala': 45, 'El Salvador': 46, 'Australia': 47,
  'Singapore': 48, 'Canada': 49, 'Nicaragua': 50, 'Taiwan ROC': 51,
  'Switzerland': 52, 'Dominican Republic': 53, 'Vietnam': 54, 'Finland': 55,
  'Thailand': 56, 'Israel': 57, 'Costa Rica': 58, 'Luxembourg': 59,
  'South Africa': 60, 'Ireland': 61, 'Mauritius': 62, 'Cyprus': 63,
  'United States of America': 64, 'Norway': 65, 'Honduras': 66, 'Guernsey': 67,
  'United Kingdom': 68, 'Malta': 69, 'Czech Republic': 70, 'Curacao': 71,
  'British Virgin Islands': 72, 'Jamaica': 73, 'The Netherlands': 74,
  'Jersey': 75, 'Hong Kong, SAR': 76, 'New Zealand': 77, 'Denmark': 78,
  'The Cayman Islands': 79
};

// Atlas's global.json "code" field -> TMF jurisdiction name. Only codes TMF
// actually covers are listed here -- deliberately no entry for Iceland (IS),
// which the 2025 GBCI edition doesn't include at all; tmfRankForCode()
// returns null for it rather than guessing.
const TMF_CODE_TO_NAME = {
  'GB': 'United Kingdom',
  'US': 'United States of America',
  'DE': 'Germany',
  'NL': 'The Netherlands',
  'NO': 'Norway',
  'CH': 'Switzerland',
  'DK': 'Denmark',
  'SE': 'Sweden',
  'ZA': 'South Africa',
  'IE': 'Ireland',
  'FR': 'France',
  'IT': 'Italy',
  'ES': 'Spain',
  'AU': 'Australia',
  'AT': 'Austria',
  'JP': 'Japan',
  'HK': 'Hong Kong, SAR',
  'SG': 'Singapore',
  'NZ': 'New Zealand',
  'CA': 'Canada',
  'FI': 'Finland'
};

function tmfRankForCode(code) {
  const name = TMF_CODE_TO_NAME[(code || '').toUpperCase()];
  if (!name) return null;
  const rank = TMF_GBCI_RANKINGS[name];
  return typeof rank === 'number' ? rank : null;
}
