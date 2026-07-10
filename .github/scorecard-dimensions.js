// Shared scorecard reference data, used by both the per-country matrix page
// (country.html) and the cross-country overview (index.html), so the two
// stay in lockstep as more countries and dimensions are added.
//
// Shape matches FINAL_Country_Matrix.xlsx: each segment gets a score of
// 1-3 on each of these 12 dimensions. "Market opportunity" is weighted x3
// in the Overall formula (matching the workbook's own =SUM(...)+(3*mo)
// calculation) since it's judged the primary driver of whether a segment
// is worth pursuing at all. See Sources & Methodology for the full writeup.

const SCORECARD_DIMENSIONS = [
  { key: 'market_opportunity', label: 'Market opportunity', weight: 3,
    question: 'Is there sufficient investable opportunity in this segment to warrant active marketing?' },
  { key: 'outsourced_management', label: 'Outsourced management', weight: 1,
    question: 'Are assets outsourced to external managers, or run internally?' },
  { key: 'pricing_impact', label: 'Pricing impact', weight: 1,
    question: 'Is there likely to be pricing pressure in this segment?' },
  { key: 'alignment_of_investment_thinking', label: 'Alignment of investment thinking', weight: 1,
    question: 'Does the market buy the kind of strategy on offer (e.g. concentrated equity, private credit)?' },
  { key: 'distribution_resources_required', label: 'Distribution resources required', weight: 1,
    question: 'How concentrated is asset ownership, and how much resource is needed to cover the market?' },
  { key: 'regulatory_complexity', label: 'Regulatory complexity', weight: 1,
    question: 'How onerous are accounting, tax and other requirements? Uses TMF Group rankings as a reference where available.' },
  { key: 'client_servicing', label: 'Client servicing', weight: 1,
    question: 'What is the expected demand for reporting and face-to-face meetings?' },
  { key: 'local_presence_required', label: 'Local presence required', weight: 1,
    question: 'Do regulations or culture require a physical office in the market?' },
  { key: 'languages_required', label: 'Languages required', weight: 1,
    question: 'Is English an acceptable language for marketing and client service?' },
  { key: 'investor_decision_making', label: 'Investor decision-making', weight: 1,
    question: 'Are decisions made by committee or by an individual, and what are typical timeframes for action?' },
  { key: 'comingled_vehicles', label: 'Comingled vehicles', weight: 1,
    question: 'Can existing pooled vehicles be deployed, or will bespoke/new funds be required?' },
  { key: 'consultant_reliant', label: 'Consultant reliant', weight: 1,
    question: 'Are decisions intermediated — will sign-off from an investment consultant be required?' }
];

// The 19-segment taxonomy in the fixed display order used across the app
// (matches Sources & Methodology), so columns line up the same way on every
// country page and in the cross-country overview.
const CANONICAL_SEGMENT_ORDER = [
  'DB Pension (Corp)', 'DC Pension (Corp)', 'DB Pension (Govt)', 'DC Pension (Govt)',
  'DB Pension (Union)', 'DC Pension (Union)',
  'DB Pension (Healthcare non-profit)', 'DC Pension (Healthcare non-profit)',
  'DB Pension (Endowments)', 'DC Pension (Endowments)',
  'DB Pension (Tax exempt)', 'DC Pension (Tax exempt)',
  'Endowments E&F', 'Tax exempt E&F', 'Foundations E&F', 'Healthcare non-profit E&F',
  'Life insurance', 'Non-life insurance', 'SWF'
];

// Five peer top-level regions, following asset-management-industry convention
// (Middle East and Africa split out from Europe) rather than the UN geoscheme
// (which folds Middle East into Asia). Added 2026-07-10 alongside the
// region/sub_region fields on each global.json country entry. Display order
// matches how the regions were first specified, not alphabetical or by AUM.
const REGION_ORDER = ['Europe', 'Middle East', 'Africa', 'Asia Pacific', 'Americas'];

function regionSortIndex(regionName) {
  const i = REGION_ORDER.indexOf(regionName);
  return i === -1 ? REGION_ORDER.length : i;
}

function segmentSortIndex(segmentName) {
  const i = CANONICAL_SEGMENT_ORDER.indexOf(segmentName);
  return i === -1 ? CANONICAL_SEGMENT_ORDER.length : i;
}

// Overall score for a segment, or null if any of the 12 dimensions hasn't
// been scored yet. Deliberately not a partial sum — an incomplete segment
// should read as "not yet scored", not as a misleadingly low real number.
// This is the mechanism that nudges completion: gaps stay visibly blank.
function computeOverallScore(scorecard) {
  if (!scorecard) return null;
  let total = 0;
  for (const dim of SCORECARD_DIMENSIONS) {
    const v = scorecard[dim.key];
    if (typeof v !== 'number') return null;
    total += v * dim.weight;
  }
  return total;
}

function scoredDimensionCount(scorecard) {
  if (!scorecard) return 0;
  return SCORECARD_DIMENSIONS.filter(d => typeof scorecard[d.key] === 'number').length;
}
