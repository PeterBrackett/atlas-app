// Shared scorecard reference data, used by both the per-country matrix page
// (country.html) and the cross-country overview (index.html), so the two
// stay in lockstep as more countries and dimensions are added.
//
// Shape matches FINAL_Country_Matrix.xlsx: each segment gets a score of
// 1-3 on each of these 12 dimensions. "Market opportunity" is weighted x3
// in the Overall formula (matching the workbook's own =SUM(...)+(3*mo)
// calculation) since it's judged the primary driver of whether a segment
// is worth pursuing at all. See Sources & Methodology for the full writeup.

// autoScoreNote is set only on the 4 dimensions that can be bulk-overwritten
// in one click from the "Bulk auto-scoring tools" panel on overview.html
// (Market opportunity from allocation size, Regulatory complexity from TMF
// rank, Distribution resources required from institution concentration,
// Languages required from EF EPI). country.html uses this to show a small
// "auto" flag next to the row label, since a value here could be an
// untouched bulk-scored figure rather than a considered per-segment
// judgment call -- worth a second look before relying on it. The other 8
// dimensions have no such tool and are always a manual entry.
const SCORECARD_DIMENSIONS = [
  { key: 'market_opportunity', label: 'Market opportunity', weight: 3,
    question: 'Is there sufficient investable opportunity in this segment to warrant active marketing?',
    autoScoreNote: 'Can be bulk-set from an allocation threshold via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'outsourced_management', label: 'Outsourced management', weight: 1,
    question: 'Are assets outsourced to external managers, or run internally?' },
  { key: 'pricing_impact', label: 'Pricing impact', weight: 1,
    question: 'Is there likely to be pricing pressure in this segment?' },
  { key: 'alignment_of_investment_thinking', label: 'Alignment of investment thinking', weight: 1,
    question: 'Does the market buy the kind of strategy on offer (e.g. concentrated equity, private credit)?' },
  { key: 'distribution_resources_required', label: 'Distribution resources required', weight: 1,
    question: 'How concentrated is asset ownership, and how much resource is needed to cover the market?',
    autoScoreNote: 'Can be bulk-set from institution concentration (top-3 AUM share) via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'regulatory_complexity', label: 'Regulatory complexity', weight: 1,
    question: 'How onerous are accounting, tax and other requirements? Uses TMF Group rankings as a reference where available.',
    autoScoreNote: 'Can be bulk-set from TMF Group rankings via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'client_servicing', label: 'Client servicing', weight: 1,
    question: 'What is the expected demand for reporting and face-to-face meetings?' },
  { key: 'local_presence_required', label: 'Local presence required', weight: 1,
    question: 'Do regulations or culture require a physical office in the market?' },
  { key: 'languages_required', label: 'Languages required', weight: 1,
    question: 'Is English an acceptable language for marketing and client service?',
    autoScoreNote: 'Can be bulk-set from the EF English Proficiency Index via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
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

// Overall's true range, given the current 12 dimensions (market_opportunity
// weighted x3, the other 11 weighted x1): min 14 (everything scored 1), max
// 42 (everything scored 3) -- not 39, which would only be right with 10
// non-market-opportunity dimensions rather than 11. Recomputed from
// SCORECARD_DIMENSIONS rather than hard-coded, so it stays correct if a
// dimension is ever added or removed.
const OVERALL_MIN = SCORECARD_DIMENSIONS.reduce((s, d) => s + 1 * d.weight, 0);
const OVERALL_MAX = SCORECARD_DIMENSIONS.reduce((s, d) => s + 3 * d.weight, 0);

// Red/amber/green banding for the Overall score, agreed with Peter
// (2026-07-10): red at or below 22, amber 23-30, green 31 and up -- a full
// partition of the possible range with no gap at the boundary. Shared so the
// cross-country overview matrix and each country page's own Overall row use
// the exact same cutoffs.
function overallScoreClass(value) {
  if (value === null || typeof value !== 'number') return 'cell-missing';
  if (value <= 22) return 'overall-red';
  if (value <= 30) return 'overall-amber';
  return 'overall-green';
}

// Looks up a segment's allocation for a given asset type, optionally narrowed
// to one finer style within it. Returns null if that segment has no data for
// the selection (e.g. Life/Non-life insurance outside Equities, or a style
// that got folded into "Other/Unspecified" for that particular segment).
// Shared by country.html (one country, segments as columns) and index.html
// (every built country, so the same asset-type/style picker can drive every
// country x segment cell in the cross-country matrix at once).
function getAllocationValue(segment, assetType, assetStyle) {
  const entry = (segment.allocation || []).find(a => a.asset_type === assetType);
  if (!entry) return null;
  if (!assetStyle) return { value_bn: entry.value_bn, pct: entry.pct_of_segment };
  const style = (entry.styles || []).find(s => s.asset_style === assetStyle);
  if (!style) return null;
  const pct = segment.aum_bn ? (style.value_bn / segment.aum_bn * 100) : 0;
  return { value_bn: style.value_bn, pct: Math.round(pct * 10) / 10 };
}

// Builds the ordered list of asset types actually present across a set of
// segments, in the preferred display order, so a dropdown never offers a
// choice with nothing behind it. Shared between the per-country and
// cross-country allocation selectors.
function assetTypesPresent(segs) {
  const typesSeen = new Set();
  segs.forEach(s => (s.allocation || []).forEach(a => typesSeen.add(a.asset_type)));
  const preferredOrder = ['Equities', 'Fixed Income', 'Cash & Short-Term', 'Real Estate', 'Alternatives', 'Other/Unclassified'];
  return preferredOrder.filter(t => typesSeen.has(t)).concat([...typesSeen].filter(t => !preferredOrder.includes(t)));
}

// Styles actually present for one asset type across a set of segments,
// excluding "Other/Unspecified" (that's the fallback, not a real style
// choice), sorted alphabetically.
function assetStylesPresent(segs, assetType) {
  const stylesSeen = new Set();
  segs.forEach(s => {
    const entry = (s.allocation || []).find(a => a.asset_type === assetType);
    (entry ? entry.styles || [] : []).forEach(st => {
      if (st.asset_style !== 'Other/Unspecified') stylesSeen.add(st.asset_style);
    });
  });
  return [...stylesSeen].sort();
}
