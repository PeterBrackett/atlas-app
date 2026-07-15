// Shared scorecard reference data, used by both the per-country matrix page
// (country.html) and the cross-country overview (index.html), so the two
// stay in lockstep as more countries and dimensions are added.
//
// Shape matches FINAL_Country_Matrix.xlsx: each segment gets a score of
// 1-3 on each of these 12 dimensions. "Market opportunity" is weighted x3
// in the Overall formula (matching the workbook's own =SUM(...)+(3*mo)
// calculation) since it's judged the primary driver of whether a segment
// is worth pursuing at all. See Sources & Methodology for the full writeup.

// autoScoreNote is set on the 11 dimensions that can be bulk-overwritten in
// one click from the "Bulk auto-scoring tools" panel on overview.html
// (Market opportunity from allocation size, Regulatory complexity from TMF
// rank, Distribution resources required from institution concentration,
// Languages required from EF EPI, Local presence required from OECD
// FDIRRI, Client servicing from Mercer Integrity, Consultant reliant from
// IC Research consultant density, Pricing impact from segment AUM size,
// and Outsourced management / Comingled vehicles / Investor
// decision-making from starting-point defaults by segment type -- see
// segment-type-defaults.js). country.html uses this to show a small "auto"
// flag next to the row label, since a value here could be an untouched
// bulk-scored figure rather than a considered per-segment judgment call --
// worth a second look before relying on it, especially for the last four,
// which are heuristics rather than real external data. Only "Alignment of
// investment thinking" has no such tool and is always a manual entry --
// no usable proxy for it was found.
const SCORECARD_DIMENSIONS = [
  { key: 'market_opportunity', label: 'Market opportunity', weight: 3,
    question: 'Is there sufficient investable opportunity in this segment to warrant active marketing?',
    autoScoreNote: 'Can be bulk-set from an allocation threshold via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'outsourced_management', label: 'Outsourced management', weight: 1,
    question: 'Are assets outsourced to external managers, or run internally?',
    autoScoreNote: 'Can be bulk-set from a starting-point default by segment type via the overview page’s auto-scoring tools (not real data — see segment-type-defaults.js) — check it reflects a real judgment call, not just an untouched default.' },
  { key: 'pricing_impact', label: 'Pricing impact', weight: 1,
    question: 'Is there likely to be pricing pressure in this segment?',
    autoScoreNote: 'Can be bulk-set from a segment-AUM-size heuristic via the overview page’s auto-scoring tools (not a real pricing dataset) — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'alignment_of_investment_thinking', label: 'Alignment of investment thinking', weight: 1,
    question: 'Does the market buy the kind of strategy on offer (e.g. concentrated equity, private credit)?' },
  { key: 'distribution_resources_required', label: 'Distribution resources required', weight: 1,
    question: 'How concentrated is asset ownership, and how much resource is needed to cover the market?',
    autoScoreNote: 'Can be bulk-set from institution concentration (top-3 AUM share) via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'regulatory_complexity', label: 'Regulatory complexity', weight: 1,
    question: 'How onerous are accounting, tax and other requirements? Uses TMF Group rankings as a reference where available.',
    autoScoreNote: 'Can be bulk-set from TMF Group rankings via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'client_servicing', label: 'Client servicing', weight: 1,
    question: 'What is the expected demand for reporting and face-to-face meetings?',
    autoScoreNote: 'Can be bulk-set from the Mercer CFA Institute Global Pension Index Integrity sub-index via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'local_presence_required', label: 'Local presence required', weight: 1,
    question: 'Do regulations or culture require a physical office in the market?',
    autoScoreNote: 'Can be bulk-set from the OECD FDI Regulatory Restrictiveness Index via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'languages_required', label: 'Languages required', weight: 1,
    question: 'Is English an acceptable language for marketing and client service?',
    autoScoreNote: 'Can be bulk-set from the EF English Proficiency Index via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' },
  { key: 'investor_decision_making', label: 'Investor decision-making', weight: 1,
    question: 'Are decisions made by committee or by an individual, and what are typical timeframes for action?',
    autoScoreNote: 'Can be bulk-set from a starting-point default by segment type via the overview page’s auto-scoring tools (not real data — see segment-type-defaults.js) — check it reflects a real judgment call, not just an untouched default.' },
  { key: 'comingled_vehicles', label: 'Comingled vehicles', weight: 1,
    question: 'Can existing pooled vehicles be deployed, or will bespoke/new funds be required?',
    autoScoreNote: 'Can be bulk-set from a starting-point default by segment type via the overview page’s auto-scoring tools (not real data — see segment-type-defaults.js) — check it reflects a real judgment call, not just an untouched default.' },
  { key: 'consultant_reliant', label: 'Consultant reliant', weight: 1,
    question: 'Are decisions intermediated — will sign-off from an investment consultant be required?',
    autoScoreNote: 'Can be bulk-set from IC Research consultant density via the overview page’s auto-scoring tools — check it reflects a real judgment call, not just an untouched auto-score.' }
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

// enabledDimensions is an optional {dimensionKey: boolean} map, sourced
// from global.json's enabled_dimensions field (see the "toggle factors
// on/off" feature added 2026-07-15). A dimension counts as enabled unless
// it's explicitly false in the map -- undefined/missing map, or a missing
// key within it, both mean "enabled", so every existing call site that
// doesn't pass this argument at all keeps its old all-12-required
// behaviour with no code changes needed there.
function isDimensionEnabled(dimKey, enabledDimensions) {
  return !enabledDimensions || enabledDimensions[dimKey] !== false;
}

function enabledDimensionCount(enabledDimensions) {
  return SCORECARD_DIMENSIONS.filter(d => isDimensionEnabled(d.key, enabledDimensions)).length;
}

// Overall score for a segment, or null if any ENABLED dimension hasn't
// been scored yet. Deliberately not a partial sum over the required set —
// an incomplete segment should read as "not yet scored", not as a
// misleadingly low real number. This is the mechanism that nudges
// completion: gaps stay visibly blank. Disabled dimensions (per
// enabledDimensions) are skipped entirely -- not required, not summed --
// so toggling something off can turn a previously-null Overall into a real
// number without needing that dimension scored.
function computeOverallScore(scorecard, enabledDimensions) {
  if (!scorecard) return null;
  let total = 0;
  for (const dim of SCORECARD_DIMENSIONS) {
    if (!isDimensionEnabled(dim.key, enabledDimensions)) continue;
    const v = scorecard[dim.key];
    if (typeof v !== 'number') return null;
    total += v * dim.weight;
  }
  return total;
}

// Count of enabled dimensions that have a score, out of enabledDimensionCount()
// (not always 12 -- see enabledDimensions above). Disabled dimensions are
// excluded from both the numerator and the denominator a caller would use.
function scoredDimensionCount(scorecard, enabledDimensions) {
  if (!scorecard) return 0;
  return SCORECARD_DIMENSIONS.filter(d => isDimensionEnabled(d.key, enabledDimensions) && typeof scorecard[d.key] === 'number').length;
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

// Monochrome line-icon glyphs for the 12 scorecard dimensions, shown next to
// each dimension's row label on country.html's scorecard table. Inner SVG
// markup only (no wrapping <svg> tag) so callers can size/style the wrapper
// themselves via the .dim-icon class in style.css; all paths assume a
// 0 0 24 24 viewBox, fill:none, stroke:currentColor so they inherit the
// row-label text color automatically (navy in the current theme). Kept here
// rather than inline in country.html so overview.html or any future page can
// reuse the same set. Agreed with Peter (2026-07-14): coins for Pricing
// impact rather than a price tag, everything else a first pass he approved.
const DIMENSION_ICONS = {
  market_opportunity: '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/>',
  outsourced_management: '<path d="M10 6H6a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"/><path d="M14 4h6v6"/><path d="M20 4 11 13"/>',
  pricing_impact: '<circle cx="15" cy="9" r="6.5"/><circle cx="9" cy="15" r="6.5"/><line x1="15" y1="6" x2="15" y2="12"/><line x1="9" y1="12" x2="9" y2="18"/>',
  alignment_of_investment_thinking: '<circle cx="12" cy="12" r="8"/><polygon points="12,7.5 14,12 12,16.5 10,12"/>',
  distribution_resources_required: '<circle cx="12" cy="9" r="5"/><polygon points="12,21 8,13 16,13"/>',
  regulatory_complexity: '<line x1="12" y1="4" x2="12" y2="20"/><line x1="4" y1="8" x2="20" y2="8"/><line x1="4" y1="8" x2="4" y2="12"/><path d="M2 12a2 2 0 0 0 4 0"/><line x1="20" y1="8" x2="20" y2="12"/><path d="M18 12a2 2 0 0 0 4 0"/><line x1="7" y1="20" x2="17" y2="20"/>',
  client_servicing: '<path d="M4 13a8 8 0 0 1 16 0"/><rect x="3" y="13" width="4" height="6" rx="2"/><rect x="17" y="13" width="4" height="6" rx="2"/><path d="M20 19v1a3 3 0 0 1-3 3h-3"/>',
  local_presence_required: '<rect x="4" y="3" width="8" height="18"/><rect x="12" y="9" width="8" height="12"/><rect x="7" y="16" width="2" height="5"/>',
  languages_required: '<circle cx="12" cy="12" r="8"/><line x1="4" y1="12" x2="20" y2="12"/><path d="M12 4a12 12 0 0 1 0 16"/><path d="M12 4a12 12 0 0 0 0 16"/>',
  investor_decision_making: '<circle cx="12" cy="12" r="8"/><line x1="12" y1="12" x2="12" y2="7"/><line x1="12" y1="12" x2="15.5" y2="14"/>',
  comingled_vehicles: '<circle cx="9" cy="12" r="6"/><circle cx="15" cy="12" r="6"/>',
  consultant_reliant: '<rect x="3" y="8" width="18" height="12" rx="2"/><path d="M8 8V6a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="3" y1="14" x2="21" y2="14"/>'
};

// Wraps a dimension's icon markup in a sized <svg>, ready to drop straight
// into a row label. Falls back to an empty string for an unknown key rather
// than throwing, so a future dimension added without an icon degrades to
// "no icon" instead of a broken page.
function dimensionIconSvg(dimKey) {
  const inner = DIMENSION_ICONS[dimKey];
  if (!inner) return '';
  return `<svg class="dim-icon" viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${inner}</svg>`;
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

// Segment-level "how much of this segment's total AUM is actually backed by
// reported allocation data" ratio -- sum of every allocation entry's
// value_bn (all asset types combined) divided by the segment's aum_bn. Not
// every institution counted in aum_bn also filed an asset-class breakdown
// (documented at source: ~96% of UK firms have at least one allocation line,
// smaller DC/money-purchase plans disproportionately file none at all), so
// this is consistently < 100% for most segments. A single ratio computed
// once per segment and applied uniformly to every asset type/style within
// it, per Peter's framing (2026-07-15) -- simpler than trying to estimate a
// separate non-reporting rate per asset class, and there's no data basis to
// do the latter more precisely anyway.
//
// A handful of segments (9 as of 2026-07-15, see fix_aum_bn_bugs.py for the
// 5 that were resolved as an aum_bn data bug) show >100% -- allocation
// summing to MORE than the segment total. That's not a reporting-gap
// situation, it's aum_bn and allocation disagreeing about the segment's true
// size, so it gets flagged rather than "scaled" (scaling by a ratio above 1
// would shrink the estimate, which is backwards for a max/ceiling figure).
function allocationCoverageRatio(segment) {
  const aum = segment.aum_bn;
  if (!aum) return null;
  const allocSum = (segment.allocation || []).reduce((sum, a) => sum + (a.value_bn || 0), 0);
  return allocSum / aum;
}

// Same lookup as getAllocationValue(), but returns a min/max range instead
// of a single figure -- an attempt to make the "how much AUM is really in
// this asset class" question honest about the fact that not every
// institution in the segment reported a breakdown:
//   min_bn  = the reported value as-is. Assumes every non-reporting
//             institution holds NONE of this asset class -- a genuine floor,
//             not a best guess.
//   max_bn  = reported value scaled up by 1/coverage, i.e. value_bn *
//             (aum_bn / total reported allocation). Assumes non-reporting
//             institutions have the same asset mix as reporting ones,
//             extrapolated across the segment's full AUM -- a ceiling, not
//             a best guess either.
// The true figure is somewhere in between and unknowable from this data
// alone -- that's the point of showing a range rather than picking one.
//
// Returns null under the same conditions getAllocationValue() would (no
// data for this type/style at all). When coverage can't be computed (no
// aum_bn) or is >100% (aum_bn/allocation disagreement -- see
// allocationCoverageRatio() above), min_bn and max_bn both just equal the
// reported value_bn and `flagged` is set so callers can surface that this
// particular figure isn't a real range, just a number with a caveat.
// 0.1% tolerance on the >100% check -- rounding a corrected aum_bn to 3
// decimal places (see fix_aum_bn_bugs.py) can leave a coverage ratio like
// 100.0009% that's really just floating-point noise, not a genuine
// aum_bn/allocation disagreement worth flagging.
const COVERAGE_ANOMALY_TOLERANCE = 1.001;

function getAllocationRange(segment, assetType, assetStyle) {
  const result = getAllocationValue(segment, assetType, assetStyle);
  if (!result) return null;

  const coverage = allocationCoverageRatio(segment);
  const coveragePct = typeof coverage === 'number' ? Math.round(coverage * 1000) / 10 : null;

  if (coverage === null || coverage <= 0 || coverage > COVERAGE_ANOMALY_TOLERANCE) {
    return {
      value_bn: result.value_bn,
      min_bn: result.value_bn,
      max_bn: result.value_bn,
      coverage_pct: coveragePct,
      flagged: coverage !== null && coverage > COVERAGE_ANOMALY_TOLERANCE
    };
  }

  const maxBn = Math.round((result.value_bn / coverage) * 1000) / 1000;
  return {
    value_bn: result.value_bn,
    min_bn: result.value_bn,
    max_bn: maxBn,
    coverage_pct: coveragePct,
    flagged: false
  };
}

// Compact display string for a range from getAllocationRange(), e.g.
// "$5.24bn" when there's no meaningful range to show (100% coverage, or
// coverage couldn't be computed), "$5.24-8.16bn (64% reported)" when there
// is one, or "$5.24bn (data flagged — see notes)" for the >100%-coverage
// anomaly cases. maximumFractionDigits kept at 2 to match the rest of the
// site's $bn formatting.
function formatAllocationRange(range) {
  if (!range) return '—';
  const fmt = (v) => v.toLocaleString(undefined, { maximumFractionDigits: 2 });
  if (range.flagged) {
    return `$${fmt(range.value_bn)}bn (data flagged — see notes)`;
  }
  if (range.coverage_pct === null || range.coverage_pct >= 100) {
    return `$${fmt(range.value_bn)}bn`;
  }
  return `$${fmt(range.min_bn)}–${fmt(range.max_bn)}bn (${range.coverage_pct}% reported)`;
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
