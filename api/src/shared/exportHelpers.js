// Server-side mirror of atlas-site/scorecard-dimensions.js. Kept as a
// deliberate duplicate rather than a shared import because the site's copy
// is a plain browser <script> (no module system) and the API is a separate
// Node deployment (api_location) with no access to atlas-site's files at
// build time. If the 12 dimensions, weights, or canonical segment order
// ever change, update both files -- this one and atlas-site/scorecard-dimensions.js.
//
// Unlike the client-side copy, this file's own .weight values never get
// mutated by a "pushed" global weighting (see applyGlobalDimensionWeights()
// in scorecard-dimensions.js) -- there's no live global.json fetch here to
// react to. Instead, country.html/picker.html always send an explicit,
// fully-resolved weight_overrides object in every export request (built
// client-side from whatever SCORECARD_DIMENSIONS' current, possibly-pushed,
// weights are, merged with any session-local override on top), so
// buildScorecardMatrix()'s weightOverrides parameter is the one and only
// place a pushed weighting reaches an export -- not this constant.
const SCORECARD_DIMENSIONS = [
  { key: 'market_opportunity', label: 'Market opportunity', weight: 3 },
  { key: 'outsourced_management', label: 'Outsourced management', weight: 1 },
  { key: 'pricing_impact', label: 'Pricing impact', weight: 1 },
  { key: 'alignment_of_investment_thinking', label: 'Alignment of investment thinking', weight: 1 },
  { key: 'distribution_resources_required', label: 'Distribution resources required', weight: 1 },
  { key: 'regulatory_complexity', label: 'Regulatory complexity', weight: 1 },
  { key: 'client_servicing', label: 'Client servicing', weight: 1 },
  { key: 'local_presence_required', label: 'Local presence required', weight: 1 },
  { key: 'languages_required', label: 'Languages required', weight: 1 },
  { key: 'investor_decision_making', label: 'Investor decision-making', weight: 1 },
  { key: 'comingled_vehicles', label: 'Comingled vehicles', weight: 1 },
  { key: 'consultant_reliant', label: 'Consultant reliant', weight: 1 }
];

const CANONICAL_SEGMENT_ORDER = [
  'DB Pension (Corp)', 'DC Pension (Corp)', 'DB Pension (Govt)', 'DC Pension (Govt)',
  'DB Pension (Union)', 'DC Pension (Union)',
  'DB Pension (Healthcare non-profit)', 'DC Pension (Healthcare non-profit)',
  'DB Pension (Endowments)', 'DC Pension (Endowments)',
  'DB Pension (Tax exempt)', 'DC Pension (Tax exempt)',
  'Endowments E&F', 'Tax exempt E&F', 'Foundations E&F', 'Healthcare non-profit E&F',
  'Life insurance', 'Non-life insurance', 'SWF'
];

function segmentSortIndex(segmentName) {
  const i = CANONICAL_SEGMENT_ORDER.indexOf(segmentName);
  return i === -1 ? CANONICAL_SEGMENT_ORDER.length : i;
}

// Client-defined custom scorecard factors (global.json's custom_dimensions
// field -- see setCustomDimensions.js) are threaded through as an explicit
// `customDimensions` argument to every function below, rather than being
// pushed onto the module-level SCORECARD_DIMENSIONS constant the way
// applyCustomDimensions() mutates the client-side copy in
// scorecard-dimensions.js. That mutate-in-place trick is safe there because
// each browser tab has its own JS heap, but this file runs inside an Azure
// Functions process that can serve several concurrent export requests (from
// different clients, with different custom factors) against the same warm
// module instance -- mutating a shared array per-request would let one
// request's custom factors leak into another's export. resolveDimensions()
// instead builds a fresh, request-scoped array on every call, leaving the
// shared constant untouched.
function resolveDimensions(customDimensions) {
  return (customDimensions && customDimensions.length) ? SCORECARD_DIMENSIONS.concat(customDimensions) : SCORECARD_DIMENSIONS;
}

// Kept in sync with the same functions in scorecard-dimensions.js (see the
// comment there for the reasoning) -- enabledDimensions is an optional
// {dimensionKey: boolean} map sourced from global.json's enabled_dimensions
// field; a dimension counts as enabled unless it's explicitly false.
function isDimensionEnabled(dimKey, enabledDimensions) {
  return !enabledDimensions || enabledDimensions[dimKey] !== false;
}

function enabledDimensionCount(enabledDimensions, customDimensions) {
  return resolveDimensions(customDimensions).filter((d) => isDimensionEnabled(d.key, enabledDimensions)).length;
}

// Mirrors atlas-site/scorecard-dimensions.js -- changed 2026-07-15 to
// always return a real number: a missing dimension contributes 0 instead
// of blocking the whole Overall. See that file's comment for the reasoning.
//
// weightOverrides is an optional {dimensionKey: number} map -- added
// 2026-07-17 so the Word/PowerPoint exports can reflect the project
// builder's per-dimension weighting column (picker.html), the same way
// atlas-site/scorecard-dimensions.js's computeOverallScore() does for the
// on-screen table. This export-side copy was the actual gap: the export
// functions never received the custom weights at all, so an export always
// used the default weighting regardless of what picker.html was showing on
// screen. A dimension missing from the map (or no map at all) keeps its own
// default weight, so every existing caller (country.html's single-country
// export, which has no weighting UI) is unaffected.
function computeOverallScore(scorecard, enabledDimensions, weightOverrides, customDimensions) {
  let total = 0;
  for (const dim of resolveDimensions(customDimensions)) {
    if (!isDimensionEnabled(dim.key, enabledDimensions)) continue;
    const v = scorecard ? scorecard[dim.key] : undefined;
    const w = (weightOverrides && typeof weightOverrides[dim.key] === 'number') ? weightOverrides[dim.key] : dim.weight;
    total += (typeof v === 'number' ? v : 0) * w;
  }
  return total;
}

// Mirrors atlas-site/scorecard-dimensions.js's computeOverallRange() --
// generalizes the fixed 14-42 range to an arbitrary weight set and
// enabled-dimension set, so overallColor()'s red/amber/green banding can be
// rescaled to match a custom weighting instead of staying calibrated for
// the default one.
function computeOverallRange(enabledDimensions, weightOverrides, customDimensions) {
  let min = 0, max = 0;
  for (const dim of resolveDimensions(customDimensions)) {
    if (!isDimensionEnabled(dim.key, enabledDimensions)) continue;
    const w = (weightOverrides && typeof weightOverrides[dim.key] === 'number') ? weightOverrides[dim.key] : dim.weight;
    min += 1 * w;
    max += 3 * w;
  }
  return { min, max };
}

function scoredDimensionCount(scorecard, enabledDimensions, customDimensions) {
  if (!scorecard) return 0;
  return resolveDimensions(customDimensions).filter((d) => isDimensionEnabled(d.key, enabledDimensions) && typeof scorecard[d.key] === 'number').length;
}

// Finds a segment's total Equities allocation figure, if present, matching
// country.html's getAllocationValue() for the (Equities, no style) case.
function equityValue(segment) {
  const entry = (segment.allocation || []).find((a) => a.asset_type === 'Equities');
  return entry ? entry.value_bn : null;
}

// Server-side mirror of scorecard-dimensions.js's allocationCoverageRatio/
// getAllocationRange/formatAllocationRange (see that file for the full
// reasoning) -- kept as a deliberate duplicate for the same reason as
// SCORECARD_DIMENSIONS above: the site's copy is a plain browser <script>,
// this is a separate Node deployment with no shared module access at build
// time. If the min/max logic ever changes, update both files.
function allocationCoverageRatio(segment) {
  const aum = segment.aum_bn;
  if (!aum) return null;
  const allocSum = (segment.allocation || []).reduce((sum, a) => sum + (a.value_bn || 0), 0);
  return allocSum / aum;
}

// Kept in sync with the same constant in scorecard-dimensions.js -- see the
// comment there for why a small tolerance is needed on the >100% check.
const COVERAGE_ANOMALY_TOLERANCE = 1.001;

function getAllocationRange(segment, assetType, assetStyle) {
  const entry = (segment.allocation || []).find((a) => a.asset_type === assetType);
  if (!entry) return null;
  let value_bn = entry.value_bn;
  if (assetStyle) {
    const style = (entry.styles || []).find((s) => s.asset_style === assetStyle);
    if (!style) return null;
    value_bn = style.value_bn;
  }

  const coverage = allocationCoverageRatio(segment);
  const coverage_pct = typeof coverage === 'number' ? Math.round(coverage * 1000) / 10 : null;

  if (coverage === null || coverage <= 0 || coverage > COVERAGE_ANOMALY_TOLERANCE) {
    return { value_bn, min_bn: value_bn, max_bn: value_bn, coverage_pct, flagged: coverage !== null && coverage > COVERAGE_ANOMALY_TOLERANCE };
  }

  const max_bn = Math.round((value_bn / coverage) * 1000) / 1000;
  return { value_bn, min_bn: value_bn, max_bn, coverage_pct, flagged: false };
}

function formatAllocationRange(range) {
  if (!range) return null;
  const fmt = (v) => v.toLocaleString(undefined, { maximumFractionDigits: 2 });
  if (range.flagged) return `$${fmt(range.value_bn)}bn (data flagged)`;
  if (range.coverage_pct === null || range.coverage_pct >= 100) return `$${fmt(range.value_bn)}bn`;
  return `$${fmt(range.min_bn)}–${fmt(range.max_bn)}bn (${range.coverage_pct}% reported)`;
}

// Rows for the "AUM by segment" table, sorted largest-first -- same order
// country.html uses. equity_range is the min/max-scaled string (or null if
// there's no Equities data at all for this segment, e.g. Life/Non-life
// insurance's OECD-sourced allocation, which has no reporting-coverage gap
// to speak of since it's an industry aggregate, not itemised institutions).
function buildAumRows(segments) {
  return (segments || [])
    .slice()
    .sort((a, b) => b.aum_bn - a.aum_bn)
    .map((s) => ({
      segment: s.segment,
      aum_bn: s.aum_bn,
      equity_bn: equityValue(s),
      equity_range: formatAllocationRange(getAllocationRange(s, 'Equities', '')),
      basis: s.basis || ''
    }));
}

// Traffic-light colors matching atlas-site/style.css's td.score-1/2/3 and
// td.overall-red/amber/green rules exactly (same hex values), so the Word
// and PowerPoint exports read as the same visual language as the site
// rather than plain black-and-white tables. Returns null for cells that
// shouldn't be colored (missing/non-numeric values).
function scoreColor(value) {
  if (value === 1) return { bg: 'FBE1E1', fg: 'A3291F' };
  if (value === 2) return { bg: 'FDEEE0', fg: '9A5A1A' };
  if (value === 3) return { bg: 'E3F2E3', fg: '1F7A34' };
  return null;
}

// Matches atlas-site/style.css's td.cell-missing (2026-07-15) -- a distinct
// yellow flag for "no data entered", separate from the red/amber/green
// score palette above, so a 0 substituted for a genuine gap doesn't read as
// a real (bad) score of 0.
const MISSING_COLOR = { bg: 'FFF3B0', fg: '7A5C00' };

// Overall score thresholds match style.css's comment: range 14-42,
// <=22 red, 23-30 amber, >=31 green.
const DEFAULT_OVERALL_MIN = SCORECARD_DIMENSIONS.reduce((s, d) => s + 1 * d.weight, 0);
const DEFAULT_OVERALL_MAX = SCORECARD_DIMENSIONS.reduce((s, d) => s + 3 * d.weight, 0);
const RED_CUTOFF_FRACTION = (22 - DEFAULT_OVERALL_MIN) / (DEFAULT_OVERALL_MAX - DEFAULT_OVERALL_MIN);
const AMBER_CUTOFF_FRACTION = (30 - DEFAULT_OVERALL_MIN) / (DEFAULT_OVERALL_MAX - DEFAULT_OVERALL_MIN);

// `range` is optional -- {min, max} from computeOverallRange(), passed when
// a custom weighting is in effect (see computeOverallScore() above) so the
// 22/30 cutoffs get rescaled onto whatever range actually applies, at the
// same relative position they sit at within the default 14-42 range.
// Callers that don't pass `range` get byte-for-byte the original behaviour.
function overallColor(value, range) {
  if (typeof value !== 'number') return null;
  if (!range) {
    if (value <= 22) return { bg: 'FBE1E1', fg: 'A3291F' };
    if (value <= 30) return { bg: 'FDEEE0', fg: '9A5A1A' };
    return { bg: 'E3F2E3', fg: '1F7A34' };
  }
  const span = range.max - range.min;
  const redCutoff = range.min + RED_CUTOFF_FRACTION * span;
  const amberCutoff = range.min + AMBER_CUTOFF_FRACTION * span;
  if (value <= redCutoff) return { bg: 'FBE1E1', fg: 'A3291F' };
  if (value <= amberCutoff) return { bg: 'FDEEE0', fg: '9A5A1A' };
  return { bg: 'E3F2E3', fg: '1F7A34' };
}

// Column-ordered segments plus row data for the opportunity scorecard
// matrix, matching renderScorecardMatrix() in country.html: AUM row, then
// one row per dimension, then Scored and Overall. Each row carries a
// parallel `colors` array (one entry per column, null or {bg,fg} hex pair)
// so the Word/PowerPoint exports can reproduce the site's red/amber/green
// traffic-light coding without re-deriving it from the display strings.
// enabledDimensions is the optional {dimensionKey: boolean} map from the
// "toggle factors on/off" feature (global.json's enabled_dimensions) --
// unlike the on-screen matrix (country.html/overview.html), which keeps a
// disabled dimension's column visible with an "Off" flag, Peter's
// 2026-07-15 request was for exports to leave it out of the table
// altogether, not just exclude it from the Scored/Overall calculation. So
// disabled dimensions get no row at all here.
// weightOverrides is an optional {dimensionKey: number} map -- see
// computeOverallScore()'s comment. Affects three things: the dimension row
// labels' "(x{weight})" suffix now reflects the effective (possibly
// overridden) weight rather than always the default, the Overall row uses
// the overridden weights, and Overall's colour banding rescales to match
// via computeOverallRange().
// `allocType`/`allocStyle` -- added 2026-07-23 after Peter noticed the
// exported scorecard table never had an allocation row at all, unlike the
// on-screen matrix (country.html/picker.html/overview.html), which shows AUM
// directly followed by an Allocation row driven by the "Allocation row
// shows" dropdown (CURRENT_ALLOC_TYPE/CURRENT_ALLOC_STYLE on-screen). Revised
// 2026-08-07 per Peter: no asset class is a real, and now the default,
// selection -- the caller not sending one (an older cached client, a direct
// API call, or simply never having opted into an asset-class breakdown for
// this export) means the output shows AUM only, with no allocation row at
// all, rather than silently defaulting to Equities.
function buildScorecardMatrix(segments, enabledDimensions, weightOverrides, allocType, allocStyle, customDimensions) {
  const cols = (segments || []).slice().sort((a, b) => segmentSortIndex(a.segment) - segmentSortIndex(b.segment));
  const enabledCount = enabledDimensionCount(enabledDimensions, customDimensions);
  // Recompute the Overall range (rather than relying on the fixed 22/30
  // cutoffs baked into overallColor()'s no-range branch) whenever either a
  // weighting override OR any custom dimension is in play -- a custom
  // factor shifts the achievable Overall range even at its default weight,
  // so the fixed cutoffs (calibrated only for the original 12 dimensions)
  // would otherwise miscolor every Overall cell once a custom factor exists,
  // not just ones with an explicit weight override.
  const overallRange = (weightOverrides || (customDimensions && customDimensions.length))
    ? computeOverallRange(enabledDimensions, weightOverrides, customDimensions)
    : null;
  const resolvedAllocType = allocType || '';
  const resolvedAllocStyle = allocStyle || '';

  const dimensionRows = resolveDimensions(customDimensions)
    .filter((dim) => isDimensionEnabled(dim.key, enabledDimensions))
    .map((dim) => {
      const effectiveWeight = (weightOverrides && typeof weightOverrides[dim.key] === 'number') ? weightOverrides[dim.key] : dim.weight;
      return {
        key: dim.key,
        type: 'dimension',
        label: dim.label + (effectiveWeight !== 1 ? ` (x${effectiveWeight})` : ''),
        values: cols.map((s) => {
          const v = s.scorecard ? s.scorecard[dim.key] : undefined;
          return typeof v === 'number' ? String(v) : '0';
        }),
        colors: cols.map((s) => {
          const v = s.scorecard ? s.scorecard[dim.key] : undefined;
          return typeof v === 'number' ? scoreColor(v) : MISSING_COLOR;
        })
      };
    });

  const scoredRow = {
    type: 'scored',
    label: 'Scored',
    values: cols.map((s) => `${scoredDimensionCount(s.scorecard, enabledDimensions, customDimensions)}/${enabledCount}`),
    colors: cols.map(() => null)
  };

  const overallRow = {
    type: 'overall',
    label: 'Overall',
    values: cols.map((s) => String(computeOverallScore(s.scorecard, enabledDimensions, weightOverrides, customDimensions))),
    colors: cols.map((s) => overallColor(computeOverallScore(s.scorecard, enabledDimensions, weightOverrides, customDimensions), overallRange))
  };

  const aumRow = {
    type: 'aum',
    label: 'AUM ($bn)',
    values: cols.map((s) => (typeof s.aum_bn === 'number' ? s.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-')),
    colors: cols.map(() => null)
  };

  // Same selection as country.html's allocLabel/allocRow, but a plain number
  // (matching the AUM row's own style) rather than formatAllocationRange()'s
  // full "$437.84–458.66bn (95.5% reported)" string -- every other column in
  // this table is capped at 6 characters wide (see buildScorecardTable() in
  // exportDocx.js/exportPptx.js), which the descriptive min-max-coverage
  // string would blow straight through, wrapping the cell across several
  // lines and, in the PowerPoint version, desyncing every dimension icon
  // below it (they're positioned assuming a fixed row height -- see
  // SCORECARD_ROW_H's comment in exportPptx.js). The fuller range/coverage
  // detail is still available in the separate "AUM by segment" table's
  // Equities range column; this row is just the reported figure at a glance.
  // No row at all when resolvedAllocType is empty (the default, and now
  // also what a user gets from the picker.html builder unless they've
  // explicitly opted into an asset class for this export) -- see this
  // function's header comment.
  const allocRow = resolvedAllocType ? {
    type: 'allocation',
    label: resolvedAllocStyle ? `${resolvedAllocType} — ${resolvedAllocStyle} ($bn)` : `${resolvedAllocType} ($bn)`,
    values: cols.map((s) => {
      const range = getAllocationRange(s, resolvedAllocType, resolvedAllocStyle);
      return range ? range.value_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-';
    }),
    colors: cols.map(() => null)
  } : null;

  return {
    columnLabels: cols.map((s) => s.segment),
    rows: [aumRow, ...(allocRow ? [allocRow] : []), ...dimensionRows, scoredRow, overallRow]
  };
}

// Character-count based column width estimator, shared by both exports so
// table columns are sized to their actual content instead of being
// stretched to fill the page/slide (Peter's 2026-07-15 feedback: "Rank only
// needs to be the width of the word Rank... Institution needs to be the
// width of the longest name unless very long ones wrap... AUM generally
// only needs to be the width of AUM ($bn)"). Neither docx nor pptxgenjs
// measures text and autofits columns on its own, so this does a cheap
// character-count approximation instead: each column's width is driven by
// its longest cell (header or body), clamped to [minChars, a per-column
// cap] so one long institution name doesn't blow out the whole table (it
// wraps within the cap instead) and a short column like "Rank" doesn't
// collapse to nothing. Returns character counts, not physical units -- each
// export file converts to its own unit (twips for docx, inches for
// pptxgenjs) with its own per-character constant, since the two render at
// different default font sizes.
function estimateColumnCharWidths(headerLabels, bodyRows, opts = {}) {
  const minChars = opts.minChars || 3;
  const defaultMax = opts.maxChars || 40;
  const maxCharsPerCol = opts.maxCharsPerCol;
  const capFor = (i) => (maxCharsPerCol && typeof maxCharsPerCol[i] === 'number') ? maxCharsPerCol[i] : defaultMax;

  const widths = headerLabels.map((h, i) => Math.min(Math.max(String(h == null ? '' : h).length, minChars), capFor(i)));
  (bodyRows || []).forEach((row) => {
    for (let i = 0; i < headerLabels.length; i++) {
      const len = row[i] == null ? 0 : String(row[i]).length;
      widths[i] = Math.max(widths[i], Math.min(Math.max(len, minChars), capFor(i)));
    }
  });
  return widths;
}

// Top institutions by AUM per segment, matching Peter's standard "top 10 +
// concentration %" report format -- mirrors atlas-site/country.html's
// segmentsWithTopInstitutions(). Only segments with
// concentration.top_institutions populated (see add_top_institutions.py /
// add_top_institutions_uk.py) are included: segments built from industry
// aggregates rather than individual institutions (e.g. Life/Non-life
// insurance), or countries not yet backfilled (currently just the US --
// its institutions file only retains a global top 10, not a per-segment
// roster), are skipped rather than guessed at.
// Kept in sync with country.html's COMMENTARY_SECTIONS -- extended
// 2026-07-23 from two to eight sections (insurance, charities, foundations,
// family offices, sovereign wealth funds and OCIO added alongside the
// original Wealth and Pensions structure). `chartSegments` mirrors the same
// field on the browser side: the canonical segment name(s) that section's
// asset-allocation chart should cover, when this country has data for them.
// Originally this export mirror didn't need chartSegments at all, since the
// donut chart was on-screen only -- added 2026-07-23 alongside
// buildSegmentAllocationChart()/commentarySectionChartSegments() below, once
// Peter asked for "everything" (including the charts) in the exported
// files too.
const COMMENTARY_SECTIONS = [
  { key: 'wealth', label: 'Wealth & key pools of capital' },
  { key: 'pensions', label: 'Pensions structure' },
  { key: 'insurance', label: 'Insurance', chartSegments: ['Life insurance', 'Non-life insurance'] },
  { key: 'charities', label: 'Charities' },
  { key: 'foundations', label: 'Foundations', chartSegments: ['Foundations E&F'] },
  { key: 'family_offices', label: 'Family offices', chartSegments: ['Family Offices (SFO)'] },
  { key: 'sovereign_wealth_funds', label: 'Sovereign wealth funds', chartSegments: ['SWF'] },
  { key: 'ocio', label: 'OCIO' }
];

// Same fixed asset-type palette as scorecard-dimensions.js's
// ASSET_TYPE_COLORS/UNREPORTED_SLICE_COLOR -- kept as a duplicate for the
// same "separate Node deployment" reason as SCORECARD_DIMENSIONS. If the
// palette ever changes on the site, update both.
const ASSET_TYPE_COLORS = {
  'Equities': '0F2540',
  'Fixed Income': '2F6F6F',
  'Cash & Short-Term': 'C98A2C',
  'Real Estate': '8A3B2F',
  'Alternatives': '5B4B8A',
  'Other/Unclassified': '9AA5AD'
};
const UNREPORTED_SLICE_COLOR = 'D6DBE0';
const ALLOCATION_PREFERRED_ORDER = ['Equities', 'Fixed Income', 'Cash & Short-Term', 'Real Estate', 'Alternatives', 'Other/Unclassified'];

function assetTypeColor(assetType) {
  return ASSET_TYPE_COLORS[assetType] || '7A8A99';
}

// Every canonical segment name in a section's chartSegments that this
// country actually has data for -- mirrors country.html's
// renderCommentaryAllocationCharts() matching logic exactly, so an export
// shows a chart for precisely the same segments the on-screen page does.
function commentarySectionChartSegments(sectionKey, segments) {
  const section = COMMENTARY_SECTIONS.find((s) => s.key === sectionKey);
  if (!section || !section.chartSegments) return [];
  return (segments || []).filter((s) => section.chartSegments.includes(s.segment));
}

// One segment's full asset-class mix, ready for either export to turn into
// its own visual -- same math as scorecard-dimensions.js's
// buildAllocationDonutSvg() (slices sized against the segment's full
// aum_bn, with an honest "Unreported" slice for the gap versus reported
// allocation, rather than rescaling the reported total up to 100%), but
// returning plain data (labels/pct/color) instead of an SVG string, since
// docx and pptxgenjs each need to build their own visual from it (a
// proportional table for Word, a native chart object for PowerPoint --
// neither speaks SVG). `sourceText` is the segment's own active source
// citation (sources[].active), the same attribution shown under each donut
// on country.html -- added 2026-07-23 after Peter pointed out the first cut
// of the chart had no sourcing at all.
function buildSegmentAllocationChart(segment) {
  const aum = segment.aum_bn || 0;
  const entries = (segment.allocation || [])
    .map((a) => ({ asset_type: a.asset_type, value_bn: a.value_bn || 0 }))
    .filter((a) => a.value_bn > 0)
    .sort((a, b) => {
      const ia = ALLOCATION_PREFERRED_ORDER.indexOf(a.asset_type), ib = ALLOCATION_PREFERRED_ORDER.indexOf(b.asset_type);
      return (ia === -1 ? ALLOCATION_PREFERRED_ORDER.length : ia) - (ib === -1 ? ALLOCATION_PREFERRED_ORDER.length : ib);
    });

  const reportedTotal = entries.reduce((s, e) => s + e.value_bn, 0);
  const unreported = aum > reportedTotal ? aum - reportedTotal : 0;

  const rawSlices = entries.map((e) => ({ label: e.asset_type, value_bn: e.value_bn, color: assetTypeColor(e.asset_type) }));
  if (unreported > 0.0001) rawSlices.push({ label: 'Unreported', value_bn: unreported, color: UNREPORTED_SLICE_COLOR });

  if (!aum || !rawSlices.length) return null;

  return {
    segmentName: segment.segment,
    aum_bn: aum,
    slices: rawSlices.map((s) => ({ label: s.label, pct: Math.round((s.value_bn / aum) * 1000) / 10, color: s.color })),
    sourceText: segmentSourceCitation(segment)
  };
}

// A segment's own active-source citation ("Basis" column / donut-chart
// attribution), extracted as its own function 2026-07-29 so it can be reused
// both by buildSegmentAllocationChart() above (its chart caption used to
// print this inline) and by buildCountrySourcesData() below (which now
// consolidates every segment's citation onto one end-of-country Sources
// page/slide instead of leaving it scattered across chart captions and the
// AUM table) -- see country.html's segmentAllocationSourceText() for the
// on-screen equivalent. Returns '' for a segment with no sources at all.
function segmentSourceCitation(segment) {
  const sources = segment.sources || [];
  const active = sources.find((s) => s.active) || sources[0];
  if (!active) return '';
  const citation = active.source || active.basis || '';
  return citation ? citation + (active.as_of ? ` (as of ${active.as_of})` : '') : '';
}

// Every source reference a country's export currently scatters across the
// document/deck -- each commentary section's own citations, plus each
// segment's AUM/allocation citation (the same one shown in the AUM table's
// "Basis" column and under each allocation chart) -- gathered into one place
// so both exports can render a single consolidated "Sources" page/slide at
// the end of that country's content instead of interrupting the text and
// charts with citation lists (Peter's 2026-07-29 "sources bleed into the
// text" feedback). `commentarySections` is buildCommentarySectionsFull()'s
// output; `segments` is the country's full segment list (same one
// buildAumRows() sorts from), so every segment gets a citation line even if
// it has no chart. Empty groups (no sources at all) are omitted so the
// caller can skip the whole page/slide when there's nothing to show.
function buildCountrySourcesData(commentarySections, segments) {
  const commentaryGroups = (commentarySections || [])
    .filter((s) => s.sources && s.sources.length)
    .map((s) => ({ sectionLabel: s.label, sources: s.sources }));

  const segmentSources = (segments || [])
    .slice()
    .sort((a, b) => (b.aum_bn || 0) - (a.aum_bn || 0))
    .map((s) => ({ segmentName: s.segment, citation: segmentSourceCitation(s) }))
    .filter((s) => s.citation);

  return { commentaryGroups, segmentSources };
}

// Turns a country's `commentary` object ({wealth: {text, sources[]}, pensions:
// {text, sources[]}}) into the shape both exports render from: one entry per
// section that actually has text, with the text split into paragraphs on
// blank lines (the same convention the picker.html/country.html edit
// textarea uses) and sources filtered down to ones with a label or url.
// Added 2026-07-23 alongside the commentary feature itself, so exports show
// the same wealth/pensions writeup as the country page rather than leaving
// it as an on-screen-only feature. Text is treated as plain paragraphs, not
// full markdown -- unlike the on-screen render (which uses marked.parse()),
// neither docx nor pptxgenjs has a markdown renderer on hand, and the
// framework's own drafted content (see uk.json) is plain prose with no
// markdown syntax, so this is a deliberate scope limit rather than an
// oversight: bullet/heading syntax typed into the edit textarea will show up
// as literal characters in an exported file, not rendered formatting.
function buildCommentarySections(commentary) {
  if (!commentary) return [];
  return COMMENTARY_SECTIONS
    .map(({ key, label }) => {
      const entry = commentary[key];
      const text = entry && typeof entry.text === 'string' ? entry.text.trim() : '';
      if (!text) return null;
      const paragraphs = text.split(/\n{2,}/).map((p) => p.trim()).filter(Boolean);
      const sources = (entry && Array.isArray(entry.sources) ? entry.sources : []).filter((s) => s && (s.label || s.url));
      return { key, label, paragraphs, sources };
    })
    .filter(Boolean);
}

// Fixed a 2026-07-23 bug: buildCommentarySections() above only keeps a
// section if it has drafted prose text, so Insurance and Foundations (real
// chart data via chartSegments, but no written text yet for most countries)
// were silently dropped from both exports entirely -- charts included --
// even though country.html renders a section's chart regardless of whether
// its prose exists (renderCommentarySection() always runs for all eight
// sections; only the text and the chart each independently decide whether
// they have anything to show). This is the export-side equivalent: a
// section survives here if it has EITHER prose OR at least one matching
// chart segment, so Insurance/Foundations/Sovereign wealth funds still
// produce their asset-allocation chart even with an empty writeup, the same
// as on screen. `segments` is this country's segment array (already passed
// into both exports); `chartSegments` on the returned object is the actual
// matched segment data (not just names), ready for
// buildSegmentAllocationChart().
// `developments` -- added 2026-07-29 alongside the two-column commentary
// layout (see country.html's renderCommentaryDevelopmentsBox()) -- is the
// optional {sectionKey: [{id, date, headline, summary, source, url}]} map
// from {code}_developments.json, sent by the client the same way commentary
// already is (country.html/picker.html read it separately from the main
// country file, same split-file reasoning as scores). A section now also
// survives if it has developments logged, even with no prose and no chart --
// matches the on-screen page, where every section shows *something* in all
// three of its text/chart/developments slots rather than only appearing when
// text or a chart happens to exist.
function buildCommentarySectionsFull(commentary, segments, developments) {
  return COMMENTARY_SECTIONS
    .map(({ key, label, chartSegments }) => {
      const entry = commentary && commentary[key];
      const text = entry && typeof entry.text === 'string' ? entry.text.trim() : '';
      const paragraphs = text ? text.split(/\n{2,}/).map((p) => p.trim()).filter(Boolean) : [];
      const sources = (entry && Array.isArray(entry.sources) ? entry.sources : []).filter((s) => s && (s.label || s.url));
      const matchedSegments = chartSegments ? (segments || []).filter((s) => chartSegments.includes(s.segment)) : [];
      const devItems = (developments && Array.isArray(developments[key]) ? developments[key] : [])
        .slice()
        .sort((a, b) => (b.date || '').localeCompare(a.date || ''));

      if (!paragraphs.length && !matchedSegments.length && !devItems.length) return null;
      return { key, label, paragraphs, sources, chartSegments: matchedSegments, hasChartSlot: !!chartSegments, developments: devItems };
    })
    .filter(Boolean);
}

function buildTopInstitutionsSections(segments) {
  return (segments || [])
    .filter((s) => s.concentration && Array.isArray(s.concentration.top_institutions) && s.concentration.top_institutions.length)
    .slice()
    .sort((a, b) => segmentSortIndex(a.segment) - segmentSortIndex(b.segment))
    .map((s) => ({
      segment: s.segment,
      top10_share_pct: s.concentration.top10_share_pct,
      n_institutions: s.concentration.n_institutions,
      institutions: s.concentration.top_institutions
    }));
}

// Shared brand palette for the Word/PowerPoint exports, added 2026-08-07
// alongside the cover page/TOC/methodology redesign -- pulled straight from
// atlas-site/style.css's :root custom properties (--navy/--navy-light/
// --accent) so the exported files read as the same visual identity as the
// site, not a separate improvised palette. Hex values without a leading '#'
// throughout (both docx's ShadingType.CLEAR fill and pptxgenjs's color
// options take bare hex), matching the convention scoreColor()/overallColor()
// already use above.
const BRAND = {
  navy: '0F2540',
  navyLight: '16344F',
  teal: '2F6F6F',
  tealTint: 'E4F0EE',   // matches style.css's .badge.bottom-up background
  navyTint: 'E8ECF2',   // a light navy tint for header rows/panels, replacing the old generic D9E2F3 blue
  ink: '222222',
  body: '333333',
  muted: '5B6B7A',
  hairline: 'D9DEE3',
  panelBg: 'F5F7F9',
  rowStripe: 'F7F9FA'
};

// One shared source of truth for the "Data structure & methodology" page/
// slide that now appears once, standardly, in every export -- condensed from
// atlas-site/data/sources.md and collateral/Atlas_Methodology_Reference.docx
// rather than written fresh, so this matches the fuller in-app/reference
// explanation instead of drifting from it. Kept as plain {heading, body}
// pairs (not docx Paragraphs or pptxgenjs text runs) so both exports can
// render it in their own native way from the same words -- added 2026-08-07
// per Peter's "a standard page that summarises the structure of the data and
// the key data sources and methodology" request.
const METHODOLOGY_CONTENT = {
  intro: 'Atlas tracks the institutional asset-owner landscape across every country covered, broken into 19 standard segments and scored on a consistent 12-factor opportunity scorecard, so figures and scores compare cleanly market to market.',
  sections: [
    {
      heading: 'Segment taxonomy',
      body: 'Every country uses the same 19 segments: DB and DC pensions (corporate, government, union, healthcare non-profit, endowment and tax-exempt sponsors), Endowments E&F, Tax exempt E&F, Foundations E&F, Healthcare non-profit E&F, Life insurance, Non-life insurance, and SWF.'
    },
    {
      heading: 'How an AUM figure is sourced',
      body: 'Each segment’s headline AUM is drawn from whichever of three routes is judged most reliable for that specific segment, not necessarily the largest: bottom-up (a mechanical sum of named institution records, primarily S&P’s Money Manager Database), top-down override (a published, authoritative figure substituted in where it is judged more reliable than the bottom-up sum — e.g. UK DB Pension (Govt), aligned to official LGPS statistics), or industry aggregate (a sector-wide published figure, used where individual institution records are not practical to collect — currently insurance). Every segment carries its full set of candidate figures, each with its own as-of date, not just the one shown.'
    },
    {
      heading: 'The opportunity scorecard',
      body: 'Every segment, in every country, is scored 1–3 against 12 standard factors — market opportunity, outsourced management, pricing impact, alignment of investment thinking, distribution resources required, regulatory complexity, client servicing, local presence required, languages required, investor decision-making, comingled vehicles, and consultant reliance. Scores combine into one weighted Overall figure (market opportunity counts 3x by default; every other factor 1x) and are colour-banded red / amber / green. Weights can be changed per project, and up to 8 custom factors added alongside the standard 12.'
    },
    {
      heading: 'Coverage & confidence',
      body: 'Figures are tracked totals within Atlas’s own data, not definitive market totals, unless a segment is specifically reconciled to an official top-down source. Non-equity allocation figures are estimated from S&P’s allocation export, scaled to each segment’s already-validated AUM — read relative proportions as more reliable than any single figure to two decimal places. Every commentary claim and every segment figure carries a traceable source.'
    }
  ],
  footer: 'Full methodology and every underlying source: Atlas — Sources & Methodology (atlas.institutionaladviser.co.uk).'
};

// Structural preview of one country's export content -- used only to build
// the table of contents (see addTocSlides() in exportPptx.js / the native TOC
// field in exportDocx.js) and, for PowerPoint, to work out how many slides
// each block will actually consume before any slide is drawn, since
// pptxgenjs has no auto-pagination and TOC entries need a real slide number
// to jump to. Every count here mirrors the real slide-building logic exactly
// (buildCommentarySectionsFull() for commentary, buildTopInstitutionsSections()
// chunked by topInstCols, buildCountrySourcesData()'s line count chunked by
// sourcesLinesPerSlide) rather than re-deriving it a different way, so the
// plan can never drift out of sync with what actually gets built. Returns []
// blocks the caller doesn't ask for via `include`, same convention as
// buildCountrySection()/addCountrySlides().
function planCountryOutline(countryName, segments, commentary, developments, includeSet, opts = {}) {
  const topInstCols = opts.topInstCols || 3;
  const sourcesLinesPerSlide = opts.sourcesLinesPerSlide || 26;
  const scorecardColsPerSlide = opts.scorecardColsPerSlide || 9;

  const commentarySections = includeSet.has('commentary') ? buildCommentarySectionsFull(commentary, segments, developments) : [];
  const blocks = [];

  if (includeSet.has('commentary') && commentarySections.length) {
    blocks.push({ type: 'commentary', label: 'Country commentary', slideCount: commentarySections.length });
  }
  if (includeSet.has('aum') && (segments || []).length) {
    blocks.push({ type: 'aum', label: 'AUM by segment', slideCount: 1 });
  }
  if (includeSet.has('scorecard') && (segments || []).length) {
    blocks.push({ type: 'scorecard', label: 'Opportunity scorecard', slideCount: Math.max(1, Math.ceil(segments.length / scorecardColsPerSlide)) });
  }
  if (includeSet.has('top_institutions')) {
    const topSections = buildTopInstitutionsSections(segments);
    if (topSections.length) blocks.push({ type: 'top_institutions', label: 'Top institutions by AUM', slideCount: Math.max(1, Math.ceil(topSections.length / topInstCols)) });
  }
  if (includeSet.has('commentary') || includeSet.has('aum')) {
    const sourcesData = buildCountrySourcesData(commentarySections, segments);
    const lineCount = sourcesData.commentaryGroups.reduce((n, g) => n + 1 + g.sources.length, 0)
      + (sourcesData.segmentSources.length ? 1 + sourcesData.segmentSources.length : 0);
    if (lineCount) blocks.push({ type: 'sources', label: 'Sources', slideCount: Math.max(1, Math.ceil(lineCount / sourcesLinesPerSlide)) });
  }

  return { countryName, blocks, commentarySections };
}

module.exports = {
  SCORECARD_DIMENSIONS,
  CANONICAL_SEGMENT_ORDER,
  segmentSortIndex,
  isDimensionEnabled,
  enabledDimensionCount,
  computeOverallScore,
  computeOverallRange,
  scoredDimensionCount,
  scoreColor,
  overallColor,
  estimateColumnCharWidths,
  buildAumRows,
  buildScorecardMatrix,
  buildCommentarySections,
  buildCommentarySectionsFull,
  commentarySectionChartSegments,
  buildSegmentAllocationChart,
  segmentSourceCitation,
  buildCountrySourcesData,
  buildTopInstitutionsSections,
  BRAND,
  METHODOLOGY_CONTENT,
  planCountryOutline
};
