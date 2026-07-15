// Server-side mirror of atlas-site/scorecard-dimensions.js. Kept as a
// deliberate duplicate rather than a shared import because the site's copy
// is a plain browser <script> (no module system) and the API is a separate
// Node deployment (api_location) with no access to atlas-site's files at
// build time. If the 12 dimensions, weights, or canonical segment order
// ever change, update both files -- this one and atlas-site/scorecard-dimensions.js.
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

// Kept in sync with the same functions in scorecard-dimensions.js (see the
// comment there for the reasoning) -- enabledDimensions is an optional
// {dimensionKey: boolean} map sourced from global.json's enabled_dimensions
// field; a dimension counts as enabled unless it's explicitly false.
function isDimensionEnabled(dimKey, enabledDimensions) {
  return !enabledDimensions || enabledDimensions[dimKey] !== false;
}

function enabledDimensionCount(enabledDimensions) {
  return SCORECARD_DIMENSIONS.filter((d) => isDimensionEnabled(d.key, enabledDimensions)).length;
}

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

function scoredDimensionCount(scorecard, enabledDimensions) {
  if (!scorecard) return 0;
  return SCORECARD_DIMENSIONS.filter((d) => isDimensionEnabled(d.key, enabledDimensions) && typeof scorecard[d.key] === 'number').length;
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

// Overall score thresholds match style.css's comment: range 14-42,
// <=22 red, 23-30 amber, >=31 green.
function overallColor(value) {
  if (typeof value !== 'number') return null;
  if (value <= 22) return { bg: 'FBE1E1', fg: 'A3291F' };
  if (value <= 30) return { bg: 'FDEEE0', fg: '9A5A1A' };
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
// disabled dimensions still get a row (so the export shows the full set,
// same as the on-screen matrix), but are excluded from the Scored
// denominator and the Overall sum, matching computeOverallScore() above.
function buildScorecardMatrix(segments, enabledDimensions) {
  const cols = (segments || []).slice().sort((a, b) => segmentSortIndex(a.segment) - segmentSortIndex(b.segment));
  const enabledCount = enabledDimensionCount(enabledDimensions);

  const dimensionRows = SCORECARD_DIMENSIONS.map((dim) => ({
    key: dim.key,
    type: 'dimension',
    label: dim.label + (dim.weight > 1 ? ` (x${dim.weight})` : '') + (isDimensionEnabled(dim.key, enabledDimensions) ? '' : ' (off)'),
    values: cols.map((s) => {
      const v = s.scorecard ? s.scorecard[dim.key] : undefined;
      return typeof v === 'number' ? String(v) : '-';
    }),
    colors: cols.map((s) => {
      const v = s.scorecard ? s.scorecard[dim.key] : undefined;
      return typeof v === 'number' ? scoreColor(v) : null;
    })
  }));

  const scoredRow = {
    type: 'scored',
    label: 'Scored',
    values: cols.map((s) => `${scoredDimensionCount(s.scorecard, enabledDimensions)}/${enabledCount}`),
    colors: cols.map(() => null)
  };

  const overallRow = {
    type: 'overall',
    label: 'Overall',
    values: cols.map((s) => {
      const overall = computeOverallScore(s.scorecard, enabledDimensions);
      return overall === null ? '-' : String(overall);
    }),
    colors: cols.map((s) => overallColor(computeOverallScore(s.scorecard, enabledDimensions)))
  };

  const aumRow = {
    type: 'aum',
    label: 'AUM ($bn)',
    values: cols.map((s) => (typeof s.aum_bn === 'number' ? s.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-')),
    colors: cols.map(() => null)
  };

  return {
    columnLabels: cols.map((s) => s.segment),
    rows: [aumRow, ...dimensionRows, scoredRow, overallRow]
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

module.exports = {
  SCORECARD_DIMENSIONS,
  CANONICAL_SEGMENT_ORDER,
  segmentSortIndex,
  isDimensionEnabled,
  enabledDimensionCount,
  computeOverallScore,
  scoredDimensionCount,
  scoreColor,
  overallColor,
  estimateColumnCharWidths,
  buildAumRows,
  buildScorecardMatrix,
  buildTopInstitutionsSections
};
