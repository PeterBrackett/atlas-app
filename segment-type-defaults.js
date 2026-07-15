// Starting-point default scores (1-3) for three scorecard dimensions that
// have no usable external index the way TMF/EF EPI/FDIRRI/Mercer/IC
// Research do for the others: "Outsourced management", "Comingled
// vehicles" and "Investor decision-making". These are judgment calls
// dressed up as a lookup table, not discovered facts -- they're typical
// industry patterns by SEGMENT TYPE (same value for e.g. every "DC Pension
// (Corp)" segment worldwide, regardless of country), not anything derived
// from Atlas's own data. Meant as a starting point to get every segment to
// a non-null Overall score, then be corrected by hand where a specific
// country/segment is known to differ. See applyOutsourcedManagementDefaults()
// etc. in overview.html -- every score here is editable in the tool's input
// boxes before it's applied, so these constants are just the pre-filled
// starting values, not a locked-in answer.
//
// Reasoning behind each column (all defensible industry generalizations,
// not hard data):
// - Outsourced management (3 = mostly outsourced to external managers, 1 =
//   often run in-house): DC arrangements are almost always fully
//   outsourced via fund platforms. Endowments & Foundations overwhelmingly
//   use external managers (well-documented industry norm). Insurers and
//   sovereign wealth funds more often carry meaningful in-house investment
//   teams, especially for liability-matching fixed income.
// - Comingled vehicles (3 = existing pooled funds work, 1 = bespoke/
//   segregated mandates typical): DC platforms are pooled-fund-native.
//   E&F segments commonly invest via commingled funds/LPs. Insurance
//   mandates often need bespoke segregated accounts for regulatory/
//   liability-matching reasons (e.g. Solvency II). Large SWFs typically
//   negotiate bespoke mandates given their scale.
// - Investor decision-making (3 = faster/simpler process, 1 = slow/
//   committee-heavy): government and union pension boards and large
//   insurer/SWF bureaucracies tend to involve multiple sign-offs and
//   political/committee processes; smaller foundations can move faster.
const SEGMENT_TYPE_DEFAULT_SCORES = {
  outsourced_management: {
    'DB Pension (Corp)': 2,
    'DC Pension (Corp)': 3,
    'DB Pension (Govt)': 2,
    'DC Pension (Govt)': 3,
    'DB Pension (Union)': 2,
    'DC Pension (Union)': 3,
    'DB Pension (Healthcare non-profit)': 3,
    'DC Pension (Healthcare non-profit)': 3,
    'DB Pension (Endowments)': 3,
    'DC Pension (Endowments)': 3,
    'DB Pension (Tax exempt)': 3,
    'DC Pension (Tax exempt)': 3,
    'Endowments E&F': 3,
    'Tax exempt E&F': 3,
    'Foundations E&F': 3,
    'Healthcare non-profit E&F': 3,
    'Life insurance': 1,
    'Non-life insurance': 1,
    'SWF': 1
  },
  comingled_vehicles: {
    'DB Pension (Corp)': 2,
    'DC Pension (Corp)': 3,
    'DB Pension (Govt)': 2,
    'DC Pension (Govt)': 3,
    'DB Pension (Union)': 2,
    'DC Pension (Union)': 3,
    'DB Pension (Healthcare non-profit)': 2,
    'DC Pension (Healthcare non-profit)': 3,
    'DB Pension (Endowments)': 2,
    'DC Pension (Endowments)': 3,
    'DB Pension (Tax exempt)': 2,
    'DC Pension (Tax exempt)': 3,
    'Endowments E&F': 3,
    'Tax exempt E&F': 3,
    'Foundations E&F': 3,
    'Healthcare non-profit E&F': 3,
    'Life insurance': 1,
    'Non-life insurance': 1,
    'SWF': 1
  },
  investor_decision_making: {
    'DB Pension (Corp)': 2,
    'DC Pension (Corp)': 2,
    'DB Pension (Govt)': 1,
    'DC Pension (Govt)': 2,
    'DB Pension (Union)': 1,
    'DC Pension (Union)': 2,
    'DB Pension (Healthcare non-profit)': 2,
    'DC Pension (Healthcare non-profit)': 2,
    'DB Pension (Endowments)': 2,
    'DC Pension (Endowments)': 2,
    'DB Pension (Tax exempt)': 2,
    'DC Pension (Tax exempt)': 2,
    'Endowments E&F': 2,
    'Tax exempt E&F': 2,
    'Foundations E&F': 3,
    'Healthcare non-profit E&F': 2,
    'Life insurance': 1,
    'Non-life insurance': 1,
    'SWF': 1
  }
};

// Segment types this file has a default for, in the same fixed display
// order as CANONICAL_SEGMENT_ORDER (scorecard-dimensions.js), so the tool
// panel table renders in a predictable order rather than object-key order.
const SEGMENT_TYPE_ORDER = [
  'DB Pension (Corp)', 'DC Pension (Corp)', 'DB Pension (Govt)', 'DC Pension (Govt)',
  'DB Pension (Union)', 'DC Pension (Union)',
  'DB Pension (Healthcare non-profit)', 'DC Pension (Healthcare non-profit)',
  'DB Pension (Endowments)', 'DC Pension (Endowments)',
  'DB Pension (Tax exempt)', 'DC Pension (Tax exempt)',
  'Endowments E&F', 'Tax exempt E&F', 'Foundations E&F', 'Healthcare non-profit E&F',
  'Life insurance', 'Non-life insurance', 'SWF'
];

function segmentTypeDefaultScore(dimensionKey, segmentName) {
  const table = SEGMENT_TYPE_DEFAULT_SCORES[dimensionKey];
  if (!table) return null;
  const score = table[segmentName];
  return typeof score === 'number' ? score : null;
}
