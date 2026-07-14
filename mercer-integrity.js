// Mercer CFA Institute Global Pension Index -- "Integrity" sub-index, an
// external reference used to auto-score the "Client servicing" scorecard
// dimension. The Index scores each system's pension framework across three
// pillars (Adequacy, Sustainability, Integrity); Integrity covers
// regulation, governance and communication -- the closest available public,
// annually-updated proxy for how rigorous a market's reporting/governance
// expectations are, which is the essence of what "Client servicing" is
// trying to capture. Runs 0-100, higher = stronger governance/regulation.
//
// Note: Peter originally flagged this index for both "Client servicing" and
// "Investor decision-making". Only wired up for Client servicing here --
// using the same number for both dimensions would make them move in
// lockstep and defeat the point of scoring them separately. Investor
// decision-making is left as a manual judgment call for now.
//
// Source: Mercer CFA Institute Global Pension Index 2025, Figure 4 ("Index
// score for each system"), main report PDF, transcribed 2026-07-14 (same
// manual-transcription pattern as tmf-rankings.js -- Mercer publishes a new
// edition annually; when that happens, update MERCER_INTEGRITY_YEAR and
// re-transcribe Figure 4's Integrity column).
const MERCER_INTEGRITY_YEAR = 2025;

// Integrity sub-index scores (0-100) for the Atlas countries the Index
// covers. Only Atlas's covered countries are listed here (unlike the TMF
// reference file) -- the Index covers 52 systems and Atlas only needs 33
// of them.
const MERCER_INTEGRITY_SCORES = {
  'GB': 79.0,
  'US': 58.0,
  'DE': 75.0,
  'NL': 86.8,
  'NO': 88.4,
  'CH': 81.6,
  'DK': 77.6,
  'SE': 83.0,
  'ZA': 75.7,
  'IE': 81.8,
  'FI': 90.6,
  'IS': 83.3,
  'CA': 80.2,
  'FR': 76.8,
  'IT': 77.8,
  'ES': 74.4,
  'AU': 86.4,
  'AT': 76.4,
  'JP': 66.8,
  'HK': 89.2,
  'SG': 90.4,
  'NZ': 81.7,
  'KR': 76.8,
  'SA': 74.2,
  'AE': 75.5,
  'OM': 71.7,
  'BE': 86.8,
  'AR': 42.4,
  'PE': 64.8,
  'CL': 86.6,
  'BW': 85.0,
  'BR': 67.3,
  'MY': 77.5
};

// Qatar is the one Atlas country the Index doesn't cover at all (52 systems
// benchmarked in 2025; Qatar isn't among them). Skipped, not guessed at.
const MERCER_INTEGRITY_UNCOVERED_CODES = ['QA'];

function mercerIntegrityScoreForCode(code) {
  const score = MERCER_INTEGRITY_SCORES[(code || '').toUpperCase()];
  return typeof score === 'number' ? score : null;
}
