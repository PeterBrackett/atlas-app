const { app } = require('@azure/functions');
const PptxGenJS = require('pptxgenjs');
const {
  buildAumRows, buildScorecardMatrix, buildCommentarySectionsFull, buildTopInstitutionsSections,
  estimateColumnCharWidths, buildSegmentAllocationChart, buildCountrySourcesData,
  BRAND, METHODOLOGY_CONTENT, planCountryOutline
} = require('../shared/exportHelpers');
const { getAtlasLogoDataUri, getAtlasLockupWhiteDataUri } = require('../shared/atlasLogo');

// ---------------------------------------------------------------------------
// Brand system -- 2026-08-07 rebuild. Peter's feedback was three things at
// once: (1) the exports were "far from a place I could share with a client",
// (2) too much dead white space on individual slides, (3) no cover page, no
// table of contents, no standard page explaining how the data is put
// together. This file addresses all three: a proper cover + TOC +
// methodology front section (see addCoverSlide/addTocSlides/
// addMethodologySlide below), a still-small-but-branded slide master (navy
// rail down the left edge, footer rule + page number, replacing the bare
// logo-only master), and per-slide-type layout fixes for the worst
// whitespace offenders found by actually rendering a sample deck and looking
// at it (LibreOffice is available in the build sandbox now, unlike when this
// file was first written) -- see the comments on addAumSlide/
// addScorecardSlide/addCountrySourcesSlides for what was actually wrong.
const C = BRAND;
const HEADER_FILL = C.navyTint;
const BORDER = { type: 'solid', color: C.hairline, pt: 0.5 };
const SLIDE_W = 13.33;
const SLIDE_H = 7.5;

// Tight cell padding (inches: [top, right, bottom, left]) applied to every
// table in this export -- pptxgenjs's default cell margin left noticeably
// more whitespace around short numeric/score cells than the content needed,
// which was Peter's original "quite a bit of wasted space" feedback.
const CELL_MARGIN = [0.03, 0.05, 0.03, 0.05];

// Converts an estimateColumnCharWidths() character count into an inch-wide
// pptxgenjs colW entry, so tables are sized to their content instead of
// stretched across the slide.
const CHAR_WIDTH_IN = 0.078;
function charsToInches(chars) {
  return Math.round((chars * CHAR_WIDTH_IN + CELL_MARGIN[1] + CELL_MARGIN[3]) * 100) / 100;
}

// Layout constants for the branded slide master. LEFT_BAR is a thin navy
// rail down the left edge of every content slide (the one consistent brand
// element every slide type shares, cover included in spirit); the logo sits
// just to the right of it, and titles/subtitles clear both.
const LEFT_BAR_W = 0.14;
const LOGO_X = 0.32;
const LOGO_Y = 0.24;
const LOGO_SIZE = 0.32;
const TITLE_X = 0.78;
const TITLE_Y = 0.28;
const SUBTITLE_Y = 0.82;
const CONTENT_X = 0.5;
const CONTENT_RIGHT_MARGIN = 0.45;
const CONTENT_W = SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN;
const FOOTER_RULE_Y = SLIDE_H - 0.42;
const ATLAS_MASTER_NAME = 'ATLAS_MASTER';

function defineAtlasMaster(pptx) {
  const logoDataUri = getAtlasLogoDataUri();
  pptx.defineSlideMaster({
    title: ATLAS_MASTER_NAME,
    background: { color: 'FFFFFF' },
    objects: [
      { rect: { x: 0, y: 0, w: LEFT_BAR_W, h: SLIDE_H, fill: { color: C.navy }, line: { type: 'none' } } },
      ...(logoDataUri ? [{ image: { x: LOGO_X, y: LOGO_Y, w: LOGO_SIZE, h: LOGO_SIZE, data: logoDataUri } }] : []),
      { line: { x: 0.5, y: FOOTER_RULE_Y, w: SLIDE_W - 1.0, h: 0, line: { color: C.hairline, width: 0.75 } } },
      { text: { text: 'Atlas — Institutional Adviser  |  Confidential', options: { x: 0.5, y: FOOTER_RULE_Y + 0.06, w: 7, h: 0.25, fontSize: 8, color: C.muted } } }
    ],
    slideNumber: { x: SLIDE_W - 1.1, y: FOOTER_RULE_Y + 0.06, w: 0.7, h: 0.25, fontSize: 8, color: C.muted, align: 'right' }
  });
}

function addAtlasSlide(pptx) {
  return pptx.addSlide({ masterName: ATLAS_MASTER_NAME });
}

// Standard content-slide header: big navy title (usually "Atlas — {country}")
// plus a smaller muted subtitle line naming the block ("Opportunity
// scorecard", "AUM by segment", etc.) -- every content slide in this file
// calls this instead of building its own two addText() calls, so the title
// treatment can never quietly drift between slide types.
function addSlideHeader(slide, title, subtitle) {
  slide.addText(title, { x: TITLE_X, y: TITLE_Y, w: SLIDE_W - TITLE_X - CONTENT_RIGHT_MARGIN, h: 0.5, fontSize: 22, bold: true, color: C.navy });
  if (subtitle) {
    slide.addText(subtitle, { x: TITLE_X, y: SUBTITLE_Y, w: SLIDE_W - TITLE_X - CONTENT_RIGHT_MARGIN, h: 0.32, fontSize: 12, color: C.muted });
  }
}

function headerCellOpts(extra) {
  return Object.assign({ bold: true, fill: { color: HEADER_FILL }, fontSize: 8, color: C.navy, margin: CELL_MARGIN }, extra);
}

// ---------------------------------------------------------------------------
// Cover slide -- new 2026-08-07. Full navy background (the one slide in the
// deck that isn't on ATLAS_MASTER, since the master's white background/left
// rail/logo treatment is specifically the *content*-slide look) with the
// white/teal reversed lockup (icon + "ATLAS" wordmark -- see atlasLogo.js),
// the report title, a one-line coverage summary, and a generated-date/
// confidentiality footer. `title`/`subtitle`/`coverageLine` are plain
// strings built by the handler below from the request payload (single
// country vs. multi-country project).
function addCoverSlide(pptx, { title, subtitle, coverageLine, generatedDate }) {
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };

  const lockupUri = getAtlasLockupWhiteDataUri();
  const lockupW = 3.6;
  const lockupH = lockupW / 3.0017; // source asset is 1801x600
  if (lockupUri) slide.addImage({ data: lockupUri, x: 0.9, y: 0.85, w: lockupW, h: lockupH });

  slide.addShape('rect', { x: 0.92, y: 2.35, w: 1.8, h: 0.035, fill: { color: C.teal }, line: { type: 'none' } });

  slide.addText(title, {
    x: 0.9, y: 2.65, w: SLIDE_W - 1.8, h: 1.4, fontSize: 32, bold: true, color: 'FFFFFF', valign: 'top'
  });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.9, y: 3.95, w: SLIDE_W - 1.8, h: 0.45, fontSize: 15, color: 'CFDAE0', valign: 'top' });
  }
  if (coverageLine) {
    slide.addText(coverageLine, { x: 0.9, y: 4.42, w: SLIDE_W - 1.8, h: 1.4, fontSize: 11.5, color: '9FB2BE', valign: 'top', lineSpacingMultiple: 1.3 });
  }

  slide.addShape('rect', { x: 0, y: SLIDE_H - 0.55, w: SLIDE_W, h: 0.0, line: { type: 'none' } });
  slide.addText(`Generated ${generatedDate}   ·   Prepared by Institutional Adviser   ·   Confidential — not for redistribution`, {
    x: 0.9, y: SLIDE_H - 0.6, w: SLIDE_W - 1.8, h: 0.35, fontSize: 9.5, color: '7E93A0'
  });
  return slide;
}

// ---------------------------------------------------------------------------
// Table of contents -- new 2026-08-07. Built from `outlineRows`, a flat list
// the handler assembles below (one 'country' header row per country/project
// section, one 'block' row per content block within it, each block row
// already carrying the real slide number it jumps to -- see
// planCountryOutline() in exportHelpers.js and the two-pass build in the
// handler for how that number is worked out before any content slide
// exists). Paginates across slides by accumulating each row's actual height
// rather than a fixed row-per-slide count, so a TOC with many short country
// sections doesn't leave a near-empty trailing slide the way the old fixed
// SOURCES_LINES_PER_SLIDE cap did for the Sources page (see
// addCountrySourcesSlides()).
const TOC_TOP_Y = 1.35;
const TOC_BOTTOM_Y = SLIDE_H - 0.55;
const TOC_COUNTRY_ROW_H = 0.42;
const TOC_BLOCK_ROW_H = 0.32;
const TOC_LABEL_X = 0.85;
const TOC_PAGE_X = SLIDE_W - 1.1;

function addOneTocSlide(pptx, rows, pageLabel) {
  const slide = addAtlasSlide(pptx);
  addSlideHeader(slide, 'Contents', pageLabel);
  let y = TOC_TOP_Y;
  rows.forEach((row) => {
    if (row.kind === 'country') {
      slide.addText(row.text, { x: CONTENT_X, y, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, h: TOC_COUNTRY_ROW_H, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
      y += TOC_COUNTRY_ROW_H;
    } else {
      const hyperlink = row.slideNumber ? { slide: row.slideNumber } : undefined;
      slide.addText(row.text, { x: TOC_LABEL_X, y, w: SLIDE_W - TOC_LABEL_X - 1.3, h: TOC_BLOCK_ROW_H, fontSize: 11, color: C.body, valign: 'middle', hyperlink });
      if (row.slideNumber) {
        slide.addText(String(row.slideNumber), { x: TOC_PAGE_X, y, w: 0.6, h: TOC_BLOCK_ROW_H, fontSize: 11, color: C.muted, align: 'right', valign: 'middle', hyperlink });
      }
      y += TOC_BLOCK_ROW_H;
    }
  });
}

function addTocSlides(pptx, outlineRows) {
  if (!outlineRows.length) return;
  const pages = [];
  let current = [];
  let y = TOC_TOP_Y;
  outlineRows.forEach((row) => {
    const rowH = row.kind === 'country' ? TOC_COUNTRY_ROW_H : TOC_BLOCK_ROW_H;
    if (y + rowH > TOC_BOTTOM_Y && current.length) {
      pages.push(current);
      current = [];
      y = TOC_TOP_Y;
    }
    current.push(row);
    y += rowH;
  });
  if (current.length) pages.push(current);

  pages.forEach((rows, i) => {
    addOneTocSlide(pptx, rows, pages.length > 1 ? `Page ${i + 1} of ${pages.length}` : undefined);
  });
  return pages.length;
}

// ---------------------------------------------------------------------------
// Methodology slide -- new 2026-08-07, Peter's "a standard page that
// summarises the structure of the data and the key data sources and
// methodology" request. Content comes entirely from
// exportHelpers.js's METHODOLOGY_CONTENT (condensed from
// atlas-site/data/sources.md and the Atlas_Methodology_Reference.docx
// collateral), not written fresh here, so it can't drift from the fuller
// reference. Laid out as a 2x2 card grid rather than a single scrolling text
// block -- with four roughly-equal sections this fills the slide evenly
// instead of leaving the same kind of dead space below a single tall text
// box that the old commentary-only sections had.
const METH_CARD_GUTTER = 0.3;
const METH_TOP_Y = 1.75;
const METH_CARD_W = (SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN - METH_CARD_GUTTER) / 2;
const METH_CARD_H = 2.35;

function addMethodologySlide(pptx) {
  const slide = addAtlasSlide(pptx);
  addSlideHeader(slide, 'Data structure & methodology', 'How Atlas figures are built, and what they do and don’t mean');
  slide.addText(METHODOLOGY_CONTENT.intro, {
    x: CONTENT_X, y: 1.15, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, h: 0.5, fontSize: 11, color: C.body, valign: 'top'
  });

  METHODOLOGY_CONTENT.sections.forEach((sec, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = CONTENT_X + col * (METH_CARD_W + METH_CARD_GUTTER);
    const y = METH_TOP_Y + row * (METH_CARD_H + 0.2);
    slide.addShape('roundRect', { x, y, w: METH_CARD_W, h: METH_CARD_H, rectRadius: 0.06, fill: { color: C.panelBg }, line: { color: C.hairline, width: 0.75 } });
    slide.addText(sec.heading, { x: x + 0.18, y: y + 0.14, w: METH_CARD_W - 0.36, h: 0.3, fontSize: 12.5, bold: true, color: C.teal });
    slide.addText(sec.body, { x: x + 0.18, y: y + 0.48, w: METH_CARD_W - 0.36, h: METH_CARD_H - 0.62, fontSize: 9.5, color: C.body, valign: 'top', lineSpacingMultiple: 1.15 });
  });

  slide.addText(METHODOLOGY_CONTENT.footer, {
    x: CONTENT_X, y: METH_TOP_Y + 2 * (METH_CARD_H + 0.2) + 0.02, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, h: 0.3, fontSize: 9, italic: true, color: C.muted
  });
}

// ---------------------------------------------------------------------------
// AUM by segment -- rebuilt 2026-08-07. The old version was just a table
// pinned to the top of an otherwise-empty slide (on a country with few
// segments, more than half the slide was blank). Adds a row of "at a
// glance" stat cards above the table (total AUM, segment count, largest
// segment, and a bottom-up/top-down/industry-aggregate source mix count) --
// genuinely useful summary content, not padding, and it fills the space a
// short table would otherwise leave empty.
function classifyBasis(basis) {
  const b = (basis || '').toLowerCase();
  if (b.startsWith('top-down')) return 'topDown';
  if (b.startsWith('industry aggregate')) return 'industryAgg';
  if (b.startsWith('bottom-up')) return 'bottomUp';
  return 'other';
}

function computeAumStats(rows) {
  const totalAum = rows.reduce((s, r) => s + (typeof r.aum_bn === 'number' ? r.aum_bn : 0), 0);
  const largest = rows.slice().sort((a, b) => (b.aum_bn || 0) - (a.aum_bn || 0))[0];
  const mix = { bottomUp: 0, topDown: 0, industryAgg: 0, other: 0 };
  rows.forEach((r) => { mix[classifyBasis(r.basis)] += 1; });
  const mixParts = [];
  if (mix.bottomUp) mixParts.push(`${mix.bottomUp} bottom-up`);
  if (mix.topDown) mixParts.push(`${mix.topDown} top-down`);
  if (mix.industryAgg) mixParts.push(`${mix.industryAgg} industry aggregate`);
  if (mix.other) mixParts.push(`${mix.other} other`);
  return { totalAum, count: rows.length, largest, mixText: mixParts.join(' · ') };
}

const STAT_CARD_GUTTER = 0.22;
function addStatCards(slide, cards, y, h) {
  const w = (SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN - (cards.length - 1) * STAT_CARD_GUTTER) / cards.length;
  cards.forEach((card, i) => {
    const x = CONTENT_X + i * (w + STAT_CARD_GUTTER);
    slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.05, fill: { color: C.panelBg }, line: { color: C.hairline, width: 0.75 } });
    slide.addText(card.label, { x: x + 0.14, y: y + 0.1, w: w - 0.28, h: 0.25, fontSize: 8.5, color: C.muted, bold: true });
    slide.addText(card.value, { x: x + 0.14, y: y + 0.34, w: w - 0.28, h: h - 0.5, fontSize: card.small ? 11 : 15, bold: true, color: C.navy, valign: 'top' });
  });
}

function buildAumTableRows(rows) {
  const headerLabels = ['Segment', 'AUM ($bn)', 'Equities ($bn)', 'Basis', 'Equities range (min-max)'];
  const bodyText = rows.map((r) => [
    r.segment,
    typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    typeof r.equity_bn === 'number' ? r.equity_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    r.basis || '',
    r.equity_range || '-'
  ]);
  const colW = estimateColumnCharWidths(headerLabels, bodyText, {
    minChars: 4,
    maxCharsPerCol: [26, 10, 10, 24, 30]
  }).map(charsToInches);

  const header = headerLabels.map((t) => ({ text: t, options: headerCellOpts() }));
  const body = bodyText.map((cells, ri) => cells.map((text, i) => ({
    text,
    options: { fontSize: i < 3 ? 9 : 8, margin: CELL_MARGIN, fill: { color: ri % 2 ? C.rowStripe : 'FFFFFF' } }
  })));
  return { rows: [header, ...body], colW };
}

const AUM_STAT_Y = 1.15;
const AUM_STAT_H = 0.85;
const AUM_TABLE_Y = AUM_STAT_Y + AUM_STAT_H + 0.25;

function addAumSlide(pptx, countryName, segments, generatedDate) {
  const rows = buildAumRows(segments);
  const stats = computeAumStats(rows);
  const slide = addAtlasSlide(pptx);
  addSlideHeader(slide, `Atlas — ${countryName}`, `AUM by segment — generated ${generatedDate}`);

  addStatCards(slide, [
    { label: 'TOTAL TRACKED AUM', value: `$${stats.totalAum.toLocaleString(undefined, { maximumFractionDigits: 1 })}bn` },
    { label: 'SEGMENTS', value: String(stats.count) },
    { label: 'LARGEST SEGMENT', value: stats.largest ? stats.largest.segment : '—', small: true },
    { label: 'SOURCE MIX', value: stats.mixText || '—', small: true }
  ], AUM_STAT_Y, AUM_STAT_H);

  const aumTable = buildAumTableRows(rows);
  slide.addTable(aumTable.rows, {
    x: CONTENT_X, y: AUM_TABLE_Y,
    colW: aumTable.colW,
    border: BORDER,
    autoPage: false
  });
}

// ---------------------------------------------------------------------------
// Opportunity scorecard -- rebuilt 2026-08-07. Two real fixes, not just
// styling: (1) a country with many segment columns (UK has 18) used to run
// the table straight off the right edge of the slide with no warning --
// real columns were simply invisible in the exported file. The table now
// chunks across multiple slides at SCORECARD_COLS_PER_SLIDE columns each
// ("segments 1-12 of 18" style subtitle), so nothing is ever silently cut
// off. (2) the per-dimension icon column (small images floating in the
// slide margin next to each row) is dropped -- pptxgenjs tables can't embed
// an icon inline with a cell's text the way the Word export can (see
// dimensionHeaderCell() in exportDocx.js), so this file used a separate
// addImage() per row positioned to line up with the table's fixed row
// height. It technically worked, but a floating-icon-in-the-margin approach
// reads as a UI element, not a "final deck" convention -- replaced with a
// plain text label plus a proper colour-key legend under the table, which
// is both clearer at deck-viewing distance and removes a layout that broke
// silently if the label column ever wrapped to two lines.
const SCORECARD_COLS_PER_SLIDE = 12;
const SCORECARD_ROW_H = 0.27;
const SCORECARD_TABLE_X = CONTENT_X;

function sliceScorecardMatrix(matrix, start, end) {
  return {
    columnLabels: matrix.columnLabels.slice(start, end),
    rows: matrix.rows.map((row) => ({
      ...row,
      values: row.values.slice(start, end),
      colors: row.colors ? row.colors.slice(start, end) : undefined
    }))
  };
}

function buildScorecardTableRows(matrix) {
  const headerLabels = ['Dimension', ...matrix.columnLabels];
  const bodyText = matrix.rows.map((row) => [row.label, ...row.values]);
  const maxCharsPerCol = [36, ...matrix.columnLabels.map(() => 7)];
  const colW = estimateColumnCharWidths(headerLabels, bodyText, { minChars: 3, maxCharsPerCol }).map(charsToInches);

  const header = headerLabels.map((t) => ({
    text: t,
    options: headerCellOpts({ fontSize: 8 })
  }));
  const body = matrix.rows.map((row) => ([
    { text: row.label, options: headerCellOpts({ fontSize: 8 }) },
    ...row.values.map((v, i) => {
      const color = row.colors ? row.colors[i] : null;
      const base = { fontSize: 8, align: 'center', margin: CELL_MARGIN };
      if (!color) return { text: v, options: base };
      return { text: v, options: Object.assign({}, base, { fill: { color: color.bg }, color: color.fg, bold: true }) };
    })
  ]));
  return { rows: [header, ...body], colW };
}

const SCORE_LEGEND = [
  { label: '1 — needs attention', color: { bg: 'FBE1E1', fg: 'A3291F' } },
  { label: '2 — moderate', color: { bg: 'FDEEE0', fg: '9A5A1A' } },
  { label: '3 — favourable', color: { bg: 'E3F2E3', fg: '1F7A34' } },
  { label: 'not yet scored', color: { bg: 'FFF3B0', fg: '7A5C00' } }
];

function addScoreLegend(slide, y) {
  let x = CONTENT_X;
  SCORE_LEGEND.forEach((item) => {
    slide.addShape('rect', { x, y, w: 0.16, h: 0.16, fill: { color: item.color.bg }, line: { color: item.color.fg, width: 0.75 } });
    slide.addText(item.label, { x: x + 0.22, y: y - 0.05, w: 1.7, h: 0.26, fontSize: 8, color: C.muted, valign: 'middle' });
    x += 2.0;
  });
}

function addScorecardSlide(pptx, countryName, segments, enabledDimensions, weightOverrides, allocType, allocStyle, customDimensions) {
  const matrix = buildScorecardMatrix(segments, enabledDimensions, weightOverrides, allocType, allocStyle, customDimensions);
  const totalCols = matrix.columnLabels.length;
  const chunkCount = Math.max(1, Math.ceil(totalCols / SCORECARD_COLS_PER_SLIDE));

  for (let c = 0; c < chunkCount; c++) {
    const start = c * SCORECARD_COLS_PER_SLIDE;
    const end = Math.min(totalCols, start + SCORECARD_COLS_PER_SLIDE);
    const chunk = chunkCount > 1 ? sliceScorecardMatrix(matrix, start, end) : matrix;
    const subtitle = chunkCount > 1 ? `Opportunity scorecard — segments ${start + 1}–${end} of ${totalCols}` : 'Opportunity scorecard';

    const slide = addAtlasSlide(pptx);
    addSlideHeader(slide, `Atlas — ${countryName}`, subtitle);
    const tableY = 1.3;
    const table = buildScorecardTableRows(chunk);
    slide.addTable(table.rows, {
      x: SCORECARD_TABLE_X, y: tableY,
      colW: table.colW,
      rowH: SCORECARD_ROW_H,
      border: BORDER,
      autoPage: false
    });
    const legendY = tableY + (chunk.rows.length + 1) * SCORECARD_ROW_H + 0.25;
    addScoreLegend(slide, legendY);
  }
}

// ---------------------------------------------------------------------------
// Top institutions by AUM -- restyled 2026-08-07 (header/colour system only;
// the 3-across layout from Peter's 2026-07-15 feedback already used the
// slide reasonably well, so the structure is unchanged).
const TOPINST_INSTITUTION_MAX_CHARS = 30;

function buildTopInstitutionsTableRows(section) {
  const headerLabels = ['Rank', 'Institution', 'AUM ($bn)'];
  const bodyText = section.institutions.map((inst, i) => [
    String(i + 1),
    inst.name,
    typeof inst.aum_bn === 'number' ? inst.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'
  ]);
  const colW = estimateColumnCharWidths(headerLabels, bodyText, {
    minChars: 4,
    maxCharsPerCol: [6, TOPINST_INSTITUTION_MAX_CHARS, 12]
  }).map(charsToInches);

  const header = headerLabels.map((t) => ({ text: t, options: headerCellOpts() }));
  const body = bodyText.map((cells, ri) => cells.map((text) => ({ text, options: { fontSize: 8.5, margin: CELL_MARGIN, fill: { color: ri % 2 ? C.rowStripe : 'FFFFFF' } } })));
  return { rows: [header, ...body], colW };
}

const TOPINST_COLS = 3;
const TOPINST_MARGIN = 0.4;
const TOPINST_GUTTER = 0.3;
const TOPINST_SLOT_W = (SLIDE_W - 2 * TOPINST_MARGIN - (TOPINST_COLS - 1) * TOPINST_GUTTER) / TOPINST_COLS;
const TOPINST_SUBTITLE_Y = 1.35;
const TOPINST_TABLE_Y = 1.8;

function addTopInstitutionsSlides(pptx, countryName, segments) {
  const sections = buildTopInstitutionsSections(segments);
  for (let i = 0; i < sections.length; i += TOPINST_COLS) {
    const group = sections.slice(i, i + TOPINST_COLS);
    const slide = addAtlasSlide(pptx);
    // No shared subtitle line here (unlike addSlideHeader's usual second
    // arg) -- each of the 3 columns below carries its own segment-name
    // subheading at TOPINST_SUBTITLE_Y, and a single "Top institutions by
    // AUM" line at the same spot the shared subtitle would use collided
    // with the first column's heading in practice (both sit at roughly the
    // same y, and LibreOffice/PowerPoint's actual glyph height doesn't
    // respect the text box's nominal h closely enough to avoid overlap at
    // this line spacing).
    addSlideHeader(slide, `Atlas — ${countryName}`);
    slide.addText('Top institutions by AUM', { x: TITLE_X, y: SUBTITLE_Y, w: SLIDE_W - TITLE_X - CONTENT_RIGHT_MARGIN, h: 0.32, fontSize: 12, color: C.muted });

    group.forEach((section, col) => {
      const slotX = TOPINST_MARGIN + col * (TOPINST_SLOT_W + TOPINST_GUTTER);
      const nText = section.n_institutions ? ` of ${section.n_institutions.toLocaleString()} identified` : '';
      slide.addText(section.segment, { x: slotX, y: TOPINST_SUBTITLE_Y, w: TOPINST_SLOT_W, fontSize: 12, bold: true, color: C.navy });
      slide.addText(
        `Top ${section.institutions.length}${nText} hold ${section.top10_share_pct}% of segment AUM`,
        { x: slotX, y: TOPINST_SUBTITLE_Y + 0.26, w: TOPINST_SLOT_W, fontSize: 9, color: C.muted }
      );
      const topTable = buildTopInstitutionsTableRows(section);
      slide.addTable(topTable.rows, {
        x: slotX, y: TOPINST_TABLE_Y,
        colW: topTable.colW,
        border: BORDER,
        autoPage: false
      });
    });
  }
}

// ---------------------------------------------------------------------------
// Country commentary -- layout unchanged from the 2026-07-29 two-column
// redesign (text left; asset allocation + recent developments stacked
// right), which already used the slide well when there's chart/developments
// content. Only the header treatment changes here to match the new system.
const COMMENTARY_BODY_Y = 1.3;
const CHART_W = 2.25;
const CHART_H = 1.41;
const CHART_CAPTION_H = 0.55;
const CHART_GUTTER = 0.25;

function addOneSegmentChart(slide, chart, x, y, w) {
  slide.addChart('doughnut', [{
    name: 'Asset allocation',
    labels: chart.slices.map((s) => s.label),
    values: chart.slices.map((s) => s.pct)
  }], {
    x, y, w: CHART_W, h: CHART_H,
    chartColors: chart.slices.map((s) => s.color),
    showLegend: true,
    legendPos: 'r',
    legendFontSize: 6,
    showValue: true,
    dataLabelFormatCode: '0"%"',
    dataLabelColor: 'FFFFFF',
    dataLabelFontSize: 6
  });
  slide.addText(chart.segmentName, {
    x, y: y + CHART_H + 0.03, w: Math.max(w, CHART_W), h: CHART_CAPTION_H,
    fontSize: 7, bold: true, color: C.muted, valign: 'top'
  });
}

// Estimates how tall a set of paragraphs will render at a given font size,
// so addCommentarySlideBody() below can pick the largest font that actually
// fits bodyH instead of a single fixed size -- added 2026-08-07 after
// testing against real data found the fixed 11pt size silently overflowed
// the text box (and the visible slide) on countries whose commentary runs
// three full paragraphs (e.g. UK's Pensions structure), with the overflow
// text simply lost off the bottom of the slide since autoFit is off. Uses a
// standard ~0.5em average character width approximation for a proportional
// font, deliberately biased slightly wide (CHAR_WIDTH_EM_FACTOR) so it
// under-counts characters-per-line and therefore over-estimates required
// height -- picking a font a touch smaller than strictly necessary is a far
// smaller problem than picking one that overflows.
const CHAR_WIDTH_EM_FACTOR = 0.52;
const LINE_HEIGHT_EM_FACTOR = 1.25;

function estimateParagraphsHeightIn(paragraphs, widthIn, fontSizePt, paraSpaceAfterPt) {
  const widthPt = widthIn * 72;
  const charsPerLine = Math.max(1, Math.floor(widthPt / (fontSizePt * CHAR_WIDTH_EM_FACTOR)));
  const lineHeightIn = (fontSizePt * LINE_HEIGHT_EM_FACTOR) / 72;
  const paraGapIn = paraSpaceAfterPt / 72;
  let totalIn = 0;
  paragraphs.forEach((p) => {
    const lines = Math.max(1, Math.ceil(p.length / charsPerLine));
    totalIn += lines * lineHeightIn + paraGapIn;
  });
  return totalIn;
}

// Candidate sizes tried largest-first; the first that fits wins, so a short
// section (e.g. one paragraph) still gets the larger, more legible size
// instead of every section defaulting to the smallest that fits the longest
// one.
const COMMENTARY_FONT_CANDIDATES = [12.5, 11.5, 11, 10, 9.5, 9, 8.5];
function pickCommentaryFontSize(paragraphs, widthIn, maxHeightIn) {
  for (const size of COMMENTARY_FONT_CANDIDATES) {
    const paraSpaceAfter = size >= 11 ? 12 : 8;
    if (estimateParagraphsHeightIn(paragraphs, widthIn, size, paraSpaceAfter) <= maxHeightIn) {
      return { fontSize: size, paraSpaceAfter };
    }
  }
  const smallest = COMMENTARY_FONT_CANDIDATES[COMMENTARY_FONT_CANDIDATES.length - 1];
  return { fontSize: smallest, paraSpaceAfter: 8 };
}

const PLACEHOLDER_TEXT_OPTS = { italic: true, fontSize: 9, color: '8A97A1' };
const RIGHT_COL_X = 7.3;
const RIGHT_COL_W = SLIDE_W - 0.8 - RIGHT_COL_X;
const RIGHT_HEADING_H = 0.22;

function addCommentarySlideBody(slide, countryName, sectionLabel, section) {
  addSlideHeader(slide, `Atlas — ${countryName}`, sectionLabel);

  const textW = 6.6;
  const bodyH = 5.6;

  if (section.paragraphs.length) {
    // Font size is picked per-section (see pickCommentaryFontSize() above),
    // not fixed -- testing against real data found a single fixed 11pt
    // silently overflowed the slide on countries with three full paragraphs
    // of commentary (UK's Pensions structure), losing content off the
    // bottom with autoFit deliberately off. A short section instead gets a
    // larger, more legible size than the old fixed 11pt, which both fills
    // the space better and never risks overflow on a longer one.
    const { fontSize, paraSpaceAfter } = pickCommentaryFontSize(section.paragraphs, textW, bodyH);
    const bodyRuns = section.paragraphs.map((p) => ({ text: p, options: { breakLine: true, paraSpaceAfter } }));
    slide.addText(bodyRuns, {
      x: CONTENT_X, y: COMMENTARY_BODY_Y, w: textW, h: bodyH,
      fontSize, valign: 'top', align: 'left', autoFit: false
    });
  } else {
    slide.addText(`No commentary written yet for ${sectionLabel}.`, {
      x: CONTENT_X, y: COMMENTARY_BODY_Y, w: textW, h: 0.4, valign: 'top', ...PLACEHOLDER_TEXT_OPTS
    });
  }

  slide.addText('Asset allocation', { x: RIGHT_COL_X, y: COMMENTARY_BODY_Y, w: RIGHT_COL_W, h: RIGHT_HEADING_H, fontSize: 11, bold: true, color: C.navy });
  const allocContentY = COMMENTARY_BODY_Y + RIGHT_HEADING_H + 0.05;
  const charts = section.chartSegments.map((seg) => buildSegmentAllocationChart(seg)).filter(Boolean);

  let allocBlockH;
  if (!section.hasChartSlot) {
    slide.addText('Not available for this section.', { x: RIGHT_COL_X, y: allocContentY, w: RIGHT_COL_W, h: 0.35, valign: 'top', ...PLACEHOLDER_TEXT_OPTS });
    allocBlockH = 0.4;
  } else if (!charts.length) {
    slide.addText('No segment-level data available yet.', { x: RIGHT_COL_X, y: allocContentY, w: RIGHT_COL_W, h: 0.35, valign: 'top', ...PLACEHOLDER_TEXT_OPTS });
    allocBlockH = 0.4;
  } else {
    charts.forEach((chart, i) => {
      const chartY = allocContentY + i * (CHART_H + CHART_CAPTION_H + CHART_GUTTER);
      addOneSegmentChart(slide, chart, RIGHT_COL_X, chartY, RIGHT_COL_W);
    });
    allocBlockH = charts.length * (CHART_H + CHART_CAPTION_H + CHART_GUTTER);
  }

  const devHeadingY = allocContentY + allocBlockH + 0.15;
  slide.addText('Recent developments', { x: RIGHT_COL_X, y: devHeadingY, w: RIGHT_COL_W, h: RIGHT_HEADING_H, fontSize: 11, bold: true, color: C.navy });
  const devContentY = devHeadingY + RIGHT_HEADING_H + 0.05;
  const devContentH = Math.max(0.3, (COMMENTARY_BODY_Y + bodyH) - devContentY);

  if (!section.developments.length) {
    slide.addText('No recent developments logged yet.', { x: RIGHT_COL_X, y: devContentY, w: RIGHT_COL_W, h: 0.35, valign: 'top', ...PLACEHOLDER_TEXT_OPTS });
  } else {
    const devRuns = [];
    section.developments.forEach((d) => {
      devRuns.push({ text: `${d.date ? `${d.date} — ` : ''}${d.headline || ''}`, options: { bold: true, breakLine: true, fontSize: 8, paraSpaceAfter: 1 } });
      if (d.summary) devRuns.push({ text: d.summary, options: { breakLine: true, fontSize: 7, paraSpaceAfter: 1 } });
      const srcLabel = d.source || (d.url ? 'Source' : '');
      if (srcLabel) devRuns.push({ text: `Source: ${srcLabel}`, options: { breakLine: true, italic: true, fontSize: 6, color: C.muted, paraSpaceAfter: 6 } });
    });
    slide.addText(devRuns, { x: RIGHT_COL_X, y: devContentY, w: RIGHT_COL_W, h: devContentH, valign: 'top', align: 'left', autoFit: false });
  }
}

function addCommentarySlides(pptx, countryName, commentarySections) {
  commentarySections.forEach((section) => {
    const slide = addAtlasSlide(pptx);
    addCommentarySlideBody(slide, countryName, section.label, section);
  });
}

// ---------------------------------------------------------------------------
// Sources -- rebuilt 2026-08-07. The old version paginated at a fixed 26
// lines/slide, which could leave a trailing slide with almost nothing on it
// (a real example: a slide with 4 short lines and the rest of the frame
// blank). Now balances lines evenly across however many slides are actually
// needed (ceil(total/perSlide), then divide total by that page count) so no
// page is left mostly empty just because it happened to be last, and lays
// each slide's lines out in two columns instead of one so the same line
// count uses the slide's width as well as its height.
const SOURCES_LINES_PER_SLIDE_TARGET = 30;
const SOURCES_COL_GUTTER = 0.4;
const SOURCES_COL_W = (SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN - SOURCES_COL_GUTTER) / 2;

function buildSourcesLines(sourcesData) {
  const lines = [];
  const { commentaryGroups, segmentSources } = sourcesData;
  if (commentaryGroups.length) {
    lines.push({ text: 'Commentary', section: true });
    commentaryGroups.forEach((g) => {
      lines.push({ text: g.sectionLabel, group: true });
      g.sources.forEach((s) => lines.push({ text: `- ${s.label || s.url}${s.label && s.url ? ` — ${s.url}` : ''}` }));
    });
  }
  if (segmentSources.length) {
    lines.push({ text: 'Segment data (AUM & asset allocation)', section: true });
    segmentSources.forEach((s) => lines.push({ text: `${s.segmentName}: ${s.citation}` }));
  }
  return lines;
}

function sourceLineRun(l) {
  return {
    text: l.text,
    options: {
      breakLine: true,
      bold: !!(l.section || l.group),
      fontSize: l.section ? 12 : 8.5,
      color: l.section ? C.navy : (l.group ? C.body : C.muted),
      paraSpaceAfter: l.section ? 6 : (l.group ? 3 : 1)
    }
  };
}

function addCountrySourcesSlides(pptx, countryName, commentarySections, segments) {
  const sourcesData = buildCountrySourcesData(commentarySections, segments);
  const lines = buildSourcesLines(sourcesData);
  if (!lines.length) return;

  const pageCount = Math.max(1, Math.ceil(lines.length / SOURCES_LINES_PER_SLIDE_TARGET));
  const perPage = Math.ceil(lines.length / pageCount);

  for (let i = 0; i < lines.length; i += perPage) {
    const group = lines.slice(i, i + perPage);
    const mid = Math.ceil(group.length / 2);
    const colA = group.slice(0, mid);
    const colB = group.slice(mid);

    const slide = addAtlasSlide(pptx);
    addSlideHeader(slide, `Atlas — ${countryName}`, 'Sources');
    slide.addText(colA.map(sourceLineRun), { x: CONTENT_X, y: 1.3, w: SOURCES_COL_W, h: 5.7, valign: 'top', align: 'left', autoFit: false });
    if (colB.length) {
      slide.addText(colB.map(sourceLineRun), { x: CONTENT_X + SOURCES_COL_W + SOURCES_COL_GUTTER, y: 1.3, w: SOURCES_COL_W, h: 5.7, valign: 'top', align: 'left', autoFit: false });
    }
  }
}

// ---------------------------------------------------------------------------
// Content-block dispatch. Which of the four blocks to add for a country --
// 'commentary', 'aum', 'scorecard', 'top_institutions'.
const ALL_CONTENT_TYPES = ['commentary', 'aum', 'scorecard', 'top_institutions'];
function resolveInclude(rawInclude) {
  const valid = Array.isArray(rawInclude) ? rawInclude.filter((k) => ALL_CONTENT_TYPES.includes(k)) : [];
  return new Set(valid.length ? valid : ALL_CONTENT_TYPES);
}

function addCountrySlides(pptx, countryName, segments, generatedDate, enabledDimensions, includeSet, weightOverrides, commentarySections, allocType, allocStyle, customDimensions) {
  if (includeSet.has('commentary') && commentarySections.length) addCommentarySlides(pptx, countryName, commentarySections);
  if (includeSet.has('aum') && (segments || []).length) addAumSlide(pptx, countryName, segments, generatedDate);
  if (includeSet.has('scorecard') && (segments || []).length) addScorecardSlide(pptx, countryName, segments, enabledDimensions, weightOverrides, allocType, allocStyle, customDimensions);
  if (includeSet.has('top_institutions')) addTopInstitutionsSlides(pptx, countryName, segments);
  if (includeSet.has('commentary') || includeSet.has('aum')) {
    addCountrySourcesSlides(pptx, countryName, commentarySections, segments);
  }
}

// Optional appendix slides built from picker.html's "Attach supporting
// evidence" picker. Restyled header only -- layout unchanged, it already
// used the slide reasonably fully.
function addEvidenceSlides(pptx, evidence) {
  if (!Array.isArray(evidence) || !evidence.length) return;

  const introSlide = addAtlasSlide(pptx);
  addSlideHeader(introSlide, 'Appendix — Supporting Evidence');
  introSlide.addText('Cross-cutting reference material (fee benchmarks, market-structure surveys) that applies across countries or regions, attached to this export separately from the country data above.', {
    x: CONTENT_X, y: 1.2, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, h: 1.2, fontSize: 14, color: C.muted, valign: 'top'
  });

  evidence.forEach((e) => {
    const slide = addAtlasSlide(pptx);
    addSlideHeader(slide, e.title || e.id || 'Untitled entry');
    const metaText = `${e.theme ? `Theme: ${e.theme}` : ''}${e.scope && e.scope.length ? `  |  Scope: ${e.scope.join(', ')}` : ''}`;
    slide.addText(metaText, { x: CONTENT_X, y: 0.85, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, fontSize: 11, color: C.muted });

    const bodyRuns = [];
    if (e.summary) bodyRuns.push({ text: e.summary, options: { breakLine: true, paraSpaceAfter: 8 } });
    if (Array.isArray(e.figures) && e.figures.length) {
      bodyRuns.push({ text: 'Figures', options: { bold: true, breakLine: true, paraSpaceAfter: 2 } });
      e.figures.forEach((f) => {
        const valueText = `${typeof f.value === 'number' ? f.value : '-'}${f.unit ? ' ' + f.unit : ''}`;
        bodyRuns.push({ text: `${f.region || '-'} — ${f.metric || '-'}: ${valueText}${f.as_of ? ` (${f.as_of})` : ''}`, options: { breakLine: true, paraSpaceAfter: 2 } });
      });
    }
    if (e.note) bodyRuns.push({ text: e.note, options: { italic: true, breakLine: true, paraSpaceAfter: 8, color: '444444' } });

    slide.addText(bodyRuns, {
      x: CONTENT_X, y: 1.3, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, h: 5.2,
      fontSize: 12, valign: 'top', align: 'left', autoFit: false
    });

    slide.addText(`Source: ${e.source || 'unknown'} (${e.as_of || 'date unknown'})${e.access ? ` — ${e.access}` : ''}`, {
      x: CONTENT_X, y: 6.7, w: SLIDE_W - CONTENT_X - CONTENT_RIGHT_MARGIN, fontSize: 9, color: C.muted
    });
  });
}

app.http('exportPptx', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'export/pptx',
  handler: async (request, context) => {
    let body;
    try {
      body = await request.json();
    } catch {
      return { status: 400, jsonBody: { error: 'Invalid or missing JSON body' } };
    }

    const isMulti = Array.isArray(body.countries);
    const countries = isMulti
      ? body.countries.filter((c) => c && Array.isArray(c.segments) && c.segments.length)
      : (Array.isArray(body.segments) && body.segments.length ? [{ country_name: body.country_name || 'Country', segments: body.segments, commentary: body.commentary, developments: body.developments }] : []);

    if (!countries.length) {
      return { status: 400, jsonBody: { error: 'No segments provided to export' } };
    }

    const safeName = isMulti
      ? `Project_${countries.length}_countries`
      : countries[0].country_name.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

    try {
      const generatedDate = new Date().toISOString().slice(0, 10);
      const include = resolveInclude(body.include);
      const hasEvidence = Array.isArray(body.evidence) && body.evidence.length;

      // ---- Pass 1: plan the whole deck's structure before drawing anything.
      // Each country's block list (and each block's slide count) comes from
      // planCountryOutline() -- the same eligibility/chunking logic the real
      // slide-building functions below use, so this plan can't drift out of
      // sync with what actually gets built (see that function's comment in
      // exportHelpers.js).
      const outlines = countries.map((c) => planCountryOutline(
        c.country_name, c.segments, c.commentary, c.developments, include,
        { topInstCols: TOPINST_COLS, sourcesLinesPerSlide: SOURCES_LINES_PER_SLIDE_TARGET, scorecardColsPerSlide: SCORECARD_COLS_PER_SLIDE }
      ));

      const outlineRows = [];
      outlines.forEach((o) => {
        if (!o.blocks.length) return;
        outlineRows.push({ kind: 'country', text: o.countryName });
        o.blocks.forEach((b) => outlineRows.push({ kind: 'block', text: b.label, blockRef: b }));
      });
      if (hasEvidence) {
        outlineRows.push({ kind: 'country', text: 'Appendix' });
        outlineRows.push({ kind: 'block', text: 'Supporting evidence', blockRef: { slideCount: 1 + body.evidence.length } });
      }

      // Work out the TOC's own page count (from row heights, same pagination
      // math addTocSlides() uses) so we know exactly which slide number
      // country content starts on -- cover (1) + TOC pages + methodology (1).
      let tocPageCount = 0;
      if (outlineRows.length) {
        let y = TOC_TOP_Y, pages = 1;
        outlineRows.forEach((row) => {
          const rowH = row.kind === 'country' ? TOC_COUNTRY_ROW_H : TOC_BLOCK_ROW_H;
          if (y + rowH > TOC_BOTTOM_Y) { pages += 1; y = TOC_TOP_Y; }
          y += rowH;
        });
        tocPageCount = pages;
      }

      let slideCursor = 1 + tocPageCount + 1; // cover + TOC + methodology
      outlineRows.forEach((row) => {
        if (row.kind !== 'block') return;
        row.slideNumber = slideCursor + 1;
        slideCursor += row.blockRef.slideCount;
      });

      // ---- Pass 2: actually build the deck, in the exact same order.
      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'ATLAS_WIDE', width: SLIDE_W, height: SLIDE_H });
      pptx.layout = 'ATLAS_WIDE';
      defineAtlasMaster(pptx);

      const title = isMulti ? `Institutional Market Report — ${countries.length} countries` : `Institutional Market Report — ${countries[0].country_name}`;
      const subtitle = isMulti ? countries.map((c) => c.country_name).join(', ') : 'Country deep-dive';
      const coverageLine = `Covers: country commentary, AUM by segment, opportunity scorecard, top institutions by AUM, and full sourcing — for ${isMulti ? `the ${countries.length} selected countries` : countries[0].country_name} only.`;
      addCoverSlide(pptx, { title, subtitle, coverageLine, generatedDate });
      addTocSlides(pptx, outlineRows);
      addMethodologySlide(pptx);

      countries.forEach((c, i) => {
        addCountrySlides(
          pptx, c.country_name, c.segments, generatedDate, body.enabled_dimensions, include,
          body.weight_overrides, outlines[i].commentarySections, body.alloc_type, body.alloc_style, body.custom_dimensions
        );
      });

      addEvidenceSlides(pptx, body.evidence);

      const buffer = await pptx.write({ outputType: 'nodebuffer' });

      return {
        status: 200,
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'Content-Disposition': `attachment; filename="Atlas_${safeName}.pptx"`
        },
        body: buffer
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Failed to generate PowerPoint file', detail: err.message } };
    }
  }
});
