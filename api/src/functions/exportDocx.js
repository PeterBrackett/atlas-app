const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak, Header, Footer, AlignmentType, BorderStyle,
  TableOfContents, PageNumber, TabStopType, TabStopPosition, TableLayoutType
} = require('docx');
const {
  buildAumRows, buildScorecardMatrix, buildCommentarySectionsFull, buildTopInstitutionsSections,
  estimateColumnCharWidths, buildSegmentAllocationChart, buildCountrySourcesData,
  BRAND, METHODOLOGY_CONTENT
} = require('../shared/exportHelpers');
const { getDimensionIconBuffer } = require('../shared/dimensionIcons');
const { getAtlasLogoBuffer, getAtlasLockupBuffer } = require('../shared/atlasLogo');

// ---------------------------------------------------------------------------
// Brand system -- 2026-08-07 rebuild, the Word-export half of the same
// "far from a place I could share with a client" feedback addressed in
// exportPptx.js (see that file's header comment for the fuller context).
// Word doesn't have the PPTX version's fixed-slide-bounds overflow risk --
// text just flows onto the next page -- so this file's changes are: a
// branded cover page, a native (Word-updatable) table of contents, the same
// standard methodology page, and a colour/heading pass so tables and
// headings read as the same visual identity as the site and the PowerPoint
// export, not Word's default blue theme headings.
const C = BRAND;
const HEADER_FILL = C.navyTint;

// Compact cell margins (dxa/twips) and font sizes (half-points), applied
// throughout the export tables so more tables fit per page -- Peter's
// original "there's quite a bit of wasted space" feedback on the default
// docx table styling.
const CELL_MARGINS = { top: 30, bottom: 30, left: 80, right: 80 };
const HEADER_FONT_SIZE = 16; // 8pt
const BODY_FONT_SIZE = 16; // 8pt

// Converts an estimateColumnCharWidths() character count into a dxa
// (twips) column width, so tables are sized to their content rather than
// stretched to fill the page.
const CHAR_WIDTH_TWIPS = 105;
function charsToDxa(chars) {
  return Math.round(chars * CHAR_WIDTH_TWIPS) + CELL_MARGINS.left + CELL_MARGINS.right;
}

function headerCell(text, widthDxa) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    margins: CELL_MARGINS,
    ...(widthDxa ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: true, size: HEADER_FONT_SIZE, color: C.navy })] })]
  });
}

// `color` is the optional {bg, fg} hex pair from exportHelpers.js's
// scoreColor()/overallColor(), reproducing the site's red/amber/green
// scorecard traffic-light coding in the Word table cells. `stripe` (added
// 2026-08-07) alternates a very light grey fill on odd body rows, matching
// the same striping added to the PowerPoint export's tables -- a small
// legibility/polish detail on the longer AUM and top-institutions tables.
function bodyCell(text, color, widthDxa, stripe) {
  const fill = color ? color.bg : (stripe ? C.rowStripe : undefined);
  return new TableCell({
    margins: CELL_MARGINS,
    ...(fill ? { shading: { type: ShadingType.CLEAR, fill, color: 'auto' } } : {}),
    ...(widthDxa ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), size: BODY_FONT_SIZE, ...(color ? { color: color.fg, bold: true } : {}) })]
    })]
  });
}

// Same as headerCell(), but for a scorecard dimension row: prepends the
// dimension's icon before the label text, when one exists for that
// dimension key.
function dimensionHeaderCell(row, widthDxa) {
  const iconBuffer = row.key ? getDimensionIconBuffer(row.key) : null;
  if (!iconBuffer) return headerCell(row.label, widthDxa);
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    margins: CELL_MARGINS,
    ...(widthDxa ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    children: [new Paragraph({
      children: [
        new ImageRun({ data: iconBuffer, type: 'png', transformation: { width: 12, height: 12 } }),
        new TextRun({ text: `  ${row.label}`, bold: true, size: HEADER_FONT_SIZE, color: C.navy })
      ]
    })]
  });
}

// Atlas logo, top-left of every page. A docx Header attached to the
// document's (single) section repeats automatically on every page.
function buildLogoHeader() {
  const logoBuffer = getAtlasLogoBuffer();
  if (!logoBuffer) return undefined;
  return new Header({
    children: [new Paragraph({
      children: [new ImageRun({ data: logoBuffer, type: 'png', transformation: { width: 26, height: 26 } })]
    })]
  });
}

// Footer -- new 2026-08-07: page number (right-aligned, via a right tab
// stop across the full text width) plus a small "Atlas — Institutional
// Adviser | Confidential" tag (left-aligned), matching the PowerPoint
// export's footer rule. Word's PageNumber.CURRENT is a live field, so it
// stays correct regardless of how the final page count shakes out.
function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.hairline, space: 4 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: 'Atlas — Institutional Adviser  |  Confidential', size: 14, color: C.muted }),
          new TextRun({ text: '\t' }),
          new TextRun({ text: 'Page ', size: 14, color: C.muted }),
          new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.muted }),
          new TextRun({ text: ' of ', size: 14, color: C.muted }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, color: C.muted })
        ]
      })
    ]
  });
}

// ---------------------------------------------------------------------------
// Cover page -- new 2026-08-07. White page (unlike the PowerPoint export's
// full-navy cover -- a solid dark background is a normal thing for a title
// slide but an unusual, ink-heavy choice for the first page of a document
// that may get printed or marked up) with the standard navy/teal lockup,
// a large navy title, a thin teal rule, and a subtitle/coverage/generated-
// date block, ending in a page break.
function buildCoverPage({ title, subtitle, coverageLine, generatedDate }) {
  const lockupBuffer = getAtlasLockupBuffer();
  const children = [];
  if (lockupBuffer) {
    // Source asset is 1801x600 -- held to a modest on-page width so it
    // reads as a masthead, not a full-bleed graphic.
    children.push(new Paragraph({
      spacing: { after: 500 },
      children: [new ImageRun({ data: lockupBuffer, type: 'png', transformation: { width: 234, height: 78 } })]
    }));
  }
  children.push(new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.teal, space: 8 } },
    spacing: { after: 200 },
    children: [new TextRun({ text: '', size: 2 })]
  }));
  children.push(new Paragraph({
    spacing: { before: 300, after: 200 },
    children: [new TextRun({ text: title, bold: true, size: 56, color: C.navy })]
  }));
  if (subtitle) {
    children.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: subtitle, size: 26, color: C.body })] }));
  }
  if (coverageLine) {
    children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: coverageLine, size: 20, color: C.muted })] }));
  }
  children.push(new Paragraph({
    spacing: { before: 3000 },
    children: [new TextRun({ text: `Generated ${generatedDate}   ·   Prepared by Institutional Adviser   ·   Confidential — not for redistribution`, italics: true, size: 18, color: C.muted })]
  }));
  children.push(new Paragraph({ children: [new PageBreak()] }));
  return children;
}

// ---------------------------------------------------------------------------
// Table of contents -- new 2026-08-07. Word's native TOC field, built from
// this document's own Heading 1/2 styled paragraphs (country titles are
// Heading 1; "Country commentary"/"AUM by segment"/"Opportunity scorecard"/
// "Top institutions by AUM"/"Sources"/"Data structure & methodology"/
// "Appendix — Supporting Evidence" are all Heading 2) -- unlike the
// PowerPoint export, there's no need to hand-compute page numbers or
// pre-plan pagination here; Word does that itself once the field is
// updated. The field shows literal placeholder text ("Right-click > Update
// Field", or Word may prompt automatically) until opened and refreshed in
// Word -- standard behaviour for any programmatically generated Word TOC,
// not a bug in this file.
function buildTocPage() {
  return [
    new Paragraph({ text: 'Contents', heading: HeadingLevel.TITLE, spacing: { after: 200 } }),
    new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ---------------------------------------------------------------------------
// Methodology page -- new 2026-08-07, Peter's "a standard page that
// summarises the structure of the data and the key data sources and
// methodology" request. Content comes entirely from exportHelpers.js's
// METHODOLOGY_CONTENT (condensed from atlas-site/data/sources.md and the
// Atlas_Methodology_Reference.docx collateral) so it can't drift from the
// fuller reference. Laid out as a borderless 2x2 table of shaded cards,
// mirroring the PowerPoint export's card grid, rather than a single long
// scrolling block of prose.
const METH_CARD_BORDER = { style: BorderStyle.SINGLE, size: 2, color: C.hairline };
const METH_CARD_BORDERS = { top: METH_CARD_BORDER, bottom: METH_CARD_BORDER, left: METH_CARD_BORDER, right: METH_CARD_BORDER, insideHorizontal: METH_CARD_BORDER, insideVertical: METH_CARD_BORDER };

function methCard(section) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: C.panelBg, color: 'auto' },
    margins: { top: 160, bottom: 160, left: 160, right: 160 },
    width: { size: 4500, type: WidthType.DXA },
    children: [
      new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: section.heading, bold: true, size: 22, color: C.teal })] }),
      new Paragraph({ children: [new TextRun({ text: section.body, size: 17, color: C.body })] })
    ]
  });
}

function buildMethodologyPage() {
  const [s1, s2, s3, s4] = METHODOLOGY_CONTENT.sections;
  return [
    new Paragraph({ text: 'Data structure & methodology', heading: HeadingLevel.HEADING_2, spacing: { before: 0, after: 40 } }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: METHODOLOGY_CONTENT.intro, size: 19, color: C.body })] }),
    new Table({
      borders: METH_CARD_BORDERS,
      columnWidths: [4500, 4500],
      width: { size: 9000, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ children: [methCard(s1), methCard(s2)] }),
        new TableRow({ children: [methCard(s3), methCard(s4)] })
      ]
    }),
    new Paragraph({ spacing: { before: 200 }, children: [new TextRun({ text: METHODOLOGY_CONTENT.footer, italics: true, size: 16, color: C.muted })] }),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// "{Type} range (min-max)" reflects that not every institution counted in a
// segment's AUM also filed an asset-class breakdown -- min is the reported
// figure as-is, max is that figure scaled up to the segment's full AUM. See
// getAllocationRange() in exportHelpers.js. allocType -- 2026-08-07: the
// asset-class columns only appear at all when the caller has actually
// chosen one (see buildAumRows()'s comment); with no allocType this is a
// plain three-column Segment/AUM/Basis table, matching "AUM only" everywhere
// else in the export.
function buildAumTable(rows, allocType) {
  const headerLabels = allocType
    ? ['Segment', 'AUM ($bn)', `${allocType} ($bn)`, 'Basis', `${allocType} range (min-max)`]
    : ['Segment', 'AUM ($bn)', 'Basis'];
  const bodyText = rows.map((r) => allocType ? [
    r.segment,
    typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    typeof r.alloc_bn === 'number' ? r.alloc_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    r.basis || '',
    r.alloc_range || '-'
  ] : [
    r.segment,
    typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    r.basis || ''
  ]);
  const widths = estimateColumnCharWidths(headerLabels, bodyText, {
    minChars: 4,
    maxCharsPerCol: allocType ? [26, 10, 10, 22, 30] : [30, 10, 26]
  }).map(charsToDxa);

  const headerRow = new TableRow({
    children: headerLabels.map((label, i) => headerCell(label, widths[i]))
  });
  const dataRows = bodyText.map((cells, ri) => new TableRow({
    children: cells.map((text, i) => bodyCell(text, null, widths[i], ri % 2 === 1))
  }));
  return new Table({
    rows: [headerRow, ...dataRows],
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED
  });
}

// Stat-card row -- new 2026-08-07, the same "at a glance" summary added
// above the PowerPoint export's AUM table (see computeAumStats()/
// addStatCards() in exportPptx.js), reused here as a borderless 1x4 table
// of shaded cells so the AUM section opens with real summary content
// instead of jumping straight into a dense table.
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

function statCard(label, value) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: C.panelBg, color: 'auto' },
    margins: { top: 120, bottom: 120, left: 140, right: 140 },
    width: { size: 2400, type: WidthType.DXA },
    children: [
      new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: label, bold: true, size: 14, color: C.muted })] }),
      new Paragraph({ children: [new TextRun({ text: value, bold: true, size: 24, color: C.navy })] })
    ]
  });
}

function buildAumStatCards(stats) {
  return new Table({
    borders: METH_CARD_BORDERS,
    columnWidths: [2400, 2400, 2400, 2400],
    width: { size: 9600, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({
      children: [
        statCard('TOTAL TRACKED AUM', `$${stats.totalAum.toLocaleString(undefined, { maximumFractionDigits: 1 })}bn`),
        statCard('SEGMENTS', String(stats.count)),
        statCard('LARGEST SEGMENT', stats.largest ? stats.largest.segment : '—'),
        statCard('SOURCE MIX', stats.mixText || '—')
      ]
    })]
  });
}

// Word tables don't scroll horizontally, so a country with many segment
// columns (e.g. UK's 18) will run wide across the page. Column widths are
// content-driven the same way as buildAumTable().
function buildScorecardTable(matrix) {
  const headerLabels = ['Dimension', ...matrix.columnLabels];
  const bodyText = matrix.rows.map((row) => [row.label, ...row.values]);
  const maxCharsPerCol = [34, ...matrix.columnLabels.map(() => 6)];
  const widths = estimateColumnCharWidths(headerLabels, bodyText, { minChars: 3, maxCharsPerCol }).map(charsToDxa);
  widths[0] += 260; // room for the dimension icon alongside the label

  const headerRow = new TableRow({
    children: [headerCell('Dimension', widths[0]), ...matrix.columnLabels.map((label, i) => headerCell(label, widths[i + 1]))]
  });
  const dataRows = matrix.rows.map((row) => new TableRow({
    children: [
      dimensionHeaderCell(row, widths[0]),
      ...row.values.map((v, i) => bodyCell(v, row.colors ? row.colors[i] : null, widths[i + 1]))
    ]
  }));
  return new Table({
    rows: [headerRow, ...dataRows],
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED
  });
}

// Colour-key legend under the scorecard table -- new 2026-08-07, matching
// the same legend added to the PowerPoint export (see SCORE_LEGEND/
// addScoreLegend() in exportPptx.js), so the red/amber/green/yellow
// colour-coding is explained on the page rather than left implicit.
const SCORE_LEGEND = [
  { label: '1 — unfavourable', color: { bg: 'FBE1E1', fg: 'A3291F' } },
  { label: '2 — moderate', color: { bg: 'FDEEE0', fg: '9A5A1A' } },
  { label: '3 — favourable', color: { bg: 'E3F2E3', fg: '1F7A34' } },
  { label: 'not yet scored', color: { bg: 'FFF3B0', fg: '7A5C00' } }
];

function buildScoreLegend() {
  const runs = [];
  SCORE_LEGEND.forEach((item, i) => {
    if (i > 0) runs.push(new TextRun({ text: '     ' }));
    runs.push(new TextRun({ text: '■ ', color: item.color.fg }));
    runs.push(new TextRun({ text: item.label, size: 15, color: C.muted }));
  });
  return new Paragraph({ children: runs, spacing: { before: 100, after: 40 } });
}

function buildTopInstitutionsTable(section) {
  const headerLabels = ['Rank', 'Institution', 'AUM ($bn)'];
  const bodyText = section.institutions.map((inst, i) => [
    String(i + 1),
    inst.name,
    typeof inst.aum_bn === 'number' ? inst.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'
  ]);
  const widths = estimateColumnCharWidths(headerLabels, bodyText, {
    minChars: 4,
    maxCharsPerCol: [6, 40, 12]
  }).map(charsToDxa);

  const headerRow = new TableRow({
    children: headerLabels.map((label, i) => headerCell(label, widths[i]))
  });
  const dataRows = bodyText.map((cells, ri) => new TableRow({
    children: cells.map((text, i) => bodyCell(text, null, widths[i], ri % 2 === 1))
  }));
  return new Table({
    rows: [headerRow, ...dataRows],
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED
  });
}

function buildTopInstitutionsPerSegment(sections) {
  return sections.flatMap((section) => {
    const nText = section.n_institutions ? ` of ${section.n_institutions.toLocaleString()} identified` : '';
    return [
      new Paragraph({
        text: `${section.segment} — top ${section.institutions.length}${nText} institutions hold ${section.top10_share_pct}% of segment AUM`,
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 150, after: 40 }
      }),
      buildTopInstitutionsTable(section)
    ];
  });
}

function buildTopInstitutionsBlock(segments) {
  const sections = buildTopInstitutionsSections(segments);
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Top institutions by AUM', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  return [heading, ...buildTopInstitutionsPerSegment(sections)];
}

// Word has no native chart object -- a segment's asset-class mix renders as
// a single-row table whose cell widths are proportional to each slice's
// share of the segment's full aum_bn, shaded with the same asset-type
// colours used everywhere else.
const ALLOCATION_BAR_TOTAL_DXA = 2400; // ~1.65in
const ALLOCATION_BAR_MIN_DXA = 60;
const ALLOCATION_BAR_LABEL_MIN_DXA = 210; // only label a cell wide enough to hold "12.3%"

function buildAllocationBarTable(chart) {
  const n = chart.slices.length;
  const remaining = Math.max(ALLOCATION_BAR_TOTAL_DXA - ALLOCATION_BAR_MIN_DXA * n, 0);
  const cellWidths = chart.slices.map((s) => Math.round(ALLOCATION_BAR_MIN_DXA + (s.pct / 100) * remaining));
  const cells = chart.slices.map((s, i) => {
    const widthDxa = cellWidths[i];
    const showLabel = widthDxa >= ALLOCATION_BAR_LABEL_MIN_DXA;
    return new TableCell({
      width: { size: widthDxa, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: s.color, color: 'auto' },
      margins: { top: 40, bottom: 40, left: 20, right: 20 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: showLabel ? [new TextRun({ text: `${s.pct}%`, size: 14, color: 'FFFFFF', bold: true })] : []
      })]
    });
  });
  return new Table({
    rows: [new TableRow({ children: cells })],
    columnWidths: cellWidths,
    width: { size: ALLOCATION_BAR_TOTAL_DXA, type: WidthType.DXA },
    layout: TableLayoutType.FIXED
  });
}

function buildAllocationLegendParagraph(chart) {
  const runs = [];
  chart.slices.forEach((s, i) => {
    if (i > 0) runs.push(new TextRun({ text: '     ' }));
    runs.push(new TextRun({ text: '■ ', color: s.color }));
    runs.push(new TextRun({ text: `${s.label} ${s.pct}%`, size: 16 }));
  });
  return new Paragraph({ children: runs, spacing: { before: 40, after: 20 } });
}

function buildSegmentAllocationVisual(segment) {
  const chart = buildSegmentAllocationChart(segment);
  if (!chart) return [];
  return [
    new Paragraph({ children: [new TextRun({ text: chart.segmentName, bold: true, size: 16 })], spacing: { before: 60, after: 20 } }),
    buildAllocationBarTable(chart),
    buildAllocationLegendParagraph(chart)
  ];
}

const PLACEHOLDER_RUN_OPTS = { italics: true, size: 15, color: '5B6B7A' };

function buildAllocationBoxContent(section) {
  if (!section.hasChartSlot) {
    return [new Paragraph({ children: [new TextRun({ text: 'Not available for this section.', ...PLACEHOLDER_RUN_OPTS })] })];
  }
  if (!section.chartSegments.length) {
    return [new Paragraph({ children: [new TextRun({ text: 'No segment-level data available yet for asset allocation in this section.', ...PLACEHOLDER_RUN_OPTS })] })];
  }
  return section.chartSegments.flatMap((seg) => buildSegmentAllocationVisual(seg));
}

function buildDevelopmentsBoxContent(items) {
  if (!items.length) {
    return [new Paragraph({ children: [new TextRun({ text: 'No recent developments logged yet.', ...PLACEHOLDER_RUN_OPTS })] })];
  }
  return items.flatMap((d) => {
    const headlineRuns = [
      ...(d.date ? [new TextRun({ text: `${d.date} — `, bold: true, size: 15 })] : []),
      new TextRun({ text: d.headline || '', size: 15 })
    ];
    const paras = [new Paragraph({ children: headlineRuns, spacing: { before: 60, after: 10 } })];
    if (d.summary) paras.push(new Paragraph({ children: [new TextRun({ text: d.summary, size: 14 })], spacing: { after: 10 } }));
    const srcLabel = d.source || (d.url ? 'Source' : '');
    if (srcLabel) paras.push(new Paragraph({ children: [new TextRun({ text: `Source: ${srcLabel}${d.url ? ` — ${d.url}` : ''}`, italics: true, size: 13, color: C.muted })], spacing: { after: 40 } }));
    return paras;
  });
}

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER, insideHorizontal: NO_BORDER, insideVertical: NO_BORDER };
const COMMENTARY_LEFT_COL_DXA = 5500;
const COMMENTARY_RIGHT_COL_DXA = 3500;

function buildCommentaryBlock(sections) {
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Country commentary', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  const body = sections.flatMap((section) => {
    const subheading = new Paragraph({ text: section.label, heading: HeadingLevel.HEADING_3, spacing: { before: 150, after: 40 } });

    const leftChildren = section.paragraphs.length
      ? section.paragraphs.map((p) => new Paragraph({ text: p, spacing: { after: 100 } }))
      : [new Paragraph({ children: [new TextRun({ text: `No commentary written yet for ${section.label}.`, ...PLACEHOLDER_RUN_OPTS })] })];

    const rightChildren = [
      new Paragraph({ children: [new TextRun({ text: 'Asset allocation', bold: true, size: 16, color: C.navy })], spacing: { after: 40 } }),
      ...buildAllocationBoxContent(section),
      new Paragraph({ text: '', spacing: { before: 80 } }),
      new Paragraph({ children: [new TextRun({ text: 'Recent developments', bold: true, size: 16, color: C.navy })], spacing: { after: 40 } }),
      ...buildDevelopmentsBoxContent(section.developments)
    ];

    const layoutTable = new Table({
      borders: NO_BORDERS,
      columnWidths: [COMMENTARY_LEFT_COL_DXA, COMMENTARY_RIGHT_COL_DXA],
      width: { size: COMMENTARY_LEFT_COL_DXA + COMMENTARY_RIGHT_COL_DXA, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      rows: [new TableRow({
        children: [
          new TableCell({ width: { size: COMMENTARY_LEFT_COL_DXA, type: WidthType.DXA }, margins: { right: 120 }, children: leftChildren }),
          new TableCell({ width: { size: COMMENTARY_RIGHT_COL_DXA, type: WidthType.DXA }, margins: { left: 120 }, children: rightChildren })
        ]
      })]
    });

    return [subheading, layoutTable, new Paragraph({ text: '', spacing: { after: 120 } })];
  });
  return [heading, ...body];
}

function buildCountrySourcesPage(sourcesData) {
  const { commentaryGroups, segmentSources } = sourcesData;
  if (!commentaryGroups.length && !segmentSources.length) return [];

  const blocks = [new Paragraph({ text: 'Sources', heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 60 } })];

  if (commentaryGroups.length) {
    blocks.push(new Paragraph({ text: 'Commentary', heading: HeadingLevel.HEADING_3, spacing: { before: 100, after: 30 } }));
    commentaryGroups.forEach((g) => {
      blocks.push(new Paragraph({ children: [new TextRun({ text: g.sectionLabel, bold: true, size: 17 })], spacing: { before: 80, after: 20 } }));
      g.sources.forEach((s) => {
        blocks.push(new Paragraph({ text: `- ${s.label || s.url}${s.label && s.url ? ` — ${s.url}` : ''}`, spacing: { after: 15 } }));
      });
    });
  }

  if (segmentSources.length) {
    blocks.push(new Paragraph({ text: 'Segment data (AUM & asset allocation)', heading: HeadingLevel.HEADING_3, spacing: { before: 140, after: 30 } }));
    segmentSources.forEach((s) => {
      blocks.push(new Paragraph({ text: `${s.segmentName}: ${s.citation}`, spacing: { after: 15 } }));
    });
  }

  return blocks;
}

const ALL_CONTENT_TYPES = ['commentary', 'aum', 'scorecard', 'top_institutions'];
function resolveInclude(rawInclude) {
  const valid = Array.isArray(rawInclude) ? rawInclude.filter((k) => ALL_CONTENT_TYPES.includes(k)) : [];
  return new Set(valid.length ? valid : ALL_CONTENT_TYPES);
}

function buildCountrySection(countryName, segments, { headingLevel = HeadingLevel.HEADING_1, pageBreakBefore = false, enabledDimensions, include, weightOverrides, commentary, developments, allocType, allocStyle, customDimensions } = {}) {
  const includeSet = include || new Set(ALL_CONTENT_TYPES);
  const body = [];

  const commentarySections = includeSet.has('commentary')
    ? buildCommentarySectionsFull(commentary, segments, developments)
    : [];

  if (includeSet.has('commentary')) {
    body.push(...buildCommentaryBlock(commentarySections));
  }
  if (includeSet.has('aum')) {
    const aumRows = buildAumRows(segments, allocType, allocStyle);
    body.push(
      new Paragraph({ text: 'AUM by segment', heading: HeadingLevel.HEADING_2, spacing: { before: 150, after: 60 } }),
      buildAumStatCards(computeAumStats(aumRows)),
      new Paragraph({ text: '', spacing: { after: 120 } }),
      buildAumTable(aumRows, allocType)
    );
  }
  if (includeSet.has('scorecard')) {
    const matrix = buildScorecardMatrix(segments, enabledDimensions, weightOverrides, allocType, allocStyle, customDimensions);
    body.push(
      new Paragraph({ text: 'Opportunity scorecard', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } }),
      buildScorecardTable(matrix),
      buildScoreLegend()
    );
  }
  if (includeSet.has('top_institutions')) {
    body.push(...buildTopInstitutionsBlock(segments));
  }

  if (includeSet.has('commentary') || includeSet.has('aum')) {
    const sourcesData = buildCountrySourcesData(commentarySections, segments);
    body.push(...buildCountrySourcesPage(sourcesData));
  }

  if (!body.length) return [];

  const heading = new Paragraph({
    heading: headingLevel,
    children: [
      ...(pageBreakBefore ? [new PageBreak()] : []),
      new TextRun({ text: `Atlas — ${countryName}` })
    ]
  });
  return [heading, ...body];
}

function buildEvidenceAppendix(evidence) {
  if (!Array.isArray(evidence) || !evidence.length) return [];

  const heading = new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new PageBreak(), new TextRun({ text: 'Appendix — Supporting Evidence' })]
  });
  const intro = new Paragraph({
    text: 'Cross-cutting reference material (fee benchmarks, market-structure surveys) that applies across countries or regions, attached to this export separately from the country data above.',
    spacing: { after: 150 }
  });

  const body = evidence.flatMap((e) => {
    const subheading = new Paragraph({ text: e.title || e.id || 'Untitled entry', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 40 } });
    const metaLine = new Paragraph({
      children: [new TextRun({
        text: `${e.theme ? `Theme: ${e.theme}` : ''}${e.scope && e.scope.length ? `  |  Scope: ${e.scope.join(', ')}` : ''}`,
        italics: true, size: 18
      })],
      spacing: { after: 60 }
    });
    const summary = e.summary ? [new Paragraph({ text: e.summary, spacing: { after: 80 } })] : [];

    const figureRows = Array.isArray(e.figures) && e.figures.length ? [
      new Paragraph({ children: [new TextRun({ text: 'Figures', italics: true, size: 18 })], spacing: { before: 40, after: 20 } }),
      (() => {
        const figWidths = [1800, 2600, 1800, 1600];
        return new Table({
          columnWidths: figWidths,
          width: { size: figWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({ children: [headerCell('Region', figWidths[0]), headerCell('Metric', figWidths[1]), headerCell('Value', figWidths[2]), headerCell('As of', figWidths[3])] }),
            ...e.figures.map((f) => new TableRow({ children: [
              bodyCell(f.region || '-', null, figWidths[0]),
              bodyCell(f.metric || '-', null, figWidths[1]),
              bodyCell(`${typeof f.value === 'number' ? f.value : '-'}${f.unit ? ' ' + f.unit : ''}`, null, figWidths[2]),
              bodyCell(f.as_of || '-', null, figWidths[3])
            ] }))
          ]
        });
      })()
    ] : [];

    const note = e.note ? [new Paragraph({ children: [new TextRun({ text: e.note, italics: true, size: 18 })], spacing: { before: 60, after: 20 } })] : [];
    const sourceLine = new Paragraph({
      children: [new TextRun({ text: `Source: ${e.source || 'unknown'} (${e.as_of || 'date unknown'})${e.access ? ` — ${e.access}` : ''}`, size: 16 })],
      spacing: { before: 40, after: 100 }
    });

    return [subheading, metaLine, ...summary, ...figureRows, ...note, sourceLine];
  });

  return [heading, intro, ...body];
}

// Document-wide style overrides so headings read as the Atlas navy palette
// instead of Word's default blue theme -- new 2026-08-07. Only colour/weight
// is touched; size/spacing keep Word's normal heading scale so the doc still
// looks like a standard Word document (and paginates/prints normally),
// just recoloured to match the site and the PowerPoint export.
const DOC_STYLES = {
  default: {
    title: { run: { color: C.navy, bold: true, size: 48 } },
    heading1: { run: { color: C.navy, bold: true }, paragraph: { spacing: { before: 240, after: 120 } } },
    heading2: { run: { color: C.navy, bold: true }, paragraph: { spacing: { before: 200, after: 100 } } },
    heading3: { run: { color: C.teal, bold: true }, paragraph: { spacing: { before: 160, after: 80 } } }
  }
};

app.http('exportDocx', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'export/docx',
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

    const title = isMulti ? `Institutional Market Report — ${countries.length} countries` : `Institutional Market Report — ${countries[0].country_name}`;
    const safeName = isMulti
      ? `Project_${countries.length}_countries`
      : countries[0].country_name.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

    try {
      const generatedDate = new Date().toISOString().slice(0, 10);

      const subtitle = isMulti ? countries.map((c) => c.country_name).join(', ') : 'Country deep-dive';
      const coverageLine = `Covers: country commentary, AUM by segment, opportunity scorecard, top institutions by AUM, and full sourcing — for ${isMulti ? `the ${countries.length} selected countries` : countries[0].country_name} only.`;

      const children = [
        ...buildCoverPage({ title, subtitle, coverageLine, generatedDate }),
        ...buildTocPage(),
        ...buildMethodologyPage()
      ];

      const include = resolveInclude(body.include);
      let emittedCount = 0;
      countries.forEach((c) => {
        const section = buildCountrySection(c.country_name, c.segments, {
          headingLevel: HeadingLevel.HEADING_1,
          pageBreakBefore: emittedCount > 0,
          enabledDimensions: body.enabled_dimensions,
          include,
          weightOverrides: body.weight_overrides,
          customDimensions: body.custom_dimensions,
          commentary: c.commentary,
          developments: c.developments,
          allocType: body.alloc_type,
          allocStyle: body.alloc_style
        });
        if (!section.length) return;
        children.push(...section);
        emittedCount += 1;
      });

      if (!emittedCount) {
        return { status: 400, jsonBody: { error: 'None of the selected countries have data for the requested content type(s).' } };
      }

      children.push(...buildEvidenceAppendix(body.evidence));

      const logoHeader = buildLogoHeader();
      const doc = new Document({
        styles: DOC_STYLES,
        // Forces Word to auto-update every field (in practice, just the
        // table of contents) the first time the document is opened, rather
        // than showing an empty TOC until the user manually right-clicks >
        // Update Field. docx can't compute real page numbers itself at
        // generation time (no layout engine), so the field's cached result
        // is always empty coming out of this API -- this setting is what
        // makes Word fill it in automatically on open instead of leaving
        // that as a manual step.
        features: { updateFields: true },
        sections: [{
          // properties.titlePage + headers.first -- 2026-08-07, Peter: the
          // small running-header logo badge (top-left of every page, via
          // logoHeader below) shouldn't also appear on the cover page,
          // which already carries the large lockup image lower down (see
          // buildCoverPage()). Word's "different first page" mechanism is
          // the standard way to do this: titlePage:true splits page 1's
          // header/footer out from the rest of the document, and giving it
          // an empty Header means page 1 gets none at all rather than
          // falling back to `default`.
          properties: { titlePage: true },
          ...(logoHeader ? { headers: { default: logoHeader, first: new Header({ children: [] }) } } : {}),
          // footers.first repeats the same footer on page 1 -- titlePage
          // splits header AND footer references together, and without an
          // explicit `first` footer here page 1 would silently lose its
          // footer (page number/confidentiality line) as a side effect of
          // suppressing just the header logo above, which isn't the intent.
          footers: { default: buildFooter(), first: buildFooter() },
          children
        }]
      });
      const buffer = await Packer.toBuffer(doc);

      return {
        status: 200,
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="Atlas_${safeName}.docx"`
        },
        body: buffer
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Failed to generate Word document', detail: err.message } };
    }
  }
});
