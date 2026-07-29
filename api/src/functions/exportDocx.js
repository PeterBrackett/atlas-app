const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak, Header, AlignmentType, BorderStyle
} = require('docx');
const {
  buildAumRows, buildScorecardMatrix, buildCommentarySectionsFull, buildTopInstitutionsSections,
  estimateColumnCharWidths, buildSegmentAllocationChart, buildCountrySourcesData
} = require('../shared/exportHelpers');
const { getDimensionIconBuffer } = require('../shared/dimensionIcons');
const { getAtlasLogoBuffer } = require('../shared/atlasLogo');

const HEADER_FILL = 'D9E2F3';

// Compact cell margins (dxa/twips) and font sizes (half-points), applied
// throughout the export tables so more tables fit per page -- Peter's
// "there's quite a bit of wasted space" feedback on the default docx table
// styling, which otherwise uses ~11pt text and generous default padding.
const CELL_MARGINS = { top: 30, bottom: 30, left: 80, right: 80 };
const HEADER_FONT_SIZE = 16; // 8pt
const BODY_FONT_SIZE = 16; // 8pt

// Converts an estimateColumnCharWidths() character count into a dxa
// (twips) column width, so tables are sized to their content rather than
// stretched to fill the page -- Peter's follow-up feedback that reducing
// font size alone just left more blank padding around short cells like
// "Rank" or a 1-3 score, since the columns were still evenly dividing 100%
// of the page width. ~105 twips/char is a rough proportional-font average
// at 8pt; the cell's left+right margins are added on top.
const CHAR_WIDTH_TWIPS = 105;
function charsToDxa(chars) {
  return Math.round(chars * CHAR_WIDTH_TWIPS) + CELL_MARGINS.left + CELL_MARGINS.right;
}

function headerCell(text, widthDxa) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    margins: CELL_MARGINS,
    ...(widthDxa ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: true, size: HEADER_FONT_SIZE })] })]
  });
}

// `color` is the optional {bg, fg} hex pair from exportHelpers.js's
// scoreColor()/overallColor(), reproducing the site's red/amber/green
// scorecard traffic-light coding (see style.css's td.score-1/2/3 and
// td.overall-red/amber/green) in the Word table cells.
function bodyCell(text, color, widthDxa) {
  return new TableCell({
    margins: CELL_MARGINS,
    ...(color ? { shading: { type: ShadingType.CLEAR, fill: color.bg, color: 'auto' } } : {}),
    ...(widthDxa ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), size: BODY_FONT_SIZE, ...(color ? { color: color.fg, bold: true } : {}) })]
    })]
  });
}

// Same as headerCell(), but for a scorecard dimension row: prepends the
// dimension's icon (see dimensionIcons.js, rasterized from the same shapes
// used on country.html) before the label text, when one exists for that
// dimension key. Falls back to a plain text header cell if the icon lookup
// comes up empty, so a future dimension added without an icon still renders
// rather than breaking the export.
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
        new TextRun({ text: `  ${row.label}`, bold: true, size: HEADER_FONT_SIZE })
      ]
    })]
  });
}

// Atlas logo, top-left of every page. A docx Header attached to the
// document's (single) section repeats automatically on every page, so this
// only needs to be built once rather than re-inserted per country/section.
function buildLogoHeader() {
  const logoBuffer = getAtlasLogoBuffer();
  if (!logoBuffer) return undefined;
  return new Header({
    children: [new Paragraph({
      children: [new ImageRun({ data: logoBuffer, type: 'png', transformation: { width: 28, height: 28 } })]
    })]
  });
}

// "Equities range (min-max)" reflects that not every institution counted in
// a segment's AUM also filed an asset-class breakdown -- min is the
// reported Equities figure as-is (assumes non-reporters hold none), max is
// that figure scaled up to the segment's full AUM (assumes non-reporters
// match reporters' mix). See getAllocationRange() in exportHelpers.js.
// Column widths are content-driven (see estimateColumnCharWidths() in
// exportHelpers.js) rather than fixed percentages of the page, so e.g.
// "AUM ($bn)" doesn't reserve more room than its numbers ever use. Basis
// and the Equities range string are the two columns most likely to run
// long, so they get the highest character caps (and wrap, rather than
// stretching the table further, past that).
function buildAumTable(rows) {
  const headerLabels = ['Segment', 'AUM ($bn)', 'Equities ($bn)', 'Basis', 'Equities range (min-max)'];
  const bodyText = rows.map((r) => [
    r.segment,
    typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    typeof r.equity_bn === 'number' ? r.equity_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-',
    r.basis || '',
    r.equity_range || '-'
  ]);
  const widths = estimateColumnCharWidths(headerLabels, bodyText, {
    minChars: 4,
    maxCharsPerCol: [26, 10, 10, 22, 30]
  }).map(charsToDxa);

  const headerRow = new TableRow({
    children: headerLabels.map((label, i) => headerCell(label, widths[i]))
  });
  const dataRows = bodyText.map((cells) => new TableRow({
    children: cells.map((text, i) => bodyCell(text, null, widths[i]))
  }));
  return new Table({ rows: [headerRow, ...dataRows] });
}

// Word tables don't scroll horizontally the way a webpage can, so a country
// with many segment columns (e.g. UK's 11) will run wide. Column widths are
// content-driven the same way as buildAumTable() -- these are all
// single-digit scores or short "x/12" strings, so they get a small char
// cap and the table only ever runs as wide as it needs to. Cells carry the
// same red/amber/green shading as the site's scorecard matrix, via
// row.colors[i] (see scoreColor()/overallColor() in exportHelpers.js).
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
  return new Table({ rows: [headerRow, ...dataRows] });
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
  const dataRows = bodyText.map((cells) => new TableRow({
    children: cells.map((text, i) => bodyCell(text, null, widths[i]))
  }));
  return new Table({ rows: [headerRow, ...dataRows] });
}

// The per-segment part of the "top institutions" block -- one H3 heading +
// one small table per segment that has institution-level data. Segments
// built from industry aggregates (e.g. Life/Non-life insurance) or
// countries not yet backfilled at institution level (currently just the US)
// are skipped, not guessed at -- see buildTopInstitutionsSections in
// exportHelpers.js.
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

// One heading + one small table per segment that has institution-level data
// -- Peter's standard "top 10 institutions by AUM, and their combined AUM as
// a % of the segment" report format. Used within a per-country section (see
// buildCountrySection() below); returns [] for a country with no
// institution-level data at all, rather than an empty heading.
function buildTopInstitutionsBlock(segments) {
  const sections = buildTopInstitutionsSections(segments);
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Top institutions by AUM', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  return [heading, ...buildTopInstitutionsPerSegment(sections)];
}

// Word has no native chart object the way PowerPoint does (see
// exportPptx.js's addChart() usage) -- docx's Table/TableCell primitives are
// the only thing available, so a segment's asset-class mix renders as a
// single-row table whose cell widths are proportional to each slice's share
// of the segment's full aum_bn, shaded with the same asset-type colors used
// everywhere else (scorecard-dimensions.js's ASSET_TYPE_COLORS / this file's
// mirror in exportHelpers.js) -- a horizontal stacked bar built out of table
// cells, in effect. Every slice gets a minimum width (ALLOCATION_BAR_MIN_DXA)
// so a very small slice (e.g. a 0.5% Equities sliver) still shows as a
// visible sliver rather than vanishing at 0 width; this means the bar's
// total width sums to slightly more than "100% worth" of dxa when there are
// many tiny slices, a deliberate legibility trade-off over exact proportionality.
// Cut to 30% of the original width (8000 dxa, ~5.5in) per Peter's 2026-07-23
// request to shrink the charts by 70% -- all three constants scaled down
// together so the minimum-slice-width and label-threshold logic still holds
// the same relative proportions at the smaller size.
const ALLOCATION_BAR_TOTAL_DXA = 2400; // ~1.65in
const ALLOCATION_BAR_MIN_DXA = 60;
const ALLOCATION_BAR_LABEL_MIN_DXA = 210; // only label a cell wide enough to hold "12.3%"

function buildAllocationBarTable(chart) {
  const n = chart.slices.length;
  const remaining = Math.max(ALLOCATION_BAR_TOTAL_DXA - ALLOCATION_BAR_MIN_DXA * n, 0);
  const cells = chart.slices.map((s) => {
    const widthDxa = Math.round(ALLOCATION_BAR_MIN_DXA + (s.pct / 100) * remaining);
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
  return new Table({ rows: [new TableRow({ children: cells })], width: { size: ALLOCATION_BAR_TOTAL_DXA, type: WidthType.DXA } });
}

// A small colored square (an actual glyph, colored via TextRun's `color` --
// docx has no inline "background swatch" primitive) next to each slice's
// label and percentage, wrapping onto one flowing paragraph rather than a
// second table, since it's just a label key for the bar above it.
function buildAllocationLegendParagraph(chart) {
  const runs = [];
  chart.slices.forEach((s, i) => {
    if (i > 0) runs.push(new TextRun({ text: '     ' }));
    runs.push(new TextRun({ text: '■ ', color: s.color }));
    runs.push(new TextRun({ text: `${s.label} ${s.pct}%`, size: 16 }));
  });
  return new Paragraph({ children: runs, spacing: { before: 40, after: 20 } });
}

// One segment's chart visual: subheading, bar table, legend -- no source
// line, unlike the original version of this function. Sources used to be
// printed inline right under each chart, which was part of Peter's
// 2026-07-29 "sources bleed into the text" feedback; every segment's
// citation is now gathered once via buildCountrySourcesData() and printed on
// a single consolidated Sources page at the end of the country's content
// instead (see buildCountrySourcesPage() below). Skips segments with
// nothing to chart (buildSegmentAllocationChart() returns null for a
// zero-aum or no-allocation segment) rather than emitting an empty heading.
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

// Right-hand "Asset allocation" box content for one commentary section --
// mirrors country.html's renderCommentaryAllocationBox() exactly: a plain
// placeholder note when this section has no chart slot at all (Wealth,
// Pensions, Charities, OCIO), a different placeholder when it has a slot but
// this country has no matching segment data yet, or one visual per matching
// segment when there's something to chart. Consistent per-section frame
// regardless of content, per Peter's "all 8 sections, placeholder if none"
// request.
function buildAllocationBoxContent(section) {
  if (!section.hasChartSlot) {
    return [new Paragraph({ children: [new TextRun({ text: 'Not available for this section.', ...PLACEHOLDER_RUN_OPTS })] })];
  }
  if (!section.chartSegments.length) {
    return [new Paragraph({ children: [new TextRun({ text: 'No segment-level data available yet for asset allocation in this section.', ...PLACEHOLDER_RUN_OPTS })] })];
  }
  return section.chartSegments.flatMap((seg) => buildSegmentAllocationVisual(seg));
}

// Right-hand "Recent developments" box content -- short, dated, sourced
// items from {code}_developments.json (see getDevelopments.js), the
// export-side counterpart to country.html's renderCommentaryDevelopmentsBox().
// Each item's own source line stays inline here (unlike commentary/segment
// sources) since it's a single short citation directly under a 1-2 sentence
// item, not a list interrupting a longer block of prose -- not the kind of
// "bleeding into text" Peter flagged.
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
    if (srcLabel) paras.push(new Paragraph({ children: [new TextRun({ text: `Source: ${srcLabel}${d.url ? ` — ${d.url}` : ''}`, italics: true, size: 13, color: '5B6B7A' })], spacing: { after: 40 } }));
    return paras;
  });
}

// No visible grid lines on the two-column layout table itself -- it's here
// purely to place the allocation/developments boxes to the right of the
// text, not to look like a data table the way buildAumTable()/
// buildScorecardTable() do.
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER, insideHorizontal: NO_BORDER, insideVertical: NO_BORDER };
const COMMENTARY_LEFT_COL_DXA = 5500;
const COMMENTARY_RIGHT_COL_DXA = 3500;

// One heading + one two-column layout (text left; asset allocation and
// recent developments stacked right) per populated commentary section
// (Wealth & key pools of capital, Pensions structure, Insurance, etc.) --
// see buildCommentarySectionsFull() in exportHelpers.js for the
// text-splitting/source-filtering/chart-matching/developments rules. A
// section is skipped only if it has NEITHER drafted text, NOR a matching
// chart segment, NOR any developments logged. Rebuilt 2026-07-29 from a flat
// vertical stack into this two-column table layout, mirroring the same
// change made to country.html's on-screen rendering -- `sections` is now
// passed in pre-built (buildCommentarySectionsFull() output) rather than
// computed here, since the caller (buildCountrySection()) also needs it to
// build the end-of-country Sources page.
function buildCommentaryBlock(sections) {
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Country commentary', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  const body = sections.flatMap((section) => {
    const subheading = new Paragraph({ text: section.label, heading: HeadingLevel.HEADING_3, spacing: { before: 150, after: 40 } });

    const leftChildren = section.paragraphs.length
      ? section.paragraphs.map((p) => new Paragraph({ text: p, spacing: { after: 100 } }))
      : [new Paragraph({ children: [new TextRun({ text: `No commentary written yet for ${section.label}.`, ...PLACEHOLDER_RUN_OPTS })] })];

    const rightChildren = [
      new Paragraph({ children: [new TextRun({ text: 'Asset allocation', bold: true, size: 16, color: '0F2540' })], spacing: { after: 40 } }),
      ...buildAllocationBoxContent(section),
      new Paragraph({ text: '', spacing: { before: 80 } }),
      new Paragraph({ children: [new TextRun({ text: 'Recent developments', bold: true, size: 16, color: '0F2540' })], spacing: { after: 40 } }),
      ...buildDevelopmentsBoxContent(section.developments)
    ];

    const layoutTable = new Table({
      borders: NO_BORDERS,
      columnWidths: [COMMENTARY_LEFT_COL_DXA, COMMENTARY_RIGHT_COL_DXA],
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

// Consolidated end-of-country Sources page -- every commentary section's own
// citations, plus every segment's AUM/allocation citation, gathered by
// buildCountrySourcesData() (exportHelpers.js) and printed once here instead
// of interrupting the text and charts above with citation lists. Added
// 2026-07-29 per Peter's feedback. Returns [] when there's nothing to show
// (e.g. a fresh country with no sources logged anywhere yet), so callers
// don't need to check emptiness themselves.
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

// Which of the four content blocks a country section should include --
// 'commentary', 'aum', 'scorecard', 'top_institutions'. Defaults to all four
// (the original "everything" behaviour, plus commentary added 2026-07-23)
// when not specified, so country.html's existing single-country export
// (which never sends `include`) is unaffected. `include` itself was added
// 2026-07-16 per Peter's request to be able to export e.g. "only scorecards
// of the countries I select" rather than always the full bundle.
const ALL_CONTENT_TYPES = ['commentary', 'aum', 'scorecard', 'top_institutions'];
function resolveInclude(rawInclude) {
  const valid = Array.isArray(rawInclude) ? rawInclude.filter((k) => ALL_CONTENT_TYPES.includes(k)) : [];
  return new Set(valid.length ? valid : ALL_CONTENT_TYPES);
}

// One country's section: heading + whichever of AUM table / scorecard table
// / top-institutions-by-segment the caller asked for (see `include` above).
// Shared by both the single-country payload (country.html's per-page
// export) and the multi-country payload (picker.html's project builder) so
// a project export is just this block repeated once per selected country,
// rather than a separate document layout to maintain. Returns [] if the
// country ends up with nothing to show under the requested `include` set
// (e.g. `include` is top_institutions-only and this country has no
// institution-level data) -- the caller should skip a country entirely in
// that case rather than emit an empty heading.
function buildCountrySection(countryName, segments, { headingLevel = HeadingLevel.HEADING_1, pageBreakBefore = false, enabledDimensions, include, weightOverrides, commentary, developments, allocType, allocStyle } = {}) {
  const includeSet = include || new Set(ALL_CONTENT_TYPES);
  const body = [];

  // Built once here (rather than inside buildCommentaryBlock()) since the
  // end-of-country Sources page below also needs each section's citations --
  // computed even when 'commentary' isn't included, so a commentary-only
  // export still gets... actually only computed when needed, to avoid doing
  // work for content that was never asked for.
  const commentarySections = includeSet.has('commentary')
    ? buildCommentarySectionsFull(commentary, segments, developments)
    : [];

  if (includeSet.has('commentary')) {
    body.push(...buildCommentaryBlock(commentarySections));
  }
  if (includeSet.has('aum')) {
    body.push(
      new Paragraph({ text: 'AUM by segment', heading: HeadingLevel.HEADING_2, spacing: { before: 150, after: 60 } }),
      buildAumTable(buildAumRows(segments))
    );
  }
  if (includeSet.has('scorecard')) {
    body.push(
      new Paragraph({ text: 'Opportunity scorecard', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } }),
      buildScorecardTable(buildScorecardMatrix(segments, enabledDimensions, weightOverrides, allocType, allocStyle))
    );
  }
  if (includeSet.has('top_institutions')) {
    body.push(...buildTopInstitutionsBlock(segments));
  }

  // Consolidated Sources page -- only when commentary and/or AUM were
  // actually included, since those are the two content types that used to
  // carry inline citations; a scorecard/top-10-only export has nothing to
  // consolidate. See buildCountrySourcesPage()/buildCountrySourcesData().
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

// Optional appendix built from picker.html's "Attach supporting evidence"
// picker (see evidence.html / evidence_library.json) -- cross-cutting
// reference material like fee-benchmark surveys that applies across
// countries/regions rather than to one segment of one country, so it can't
// live inside buildCountrySection() above. `evidence` is the array of full
// entry objects the client already has (picker.html sends them directly,
// see downloadProjectExport() there) -- this function does no fetching of
// its own, just formatting. Returns [] for an empty/missing array, so a
// request that never sends `evidence` (e.g. country.html's single-country
// export) is completely unaffected -- purely additive.
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
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [headerCell('Region'), headerCell('Metric'), headerCell('Value'), headerCell('As of')] }),
          ...e.figures.map((f) => new TableRow({ children: [
            bodyCell(f.region || '-'),
            bodyCell(f.metric || '-'),
            bodyCell(`${typeof f.value === 'number' ? f.value : '-'}${f.unit ? ' ' + f.unit : ''}`),
            bodyCell(f.as_of || '-')
          ] }))
        ]
      })
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

    // Two accepted shapes: the original single-country payload from
    // country.html ({country_name, segments}), and picker.html's
    // multi-country project payload ({countries: [{country_name, segments}, ...]}).
    const isMulti = Array.isArray(body.countries);
    const countries = isMulti
      ? body.countries.filter((c) => c && Array.isArray(c.segments) && c.segments.length)
      : (Array.isArray(body.segments) && body.segments.length ? [{ country_name: body.country_name || 'Country', segments: body.segments, commentary: body.commentary, developments: body.developments }] : []);

    if (!countries.length) {
      return { status: 400, jsonBody: { error: 'No segments provided to export' } };
    }

    const docTitle = isMulti ? `Atlas — Project (${countries.length} countries)` : `Atlas — ${countries[0].country_name}`;
    const safeName = isMulti
      ? `Project_${countries.length}_countries`
      : countries[0].country_name.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

    try {
      const generatedDate = new Date().toISOString().slice(0, 10);

      const children = [];
      if (isMulti) {
        children.push(new Paragraph({ text: docTitle, heading: HeadingLevel.TITLE }));
        children.push(new Paragraph({ text: `Generated ${generatedDate} — ${countries.map((c) => c.country_name).join(', ')}` }));
      } else {
        children.push(new Paragraph({ text: `Generated ${generatedDate}` }));
      }

      // Which content types to include per country -- 'aum', 'scorecard',
      // 'top_institutions', any combination, defaulting to all three.
      // Peter's 2026-07-16 request: be able to export e.g. only scorecards,
      // or only top 10s, of the selected countries, rather than always the
      // full bundle. A country that ends up with nothing to show under the
      // requested set (e.g. top_institutions-only, and this country has no
      // institution-level data) is skipped entirely -- pageBreakBefore
      // therefore tracks the first country actually emitted, not raw array
      // position, so a skipped country doesn't leave a stray leading blank
      // page.
      const include = resolveInclude(body.include);
      let emittedCount = 0;
      countries.forEach((c) => {
        const section = buildCountrySection(c.country_name, c.segments, {
          headingLevel: HeadingLevel.HEADING_1,
          pageBreakBefore: isMulti && emittedCount > 0,
          enabledDimensions: body.enabled_dimensions,
          include,
          // weight_overrides -- picker.html's project builder weighting
          // column (see exportHelpers.js's computeOverallScore() comment).
          // Optional; country.html's single-country export never sends
          // this, so Overall there is unaffected.
          weightOverrides: body.weight_overrides,
          commentary: c.commentary,
          developments: c.developments,
          // alloc_type/alloc_style -- 2026-07-23, matches whatever the
          // "Allocation row shows" dropdown was set to on-screen when the
          // export was triggered (country.html's CURRENT_ALLOC_TYPE/
          // CURRENT_ALLOC_STYLE, picker.html's PROJECT_ALLOC_TYPE/
          // PROJECT_ALLOC_STYLE). Defaults to Equities/no-style in
          // buildScorecardMatrix() if omitted.
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

      // Supporting-evidence appendix -- see buildEvidenceAppendix() above.
      // body.evidence is only ever populated by picker.html's "Attach
      // supporting evidence" picker; country.html's single-country export
      // never sends it, so this is a no-op there.
      children.push(...buildEvidenceAppendix(body.evidence));

      children.push(new Paragraph({
        text: 'Source: Atlas. See Sources & Methodology on the site for how these figures are derived.',
        spacing: { before: 300 }
      }));

      const logoHeader = buildLogoHeader();
      const doc = new Document({
        sections: [{
          ...(logoHeader ? { headers: { default: logoHeader } } : {}),
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
