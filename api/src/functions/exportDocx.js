const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak, Header, AlignmentType
} = require('docx');
const {
  buildAumRows, buildScorecardMatrix, buildCommentarySections, buildTopInstitutionsSections,
  estimateColumnCharWidths, commentarySectionChartSegments, buildSegmentAllocationChart
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
const ALLOCATION_BAR_TOTAL_DXA = 8000; // ~5.5in
const ALLOCATION_BAR_MIN_DXA = 200;
const ALLOCATION_BAR_LABEL_MIN_DXA = 700; // only label a cell wide enough to hold "12.3%"

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

// One segment's chart block: subheading, bar table, legend, and its own
// source line -- this is the segment's AUM/allocation provenance
// (sources[].active), not one of the section's own commentary sources, same
// distinction country.html's donut cards draw. Skips segments with nothing
// to chart (buildSegmentAllocationChart() returns null for a zero-aum or
// no-allocation segment) rather than emitting an empty heading.
function buildSegmentAllocationBlock(segment) {
  const chart = buildSegmentAllocationChart(segment);
  if (!chart) return [];
  return [
    new Paragraph({ text: `${chart.segmentName} — asset allocation`, heading: HeadingLevel.HEADING_4, spacing: { before: 120, after: 30 } }),
    buildAllocationBarTable(chart),
    buildAllocationLegendParagraph(chart),
    new Paragraph({
      children: [new TextRun({ text: `Source: ${chart.sourceText || 'not recorded for this segment'}`, italics: true, size: 14, color: '5B6B7A' })],
      spacing: { after: 100 }
    })
  ];
}

// One heading + one subheading/paragraphs/asset-allocation-charts/sources
// block per populated commentary section (Wealth & key pools of capital,
// Pensions structure, Insurance, etc.) -- see buildCommentarySections() in
// exportHelpers.js for the text-splitting and source-filtering rules.
// Returns [] if the country has no commentary text at all yet (most
// countries, until written), same "contribute nothing rather than an empty
// heading" convention as buildTopInstitutionsBlock(). `segments` is needed
// alongside `commentary` so sections with chartSegments (Insurance,
// Foundations, Sovereign wealth funds) can look up this country's matching
// segment(s) for their asset-allocation chart -- added 2026-07-23 alongside
// the chart feature itself, per Peter's "everything needs to be included in
// the outputs" (the chart was on-screen only until then).
function buildCommentaryBlock(commentary, segments) {
  const sections = buildCommentarySections(commentary);
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Country commentary', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  const body = sections.flatMap((section) => {
    const subheading = new Paragraph({ text: section.label, heading: HeadingLevel.HEADING_3, spacing: { before: 150, after: 40 } });
    const paragraphs = section.paragraphs.map((p) => new Paragraph({ text: p, spacing: { after: 100 } }));

    const chartSegments = commentarySectionChartSegments(section.key, segments);
    const chartBlocks = chartSegments.flatMap((seg) => buildSegmentAllocationBlock(seg));

    const sourcesBlock = section.sources.length ? [
      new Paragraph({ children: [new TextRun({ text: 'Sources', italics: true, size: 18 })], spacing: { before: 60, after: 20 } }),
      ...section.sources.map((s) => new Paragraph({
        // Plain "- " prefix rather than docx's bullet:{level} numbering
        // feature -- the latter needs an explicit numbering config on the
        // Document to render reliably, which nothing else in this file uses,
        // so a literal dash keeps this consistent with the rest of the export
        // rather than risking an unstyled/silently-ignored bullet.
        text: `- ${s.label || s.url}${s.label && s.url ? ` — ${s.url}` : ''}`,
        spacing: { after: 20 }
      }))
    ] : [];
    return [subheading, ...paragraphs, ...chartBlocks, ...sourcesBlock];
  });
  return [heading, ...body];
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
function buildCountrySection(countryName, segments, { headingLevel = HeadingLevel.HEADING_1, pageBreakBefore = false, enabledDimensions, include, weightOverrides, commentary } = {}) {
  const includeSet = include || new Set(ALL_CONTENT_TYPES);
  const body = [];

  if (includeSet.has('commentary')) {
    body.push(...buildCommentaryBlock(commentary, segments));
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
      buildScorecardTable(buildScorecardMatrix(segments, enabledDimensions, weightOverrides))
    );
  }
  if (includeSet.has('top_institutions')) {
    body.push(...buildTopInstitutionsBlock(segments));
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
      : (Array.isArray(body.segments) && body.segments.length ? [{ country_name: body.country_name || 'Country', segments: body.segments, commentary: body.commentary }] : []);

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
          commentary: c.commentary
        });
        if (!section.length) return;
        children.push(...section);
        emittedCount += 1;
      });

      if (!emittedCount) {
        return { status: 400, jsonBody: { error: 'None of the selected countries have data for the requested content type(s).' } };
      }

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
