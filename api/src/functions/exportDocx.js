const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak, Header
} = require('docx');
const { buildAumRows, buildScorecardMatrix, buildTopInstitutionsSections } = require('../shared/exportHelpers');
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

function headerCell(text, widthPct) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    margins: CELL_MARGINS,
    ...(widthPct ? { width: { size: widthPct, type: WidthType.PERCENTAGE } } : {}),
    children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: true, size: HEADER_FONT_SIZE })] })]
  });
}

// `color` is the optional {bg, fg} hex pair from exportHelpers.js's
// scoreColor()/overallColor(), reproducing the site's red/amber/green
// scorecard traffic-light coding (see style.css's td.score-1/2/3 and
// td.overall-red/amber/green) in the Word table cells.
function bodyCell(text, color) {
  return new TableCell({
    margins: CELL_MARGINS,
    ...(color ? { shading: { type: ShadingType.CLEAR, fill: color.bg, color: 'auto' } } : {}),
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
function dimensionHeaderCell(row) {
  const iconBuffer = row.key ? getDimensionIconBuffer(row.key) : null;
  if (!iconBuffer) return headerCell(row.label);
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    margins: CELL_MARGINS,
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
// Column widths tuned to content (numeric columns don't need as much room
// as Segment/Basis text), rather than five equal-width columns leaving
// visible whitespace in the short numeric ones.
const AUM_COL_WIDTHS = [22, 12, 12, 16, 38];

function buildAumTable(rows) {
  const headerRow = new TableRow({
    children: [
      headerCell('Segment', AUM_COL_WIDTHS[0]),
      headerCell('AUM ($bn)', AUM_COL_WIDTHS[1]),
      headerCell('Equities ($bn)', AUM_COL_WIDTHS[2]),
      headerCell('Basis', AUM_COL_WIDTHS[3]),
      headerCell('Equities range (min-max)', AUM_COL_WIDTHS[4])
    ]
  });
  const dataRows = rows.map((r) => new TableRow({
    children: [
      bodyCell(r.segment),
      bodyCell(typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'),
      bodyCell(typeof r.equity_bn === 'number' ? r.equity_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'),
      bodyCell(r.basis || ''),
      bodyCell(r.equity_range || '-')
    ]
  }));
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

// Word tables don't scroll horizontally the way a webpage can, so a country
// with many segment columns (e.g. UK's 11) will run wide. Deliberately left
// at AUTO width (rather than stretched to 100% of the page) so Word sizes
// each data column to its actual content -- these are all single-digit
// scores, so a full-width table here was mostly empty padding. Cells carry
// the same red/amber/green shading as the site's scorecard matrix, via
// row.colors[i] (see scoreColor()/overallColor() in exportHelpers.js).
function buildScorecardTable(matrix) {
  const headerRow = new TableRow({
    children: [headerCell('Dimension'), ...matrix.columnLabels.map((label) => headerCell(label))]
  });
  const dataRows = matrix.rows.map((row) => new TableRow({
    children: [dimensionHeaderCell(row), ...row.values.map((v, i) => bodyCell(v, row.colors ? row.colors[i] : null))]
  }));
  return new Table({ rows: [headerRow, ...dataRows] });
}

function buildTopInstitutionsTable(section) {
  const headerRow = new TableRow({
    children: [headerCell('Rank', 12), headerCell('Institution', 58), headerCell('AUM ($bn)', 30)]
  });
  const dataRows = section.institutions.map((inst, i) => new TableRow({
    children: [
      bodyCell(String(i + 1)),
      bodyCell(inst.name),
      bodyCell(typeof inst.aum_bn === 'number' ? inst.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-')
    ]
  }));
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

// One heading + one small table per segment that has institution-level data
// -- Peter's standard "top 10 institutions by AUM, and their combined AUM as
// a % of the segment" report format. Segments built from industry
// aggregates (e.g. Life/Non-life insurance) or countries not yet backfilled
// at institution level (currently just the US) are skipped, not guessed at
// -- see buildTopInstitutionsSections in exportHelpers.js.
function buildTopInstitutionsBlock(segments) {
  const sections = buildTopInstitutionsSections(segments);
  if (!sections.length) return [];

  const heading = new Paragraph({ text: 'Top institutions by AUM', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } });
  const perSegment = sections.flatMap((section) => {
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
  return [heading, ...perSegment];
}

// One country's section: heading + AUM table + scorecard table + top
// institutions by segment. Shared by both the single-country payload
// (country.html's per-page export) and the multi-country payload
// (picker.html's project builder) so a project export is just this block
// repeated once per selected country, rather than a separate document
// layout to maintain.
function buildCountrySection(countryName, segments, { headingLevel = HeadingLevel.HEADING_1, pageBreakBefore = false } = {}) {
  const aumRows = buildAumRows(segments);
  const matrix = buildScorecardMatrix(segments);
  const heading = new Paragraph({
    heading: headingLevel,
    children: [
      ...(pageBreakBefore ? [new PageBreak()] : []),
      new TextRun({ text: `Atlas — ${countryName}` })
    ]
  });
  return [
    heading,
    new Paragraph({ text: 'AUM by segment', heading: HeadingLevel.HEADING_2, spacing: { before: 150, after: 60 } }),
    buildAumTable(aumRows),
    new Paragraph({ text: 'Opportunity scorecard', heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 60 } }),
    buildScorecardTable(matrix),
    ...buildTopInstitutionsBlock(segments)
  ];
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
      : (Array.isArray(body.segments) && body.segments.length ? [{ country_name: body.country_name || 'Country', segments: body.segments }] : []);

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

      countries.forEach((c, i) => {
        children.push(...buildCountrySection(c.country_name, c.segments, {
          headingLevel: isMulti ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_1,
          pageBreakBefore: isMulti && i > 0
        }));
      });

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
