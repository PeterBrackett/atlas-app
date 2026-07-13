const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak
} = require('docx');
const { buildAumRows, buildScorecardMatrix } = require('../shared/exportHelpers');

const HEADER_FILL = 'D9E2F3';

function headerCell(text) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL, color: 'auto' },
    children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: true })] })]
  });
}

function bodyCell(text) {
  return new TableCell({ children: [new Paragraph({ text: String(text) })] });
}

function buildAumTable(rows) {
  const headerRow = new TableRow({
    children: [headerCell('Segment'), headerCell('AUM ($bn)'), headerCell('Equities ($bn)'), headerCell('Basis')]
  });
  const dataRows = rows.map((r) => new TableRow({
    children: [
      bodyCell(r.segment),
      bodyCell(typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'),
      bodyCell(typeof r.equity_bn === 'number' ? r.equity_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-'),
      bodyCell(r.basis || '')
    ]
  }));
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

// Word tables don't scroll horizontally the way a webpage can, so a country
// with many segment columns (e.g. UK's 11) will run wide. Accepted for v1 --
// Word will shrink/wrap it onto the page reasonably well, and this is the
// same shape as the on-screen matrix. Revisit if segment counts grow much
// further (e.g. once a 19-segment country is fully built).
function buildScorecardTable(matrix) {
  const headerRow = new TableRow({
    children: [headerCell('Dimension'), ...matrix.columnLabels.map((label) => headerCell(label))]
  });
  const dataRows = matrix.rows.map((row) => new TableRow({
    children: [headerCell(row.label), ...row.values.map((v) => bodyCell(v))]
  }));
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

// One country's section: heading + AUM table + scorecard table. Shared by
// both the single-country payload (country.html's per-page export) and the
// multi-country payload (picker.html's project builder) so a project export
// is just this block repeated once per selected country, rather than a
// separate document layout to maintain.
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
    new Paragraph({ text: 'AUM by segment', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }),
    buildAumTable(aumRows),
    new Paragraph({ text: 'Opportunity scorecard', heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 100 } }),
    buildScorecardTable(matrix)
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

      const doc = new Document({ sections: [{ children }] });
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
