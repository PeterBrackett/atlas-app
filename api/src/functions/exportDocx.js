const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType
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

    const countryName = body.country_name || 'Country';
    const segments = Array.isArray(body.segments) ? body.segments : [];

    if (!segments.length) {
      return { status: 400, jsonBody: { error: 'No segments provided to export' } };
    }

    try {
      const aumRows = buildAumRows(segments);
      const matrix = buildScorecardMatrix(segments);
      const generatedDate = new Date().toISOString().slice(0, 10);

      const doc = new Document({
        sections: [{
          children: [
            new Paragraph({ text: `Atlas — ${countryName}`, heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ text: `Generated ${generatedDate}` }),
            new Paragraph({ text: 'AUM by segment', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }),
            buildAumTable(aumRows),
            new Paragraph({ text: 'Opportunity scorecard', heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 100 } }),
            buildScorecardTable(matrix),
            new Paragraph({
              text: 'Source: Atlas. See Sources & Methodology on the site for how these figures are derived.',
              spacing: { before: 300 }
            })
          ]
        }]
      });

      const buffer = await Packer.toBuffer(doc);
      const safeName = countryName.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

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
