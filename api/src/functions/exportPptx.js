const { app } = require('@azure/functions');
const PptxGenJS = require('pptxgenjs');
const { buildAumRows, buildScorecardMatrix } = require('../shared/exportHelpers');

const HEADER_FILL = 'D9E2F3';
const BORDER = { type: 'solid', color: 'CCCCCC', pt: 0.5 };

function headerCellOpts(extra) {
  return Object.assign({ bold: true, fill: { color: HEADER_FILL }, fontSize: 9 }, extra);
}

function buildAumTableRows(rows) {
  const header = ['Segment', 'AUM ($bn)', 'Equities ($bn)', 'Basis'].map((t) => ({
    text: t,
    options: headerCellOpts()
  }));
  const body = rows.map((r) => ([
    { text: r.segment, options: { fontSize: 9 } },
    { text: typeof r.aum_bn === 'number' ? r.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-', options: { fontSize: 9 } },
    { text: typeof r.equity_bn === 'number' ? r.equity_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) : '-', options: { fontSize: 9 } },
    { text: r.basis || '', options: { fontSize: 8 } }
  ]));
  return [header, ...body];
}

// Same "runs wide with many segments" caveat as the Word export -- v1
// accepts a tight fit on countries with a lot of scored segments (e.g. UK's
// 11 columns) rather than splitting the matrix across multiple slides.
function buildScorecardTableRows(matrix) {
  const header = ['Dimension', ...matrix.columnLabels].map((t) => ({
    text: t,
    options: headerCellOpts({ fontSize: 8 })
  }));
  const body = matrix.rows.map((row) => ([
    { text: row.label, options: headerCellOpts({ fontSize: 8 }) },
    ...row.values.map((v) => ({ text: v, options: { fontSize: 8, align: 'center' } }))
  ]));
  return [header, ...body];
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

    const countryName = body.country_name || 'Country';
    const segments = Array.isArray(body.segments) ? body.segments : [];

    if (!segments.length) {
      return { status: 400, jsonBody: { error: 'No segments provided to export' } };
    }

    try {
      const aumRows = buildAumRows(segments);
      const matrix = buildScorecardMatrix(segments);
      const generatedDate = new Date().toISOString().slice(0, 10);

      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'ATLAS_WIDE', width: 13.33, height: 7.5 });
      pptx.layout = 'ATLAS_WIDE';

      const aumSlide = pptx.addSlide();
      aumSlide.addText(`Atlas — ${countryName}`, { x: 0.4, y: 0.25, fontSize: 24, bold: true });
      aumSlide.addText(`AUM by segment — generated ${generatedDate}`, { x: 0.4, y: 0.85, fontSize: 12, color: '666666' });
      aumSlide.addTable(buildAumTableRows(aumRows), {
        x: 0.4, y: 1.3, w: 12.5,
        border: BORDER,
        autoPage: false
      });

      const scorecardSlide = pptx.addSlide();
      scorecardSlide.addText(`Atlas — ${countryName}`, { x: 0.4, y: 0.25, fontSize: 24, bold: true });
      scorecardSlide.addText('Opportunity scorecard', { x: 0.4, y: 0.85, fontSize: 12, color: '666666' });
      scorecardSlide.addTable(buildScorecardTableRows(matrix), {
        x: 0.3, y: 1.3, w: 12.7,
        border: BORDER,
        autoPage: false
      });

      const buffer = await pptx.write({ outputType: 'nodebuffer' });
      const safeName = countryName.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

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
