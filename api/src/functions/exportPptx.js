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

// One country's slide pair (AUM + scorecard) — shared between the
// single-country payload (country.html) and the multi-country payload
// (picker.html's project builder), so a project export is just this
// repeated once per selected country.
function addCountrySlides(pptx, countryName, segments, generatedDate) {
  const aumRows = buildAumRows(segments);
  const matrix = buildScorecardMatrix(segments);

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

    const safeName = isMulti
      ? `Project_${countries.length}_countries`
      : countries[0].country_name.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

    try {
      const generatedDate = new Date().toISOString().slice(0, 10);

      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'ATLAS_WIDE', width: 13.33, height: 7.5 });
      pptx.layout = 'ATLAS_WIDE';

      if (isMulti) {
        const titleSlide = pptx.addSlide();
        titleSlide.addText(`Atlas — Project (${countries.length} countries)`, { x: 0.4, y: 2.8, fontSize: 32, bold: true });
        titleSlide.addText(`Generated ${generatedDate} — ${countries.map((c) => c.country_name).join(', ')}`, { x: 0.4, y: 3.6, fontSize: 14, color: '666666' });
      }

      countries.forEach((c) => addCountrySlides(pptx, c.country_name, c.segments, generatedDate));

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
