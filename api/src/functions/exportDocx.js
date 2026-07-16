const { app } = require('@azure/functions');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, ShadingType, PageBreak, Header
} = require('docx');
const { buildAumRows, buildScorecardMatrix, buildTopInstitutionsSections, estimateColumnCharWidths } = require('../shared/exportHelpers');
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
// one small table per segment that has institution-level data. Factored out
// of buildTopInstitutionsBlock() so the same per-segment content can sit
// under either a per-country H2 (the original "everything for this country
// together" layout) or a per-country H2 nested inside a single shared "Top
// institutions by AUM" H1 (the "everything of this type together" layout,
// added 2026-07-16 -- see buildTopInstitutionsSection() below). Segments
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

// One country's section: heading + AUM table + scorecard table + top
// institutions by segment. Shared by both the single-country payload
// (country.html's per-page export) and the multi-country payload
// (picker.html's project builder) so a project export is just this block
// repeated once per selected country, rather than a separate document
// layout to maintain.
function buildCountrySection(countryName, segments, { headingLevel = HeadingLevel.HEADING_1, pageBreakBefore = false, enabledDimensions } = {}) {
  const aumRows = buildAumRows(segments);
  const matrix = buildScorecardMatrix(segments, enabledDimensions);
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

// "Group by content type" layout for a multi-country project export (added
// 2026-07-16, Peter's request) -- an alternative to buildCountrySection()'s
// "everything for this country together" grouping, for when what's actually
// wanted is "every country's AUM table together, then every country's
// scorecard together, then every country's top 10s together" (e.g. to lift
// just the scorecard section into a different document). Each of the three
// functions below is one top-level H1 section; a country gets its own H2
// subheading inside it. Only used when the request explicitly asks for it
// (body.group_by === 'content') -- buildCountrySection() stays the default.
function sectionHeading(text, pageBreakBefore) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 200, after: 100 },
    children: [
      ...(pageBreakBefore ? [new PageBreak()] : []),
      new TextRun({ text })
    ]
  });
}

function countrySubheading(countryName, isFirst) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: isFirst ? 0 : 200, after: 60 },
    children: [new TextRun({ text: countryName })]
  });
}

function buildAumSection(countries, pageBreakBefore) {
  const children = [sectionHeading('AUM by segment', pageBreakBefore)];
  countries.forEach((c, i) => {
    children.push(countrySubheading(c.country_name, i === 0));
    children.push(buildAumTable(buildAumRows(c.segments)));
  });
  return children;
}

function buildScorecardSection(countries, enabledDimensions, pageBreakBefore) {
  const children = [sectionHeading('Opportunity scorecard', pageBreakBefore)];
  countries.forEach((c, i) => {
    children.push(countrySubheading(c.country_name, i === 0));
    children.push(buildScorecardTable(buildScorecardMatrix(c.segments, enabledDimensions)));
  });
  return children;
}

// Skips countries with no institution-level data entirely (no empty H2 with
// nothing under it), same as buildTopInstitutionsBlock() does per-country.
// "isFirst" for spacing purposes tracks the first country actually emitted,
// not raw array position, since an earlier country may have been skipped.
function buildTopInstitutionsSection(countries, pageBreakBefore) {
  const perCountry = [];
  let emitted = 0;
  countries.forEach((c) => {
    const sections = buildTopInstitutionsSections(c.segments);
    if (!sections.length) return;
    perCountry.push(countrySubheading(c.country_name, emitted === 0), ...buildTopInstitutionsPerSegment(sections));
    emitted += 1;
  });
  if (!perCountry.length) return [];
  return [sectionHeading('Top institutions by AUM', pageBreakBefore), ...perCountry];
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

      // Two layouts for a multi-country project export: the default groups
      // everything by country (each country's own AUM/scorecard/top-10s
      // together); "content" groups by content type instead (every
      // country's AUM together, then every country's scorecard, then every
      // country's top 10s) -- Peter's 2026-07-16 request, for pulling just
      // one type of table across a project's countries into a report.
      // Single-country exports (country.html) are unaffected either way,
      // since there's only one country's worth of content regardless.
      if (isMulti && body.group_by === 'content') {
        children.push(...buildAumSection(countries, false));
        children.push(...buildScorecardSection(countries, body.enabled_dimensions, true));
        children.push(...buildTopInstitutionsSection(countries, true));
      } else {
        countries.forEach((c, i) => {
          children.push(...buildCountrySection(c.country_name, c.segments, {
            headingLevel: isMulti ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_1,
            pageBreakBefore: isMulti && i > 0,
            enabledDimensions: body.enabled_dimensions
          }));
        });
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
