const { app } = require('@azure/functions');
const PptxGenJS = require('pptxgenjs');
const { buildAumRows, buildScorecardMatrix, buildCommentarySections, buildTopInstitutionsSections, estimateColumnCharWidths } = require('../shared/exportHelpers');
const { getDimensionIconDataUri } = require('../shared/dimensionIcons');
const { getAtlasLogoDataUri } = require('../shared/atlasLogo');

const HEADER_FILL = 'D9E2F3';
const BORDER = { type: 'solid', color: 'CCCCCC', pt: 0.5 };
const SLIDE_W = 13.33;
const SLIDE_H = 7.5;

// Tight cell padding (inches: [top, right, bottom, left]) applied to every
// table in this export -- pptxgenjs's default cell margin left noticeably
// more whitespace around short numeric/score cells than the content needed,
// which was Peter's "quite a bit of wasted space" feedback.
const CELL_MARGIN = [0.03, 0.05, 0.03, 0.05];

// Converts an estimateColumnCharWidths() character count into an inch-wide
// pptxgenjs colW entry, so tables are sized to their content instead of
// stretched across the slide -- Peter's follow-up feedback: "Rank only
// needs to be the width of the word Rank... AUM generally only needs to be
// the width of AUM ($bn)". ~0.078in/char is a rough proportional-font
// average at 7-8pt; the cell's left+right margins are added on top.
const CHAR_WIDTH_IN = 0.078;
function charsToInches(chars) {
  return Math.round((chars * CHAR_WIDTH_IN + CELL_MARGIN[1] + CELL_MARGIN[3]) * 100) / 100;
}

// Atlas logo, top-left of every slide via a slide master (see
// ATLAS_MASTER_NAME / buildSlideMaster() below) rather than a per-slide
// addImage() call, so every addSlide({masterName}) call gets it for free.
// Title text on data slides is shifted right (TITLE_X) to clear the logo.
const LOGO_X = 0.15;
const LOGO_Y = 0.15;
const LOGO_SIZE = 0.35;
const TITLE_X = 0.65;
const ATLAS_MASTER_NAME = 'ATLAS_MASTER';

function defineAtlasMaster(pptx) {
  const logoDataUri = getAtlasLogoDataUri();
  pptx.defineSlideMaster({
    title: ATLAS_MASTER_NAME,
    objects: logoDataUri ? [{ image: { x: LOGO_X, y: LOGO_Y, w: LOGO_SIZE, h: LOGO_SIZE, data: logoDataUri } }] : []
  });
}

function addAtlasSlide(pptx) {
  return pptx.addSlide({ masterName: ATLAS_MASTER_NAME });
}

// pptxgenjs table cells don't support an embedded image alongside text (only
// the Word export can do a true inline icon+text cell -- see
// dimensionHeaderCell() in exportDocx.js), so the icons here are separate
// addImage() calls positioned in the slide margin just to the left of the
// table, one per dimension row. This only lines up correctly if the table's
// rows can't grow taller than ROW_H from text wrapping, which is why the
// scorecard table's label column is capped (rather than left unbounded) at
// a character count comfortably wide enough for the longest label
// ("Distribution resources required (x1)") to fit on one line at this font
// size -- see the maxCharsPerCol[0] in buildScorecardTableRows().
const SCORECARD_ROW_H = 0.26;
const SCORECARD_ICON_SIZE = 0.14;
const SCORECARD_TABLE_X = 0.5; // leaves a gutter at 0.3-0.46 for the icons
const SCORECARD_ICON_X = 0.28;

function headerCellOpts(extra) {
  return Object.assign({ bold: true, fill: { color: HEADER_FILL }, fontSize: 8, margin: CELL_MARGIN }, extra);
}

// "Equities range (min-max)" column mirrors the Word export and
// country.html -- see the comment above getAllocationRange() in
// exportHelpers.js for what min/max mean (not every institution counted in
// AUM also reported an asset-class breakdown). Column widths are
// content-driven (see estimateColumnCharWidths() in exportHelpers.js)
// rather than stretched to fill the slide -- Basis and the Equities range
// string are the two columns most likely to run long, so they get the
// highest character caps (and wrap, rather than widening the table
// further, past that). Returns {rows, colW} rather than just the row data,
// since pptxgenjs needs the column widths as a separate addTable() option.
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
    maxCharsPerCol: [26, 10, 10, 22, 30]
  }).map(charsToInches);

  const header = headerLabels.map((t) => ({ text: t, options: headerCellOpts() }));
  const body = bodyText.map((cells) => cells.map((text, i) => ({
    text,
    options: { fontSize: i < 3 ? 8 : 7, margin: CELL_MARGIN }
  })));
  return { rows: [header, ...body], colW };
}

// Same "runs wide with many segments" caveat as the Word export -- v1
// accepts a tight fit on countries with a lot of scored segments (e.g. UK's
// 11 columns) rather than splitting the matrix across multiple slides. Data
// columns hold short values (1-3, "0" for unscored, "x/12", two-digit
// overall scores), so they get a small character cap and the table only
// ever runs as wide as it needs to, rather than stretching to fill the
// slide. Cells carry the same red/amber/green shading as the site's
// scorecard matrix, via row.colors[i] (see scoreColor()/overallColor() in
// exportHelpers.js) -- unscored cells get MISSING_COLOR's yellow instead.
function buildScorecardTableRows(matrix) {
  const headerLabels = ['Dimension', ...matrix.columnLabels];
  const bodyText = matrix.rows.map((row) => [row.label, ...row.values]);
  const maxCharsPerCol = [40, ...matrix.columnLabels.map(() => 6)];
  const colW = estimateColumnCharWidths(headerLabels, bodyText, { minChars: 3, maxCharsPerCol }).map(charsToInches);

  const header = headerLabels.map((t) => ({
    text: t,
    options: headerCellOpts({ fontSize: 7 })
  }));
  const body = matrix.rows.map((row) => ([
    { text: row.label, options: headerCellOpts({ fontSize: 7 }) },
    ...row.values.map((v, i) => {
      const color = row.colors ? row.colors[i] : null;
      const base = { fontSize: 7, align: 'center', margin: CELL_MARGIN };
      if (!color) return { text: v, options: base };
      return { text: v, options: Object.assign({}, base, { fill: { color: color.bg }, color: color.fg, bold: true }) };
    })
  ]));
  return { rows: [header, ...body], colW };
}

// Places one small icon per dimension row (see dimensionIcons.js) in the
// gutter just left of the scorecard table, at the vertical center of that
// row. Relies on every table row actually being SCORECARD_ROW_H tall --
// see the comment above SCORECARD_ROW_H for why the label column is
// content-capped rather than left unbounded, so labels never wrap onto a
// second line and break this alignment. Rows without a `key` (AUM, Scored,
// Overall) are skipped, same as the docx export.
function addScorecardDimensionIcons(slide, matrix, tableY) {
  matrix.rows.forEach((row, i) => {
    if (!row.key) return;
    const dataUri = getDimensionIconDataUri(row.key);
    if (!dataUri) return;
    const rowIndexInTable = i + 1; // +1 for the header row above matrix.rows
    const rowTop = tableY + rowIndexInTable * SCORECARD_ROW_H;
    const iconY = rowTop + (SCORECARD_ROW_H - SCORECARD_ICON_SIZE) / 2;
    slide.addImage({
      data: dataUri,
      x: SCORECARD_ICON_X,
      y: iconY,
      w: SCORECARD_ICON_SIZE,
      h: SCORECARD_ICON_SIZE
    });
  });
}

// Three top-institutions tables side by side per slide (TOPINST_COLS) --
// Peter's 2026-07-15 feedback that one segment per slide, stretched across
// it, wasted the space a now-narrow table leaves. Institution's char cap is
// tighter here (30, vs. the ~40 that would suit a single full-width table)
// so three tables' worth of columns reliably fit within one slot each --
// see TOPINST_SLOT_W below for the budget this is sized against.
const TOPINST_INSTITUTION_MAX_CHARS = 30;

// Column widths are content-driven the same way as buildAumTableRows() --
// "Rank" only needs to fit "Rank" (or a 2-digit number), "AUM ($bn)" only
// needs to fit its header/numbers, and "Institution" gets the rest, capped
// so one very long name wraps instead of stretching the table further.
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
  const body = bodyText.map((cells) => cells.map((text) => ({ text, options: { fontSize: 8, margin: CELL_MARGIN } })));
  return { rows: [header, ...body], colW };
}

// Layout for the 3-across top-institutions slides: three equal slots
// (TOPINST_SLOT_W) across the slide width, with a margin on each outer edge
// and a gutter between slots.
const TOPINST_COLS = 3;
const TOPINST_MARGIN = 0.4;
const TOPINST_GUTTER = 0.3;
const TOPINST_SLOT_W = (SLIDE_W - 2 * TOPINST_MARGIN - (TOPINST_COLS - 1) * TOPINST_GUTTER) / TOPINST_COLS;
const TOPINST_SUBTITLE_Y = 0.95;
const TOPINST_TABLE_Y = 1.55;

// Three segments' top-institutions tables per slide -- Peter's standard
// "top 10 institutions by AUM, and their combined AUM as a % of the
// segment" report format, one column per segment rather than one slide per
// segment (which left a now-narrow table stretched across an otherwise
// empty slide). Segments built from industry aggregates (e.g. Life/Non-life
// insurance) or countries not yet backfilled at institution level
// (currently just the US) are skipped, not guessed at -- see
// buildTopInstitutionsSections in exportHelpers.js.
function addTopInstitutionsSlides(pptx, countryName, segments) {
  const sections = buildTopInstitutionsSections(segments);
  for (let i = 0; i < sections.length; i += TOPINST_COLS) {
    const group = sections.slice(i, i + TOPINST_COLS);
    const slide = addAtlasSlide(pptx);
    slide.addText(`Atlas — ${countryName}`, { x: TITLE_X, y: 0.25, fontSize: 24, bold: true });
    slide.addText('Top institutions by AUM', { x: 0.4, y: 0.6, fontSize: 12, color: '666666' });

    group.forEach((section, col) => {
      const slotX = TOPINST_MARGIN + col * (TOPINST_SLOT_W + TOPINST_GUTTER);
      const nText = section.n_institutions ? ` of ${section.n_institutions.toLocaleString()} identified` : '';
      slide.addText(
        `${section.segment} — top ${section.institutions.length}${nText} institutions hold ${section.top10_share_pct}% of segment AUM`,
        { x: slotX, y: TOPINST_SUBTITLE_Y, w: TOPINST_SLOT_W, fontSize: 9, color: '666666' }
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

// One slide per populated commentary section (Wealth & key pools of
// capital, Pensions structure) -- see buildCommentarySections() in
// exportHelpers.js for the text-splitting/source-filtering rules. A country
// with no commentary text yet (most countries, until written) contributes no
// slides at all, same convention as addTopInstitutionsSlides() skipping
// segments with no institution-level data. Body text and sources are two
// separate addText() calls in fixed slots rather than one flowing text box,
// since pptxgenjs doesn't auto-grow a box to fit content -- long commentary
// may run past COMMENTARY_BODY_H and get clipped, same known limitation as
// the scorecard table's fixed row height.
const COMMENTARY_BODY_Y = 1.3;
const COMMENTARY_BODY_H = 4.6;
const COMMENTARY_SOURCES_Y = 6.05;
const COMMENTARY_SOURCES_H = 1.2;

function addCommentarySlides(pptx, countryName, commentary) {
  const sections = buildCommentarySections(commentary);
  sections.forEach((section) => {
    const slide = addAtlasSlide(pptx);
    slide.addText(`Atlas — ${countryName}`, { x: TITLE_X, y: 0.25, fontSize: 24, bold: true });
    slide.addText(section.label, { x: 0.4, y: 0.85, fontSize: 12, color: '666666' });

    const bodyRuns = section.paragraphs.map((p) => ({ text: p, options: { breakLine: true, paraSpaceAfter: 10 } }));
    slide.addText(bodyRuns, {
      x: 0.4, y: COMMENTARY_BODY_Y, w: SLIDE_W - 0.8, h: COMMENTARY_BODY_H,
      fontSize: 11, valign: 'top', align: 'left', autoFit: false
    });

    if (section.sources.length) {
      const sourceRuns = [
        { text: 'Sources', options: { bold: true, breakLine: true, color: '666666' } },
        ...section.sources.map((s) => ({
          text: s.label || s.url,
          options: {
            breakLine: true,
            color: '666666',
            ...(s.url ? { hyperlink: { url: s.url } } : {})
          }
        }))
      ];
      slide.addText(sourceRuns, {
        x: 0.4, y: COMMENTARY_SOURCES_Y, w: SLIDE_W - 0.8, h: COMMENTARY_SOURCES_H,
        fontSize: 8, valign: 'top', align: 'left'
      });
    }
  });
}

// One country's AUM slide. Factored out of addCountrySlides() (below) so
// each content type's slide-building logic is independently callable --
// needed for the `include` filter (added 2026-07-16, Peter's request:
// export e.g. only scorecards, or only top 10s, of the selected countries,
// rather than always the full AUM+scorecard+top10 bundle).
function addAumSlide(pptx, countryName, segments, generatedDate) {
  const aumSlide = addAtlasSlide(pptx);
  aumSlide.addText(`Atlas — ${countryName}`, { x: TITLE_X, y: 0.25, fontSize: 24, bold: true });
  aumSlide.addText(`AUM by segment — generated ${generatedDate}`, { x: 0.4, y: 0.85, fontSize: 12, color: '666666' });
  const aumTable = buildAumTableRows(buildAumRows(segments));
  aumSlide.addTable(aumTable.rows, {
    x: 0.4, y: 1.3,
    colW: aumTable.colW,
    border: BORDER,
    autoPage: false
  });
}

// One country's scorecard slide. Same factoring-out reasoning as
// addAumSlide() above.
function addScorecardSlide(pptx, countryName, segments, enabledDimensions, weightOverrides) {
  const matrix = buildScorecardMatrix(segments, enabledDimensions, weightOverrides);
  const scorecardSlide = addAtlasSlide(pptx);
  scorecardSlide.addText(`Atlas — ${countryName}`, { x: TITLE_X, y: 0.25, fontSize: 24, bold: true });
  scorecardSlide.addText('Opportunity scorecard', { x: 0.4, y: 0.85, fontSize: 12, color: '666666' });
  const scorecardTableY = 1.3;
  const scorecardTable = buildScorecardTableRows(matrix);
  scorecardSlide.addTable(scorecardTable.rows, {
    x: SCORECARD_TABLE_X, y: scorecardTableY,
    colW: scorecardTable.colW,
    rowH: SCORECARD_ROW_H,
    border: BORDER,
    autoPage: false
  });
  addScorecardDimensionIcons(scorecardSlide, matrix, scorecardTableY);
}

// Which of the four content blocks to add for a country -- 'commentary',
// 'aum', 'scorecard', 'top_institutions'. Defaults to all four (the original
// "everything" behaviour, plus commentary added 2026-07-23) when not
// specified, so country.html's existing single-country export (which never
// sends `include`) is unaffected.
const ALL_CONTENT_TYPES = ['commentary', 'aum', 'scorecard', 'top_institutions'];
function resolveInclude(rawInclude) {
  const valid = Array.isArray(rawInclude) ? rawInclude.filter((k) => ALL_CONTENT_TYPES.includes(k)) : [];
  return new Set(valid.length ? valid : ALL_CONTENT_TYPES);
}

// One country's full slide set -- whichever of the AUM slide / scorecard
// slide / top-10-per-segment slides `include` asks for. Shared between the
// single-country payload (country.html) and the multi-country payload
// (picker.html's project builder), so a project export is just this
// repeated once per selected country.
function addCountrySlides(pptx, countryName, segments, generatedDate, enabledDimensions, include, weightOverrides, commentary) {
  const includeSet = include || new Set(ALL_CONTENT_TYPES);
  if (includeSet.has('commentary')) addCommentarySlides(pptx, countryName, commentary);
  if (includeSet.has('aum')) addAumSlide(pptx, countryName, segments, generatedDate);
  if (includeSet.has('scorecard')) addScorecardSlide(pptx, countryName, segments, enabledDimensions, weightOverrides);
  if (includeSet.has('top_institutions')) addTopInstitutionsSlides(pptx, countryName, segments);
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
      : (Array.isArray(body.segments) && body.segments.length ? [{ country_name: body.country_name || 'Country', segments: body.segments, commentary: body.commentary }] : []);

    if (!countries.length) {
      return { status: 400, jsonBody: { error: 'No segments provided to export' } };
    }

    const safeName = isMulti
      ? `Project_${countries.length}_countries`
      : countries[0].country_name.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');

    try {
      const generatedDate = new Date().toISOString().slice(0, 10);

      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'ATLAS_WIDE', width: SLIDE_W, height: SLIDE_H });
      pptx.layout = 'ATLAS_WIDE';
      defineAtlasMaster(pptx);

      if (isMulti) {
        const titleSlide = addAtlasSlide(pptx);
        titleSlide.addText(`Atlas — Project (${countries.length} countries)`, { x: 0.4, y: 2.8, fontSize: 32, bold: true });
        titleSlide.addText(`Generated ${generatedDate} — ${countries.map((c) => c.country_name).join(', ')}`, { x: 0.4, y: 3.6, fontSize: 14, color: '666666' });
      }

      // Which content types to include per country -- 'aum', 'scorecard',
      // 'top_institutions', any combination, defaulting to all three.
      // Peter's 2026-07-16 request: be able to export e.g. only scorecards,
      // or only top 10s, of the selected countries, rather than always the
      // full bundle. A country with nothing to show under the requested set
      // (e.g. top_institutions-only, and this country has no
      // institution-level data) simply contributes no slides.
      const include = resolveInclude(body.include);
      // weight_overrides -- picker.html's project builder weighting column
      // (see exportHelpers.js's computeOverallScore() comment). Optional;
      // country.html's single-country export never sends this, so Overall
      // there is unaffected.
      countries.forEach((c) => addCountrySlides(pptx, c.country_name, c.segments, generatedDate, body.enabled_dimensions, include, body.weight_overrides, c.commentary));

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
