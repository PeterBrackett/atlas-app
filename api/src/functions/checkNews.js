const { app } = require('@azure/functions');

// Same authorisation model as the other write-adjacent editor tools (see
// setDevelopments.js) -- this doesn't write anything to SharePoint itself,
// but it does trigger an external Google News fetch plus a paid Anthropic
// API call per click, so it's gated to the same signed-in editor rather than
// left open to anyone who can reach the URL.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Requires a separate Anthropic API key (an Atlas-specific key, not shared
// with anything else) as an Azure Static Web App application setting --
// unlike the SharePoint/Graph calls elsewhere in this project, there's no
// existing credential this can reuse. If this isn't set, the endpoint fails
// fast with a clear error rather than a confusing downstream 401 from
// Anthropic.
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const ANTHROPIC_MODEL = 'claude-haiku-4-5-20251001';

// Same 8 keys as setDevelopments.js's VALID_SECTIONS / country.html's
// COMMENTARY_SECTIONS -- kept as a duplicate for the same "separate Node
// deployment, no shared module at build time" reason as SCORECARD_DIMENSIONS
// elsewhere. Descriptions are given to Claude below so it has enough context
// to disambiguate e.g. an "insurer's staff pension scheme" story (pensions,
// not insurance) or a "university endowment" story (charities/foundations,
// depending on phrasing) -- plain section keys alone aren't enough context.
const SECTIONS = [
  { key: 'wealth', label: 'Wealth & key pools of capital', description: 'General institutional/private wealth pools, asset managers, wealth managers -- not a specific segment below.' },
  { key: 'pensions', label: 'Pensions structure', description: 'Occupational or state pension schemes (DB or DC), pension funds, pension regulation, buyouts, superfunds.' },
  { key: 'insurance', label: 'Insurance', description: 'Life or non-life insurance companies\' own balance-sheet investment activity, not insurance products sold to consumers.' },
  { key: 'charities', label: 'Charities', description: 'Charitable trusts/organisations and their investable reserves, not day-to-day charitable services.' },
  { key: 'foundations', label: 'Foundations', description: 'Grant-making foundations, university/hospital endowments and their investment pools.' },
  { key: 'family_offices', label: 'Family offices', description: 'Single- or multi-family offices managing private/family wealth.' },
  { key: 'sovereign_wealth_funds', label: 'Sovereign wealth funds', description: 'State-owned investment funds/reserve funds.' },
  { key: 'ocio', label: 'OCIO', description: 'Outsourced chief investment officer providers and mandates.' }
];

function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(url, Object.assign({}, options, { signal: controller.signal })).finally(() => clearTimeout(timer));
}

// Minimal hand-rolled RSS <item> extractor -- deliberately not a new npm
// dependency (this project has stayed dependency-light, see docx/pptxgenjs
// as the only two additions so far), and Google News RSS's item shape is
// simple and consistent enough that a couple of regexes cover it reliably.
function decodeXmlEntities(s) {
  return String(s || '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'").replace(/&amp;/g, '&');
}

function parseRssItems(xml) {
  const items = [];
  const itemBlocks = xml.match(/<item>[\s\S]*?<\/item>/g) || [];
  for (const block of itemBlocks) {
    const title = (block.match(/<title>([\s\S]*?)<\/title>/) || [])[1];
    const link = (block.match(/<link>([\s\S]*?)<\/link>/) || [])[1];
    const pubDate = (block.match(/<pubDate>([\s\S]*?)<\/pubDate>/) || [])[1];
    const sourceMatch = block.match(/<source url="[^"]*">([\s\S]*?)<\/source>/);
    if (!title || !link) continue;
    items.push({
      title: decodeXmlEntities(title).replace(/<!\[CDATA\[|\]\]>/g, '').trim(),
      link: decodeXmlEntities(link).trim(),
      pubDate: pubDate ? pubDate.trim() : '',
      source: sourceMatch ? decodeXmlEntities(sourceMatch[1]).replace(/<!\[CDATA\[|\]\]>/g, '').trim() : ''
    });
  }
  return items;
}

async function fetchGoogleNews(query) {
  const url = `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=en-GB&gl=GB&ceid=GB:en`;
  try {
    const res = await fetchWithTimeout(url, undefined, 12000);
    if (!res.ok) return [];
    const xml = await res.text();
    return parseRssItems(xml);
  } catch (e) {
    return [];
  }
}

// Google News RSS links are redirect wrappers (news.google.com/rss/articles/...),
// not the publisher's actual URL -- following the redirect server-side gives
// a real, clickable link for the developments box rather than a Google
// interstitial. Falls back to the original wrapper link if the redirect
// can't be resolved in time, so a slow/broken publisher doesn't drop the
// candidate entirely, just leaves a less ideal URL.
async function resolveArticleUrl(link) {
  try {
    const res = await fetchWithTimeout(link, { redirect: 'follow' }, 8000);
    return res.url || link;
  } catch (e) {
    return link;
  }
}

// Mirrors mapWithConcurrency in atlas-site/scorecard-dimensions.js -- caps
// how many URL-resolution fetches run at once so a country with lots of
// candidate hits doesn't fire dozens of simultaneous outbound requests.
async function mapWithConcurrency(items, limit, asyncFn) {
  const results = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const i = next++;
      results[i] = await asyncFn(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

function formatPubDateToMonth(pubDate) {
  const d = new Date(pubDate);
  if (isNaN(d.getTime())) return '';
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

// Sends the candidate headlines to Claude for the judgement call this
// endpoint actually needs a model for: is this genuinely specific to
// {countryName}'s institutional-investor market (not just any mention of the
// country), and if so which of the 8 sections does it belong to, plus a
// short neutral summary. Returns an array the same length as `candidates`,
// each either null (not relevant) or { section, summary }.
async function classifyCandidates(countryName, candidates) {
  if (!ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY is not configured on the server.');
  }
  if (!candidates.length) return [];

  const sectionList = SECTIONS.map((s) => `- ${s.key}: ${s.label} -- ${s.description}`).join('\n');
  const itemList = candidates.map((c, i) => `${i}. "${c.title}" (source: ${c.source || 'unknown'}, published: ${c.pubDate || 'unknown'})`).join('\n');

  const prompt = `You are screening news headlines for "Atlas", an institutional-investor competitive-intelligence tool covering ${countryName}'s asset-owner market. \
For each numbered headline below, decide whether it is a genuinely specific, substantive development about ${countryName}'s institutional investors in one of these sections:
${sectionList}

Reject anything that only mentions ${countryName} incidentally, is generic global news, is about retail/consumer products rather than institutional investors, or is too vague to summarise usefully. When in doubt, reject rather than include.

Headlines:
${itemList}

Respond with ONLY a JSON array (no other text), one entry per headline in the same order, each either:
- null, if not relevant, or
- {"section": "<one of the section keys above>", "summary": "<one neutral sentence, professional briefing tone, under 200 characters>"}`;

  const res = await fetchWithTimeout('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      model: ANTHROPIC_MODEL,
      max_tokens: 2048,
      messages: [{ role: 'user', content: prompt }]
    })
  }, 30000);

  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(`Anthropic API request failed (${res.status}): ${detail.slice(0, 300)}`);
  }

  const payload = await res.json();
  const text = (payload.content || []).map((block) => block.text || '').join('');
  // Claude is asked for JSON-only, but strip any accidental code-fence
  // wrapping or leading/trailing prose defensively rather than trusting the
  // model's output format is always perfectly bare.
  const jsonMatch = text.match(/\[[\s\S]*\]/);
  if (!jsonMatch) throw new Error('Could not parse a JSON array out of the classification response.');
  const parsed = JSON.parse(jsonMatch[0]);
  if (!Array.isArray(parsed)) throw new Error('Classification response was not a JSON array.');
  return parsed;
}

function getClientPrincipal(request) {
  const header = request.headers.get('x-ms-client-principal');
  if (!header) return null;
  try {
    return JSON.parse(Buffer.from(header, 'base64').toString('utf-8'));
  } catch {
    return null;
  }
}

// One button on country.html ("Check for news") calls this once per click --
// no scheduled/background polling (Peter's 2026-07-30 preference: on-demand
// only, reviewed before anything is saved). Searches Google News (free, no
// API key) across two broad queries to cover all 8 sections reasonably well
// in one pass, resolves each result's real article URL, then hands the
// batch to Claude to judge relevance, assign a section, and draft a summary.
// Returns candidates only -- nothing is written to {code}_developments.json
// here; the client presents each candidate for Approve/Decline, and Approve
// calls the existing /api/developments/set endpoint per item.
app.http('checkNews', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'news/check',
  handler: async (request, context) => {
    const principal = getClientPrincipal(request);
    const userEmail = ((principal && principal.userDetails) || '').toLowerCase();

    if (!ALLOWED_EDITOR_EMAIL) {
      return { status: 500, jsonBody: { error: 'Server is not configured with an authorised editor email.' } };
    }
    if (!principal || userEmail !== ALLOWED_EDITOR_EMAIL) {
      return { status: 403, jsonBody: { error: 'Not authorised to use this tool.' } };
    }
    if (!ANTHROPIC_API_KEY) {
      return { status: 500, jsonBody: { error: 'Server is not configured with an Anthropic API key (ANTHROPIC_API_KEY).' } };
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return { status: 400, jsonBody: { error: 'Invalid request body' } };
    }

    const countryName = typeof body.country_name === 'string' ? body.country_name.trim() : '';
    if (!countryName) {
      return { status: 400, jsonBody: { error: 'country_name is required' } };
    }

    try {
      // Two queries -- one tuned toward pensions/insurance/general
      // institutional investors, one toward the smaller/rarer sections
      // (SWF/family offices/foundations/charities/OCIO) -- a single generic
      // query tended to be dominated by pensions/insurance coverage and
      // rarely surfaced the rarer sections at all.
      const [batchA, batchB] = await Promise.all([
        fetchGoogleNews(`"${countryName}" (pension fund OR pension scheme OR "institutional investor" OR "asset owner")`),
        fetchGoogleNews(`"${countryName}" (sovereign wealth fund OR "family office" OR foundation endowment OR charity investment OR OCIO)`)
      ]);

      const seen = new Set();
      const merged = [];
      for (const item of [...batchA, ...batchB]) {
        const dedupeKey = item.link;
        if (seen.has(dedupeKey)) continue;
        seen.add(dedupeKey);
        merged.push(item);
      }

      // Cap how many go to the classifier -- keeps the prompt (and cost)
      // bounded even if both queries return a full page of results each.
      const capped = merged.slice(0, 30);

      const resolvedUrls = await mapWithConcurrency(capped, 6, (item) => resolveArticleUrl(item.link));

      const classifications = await classifyCandidates(countryName, capped);

      const candidates = capped
        .map((item, i) => {
          const cls = classifications[i];
          if (!cls || !cls.section || !SECTIONS.some((s) => s.key === cls.section)) return null;
          return {
            section: cls.section,
            headline: item.title,
            date: formatPubDateToMonth(item.pubDate),
            summary: typeof cls.summary === 'string' ? cls.summary : '',
            source: item.source,
            url: resolvedUrls[i] || item.link
          };
        })
        .filter(Boolean);

      return { status: 200, jsonBody: { success: true, candidates } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error checking for news', detail: err.message } };
    }
  }
});
