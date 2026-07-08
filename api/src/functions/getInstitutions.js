const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Same credentials/site/folder as getData.js and getNotes.js — kept as
// separate constants here so this function has no dependency on the others.
const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

const MAX_PAGE_SIZE = 200;
const DEFAULT_PAGE_SIZE = 50;

// How long a warm Function instance will keep serving a country's dataset
// from memory before re-fetching it from SharePoint. This is the load-bearing
// piece of this endpoint's design: institutions.html no longer downloads the
// whole country file itself — the Function does, once per cold cache, and
// only ever ships the caller a single page back to the browser. A short TTL
// (rather than caching forever) means an edited/re-pulled SharePoint file is
// picked up within a few minutes without needing a redeploy.
const CACHE_TTL_MS = 5 * 60 * 1000;

// Module-level, so it survives across invocations on the same warm instance
// (standard Consumption-plan behaviour) but not across cold starts or
// separate instances — this is a perf/cost optimisation, not a source of
// truth. Keyed by country code, since each Function instance may serve
// requests for more than one country.
const cache = new Map(); // country -> { fetchedAt, institutions, country_name, extraction_date, source_files }

let msalClient;
function getMsalClient() {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET
      }
    });
  }
  return msalClient;
}

async function getAccessToken() {
  const client = getMsalClient();
  const result = await client.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default']
  });
  return result.accessToken;
}

// Fetches and caches the full institutions file for a country. This is the
// piece that will eventually need to become a real indexed query (Cosmos DB /
// Azure SQL / Table Storage) rather than "download and JSON.parse the whole
// file" once a large-country dataset (e.g. a US-scale pull, tens of thousands
// of records) is in play — a single-digit-MB file like the UK's is fine to
// hold in Function memory, a >100MB one is not. Flagging in code, not just in
// conversation, so this doesn't get forgotten when that country gets scoped.
async function loadInstitutions(country, context) {
  const cached = cache.get(country);
  if (cached && (Date.now() - cached.fetchedAt) < CACHE_TTL_MS) {
    return cached;
  }

  const fileName = `${country}_institutions_latest.json`;
  const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}:/content`;

  const token = await getAccessToken();
  const graphResponse = await fetch(graphUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (graphResponse.status === 404) {
    return null;
  }

  if (!graphResponse.ok) {
    const detail = await graphResponse.text();
    context.error(`Graph error ${graphResponse.status}: ${detail}`);
    throw new Error(`SharePoint fetch failed with status ${graphResponse.status}`);
  }

  const data = await graphResponse.json();
  const entry = {
    fetchedAt: Date.now(),
    institutions: data.institutions || [],
    country_name: data.country_name,
    extraction_date: data.extraction_date,
    source_files: data.source_files || []
  };
  cache.set(country, entry);
  return entry;
}

function matchesSearch(record, q) {
  if (!q) return true;
  return record.entity_name && record.entity_name.toLowerCase().includes(q);
}

function hasRelationship(record, kind) {
  const gov = record.governance || {};
  switch (kind) {
    case 'manager': return !!(gov.asset_managers_used && gov.asset_managers_used.length);
    case 'custodian': return !!(gov.custodian && gov.custodian.length);
    case 'actuary': return !!(gov.actuary && gov.actuary.length);
    case 'consultant': return !!(gov.investment_consultant && gov.investment_consultant.length);
    default: return true;
  }
}

const SORTABLE_FIELDS = {
  entity_name: r => (r.entity_name || '').toLowerCase(),
  total_aum: r => (r.financials && typeof r.financials.total_aum === 'number') ? r.financials.total_aum : null,
  plan_count: r => (r.plans || []).length
};

app.http('getInstitutions', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'institutions/{country}',
  handler: async (request, context) => {
    const country = (request.params.country || '').toLowerCase();

    if (!/^[a-z0-9-]+$/.test(country)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const q = request.query.get('search');
    const search = q ? q.trim().toLowerCase() : '';
    const entityType = request.query.get('entityType') || '';
    const has = request.query.get('has') || ''; // 'manager' | 'custodian' | 'actuary' | 'consultant' | 'none' | ''
    const sortKeyRaw = request.query.get('sortKey') || 'total_aum';
    const sortKey = SORTABLE_FIELDS[sortKeyRaw] ? sortKeyRaw : 'total_aum';
    const sortDir = (request.query.get('sortDir') || 'desc') === 'asc' ? 1 : -1;

    let page = parseInt(request.query.get('page') || '0', 10);
    if (!Number.isFinite(page) || page < 0) page = 0;
    let pageSize = parseInt(request.query.get('pageSize') || String(DEFAULT_PAGE_SIZE), 10);
    if (!Number.isFinite(pageSize) || pageSize <= 0) pageSize = DEFAULT_PAGE_SIZE;
    pageSize = Math.min(pageSize, MAX_PAGE_SIZE);

    let dataset;
    try {
      dataset = await loadInstitutions(country, context);
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching institutions', detail: err.message } };
    }

    if (!dataset) {
      return { status: 404, jsonBody: { error: `No institutions file found for '${country}'` } };
    }

    let filtered = dataset.institutions.filter(r => {
      if (!matchesSearch(r, search)) return false;
      if (entityType && r.entity_type !== entityType) return false;
      if (has === 'none') {
        if (hasRelationship(r, 'manager') || hasRelationship(r, 'custodian') || hasRelationship(r, 'actuary') || hasRelationship(r, 'consultant')) return false;
      } else if (has) {
        if (!hasRelationship(r, has)) return false;
      }
      return true;
    });

    const getSortVal = SORTABLE_FIELDS[sortKey];
    filtered.sort((a, b) => {
      let av = getSortVal(a);
      let bv = getSortVal(b);
      if (av === null || av === undefined) av = sortDir === 1 ? Infinity : -Infinity;
      if (bv === null || bv === undefined) bv = sortDir === 1 ? Infinity : -Infinity;
      if (av < bv) return -1 * sortDir;
      if (av > bv) return 1 * sortDir;
      return 0;
    });

    const totalMatching = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalMatching / pageSize));
    const start = page * pageSize;
    const pageItems = filtered.slice(start, start + pageSize);

    // entity_type facet list is computed off the *full* dataset (not the
    // filtered slice) so the frontend's type dropdown doesn't shrink to only
    // whatever's on the current page/filter — computed here rather than
    // shipped as a separate endpoint since it's cheap once the file's cached.
    const entityTypes = [...new Set(dataset.institutions.map(r => r.entity_type).filter(Boolean))].sort();

    return {
      jsonBody: {
        country_code: country.toUpperCase(),
        country_name: dataset.country_name,
        extraction_date: dataset.extraction_date,
        source_files: dataset.source_files,
        total_institutions: dataset.institutions.length,
        total_matching: totalMatching,
        page,
        page_size: pageSize,
        total_pages: totalPages,
        entity_types: entityTypes,
        institutions: pageItems
      },
      headers: { 'Cache-Control': 'no-store' }
    };
  }
});
