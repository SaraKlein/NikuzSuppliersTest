const { app } = require('@azure/functions');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const sitePath = '/sites/SaraTest';
const siteHost = 'crdatasystems.sharepoint.com';
const listName = 'Suppliers';

function jsonResponse(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    jsonBody: body
  };
}

function normalize(value) {
  if (typeof value !== 'string' && typeof value !== 'number') {
    return '';
  }

  return String(value).trim().toLowerCase();
}

function getCredentials() {
  const { SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET } = process.env;
  if (!SHAREPOINT_TENANT_ID || !SHAREPOINT_CLIENT_ID || !SHAREPOINT_CLIENT_SECRET) {
    return null;
  }

  return {
    tenantId: SHAREPOINT_TENANT_ID,
    clientId: SHAREPOINT_CLIENT_ID,
    clientSecret: SHAREPOINT_CLIENT_SECRET
  };
}

async function getAccessToken(credentials) {
  const body = new URLSearchParams({
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(credentials.tenantId)}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
      signal: AbortSignal.timeout(10000)
    }
  );

  if (!response.ok) {
    throw new Error(`Microsoft identity token request failed (${response.status})`);
  }

  const tokenResponse = await response.json();
  if (typeof tokenResponse.access_token !== 'string') {
    throw new Error('Microsoft identity token response did not include an access token');
  }

  return tokenResponse.access_token;
}

async function graphGet(url, accessToken, operation) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    signal: AbortSignal.timeout(10000)
  });

  if (!response.ok) {
    throw new Error(`Microsoft Graph ${operation} request failed (${response.status})`);
  }

  return response.json();
}

async function supplierExists(values, accessToken) {
  const site = await graphGet(
    `${graphBaseUrl}/sites/${siteHost}:${sitePath}`,
    accessToken,
    'site lookup'
  );

  const listsUrl = new URL(`${graphBaseUrl}/sites/${site.id}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${listName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphGet(listsUrl, accessToken, 'list lookup');
  const list = lists.value.find((item) => item.displayName === listName);

  if (!list) {
    throw new Error(`SharePoint list "${listName}" was not found`);
  }

  let nextUrl = new URL(`${graphBaseUrl}/sites/${site.id}/lists/${list.id}/items`);
  nextUrl.searchParams.set('$expand', 'fields($select=field_1,Title,field_7)');
  nextUrl.searchParams.set('$top', '200');

  while (nextUrl) {
    const page = await graphGet(nextUrl, accessToken, 'supplier lookup');
    const found = page.value.some((item) => {
      const fields = item.fields || {};
      return normalize(fields.field_1) === values.supplierNumber
        && normalize(fields.Title) === values.supplierName
        && normalize(fields.fileld_7) === values.supplierEmail;
    });

    if (found) {
      return true;
    }

    if (page['@odata.nextLink']) {
      nextUrl = new URL(page['@odata.nextLink']);
      if (nextUrl.origin !== graphBaseUrl) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      nextUrl = null;
    }
  }

  return false;
}

app.http('validateSupplier', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const credentials = getCredentials();
    if (!credentials) {
      context.error('Supplier validation is not configured: SharePoint credentials are missing.');
      return jsonResponse(503, { error: 'בדיקת הספק אינה זמינה כרגע.' });
    }

    let payload;
    try {
      const body = await request.text();
      if (body.length > 4096) {
        return jsonResponse(413, { error: 'הבקשה גדולה מדי.' });
      }
      payload = JSON.parse(body);
    } catch {
      return jsonResponse(400, { error: 'הבקשה אינה תקינה.' });
    }

    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      return jsonResponse(400, { error: 'הבקשה אינה תקינה.' });
    }

    const { supplierNumber, supplierName, supplierEmail } = payload;
    if (
      typeof supplierNumber !== 'string'
      || typeof supplierName !== 'string'
      || typeof supplierEmail !== 'string'
      || supplierNumber.trim().length < 1
      || supplierNumber.trim().length > 64
      || supplierName.trim().length < 1
      || supplierName.trim().length > 200
      || supplierEmail.trim().length > 254
      || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supplierEmail.trim())
    ) {
      return jsonResponse(400, { error: 'נא למלא מספר ספק, שם ספק וכתובת דואר אלקטרוני תקינים.' });
    }

    try {
      const accessToken = await getAccessToken(credentials);
      const valid = await supplierExists({
        supplierNumber: normalize(supplierNumber),
        supplierName: normalize(supplierName),
        supplierEmail: normalize(supplierEmail)
      }, accessToken);

      return jsonResponse(200, { valid });
    } catch (error) {
      context.error('Supplier validation failed while contacting Microsoft Graph.', error);
      return jsonResponse(502, { error: 'לא ניתן לבדוק את פרטי הספק כרגע.' });
    }
  }
});
