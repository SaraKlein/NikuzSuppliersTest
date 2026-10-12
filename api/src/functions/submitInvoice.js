const { app } = require('@azure/functions');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const siteHost = 'crdatasystems.sharepoint.com';
const sitePath = '/sites/SaraTest';
const siteUrl = `https://${siteHost}${sitePath}`;
const listName = 'Invoices';
const supplierListName = 'Suppliers';
const maxUploadBytes = 50 * 1024 * 1024;
const invoiceColumns = [
  'SupplierNumber',
  'SupplierName',
  'SupplierEmail',
  'InvoiceNumber',
  'InvoiceDate',
  'NetAmount',
  'VATAmount',
  'TotalAmount',
  'AllocationNumber',
  'DocumentType',
  'SupplierNotes'
];

function jsonResponse(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    jsonBody: body
  };
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

async function getAccessToken(credentials, scope) {
  const body = new URLSearchParams({
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
    scope,
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

async function graphRequest(url, accessToken, options, operation) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options && options.headers)
    },
    signal: AbortSignal.timeout(20000)
  });

  if (!response.ok) {
    throw new Error(`Microsoft Graph ${operation} request failed (${response.status})`);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

async function getInvoicesList(accessToken) {
  const site = await graphRequest(
    `${graphBaseUrl}/sites/${siteHost}:${sitePath}`,
    accessToken,
    null,
    'site lookup'
  );
  const listsUrl = new URL(`${graphBaseUrl}/sites/${site.id}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${listName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'list lookup');
  const list = lists.value.find((item) => item.displayName === listName);

  if (!list) {
    throw new Error(`SharePoint list "${listName}" was not found`);
  }

  return { site, list };
}

async function supplierExists(siteId, values, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${supplierListName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'supplier list lookup');
  const list = lists.value.find((item) => item.displayName === supplierListName);
  if (!list) {
    throw new Error(`SharePoint list "${supplierListName}" was not found`);
  }

  let itemsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${list.id}/items`);
  itemsUrl.searchParams.set('$expand', 'fields($select=field_1,Title,field_7)');
  itemsUrl.searchParams.set('$top', '200');
  while (itemsUrl) {
    const page = await graphRequest(itemsUrl, accessToken, null, 'supplier verification');
    const found = page.value.some((item) => {
      const fields = item.fields || {};
      return String(fields.field_1 || '').trim().toLowerCase() === values.supplierNumber.toLowerCase()
        && String(fields.Title || '').trim().toLowerCase() === values.supplierName.toLowerCase()
        && String(fields.field_7 || '').trim().toLowerCase() === values.supplierEmail.toLowerCase();
    });
    if (found) {
      return true;
    }

    if (page['@odata.nextLink']) {
      itemsUrl = new URL(page['@odata.nextLink']);
      if (itemsUrl.origin !== new URL(graphBaseUrl).origin) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      itemsUrl = null;
    }
  }

  return false;
}

async function getListColumns(siteId, listId, accessToken) {
  const url = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${listId}/columns`);
  url.searchParams.set('$select', 'name,displayName');
  const result = await graphRequest(url, accessToken, null, 'list columns lookup');
  const columns = new Map();
  for (const column of result.value) {
    columns.set(column.displayName.toLowerCase(), column.name);
  }
  return columns;
}

async function getRequestDigest(accessToken) {
  const response = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=verbose'
    },
    signal: AbortSignal.timeout(20000)
  });

  if (!response.ok) {
    throw new Error(`SharePoint request digest failed (${response.status})`);
  }

  const result = await response.json();
  const digest = result.d && result.d.GetContextWebInformation
    ? result.d.GetContextWebInformation.FormDigestValue
    : null;
  if (typeof digest !== 'string') {
    throw new Error('SharePoint request digest response was invalid');
  }

  return digest;
}

function getTextField(formData, name) {
  const value = formData.get(name);
  return typeof value === 'string' ? value.trim() : '';
}

function validDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }
  const date = new Date(`${value}T00:00:00.000Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function getInvoiceValues(formData) {
  const supplierNumber = getTextField(formData, 'verifiedNumber');
  const supplierName = getTextField(formData, 'verifiedName');
  const supplierEmail = getTextField(formData, 'verifiedEmail');
  const invoiceNumber = getTextField(formData, 'invoiceNumber');
  const invoiceDate = getTextField(formData, 'invoiceDate');
  const netText = getTextField(formData, 'invoiceNet');
  const vatText = getTextField(formData, 'invoiceVat');
  const allocationNumber = getTextField(formData, 'allocationNumber');
  const documentType = getTextField(formData, 'documentType');
  const supplierNotes = getTextField(formData, 'supplierNotes');
  const purchaseOrder = getTextField(formData, 'purchaseOrder');
  const netAmount = Number(netText);
  const vatAmount = Number(vatText);

  if (
    !supplierNumber
    || supplierNumber.length > 64
    || !supplierName
    || supplierName.length > 200
    || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supplierEmail)
    || !invoiceNumber
    || invoiceNumber.length > 64
    || !validDate(invoiceDate)
    || !netText
    || !Number.isFinite(netAmount)
    || netAmount < 0.01
    || !vatText
    || !Number.isFinite(vatAmount)
    || vatAmount < 0
    || !['tax-invoice', 'credit-invoice', 'invoice-receipt'].includes(documentType)
    || allocationNumber.length > 64
    || supplierNotes.length > 5000
    || purchaseOrder.length > 128
  ) {
    return null;
  }

  const totalAmount = Math.round((netAmount + vatAmount + Number.EPSILON) * 100) / 100;
  if (!Number.isFinite(totalAmount)) {
    return null;
  }

  return {
    supplierNumber,
    supplierName,
    supplierEmail,
    invoiceNumber,
    invoiceDate,
    netAmount,
    vatAmount,
    totalAmount,
    allocationNumber,
    documentType: {
      'tax-invoice': 'חשבונית מס',
      'credit-invoice': 'חשבונית זיכוי',
      'invoice-receipt': 'חשבונית מס / קבלה'
    }[documentType],
    supplierNotes,
    purchaseOrder
  };
}

function isFile(value) {
  return value && typeof value === 'object'
    && typeof value.name === 'string'
    && typeof value.arrayBuffer === 'function';
}

function safeFileName(name, index, usedNames) {
  const cleaned = name
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/[. ]+$/g, '')
    .slice(-180) || `attachment-${index + 1}`;
  const dot = cleaned.lastIndexOf('.');
  const stem = dot > 0 ? cleaned.slice(0, dot) : cleaned;
  const extension = dot > 0 ? cleaned.slice(dot) : '';
  let candidate = cleaned;
  let suffix = 2;
  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${stem}-${suffix}${extension}`;
    suffix += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

async function addAttachment(itemId, file, fileName, accessToken, digest) {
  const escapedFileName = encodeURIComponent(fileName.replace(/'/g, "''")).replace(/'/g, '%27');
  const response = await fetch(
    `${siteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})/AttachmentFiles/add(FileName='${escapedFileName}')`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'X-RequestDigest': digest,
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/octet-stream'
      },
      body: Buffer.from(await file.arrayBuffer()),
      signal: AbortSignal.timeout(30000)
    }
  );

  if (!response.ok) {
    throw new Error(`SharePoint attachment upload failed (${response.status})`);
  }
}

app.http('submitInvoice', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const credentials = getCredentials();
    if (!credentials) {
      context.error('Invoice submission is not configured: SharePoint credentials are missing.');
      return jsonResponse(503, { error: 'שליחת החשבונית אינה זמינה כרגע.' });
    }

    let formData;
    try {
      formData = await request.formData();
    } catch {
      return jsonResponse(400, { error: 'הטופס לא התקבל בפורמט תקין.' });
    }

    const values = getInvoiceValues(formData);
    const files = formData.getAll('invoiceFile').filter((file) => isFile(file) && file.name);
    const supportingFiles = formData.getAll('supportingFiles').filter((file) => isFile(file) && file.name);
    if (!values || files.length !== 1 || !/\.pdf$/i.test(files[0].name)) {
      return jsonResponse(400, { error: 'נא למלא את פרטי החשבונית ולצרף חשבונית בפורמט PDF.' });
    }

    const attachments = [files[0], ...supportingFiles];
    const totalFileSize = attachments.reduce((total, file) => total + file.size, 0);
    if (attachments.some((file) => file.size === 0) || totalFileSize > maxUploadBytes) {
      return jsonResponse(413, { error: 'הקבצים ריקים או שגודל הקבצים המצטבר עולה על 50MB.' });
    }

    try {
      const pdfHeader = Buffer.from(await files[0].slice(0, 5).arrayBuffer()).toString('ascii');
      if (pdfHeader !== '%PDF-') {
        return jsonResponse(400, { error: 'הקובץ המצורף אינו קובץ PDF תקין.' });
      }

      const [graphToken, sharePointToken] = await Promise.all([
        getAccessToken(credentials, 'https://graph.microsoft.com/.default'),
        getAccessToken(credentials, `https://${siteHost}/.default`)
      ]);
      const { site, list } = await getInvoicesList(graphToken);
      if (!await supplierExists(site.id, values, graphToken)) {
        return jsonResponse(403, { error: 'פרטי הספק אינם תואמים לרשימת הספקים.' });
      }
      const columns = await getListColumns(site.id, list.id, graphToken);
      const missingColumns = invoiceColumns.filter((name) => !columns.has(name.toLowerCase()));
      if (values.purchaseOrder && !columns.has('purchase order')) {
        missingColumns.push('Purchase Order');
      }
      if (missingColumns.length) {
        throw new Error(`Invoices list is missing columns: ${missingColumns.join(', ')}`);
      }

      const digest = await getRequestDigest(sharePointToken);
      const fields = {
        Title: values.invoiceNumber
      };
      const fieldValues = {
        'Supplier Number': values.supplierNumber,
        'Supplier Name': values.supplierName,
        'Supplier Email': values.supplierEmail,
        'Invoice Number': values.invoiceNumber,
        'Invoice Date': values.invoiceDate,
        'Net Amount': values.netAmount,
        'VAT Amount': values.vatAmount,
        'Total Amount': values.totalAmount,
        'Allocation Number': values.allocationNumber,
        'Document Type': values.documentType,
        'Supplier Notes': values.supplierNotes
      };
      if (values.purchaseOrder) {
        fieldValues['Purchase Order'] = values.purchaseOrder;
      }
      for (const [displayName, value] of Object.entries(fieldValues)) {
        fields[columns.get(displayName.toLowerCase())] = value;
      }

      const createdItem = await graphRequest(
        `${graphBaseUrl}/sites/${site.id}/lists/${list.id}/items`,
        graphToken,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fields })
        },
        'invoice creation'
      );

      const usedNames = new Set();
      try {
        for (let index = 0; index < attachments.length; index += 1) {
          const fileName = safeFileName(attachments[index].name, index, usedNames);
          await addAttachment(createdItem.id, attachments[index], fileName, sharePointToken, digest);
        }
      } catch (error) {
        try {
          await graphRequest(
            `${graphBaseUrl}/sites/${site.id}/lists/${list.id}/items/${createdItem.id}`,
            graphToken,
            { method: 'DELETE' },
            'incomplete invoice cleanup'
          );
        } catch (cleanupError) {
          context.error('Failed to remove an invoice with incomplete attachments.', cleanupError);
        }
        throw error;
      }

      return jsonResponse(201, { success: true, itemId: createdItem.id });
    } catch (error) {
      context.error('Invoice submission failed.', error);
      return jsonResponse(502, {
        error: 'לא ניתן לשמור את החשבונית כרגע. נא לבדוק את הגדרות SharePoint ולנסות שוב.'
      });
    }
  }
});
