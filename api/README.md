# Supplier validation API

The `POST /api/validateSupplier` endpoint checks that the supplier number, name,
and email match one item in the `Suppliers` list on the
[SaraTest SharePoint site](https://crdatasystems.sharepoint.com/sites/SaraTest/).
The matching SharePoint internal field names are `field_1`, `Title`, and
`fileld_7`, respectively.

Configure these application settings for the Azure Static Web App API:

- `SHAREPOINT_TENANT_ID`: Microsoft Entra tenant ID
- `SHAREPOINT_CLIENT_ID`: app registration's application (client) ID
- `SHAREPOINT_CLIENT_SECRET`: app registration secret

The app registration needs Microsoft Graph **application** permission
`Sites.Selected`, admin consent, and an explicit read grant to the SaraTest
site. Store the secret only in Azure application settings; do not add it to
the repository or send it to the browser.

The endpoint returns only whether all three values match. This checks supplied
details against the list; it does not authenticate the person or establish a
logged-in session. Use a second authentication step before granting access to
sensitive supplier data or actions.
