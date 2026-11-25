# SharePoint Embedded test workspace

A devcontainer and a small, dependency-free Node.js web app for evaluating
[SharePoint Embedded end to end](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/build-app-end-to-end).
The app creates/lists containers and uploads/lists files in the selected container's root.
It uses a confidential backend with **app-only** Microsoft Graph access; secrets and
tokens never go to the browser.

## Development environment

Open this repository in VS Code and choose **Dev Containers: Reopen in Container**
(requires Docker and the Dev Containers extension), or create a GitHub Codespace.
The container includes:

- Node.js 22 and npm.
- Azure CLI (`az`) for Entra app registration, permissions, and tenant sign-in.
- `curl`, `jq`, and CA certificates for Graph requests and JSON inspection.
- The SharePoint Embedded VS Code extension for guided container-type creation
  and registration. Complete its browser sign-in when prompted.
- Private forwarding of port 3000.

The post-create command checks tool availability. No npm dependencies need installing.
Azure CLI cannot replace container-type creation/registration: those are separate
SharePoint Embedded operations. PowerShell is not required for this trial path;
Microsoft's Graph APIs and VS Code extension handle it.

## 1. Prepare a test tenant and owning app

You need a Microsoft 365 tenant with SharePoint available and an administrator
who can grant consent and perform SharePoint Embedded setup. This sample targets
the public Microsoft cloud and a **single tenant**, not a multi-geo tenant.
An Azure subscription is not needed for the trial.

Follow the [VS Code quickstart](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/quickstart-vscode)
in the SharePoint Embedded view:

1. Sign in with your test tenant administrator.
2. Select **Create Container Type**, choose **trial**, and create a new owning app
   (or select a dedicated test app created with Azure CLI below).
3. Record the owning app's **Application (client) ID**, **Directory (tenant) ID**,
   and **container type ID**.
4. Register the type in the local/consuming tenant. Even when it is the same
   tenant, registration is mandatory.

One owning app can own exactly one container type. A trial lasts **30 days**, has
at most **five containers including recycled containers**, and allows **1 GB per
container**. It cannot be converted to production. Do not create more types or
containers on every restart; the app only creates a container when you request it.

### Optional: create/configure the owning app with Azure CLI

Run in the devcontainer terminal, substituting your tenant GUID:

```bash
az login --tenant YOUR_TENANT_GUID --allow-no-subscriptions --use-device-code
CLIENT_ID=$(az ad app create --display-name "SharePoint Embedded local test" \
  --sign-in-audience AzureADMyOrg --query appId --output tsv)
az ad sp create --id "$CLIENT_ID"
echo "$CLIENT_ID"
```

Select this existing app in the extension's create-type flow. Do not select a
production app: the extension updates its configuration.

After creating/selecting the app, verify it has these Graph **application**
permissions (not just delegated permissions):

| Permission | Purpose |
| --- | --- |
| `FileStorageContainer.Selected` | Sample's container and file operations |
| `FileStorageContainerTypeReg.Selected` | One-time registration by the owning app |

To add these permissions without hardcoding permission GUIDs:

```bash
# Set CLIENT_ID to the owning app ID if the extension created it.
GRAPH_APP_ID=00000003-0000-0000-c000-000000000000
GRAPH_SP=$(az ad sp show --id "$GRAPH_APP_ID" --output json)
for permission in FileStorageContainer.Selected FileStorageContainerTypeReg.Selected; do
  ROLE_ID=$(printf '%s' "$GRAPH_SP" | jq -er --arg permission "$permission" \
    '.appRoles[] | select(.value == $permission and .isEnabled and (.allowedMemberTypes | index("Application"))) | .id')
  az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_APP_ID" \
    --api-permissions "$ROLE_ID=Role"
done
az ad app permission admin-consent --id "$CLIENT_ID"
```

An authorized tenant administrator must run consent. Your Entra policy may require
browser/admin approval instead. Azure CLI sign-in alone does **not** consent the
sample app's permissions.

Container-type creation uses **delegated** `FileStorageContainerType.Manage.All`,
not client credentials; the extension handles that sign-in. Do not attempt to
create a type with the sample's app-only token or assume an Azure CLI Graph token
has the extension's delegated permissions.

## 2. Register container-type application permissions

In the extension, register the type and verify the owning app's **application**
permission grants cover `create`, `read`, `write`, `readContent`, and `writeContent`
(or `full` for this isolated trial). No delegated grants are needed by this sample.
Graph consent alone does not grant container access.

In Entra **App registrations > your owning app > Certificates & secrets**, create
a short-lived client secret for local testing. Copy its **Value**, not its secret
ID. Use certificates or managed identity for a production backend.

## 3. Configure and run

From the repository root:

```bash
cp .env.example .env
chmod 600 .env
```

Edit `.env` locally and fill in `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, and
`CONTAINER_TYPE_ID`. GUIDs must be actual IDs, not the SharePoint site URL.
The `.env` file is ignored by Git. Do not paste credentials into source, terminal
recordings, issues, or screenshots.

```bash
npm start
```

Open **http://localhost:3000** using a private localhost port forward. Create a
container, select it, upload a small file, and confirm its name and size appear.
Uploads are limited to **1 MiB** and replace existing content with the same name.
The UI lists root folders but does not navigate them.

**This is not a production app.** It has no user sign-in: every caller acts as the
owning app. Keep the forwarded port private and never expose it publicly, on your
LAN, or behind a shared proxy. The server binds to loopback and accepts only
localhost hosts and same-origin browser mutations. In Codespaces, use VS Code's
local port forwarding; the hosted `*.app.github.dev` URL is deliberately rejected.
Use delegated authentication and container roles for a real user-facing app.

## Direct Graph smoke test with curl

For the same app-only identity, load your trusted local `.env` in Bash and acquire
a token (do not enable shell tracing or echo the token):

```bash
set -a
source .env
set +a
TOKEN=$(curl --fail-with-body --silent --show-error \
  -X POST "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" \
  --data-urlencode "client_id=$CLIENT_ID" \
  --data-urlencode "client_secret=$CLIENT_SECRET" \
  --data-urlencode "scope=https://graph.microsoft.com/.default" \
  --data-urlencode "grant_type=client_credentials" | jq -er .access_token)
curl --fail-with-body --silent --show-error --get \
  "https://graph.microsoft.com/v1.0/storage/fileStorage/containers" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  --data-urlencode "\$filter=containerTypeId eq $CONTAINER_TYPE_ID" | jq .
```

If the extension's registration did not include app-only grants, the owning app
can explicitly register them using the same token:

```bash
jq -n --arg appId "$CLIENT_ID" \
  '{applicationPermissionGrants: [{appId: $appId, delegatedPermissions: ["none"],
    applicationPermissions: ["create", "read", "write", "readContent", "writeContent"]}]}' |
  curl --fail-with-body --silent --show-error \
    -X PUT "https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypeRegistrations/$CONTAINER_TYPE_ID" \
    -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" -H "Content-Type: application/json" \
    --data-binary @- | jq .
unset TOKEN CLIENT_SECRET
```

**PUT replaces the whole registration**, including guest apps and delegated
grants. Use only for this dedicated test type; review existing grants before
overwriting them. See [registration guidance](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/register-application-permissions).
Registration changes can take up to **one hour** to propagate; avoid repeated
container creation while waiting.

## Validation and troubleshooting

```bash
npm run check
npm test
```

These use only Node's built-in syntax checker and test runner. Tests mock Entra
and Graph; no live tenant or credentials are required, and there is no build step.

- Missing configuration: fill in all four values in `.env`, then restart.
- Authentication/401: check tenant, client ID, secret value/expiry, and consent.
- Access denied/403: verify Graph **application** permissions, admin consent,
  container-type registration, and its application grants.
- No containers: create one after registration has propagated.
- Trial expired or quota exceeded: inspect the trial in the extension; recycled
  containers still count. Clean up test containers there and revoke/delete the
  test app's credentials when finished.
- Standard/production types: configure Azure billing before creating containers;
  follow Microsoft's end-to-end guide rather than using this free-trial path.
