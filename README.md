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
set -o pipefail
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
```

**PUT replaces the whole registration**, including guest apps and delegated
grants. Use only for this dedicated test type; review existing grants before
overwriting them. See [registration guidance](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/register-application-permissions).
Registration changes can take up to **one hour** to propagate; avoid repeated
container creation while waiting.

### Manage files and folders

These [file management](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/manage-files)
examples use the app-only `TOKEN` above and the existing `readContent` /
`writeContent` grants. Choose an existing **test** container from the list or UI;
its `b!` container ID is also its drive ID. Run the examples in order in the same
Bash session, stopping if any request fails.

```bash
CONTAINER_ID='b!REPLACE_WITH_TEST_CONTAINER_ID'

# Create a folder; fail rather than silently rename an existing folder.
FOLDER_ID=$(curl --fail-with-body --silent --show-error \
  -X POST "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/root/children" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  -H "Content-Type: application/json" \
  --data-binary '{"name":"curl-samples","folder":{},"@microsoft.graph.conflictBehavior":"fail"}' \
  | jq -er .id)

# Upload a small text file and save its durable DriveItem ID.
ITEM_ID=$(printf '%s\n' 'SharePoint Embedded curl sample' |
  curl --fail-with-body --silent --show-error \
    -X PUT "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$FOLDER_ID:/curl-sample.txt:/content" \
    -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
    -H "Content-Type: text/plain" --data-binary @- | jq -er .id)

# List the folder's children and read the uploaded file's metadata.
curl --fail-with-body --silent --show-error \
  "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$FOLDER_ID/children" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" | jq .
curl --fail-with-body --silent --show-error --get \
  "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  --data-urlencode '$select=id,name,size,eTag,parentReference' | jq .
```

For an existing local file, replace the `printf` pipeline with
`--data-binary @./your-file.txt` on the upload request. Simple Graph uploads allow
up to **250 MB** (the web app's separate 1 MiB limit does not apply to curl).
For larger files, use a
[chunked upload session](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0).
Do not overwrite real content: uploading to an existing file path replaces it.

Download via the short-lived, preauthenticated URL returned by Graph. Treat that
URL as a secret: do not print or share it, and do not send the Graph bearer token
to the download host. These commands print the sample file's text to the terminal;
use `--output ./download.txt` instead for binary files.

```bash
DOWNLOAD_URL=$(curl --fail-with-body --silent --show-error \
  "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  | jq -er '."@microsoft.graph.downloadUrl"')
curl --fail-with-body --silent --show-error --proto '=https' "$DOWNLOAD_URL"
unset DOWNLOAD_URL

# Replace content while retaining the item ID.
printf '%s\n' 'Updated SharePoint Embedded curl sample' |
  curl --fail-with-body --silent --show-error \
    -X PUT "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID/content" \
    -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
    -H "Content-Type: text/plain" --data-binary @- | jq .

# Rename and move the sample file from its folder to the container root.
ROOT_ID=$(curl --fail-with-body --silent --show-error \
  "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/root" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" | jq -er .id)
jq -n --arg parentId "$ROOT_ID" \
  '{name: "curl-sample-renamed.txt", parentReference: {id: $parentId}}' |
  curl --fail-with-body --silent --show-error \
    -X PATCH "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID" \
    -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
    -H "Content-Type: application/json" --data-binary @- | jq .

# Inspect available versions.
curl --fail-with-body --silent --show-error \
  "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID/versions" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" | jq .
```

**Delete only the sample file you just created.** Delete returns `204 No Content`,
so do not pipe its response to `jq`. The restore example uses the original
DriveItem ID as an alternate key, which is available in Graph **beta**; it is not
the recycle-bin item's ID. Deleted content is retained for up to 93 days.

```bash
curl --fail-with-body --silent --show-error \
  -X DELETE "https://graph.microsoft.com/v1.0/drives/$CONTAINER_ID/items/$ITEM_ID" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")"

# Inspect deleted items, then restore the sample by its original DriveItem ID.
curl --fail-with-body --silent --show-error \
  "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/$CONTAINER_ID/recycleBin/items" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" | jq .
curl --fail-with-body --silent --show-error \
  -X POST "https://graph.microsoft.com/beta/storage/fileStorage/containers/$CONTAINER_ID/recycleBin/items(driveItemId='$ITEM_ID')/restore" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  -H "Content-Type: application/json" | jq .
```

### Search containers and files (delegated, preview)

[SharePoint Embedded search](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/search-containers-files)
does **not** support the app-only `TOKEN`. Use a Microsoft Graph delegated token
from your app's user sign-in flow with `Files.Read.All` and
`FileStorageContainer.Selected`, appropriate consent, container-type delegated
grants, and user access to the target content. The web app has no such sign-in
flow; these samples require separate delegated authentication. Do not assume
`az account get-access-token` has your app's delegated permissions.

Read that token without echoing it or putting it in shell history. Both searches
use `/beta` because `includeHiddenContent` is beta-only. Always include
`ContainerTypeId` to restrict results to your app; search otherwise returns
matching content the user can access even outside your app's container type.

```bash
read -rs -p 'Delegated Microsoft Graph token: ' DELEGATED_TOKEN
printf '\n'

# Search all containers of this application's type.
jq -n --arg typeId "$CONTAINER_TYPE_ID" \
  '{requests: [{
    entityTypes: ["drive"],
    query: {queryString: ("ContainerTypeId:" + $typeId)},
    sharePointOneDriveOptions: {includeHiddenContent: true},
    from: 0, size: 25
  }]}' |
  curl --fail-with-body --silent --show-error \
    -X POST "https://graph.microsoft.com/beta/search/query" \
    -H "$(printf 'Authorization: %s %s' Bearer "$DELEGATED_TOKEN")" \
    -H "Content-Type: application/json" --data-binary @- | jq .

# Search files/folders in the selected container by title.
jq -n --arg typeId "$CONTAINER_TYPE_ID" --arg containerId "$CONTAINER_ID" \
  '{requests: [{
    entityTypes: ["driveItem"],
    query: {queryString: ("Title:curl-sample AND ContainerId:" + $containerId +
      " AND ContainerTypeId:" + $typeId)},
    sharePointOneDriveOptions: {includeHiddenContent: true},
    from: 0, size: 25
  }]}' |
  curl --fail-with-body --silent --show-error \
    -X POST "https://graph.microsoft.com/beta/search/query" \
    -H "$(printf 'Authorization: %s %s' Bearer "$DELEGATED_TOKEN")" \
    -H "Content-Type: application/json" --data-binary @- | jq .
unset DELEGATED_TOKEN
```

For full-text search across the type, replace the second query with
`curl-sample AND ContainerTypeId:<your-type-id>`. New or changed content may not
be indexed immediately. Page with `from` and `size`, checking
`hitsContainers[].moreResultsAvailable` and `total` before requesting another page.

### Archive and reactivate containers (preview)

[Archival](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/archive-restore-containers)
requires Microsoft 365 Archive availability and billing configuration in the
tenant. Do not assume the free trial supports it. A SharePoint Embedded
administrator must first enable archival on the container type using admin
PowerShell (not included in this devcontainer):

```powershell
Set-SPOContainerTypeConfiguration -ContainerTypeId <your-type-id> -IsArchiveEnabled $true
```

These calls can use the app-only `TOKEN` with application
`FileStorageContainer.Selected` and container-type `write` permission. With
delegated authentication, the caller must be an `Owner` / `PrincipalOwner` or an
authorized administrator. **Archiving makes the entire container inaccessible**;
use only the chosen test container, stop file operations first, and run each
lifecycle command deliberately. Archive and unarchive return `204 No Content`.

```bash
curl --fail-with-body --silent --show-error \
  -X POST "https://graph.microsoft.com/beta/storage/fileStorage/containers/$CONTAINER_ID/archive" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")"

# Inspect archive state (archivalDetails is beta-only).
curl --fail-with-body --silent --show-error --get \
  "https://graph.microsoft.com/beta/storage/fileStorage/containers/$CONTAINER_ID" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  --data-urlencode '$select=id,displayName,archivalDetails' | jq .

# List archived containers belonging to this type.
curl --fail-with-body --silent --show-error --get \
  "https://graph.microsoft.com/beta/storage/fileStorage/containers" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")" \
  --data-urlencode "\$filter=containerTypeId eq $CONTAINER_TYPE_ID and archivalDetails ne null" \
  | jq .

# Request reactivation; inspect archivalDetails again before accessing files.
curl --fail-with-body --silent --show-error \
  -X POST "https://graph.microsoft.com/beta/storage/fileStorage/containers/$CONTAINER_ID/unarchive" \
  -H "$(printf 'Authorization: %s %s' Bearer "$TOKEN")"
```

Reactivation is instantaneous in the first seven days after archival; later it
can take up to 24 hours. Keep reads and writes disabled until the service reports
the container is active again. Archival is not deletion or recycle-bin restore.
Graph beta APIs are preview-only and subject to change; these direct curl
examples do not enable beta endpoints in the web app.

When finished with all curl samples, clear credentials and temporary IDs:

```bash
unset TOKEN CLIENT_SECRET DELEGATED_TOKEN DOWNLOAD_URL CONTAINER_ID FOLDER_ID ITEM_ID ROOT_ID
```

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
