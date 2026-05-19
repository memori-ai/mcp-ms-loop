# mcp-ms-loop — Complete guide

MCP server that allows AI Agents to read **Microsoft Loop** files via Microsoft Graph API, using **app-only authentication** (client credentials). Works for Loop files stored anywhere:

- Personal OneDrive
- SharePoint sites (Teams channels)
- **Microsoft Loop app** (SharePoint Embedded, `CSP_…` containers)

---

## Table of Contents

1. [Architecture Overview](#1-architecture-overview)
2. [Prerequisites](#2-prerequisites)
3. [Create the App Registration on Azure](#3-create-the-app-registration-on-azure)
4. [Assign API Permissions](#4-assign-api-permissions)
5. [SPE Registration — required for the Loop app](#5-spe-registration--required-for-the-loop-app)
6. [Configure the MCP in the gateway](#6-configure-the-mcp-in-the-gateway)
7. [Available Tools](#7-available-tools)
8. [Minimum Required Permissions](#8-minimum-required-permissions)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Architecture Overview

```
AI Agent
   │
   ▼
MCP Gateway (Rails)
   │  injects TENANT_ID, CLIENT_ID, CLIENT_SECRET
   │  as environment variables
   ▼
npx mcp-ms-loop   ← this package
   │
   ▼
Microsoft Graph API
   │
   ├── /drives/{driveId}/items/{itemId}/content?format=html
   │     ↓ redirect to westeurope1-mediap.svc.ms (HTML converter)
   │
   └── /v1.0/search/query  (for workspace discovery)
```

Credentials **never pass through the LLM**: they are injected by the gateway as environment variables, invisible to the Agent.

---

## 2. Prerequisites

| Requirement | Details |
|---|---|
| Azure Portal access | **Application Administrator** role or higher on the tenant |
| Global Admin access | Required only for the PowerShell step (§5), one time only |
| PowerShell (`pwsh`) | For the SPE step. Works on Windows (native) and Linux/macOS via GitHub Codespaces or Azure Cloud Shell |
| Microsoft 365 tenant | With a license that includes Microsoft Loop (M365 Business Standard/Premium, E3/E5) |

---

## 3. Create the App Registration on Azure

1. Go to **[portal.azure.com](https://portal.azure.com)** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Settings:
   - **Name**: e.g. `MCP MS Loop`
   - **Supported account types**: `Accounts in this organizational directory only`
   - **Redirect URI**: leave empty (not needed for app-only)
3. Click **Register**
4. From the app page, note:
   - **Application (client) ID** → `CLIENT_ID`
   - **Directory (tenant) ID** → `TENANT_ID`
5. Go to **Certificates & secrets** → **New client secret**
   - Description: e.g. `mcp-loop-secret`
   - Expires: choose an appropriate expiration (e.g. 24 months)
   - Click **Add** and immediately copy the **Value** → `CLIENT_SECRET`

> ⚠️ The `CLIENT_SECRET` is only visible at creation time. Copy it immediately.

---

## 4. Assign API Permissions

From the app page → **API permissions** → **Add a permission**:

### Required permissions (minimum)

| API | Type | Permission | Admin consent |
|---|---|---|---|
| Microsoft Graph | Application | `Files.Read.All` | Yes |
| Microsoft Graph | Application | `FileStorageContainer.Selected` | Yes |

### Temporary permission (only for §5, removable afterwards)

| API | Type | Permission | Admin consent |
|---|---|---|---|
| SharePoint | Application | `Sites.FullControl.All` | Yes |

After adding all permissions:
- Click **Grant admin consent for \<tenant\>**
- Verify all show status ✅ **Granted**

> ℹ️ `Sites.FullControl.All` is only needed to run `Connect-SPOService` in PowerShell (§5). Once that step is complete, it can be **removed**.

---

## 5. SPE Registration — required for the Loop app

Files created in the **Microsoft Loop app** live in **SharePoint Embedded (SPE)** containers. These containers have a separate authorization model: Graph permissions alone are not enough — you must register the app as a "guest application" on the Loop container type via PowerShell.

This step:
- Is performed **once per tenant**
- Requires a **Global Admin** account
- Cannot be replaced by Graph API or the Azure interface

### Loop container type ID (constant across all tenants)

```
a187e399-0c36-4b98-8f04-1edc167a0996
```

---

### Option A — Windows (simplest)

Open PowerShell as administrator:

```powershell
# Install the SPO module
Install-Module Microsoft.Online.SharePoint.PowerShell -Force
Import-Module Microsoft.Online.SharePoint.PowerShell

# Connect with interactive login (opens browser)
Connect-SPOService -Url "https://<TENANT_NAME>-admin.sharepoint.com" -Interactive

# Register the app as guest on the Loop container type
Set-SPOApplicationPermission `
  -OwningApplicationId "a187e399-0c36-4b98-8f04-1edc167a0996" `
  -GuestApplicationId  "<CLIENT_ID>" `
  -PermissionAppOnly   "readcontent"
```

Replace:
- `<TENANT_NAME>` with your tenant prefix (e.g. `contoso`)
- `<CLIENT_ID>` with the Application ID created in §3

No output from the command → it worked. Absence of output means success.

---

### Option B — Linux / macOS / GitHub Codespaces

`Connect-SPOService -Interactive` does not work on Linux. Use **certificate-based** authentication instead.

#### B1. Generate a self-signed certificate and upload it to Azure

```bash
# Generate private key and certificate (once)
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem \
  -days 730 -nodes -subj "/CN=mcp-loop"

# Create .pfx (with password) for PowerShell
openssl pkcs12 -export \
  -inkey key.pem -in cert.pem \
  -out cert.pfx -passout pass:temp123

# Show SHA1 fingerprint to verify against Azure
openssl x509 -in cert.pem -fingerprint -sha1 -noout
```

On **Azure Portal** → App Registration → **Certificates & secrets** → **Certificates** → **Upload certificate**:
- Upload `cert.pem`
- Verify the thumbprint shown on Azure matches the `openssl x509 … -fingerprint` output

#### B2. Install PowerShell and the SPO module

```bash
# Install pwsh (if not present)
# On Debian/Ubuntu:
sudo apt-get install -y powershell

# In GitHub Codespaces it is already available, just start the shell:
pwsh
```

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Force
Import-Module Microsoft.Online.SharePoint.PowerShell -Force
```

#### B3. Fix the MSAL dependency (Linux only)

The SPO module on Linux requires a DLL that is not included automatically:

```bash
# From bash (not pwsh), find where the module is installed:
SPOMOD=$(pwsh -c "Split-Path (Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1).Path")
echo "Module path: $SPOMOD"

# Create a temporary dotnet project to download the DLL
mkdir -p /tmp/getmsal && cd /tmp/getmsal
dotnet new console -n getmsal --force
cd getmsal
dotnet add package Microsoft.Identity.Client --version 4.74.1
dotnet build -o out

# Copy the DLL into the SPO module directory
cp out/Microsoft.Identity.Client.dll "$SPOMOD/"
```

#### B4. Connect and register

```powershell
Import-Module Microsoft.Online.SharePoint.PowerShell -Force

Connect-SPOService `
  -Url             "https://<TENANT_NAME>-admin.sharepoint.com" `
  -ClientId        "<CLIENT_ID>" `
  -TenantId        "<TENANT_ID>" `
  -CertificatePath "/path/to/cert.pfx" `
  -CertificatePassword (ConvertTo-SecureString "temp123" -AsPlainText -Force)

Set-SPOApplicationPermission `
  -OwningApplicationId "a187e399-0c36-4b98-8f04-1edc167a0996" `
  -GuestApplicationId  "<CLIENT_ID>" `
  -PermissionAppOnly   "readcontent"
```

> ℹ️ `Connect-SPOService` requires `Sites.FullControl.All` (Application) on the app (added in §4). Once this step is complete, you can **remove** that permission from Azure to reduce attack surface.

---

## 6. Configure the MCP in the gateway

### Startup command (url_command_hidden)

```
npx -y mcp-ms-loop
```

The `-y` flag ensures npx always downloads the latest published version without prompting.

### Custom parameters (JSON to enter in the gateway "Parameters" section)

```json
{
  "TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "CLIENT_SECRET": "the-secret-value-NOT-the-id"
}
```

> ⚠️ `CLIENT_SECRET` is the **Value** of the secret (the long string like `abc123~XYZ...`), not the UUID/ID shown in the secrets list on Azure.

These values are injected as environment variables into the Node.js process. They never appear in messages exchanged with the LLM.

---

## 7. Available Tools

### `getLoopByShareUrl`

Retrieves the full content of a Loop file starting from its sharing URL.

**Parameters:**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `shareUrl` | string | Yes | Full Loop file URL (the link you open in the browser) |
| `includeText` | boolean | No (default `true`) | If `true`, also returns a plain-text version (HTML stripped) |

**Response:**

```json
{
  "name": "Meeting Notes.loop",
  "driveId": "b!XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  "itemId": "01XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  "webUrl": "https://contoso.sharepoint.com/...",
  "lastModified": "2026-01-15T10:30:00Z",
  "size": 48200,
  "mimeType": "application/octet-stream",
  "html": "<html><meta name=\"GENERATOR\" content=\"Microsoft Loop\"><body>...</body></html>",
  "text": "Meeting Notes\n\nAgenda\n1. Project updates...",
  "truncated": false,
  "textTruncated": false
}
```

---

### `getLoopByItemId`

Direct variant: uses `driveId` and `itemId` if already known (faster, skips share URL resolution).

**Parameters:**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `driveId` | string | Yes | ID of the drive containing the file |
| `itemId` | string | Yes | ID of the file |
| `includeText` | boolean | No (default `true`) | Include plain-text version |

---

### `listLoopContainers`

Discovers all Microsoft Loop workspaces (drives) accessible to the app. Uses the Microsoft 365 Search API to find all `.loop` files and group them by drive/workspace.

**No parameters required.**

**Response:**

```json
[
  {
    "driveId": "b!XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
    "name": "Team Workspace",
    "webUrl": "https://contoso.sharepoint.com/contentstorage/CSP_.../Document%20Library",
    "loopFileCount": 8
  },
  {
    "driveId": "b!YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY",
    "name": "OneDrive",
    "webUrl": "https://contoso-my.sharepoint.com/personal/.../Documents",
    "loopFileCount": 2
  }
]
```

---

### `listLoopFilesInDrive`

Lists all `.loop` and `.fluid` files inside a drive. The scan is **recursive** (Loop files in the Loop app are stored inside the `LoopAppData/` subfolder).

**Parameters:**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `driveId` | string | Yes | Drive ID to scan (obtained from `listLoopContainers`) |
| `path` | string | No (default `root`) | Starting path inside the drive |

**Response:**

```json
[
  {
    "id": "01XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
    "name": "Meeting Notes.loop",
    "driveId": "b!XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
    "size": 48200,
    "webUrl": "https://contoso.sharepoint.com/...",
    "lastModified": "2026-01-15T10:30:00Z",
    "mimeType": "application/octet-stream"
  }
]
```

---

## 8. Minimum Required Permissions

Once the SPE registration (§5) is complete, the **final** app configuration only needs these permissions:

| API | Type | Permission | Reason |
|---|---|---|---|
| Microsoft Graph | Application | `Files.Read.All` | Read Loop files |
| Microsoft Graph | Application | `FileStorageContainer.Selected` | Access SPE containers (Loop workspaces) |

Everything else (delegated permissions, `Sites.Read.All`, `User.Read`, `offline_access`, `Sites.FullControl.All`) can be removed.

---

## 9. Troubleshooting

### `403 accessDenied` on Loop app files

The §5 step (`Set-SPOApplicationPermission`) was not executed or did not take effect. Verify that the correct `CLIENT_ID` was used and that you were connected as Global Admin. Propagation may take a few minutes.

### `400 failed to parse filter parameter` on `listLoopContainers`

Old package version (≤ 1.0.3). Update to the latest version: the tool now uses the Search API instead of the SPE container API, which is not accessible with minimum permissions.

### `Connect-SPOService: Object reference not set to an instance of an object`

You are using `Connect-SPOService -Interactive` on Linux. It does not work. Use certificate-based authentication (§5 Option B).

### `Could not load file or assembly 'Microsoft.Identity.Client'`

The MSAL DLL is missing in the Linux environment. Follow step B3 in §5.

### `Connect-SPOService: (401) Unauthorized` (with certificate)

The app is missing `SharePoint > Sites.FullControl.All` (Application) with admin consent. Add it temporarily, run `Set-SPOApplicationPermission`, then remove it.

### `listLoopContainers` returns empty array `[]`

The Search API found no accessible Loop files. Possible causes:
1. Step §5 was not completed → SPE files not accessible
2. There are no `.loop` files in the tenant
3. The tenant uses a region other than EMEA/NAM/APAC — contact Microsoft support

### `CLIENT_SECRET` has expired

Azure shows the expiration date in the **Certificates & secrets** section. If expired, generate a new secret and update the JSON parameters in the gateway.

---

## References

- [Microsoft Graph — Drive Items](https://learn.microsoft.com/en-us/graph/api/driveitem-get)
- [SharePoint Embedded — Loop HTML export technical guide](https://wals.pro/blogs/news/microsoft-loop-to-html-export-technical-guide)
- [Set-SPOApplicationPermission (Microsoft Docs)](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spoapplicationpermission)
- [Microsoft 365 Search API — App-only](https://learn.microsoft.com/en-us/graph/search-concept-searchall)
