# mcp-ms-loop — Guida completa per system integrator

Server MCP che permette agli Agenti AI di leggere i file **Microsoft Loop** via Microsoft Graph API, con autenticazione **app-only** (client credentials). Funziona per file Loop ovunque siano salvati:

- OneDrive personale
- Siti SharePoint (canali Teams)
- **App Microsoft Loop** (SharePoint Embedded, container `CSP_…`)

---

## Indice

1. [Panoramica architetturale](#1-panoramica-architetturale)
2. [Prerequisiti](#2-prerequisiti)
3. [Creare l'App Registration su Azure](#3-creare-lapp-registration-su-azure)
4. [Assegnare i permessi API](#4-assegnare-i-permessi-api)
5. [Registrazione SPE — passaggio obbligatorio per l'app Loop](#5-registrazione-spe--passaggio-obbligatorio-per-lapp-loop)
6. [Configurare l'MCP nel gateway](#6-configurare-lmcp-nel-gateway)
7. [Tool disponibili](#7-tool-disponibili)
8. [Permessi minimi necessari](#8-permessi-minimi-necessari)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Panoramica architetturale

```
Agente AI
   │
   ▼
MCP Gateway (Rails)
   │  inietta TENANT_ID, CLIENT_ID, CLIENT_SECRET
   │  come variabili d'ambiente
   ▼
npx mcp-ms-loop   ← questo pacchetto
   │
   ▼
Microsoft Graph API
   │
   ├── /drives/{driveId}/items/{itemId}/content?format=html
   │     ↓ redirect a westeurope1-mediap.svc.ms (converter HTML)
   │
   └── /v1.0/search/query  (per discovery dei workspace)
```

Le credenziali **non passano mai dall'LLM**: vengono iniettate dal gateway come variabili d'ambiente, invisibili all'Agente.

---

## 2. Prerequisiti

| Requisito | Dettaglio |
|---|---|
| Accesso Azure Portal | Ruolo **Application Administrator** o superiore sul tenant |
| Accesso Global Admin | Necessario solo per il passaggio PowerShell (§5), una volta sola |
| PowerShell con `pwsh` | Per il passaggio SPE. Funziona su Windows (nativo) e Linux/macOS via GitHub Codespaces o Azure Cloud Shell |
| Tenant Microsoft 365 | Con licenza che include Microsoft Loop (M365 Business Standard/Premium, E3/E5) |

---

## 3. Creare l'App Registration su Azure

1. Vai su **[portal.azure.com](https://portal.azure.com)** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Impostazioni:
   - **Name**: es. `MCP MS Loop`
   - **Supported account types**: `Accounts in this organizational directory only`
   - **Redirect URI**: lascia vuoto (non serve per app-only)
3. Clicca **Register**
4. Dalla pagina dell'app, annota:
   - **Application (client) ID** → `CLIENT_ID`
   - **Directory (tenant) ID** → `TENANT_ID`
5. Vai su **Certificates & secrets** → **New client secret**
   - Description: es. `mcp-loop-secret`
   - Expires: scegli la scadenza appropriata (es. 24 mesi)
   - Clicca **Add** e copia subito il **Value** → `CLIENT_SECRET`

> ⚠️ Il `CLIENT_SECRET` è visibile solo al momento della creazione. Copialo subito.

---

## 4. Assegnare i permessi API

Dalla pagina dell'app → **API permissions** → **Add a permission**:

### Permessi richiesti (minimi)

| API | Tipo | Permesso | Admin consent |
|---|---|---|---|
| Microsoft Graph | Application | `Files.Read.All` | Sì |
| Microsoft Graph | Application | `FileStorageContainer.Selected` | Sì |

### Permesso temporaneo (solo per §5, poi rimuovibile)

| API | Tipo | Permesso | Admin consent |
|---|---|---|---|
| SharePoint | Application | `Sites.FullControl.All` | Sì |

Dopo aver aggiunto tutti i permessi:
- Clicca **Grant admin consent for \<tenant\>**
- Verifica che tutti mostrino lo stato ✅ **Granted**

> ℹ️ `Sites.FullControl.All` serve solo per eseguire `Connect-SPOService` in PowerShell (§5). Una volta completato quel passaggio, può essere **rimosso**.

---

## 5. Registrazione SPE — passaggio obbligatorio per l'app Loop

I file creati nell'**app Microsoft Loop** vivono in container **SharePoint Embedded (SPE)**. Questi container hanno un modello di autorizzazione separato: non basta avere permessi Graph, bisogna registrare l'app come "guest application" sul container type di Loop tramite PowerShell.

Questo passaggio:
- Si esegue **una sola volta** per tenant
- Richiede un account **Global Admin**
- Non è sostituibile via Graph API o interfaccia Azure

### ID del container type di Loop (costante per tutti i tenant)

```
a187e399-0c36-4b98-8f04-1edc167a0996
```

---

### Opzione A — Windows (più semplice)

Apri PowerShell come amministratore:

```powershell
# Installa il modulo SPO
Install-Module Microsoft.Online.SharePoint.PowerShell -Force
Import-Module Microsoft.Online.SharePoint.PowerShell

# Connetti con login interattivo (apre il browser)
Connect-SPOService -Url "https://<TENANT_NAME>-admin.sharepoint.com" -Interactive

# Registra l'app come guest sul container type di Loop
Set-SPOApplicationPermission `
  -OwningApplicationId "a187e399-0c36-4b98-8f04-1edc167a0996" `
  -GuestApplicationId  "<CLIENT_ID>" `
  -PermissionAppOnly   "readcontent"
```

Sostituisci:
- `<TENANT_NAME>` con il prefisso del tuo tenant (es. `contoso`)
- `<CLIENT_ID>` con l'Application ID dell'app creata al §3

Se il comando non dà output → ha funzionato. L'assenza di output è il successo.

---

### Opzione B — Linux / macOS / GitHub Codespaces

`Connect-SPOService -Interactive` non funziona su Linux. Usa l'autenticazione via **certificato**.

#### B1. Genera un certificato self-signed e caricalo su Azure

```bash
# Genera chiave privata e certificato (1 sola volta)
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem \
  -days 730 -nodes -subj "/CN=mcp-loop"

# Crea il .pfx (con password) per PowerShell
openssl pkcs12 -export \
  -inkey key.pem -in cert.pem \
  -out cert.pfx -passout pass:temp123

# Mostra il fingerprint SHA1 per confrontarlo su Azure
openssl x509 -in cert.pem -fingerprint -sha1 -noout
```

Su **Azure Portal** → App Registration → **Certificates & secrets** → **Certificates** → **Upload certificate**:
- Carica `cert.pem`
- Verifica che il thumbprint mostrato su Azure corrisponda all'output del comando `openssl x509 … -fingerprint`

#### B2. Installa PowerShell e il modulo SPO

```bash
# Installa pwsh (se non presente)
# Su Debian/Ubuntu:
sudo apt-get install -y powershell

# In GitHub Codespaces è già disponibile, avvia la shell:
pwsh
```

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Force
Import-Module Microsoft.Online.SharePoint.PowerShell -Force
```

#### B3. Risolvi la dipendenza MSAL (solo Linux)

Il modulo SPO su Linux necessita di una DLL che non viene inclusa automaticamente:

```bash
# Dalla bash (non da pwsh), individua dove è installato il modulo:
SPOMOD=$(pwsh -c "Split-Path (Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1).Path")
echo "Modulo in: $SPOMOD"

# Crea un progetto dotnet temporaneo per scaricare la DLL
mkdir -p /tmp/getmsal && cd /tmp/getmsal
dotnet new console -n getmsal --force
cd getmsal
dotnet add package Microsoft.Identity.Client --version 4.74.1
dotnet build -o out

# Copia la DLL nella directory del modulo SPO
cp out/Microsoft.Identity.Client.dll "$SPOMOD/"
```

#### B4. Connetti e registra

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

> ℹ️ `Connect-SPOService` richiede `Sites.FullControl.All` (Application) sull'app (aggiunto al §4). Dopo aver completato questo passaggio, puoi **rimuovere** quel permesso da Azure se vuoi ridurre la superficie di attacco.

---

## 6. Configurare l'MCP nel gateway

### Comando di avvio (url_command_hidden)

```
npx -y mcp-ms-loop
```

L'opzione `-y` fa sì che npx scarichi sempre l'ultima versione pubblicata senza chiedere conferma.

### Parametri personalizzati (JSON da inserire nella sezione "Parametri" del gateway)

```json
{
  "TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "CLIENT_SECRET": "il-valore-del-secret-NON-l-id"
}
```

> ⚠️ `CLIENT_SECRET` è il **Value** del secret (la stringa lunga come `abc123~XYZ...`), non l'ID/UUID che compare nell'elenco secrets su Azure.

Questi valori vengono iniettati come variabili d'ambiente nel processo Node.js. Non compaiono mai nei messaggi scambiati con l'LLM.

---

## 7. Tool disponibili

### `getLoopByShareUrl`

Recupera il contenuto completo di un file Loop a partire dal link di condivisione.

**Parametri:**

| Parametro | Tipo | Obbligatorio | Descrizione |
|---|---|---|---|
| `shareUrl` | string | Sì | Link completo del file Loop (quello che apri nel browser) |
| `includeText` | boolean | No (default `true`) | Se `true`, restituisce anche il testo plain (HTML strippato) |

**Risposta:**

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

Variante diretta: usa `driveId` e `itemId` se già noti (più veloce, salta la risoluzione del link).

**Parametri:**

| Parametro | Tipo | Obbligatorio | Descrizione |
|---|---|---|---|
| `driveId` | string | Sì | ID del drive contenente il file |
| `itemId` | string | Sì | ID del file |
| `includeText` | boolean | No (default `true`) | Include versione testuale |

---

### `listLoopContainers`

Scopre tutti i workspace Microsoft Loop accessibili all'app. Usa la Microsoft 365 Search API per trovare tutti i file `.loop` e raggrupparli per drive/workspace.

**Nessun parametro richiesto.**

**Risposta:**

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

Elenca tutti i file `.loop` e `.fluid` dentro un drive. La scansione è **ricorsiva** (i file Loop nell'app Loop sono salvati dentro la sottocartella `LoopAppData/`).

**Parametri:**

| Parametro | Tipo | Obbligatorio | Descrizione |
|---|---|---|---|
| `driveId` | string | Sì | ID del drive da scansionare (ottenuto da `listLoopContainers`) |
| `path` | string | No (default `root`) | Percorso di partenza dentro il drive |

**Risposta:**

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

## 8. Permessi minimi necessari

Una volta completata la registrazione SPE (§5), la configurazione **finale** dell'app deve avere solo questi permessi:

| API | Tipo | Permesso | Motivo |
|---|---|---|---|
| Microsoft Graph | Application | `Files.Read.All` | Leggere i file Loop |
| Microsoft Graph | Application | `FileStorageContainer.Selected` | Accedere ai container SPE (workspace Loop) |

Tutto il resto (delegated, `Sites.Read.All`, `User.Read`, `offline_access`, `Sites.FullControl.All`) può essere rimosso.

---

## 9. Troubleshooting

### `403 accessDenied` su file nell'app Loop

Il passaggio §5 (`Set-SPOApplicationPermission`) non è stato eseguito o non ha avuto effetto. Verificare di aver usato il `CLIENT_ID` corretto e di essersi connessi come Global Admin. La propagazione può richiedere qualche minuto.

### `400 failed to parse filter parameter` su `listLoopContainers`

Versione vecchia del pacchetto (≤ 1.0.3). Aggiorna all'ultima versione: il tool ora usa la Search API invece dell'API SPE container, che non è accessibile con i permessi minimi.

### `Connect-SPOService: Object reference not set to an instance of an object`

Stai usando `Connect-SPOService -Interactive` su Linux. Non funziona. Usa l'autenticazione via certificato (§5 Opzione B).

### `Could not load file or assembly 'Microsoft.Identity.Client'`

La DLL MSAL manca nell'ambiente Linux. Segui il passaggio B3 della §5.

### `Connect-SPOService: (401) Unauthorized` (con certificato)

L'app non ha il permesso `SharePoint > Sites.FullControl.All` (Application) con admin consent. Aggiungilo temporaneamente, esegui `Set-SPOApplicationPermission`, poi rimuovilo.

### `listLoopContainers` restituisce array vuoto `[]`

La Search API non ha trovato file Loop accessibili. Cause possibili:
1. Il passaggio §5 non è stato fatto → file SPE non accessibili
2. Non ci sono file `.loop` nel tenant
3. Il tenant usa una region diversa da EMEA/NAM/APAC — contatta supporto Microsoft

### Il `CLIENT_SECRET` è scaduto

Azure mostra la data di scadenza nella sezione **Certificates & secrets**. Se scaduto, genera un nuovo secret, aggiorna il JSON dei parametri nel gateway.

---

## Riferimenti

- [Microsoft Graph — Drive Items](https://learn.microsoft.com/en-us/graph/api/driveitem-get)
- [SharePoint Embedded — Guida tecnica Loop export](https://wals.pro/blogs/news/microsoft-loop-to-html-export-technical-guide)
- [Set-SPOApplicationPermission (Microsoft Docs)](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spoapplicationpermission)
- [Microsoft 365 Search API — App-only](https://learn.microsoft.com/en-us/graph/search-concept-searchall)
