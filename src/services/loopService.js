import { getAccessToken } from "./authService.js";
import axios from 'axios';

// Limite caratteri restituiti nel testo per evitare risposte oversize
const MAX_TOKENS = 200000;
const CHARS_PER_TOKEN = 4;
const SAFETY_FACTOR = 0.5;
const ALLOWED_CHARS = Math.floor(MAX_TOKENS * CHARS_PER_TOKEN * SAFETY_FACTOR);

/**
 * Encode a sharing URL into the base64url format required by the
 * Graph API /shares/{encoded}/driveItem endpoint.
 *
 * @param {string} url - The full sharing URL.
 * @returns {string} - The encoded share token (e.g. "u!aHR0cHM6...").
 */
function encodeShareUrl(url) {
  const b64 = Buffer.from(url, 'utf-8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
  return 'u!' + b64;
}

/**
 * Detect whether a sharing URL points to a SharePoint Embedded (SPE)
 * container (Loop app workspace). These URLs contain "contentstorage/CSP_"
 * and cannot be resolved via the /shares/{token}/driveItem Graph endpoint
 * (which returns 422). Instead, driveId and itemId must be extracted from
 * the "nav" query parameter embedded in the URL itself.
 *
 * @param {string} shareUrl
 * @returns {boolean}
 */
function isSpeUrl(shareUrl) {
  return shareUrl.includes('contentstorage/CSP_');
}

/**
 * Parse the "nav" query parameter from a SharePoint Embedded sharing URL
 * to extract the driveId ("d") and itemId ("f").
 *
 * The nav value is base64-encoded URL-encoded query string, e.g.:
 *   base64("s=%2Fcontentstorage%2FCSP_...&d=b!Fx9...&f=01GO3C...")
 *
 * @param {string} navParam - The raw value of the "nav" query parameter.
 * @returns {{ driveId: string|null, itemId: string|null }}
 */
function parseSpeNavParam(navParam) {
  try {
    const decoded = Buffer.from(navParam, 'base64').toString('utf-8');
    const params = new URLSearchParams(decoded);
    return {
      driveId: params.get('d') || null,
      itemId: params.get('f') || null,
    };
  } catch {
    return { driveId: null, itemId: null };
  }
}

/**
 * Strip HTML tags and collapse whitespace, returning a plain-text
 * version of the content.
 *
 * @param {string} html
 * @returns {string}
 */
function htmlToText(html) {
  if (!html) return '';
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Truncate a string to ALLOWED_CHARS, flagging if truncation occurred.
 *
 * @param {string} str
 * @returns {{ value: string, truncated: boolean }}
 */
function truncate(str) {
  if (!str) return { value: '', truncated: false };
  if (str.length <= ALLOWED_CHARS) {
    return { value: str, truncated: false };
  }
  return {
    value:
      str.substring(0, ALLOWED_CHARS) +
      "\n[... ATTENZIONE: contenuto troncato per superamento limite token ...]",
    truncated: true,
  };
}

/**
 * Fetch the HTML content of a Loop file given its driveId and itemId.
 * Internal helper used by the public functions below.
 *
 * @param {string} accessToken
 * @param {string} driveId
 * @param {string} itemId
 * @returns {Promise<string>} - The HTML content as a string.
 */
async function fetchLoopHtml(accessToken, driveId, itemId) {
  // First attempt: format=html conversion (works for standard Office docs)
  const urlHtml = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content?format=html`;
  console.error(`[fetchLoopHtml] GET ${urlHtml}`);
  try {
    const response = await axios.get(urlHtml, {
      headers: { Authorization: `Bearer ${accessToken}` },
      responseType: 'text',
      maxRedirects: 10,
      transformResponse: [(data) => data],
    });
    return typeof response.data === 'string'
      ? response.data
      : Buffer.from(response.data).toString('utf-8');
  } catch (err) {
    const status = err.response?.status;
    const detail = JSON.stringify(err.response?.data ?? err.message);
    console.error(`[fetchLoopHtml] format=html failed (${status}): ${detail}`);

    // Fallback: download raw content for .loop/.fluid files (Fluid Framework)
    // These are binary packages; we return them base64-encoded so the caller
    // at least gets something actionable instead of a hard error.
    if (status === 400 || status === 415) {
      console.error(`[fetchLoopHtml] Trying raw content download as fallback`);
      const urlRaw = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
      const rawResp = await axios.get(urlRaw, {
        headers: { Authorization: `Bearer ${accessToken}` },
        responseType: 'arraybuffer',
        maxRedirects: 10,
      });
      const b64 = Buffer.from(rawResp.data).toString('base64');
      return `<!-- Loop file raw content (base64, Fluid Framework binary). format=html not supported for this file type. -->\n<pre data-encoding="base64">${b64.substring(0, 2000)}${b64.length > 2000 ? '...[truncated]' : ''}</pre>`;
    }

    throw err;
  }
}

/**
 * Fetch driveItem metadata.
 *
 * @param {string} accessToken
 * @param {string} driveId
 * @param {string} itemId
 * @returns {Promise<Object>}
 */
async function fetchDriveItemMetadata(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=id,name,size,webUrl,lastModifiedDateTime,file,parentReference`;
  const response = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return response.data;
}

/**
 * Retrieve the HTML content of a Loop file starting from its sharing URL.
 *
 * @param {string} tenantId
 * @param {string} clientId
 * @param {string} clientSecret
 * @param {string} shareUrl - The full sharing URL of the Loop file.
 * @param {boolean} includeText - If true, also returns a plain-text version.
 * @returns {Promise<Object>}
 */
export async function getLoopByShareUrl(
  tenantId,
  clientId,
  clientSecret,
  shareUrl,
  includeText = true
) {
  try {
    if (!shareUrl) throw new Error("shareUrl is required");
    const accessToken = await getAccessToken(tenantId, clientId, clientSecret);

    let driveId, itemId, itemMeta;

    if (isSpeUrl(shareUrl)) {
      // SharePoint Embedded (Loop app workspace) URL:
      // Graph's /shares/{token}/driveItem returns 422 for these URLs.
      // Extract driveId and itemId directly from the "nav" parameter.
      const urlObj = new URL(shareUrl);
      const navParam = urlObj.searchParams.get('nav');
      if (!navParam) {
        throw new Error(
          'SharePoint Embedded URL detected but "nav" parameter is missing. ' +
          'Use listLoopContainers + listLoopFilesInDrive to browse files, ' +
          'then call getLoopByItemId with the driveId and itemId.'
        );
      }
      const parsed = parseSpeNavParam(navParam);
      console.error(`[getLoopByShareUrl] SPE nav parsed → driveId=${parsed.driveId} itemId=${parsed.itemId}`);
      if (!parsed.driveId || !parsed.itemId) {
        throw new Error(
          'Unable to extract driveId/itemId from the SharePoint Embedded nav parameter. ' +
          'Use listLoopContainers + listLoopFilesInDrive instead.'
        );
      }
      driveId = parsed.driveId;
      itemId = parsed.itemId;
      // Fetch metadata separately since we skipped the /shares resolution
      itemMeta = await fetchDriveItemMetadata(accessToken, driveId, itemId);
    } else {
      // Standard OneDrive / SharePoint sharing URL:
      // Resolve via /shares/{token}/driveItem
      const encoded = encodeShareUrl(shareUrl);
      const driveItemUrl = `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem?$select=id,name,size,webUrl,lastModifiedDateTime,file,parentReference`;
      const driveItemResp = await axios.get(driveItemUrl, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      itemMeta = driveItemResp.data;
      driveId = itemMeta.parentReference?.driveId;
      itemId = itemMeta.id;
      if (!driveId || !itemId) {
        throw new Error("Unable to resolve driveId/itemId from the share URL");
      }
    }

    // Scarica il contenuto HTML
    const html = await fetchLoopHtml(accessToken, driveId, itemId);
    const htmlTruncated = truncate(html);

    const result = {
      name: itemMeta.name,
      driveId,
      itemId,
      webUrl: itemMeta.webUrl,
      lastModified: itemMeta.lastModifiedDateTime,
      size: itemMeta.size,
      mimeType: itemMeta.file?.mimeType,
      html: htmlTruncated.value,
      truncated: htmlTruncated.truncated,
    };

    if (includeText) {
      const text = htmlToText(html);
      const textTruncated = truncate(text);
      result.text = textTruncated.value;
      result.textTruncated = textTruncated.truncated;
    }

    return result;
  } catch (error) {
    const detail = JSON.stringify(error.response?.data ?? error.message);
    const status = error.response?.status ?? 'N/A';
    const msg = `getLoopByShareUrl failed (HTTP ${status}): ${detail}`;
    console.error(msg);
    throw new Error(msg);
  }
}

/**
 * Retrieve the HTML content of a Loop file using its driveId and itemId.
 *
 * @param {string} tenantId
 * @param {string} clientId
 * @param {string} clientSecret
 * @param {string} driveId
 * @param {string} itemId
 * @param {boolean} includeText
 * @returns {Promise<Object>}
 */
export async function getLoopByItemId(
  tenantId,
  clientId,
  clientSecret,
  driveId,
  itemId,
  includeText = true
) {
  try {
    if (!driveId || !itemId) {
      throw new Error("driveId and itemId are required");
    }
    const accessToken = await getAccessToken(tenantId, clientId, clientSecret);

    const [metadata, html] = await Promise.all([
      fetchDriveItemMetadata(accessToken, driveId, itemId),
      fetchLoopHtml(accessToken, driveId, itemId),
    ]);

    const htmlTruncated = truncate(html);

    const result = {
      name: metadata.name,
      driveId,
      itemId,
      webUrl: metadata.webUrl,
      lastModified: metadata.lastModifiedDateTime,
      size: metadata.size,
      mimeType: metadata.file?.mimeType,
      html: htmlTruncated.value,
      truncated: htmlTruncated.truncated,
    };

    if (includeText) {
      const text = htmlToText(html);
      const textTruncated = truncate(text);
      result.text = textTruncated.value;
      result.textTruncated = textTruncated.truncated;
    }

    return result;
  } catch (error) {
    const detail = JSON.stringify(error.response?.data ?? error.message);
    const status = error.response?.status ?? 'N/A';
    const msg = `getLoopByItemId failed (HTTP ${status}): ${detail}`;
    console.error(msg);
    throw new Error(msg);
  }
}

/**
 * List all Loop workspaces (drives) accessible to the app by searching for
 * Loop files via the Microsoft 365 Search API and grouping results by drive.
 *
 * Note: the SPE container-listing endpoint (/storage/fileStorage/containers)
 * requires admin-level permissions that are incompatible with the minimal
 * FileStorageContainer.Selected app permission used by this MCP. The Search
 * API approach discovers all Loop workspaces without elevated permissions.
 *
 * @param {string} tenantId
 * @param {string} clientId
 * @param {string} clientSecret
 * @returns {Promise<Array<{driveId, name, webUrl, loopFileCount}>>}
 */
export async function listLoopContainers(tenantId, clientId, clientSecret) {
  try {
    const accessToken = await getAccessToken(tenantId, clientId, clientSecret);

    // The Search API with app-only auth requires an explicit region. Try the
    // most common regions in order; the tenant error message tells us which
    // one is valid ("Only valid regions are X").
    const REGIONS = ['EMEA', 'NAM', 'APAC', 'GCC', 'GCCHIGH'];
    let hits = [];

    for (const region of REGIONS) {
      try {
        const response = await axios.post(
          'https://graph.microsoft.com/v1.0/search/query',
          {
            requests: [{
              entityTypes: ['driveItem'],
              query: { queryString: 'filetype:loop OR filetype:fluid' },
              from: 0,
              size: 500,
              region,
            }],
          },
          { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
        );
        hits = response.data?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
        console.error(`[listLoopContainers] Search region=${region} → ${hits.length} files`);
        break;
      } catch (err) {
        const msg = err.response?.data?.error?.message ?? '';
        if (msg.includes('Only valid regions are') || msg.includes('Region is required')) {
          // Extract the valid region(s) from the error message if possible
          const match = msg.match(/Only valid regions are\s+(\S+)/);
          if (match) {
            const validRegion = match[1].replace(/[^A-Z]/g, '');
            console.error(`[listLoopContainers] Tenant requires region=${validRegion}, retrying`);
            try {
              const r2 = await axios.post(
                'https://graph.microsoft.com/v1.0/search/query',
                { requests: [{ entityTypes:['driveItem'], query:{ queryString:'filetype:loop OR filetype:fluid' }, from:0, size:500, region: validRegion }] },
                { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
              );
              hits = r2.data?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
              break;
            } catch (innerErr) {
              throw innerErr;
            }
          }
          continue;
        }
        throw err;
      }
    }

    // Group hits by driveId → one entry per workspace
    const driveMap = new Map();
    for (const hit of hits) {
      const driveId = hit.resource?.parentReference?.driveId;
      if (!driveId) continue;
      if (!driveMap.has(driveId)) {
        driveMap.set(driveId, { driveId, name: null, webUrl: null, loopFileCount: 0 });
      }
      driveMap.get(driveId).loopFileCount++;
    }

    // Enrich each drive with its display name and URL
    await Promise.all(
      Array.from(driveMap.values()).map(async (ws) => {
        try {
          const r = await axios.get(
            `https://graph.microsoft.com/v1.0/drives/${ws.driveId}`,
            { headers: { Authorization: `Bearer ${accessToken}` } }
          );
          ws.name = r.data.name ?? ws.driveId;
          ws.webUrl = r.data.webUrl ?? null;
        } catch {
          ws.name = ws.driveId;
        }
      })
    );

    return Array.from(driveMap.values());
  } catch (error) {
    const detail = JSON.stringify(error.response?.data ?? error.message);
    const status = error.response?.status ?? 'N/A';
    const msg = `listLoopContainers failed (HTTP ${status}): ${detail}`;
    console.error(msg);
    throw new Error(msg);
  }
}

/**
 * Recursively list all Loop files (.loop and .fluid) inside a drive.
 * Loop files in SPE workspaces are stored inside a "LoopAppData" subfolder,
 * so a recursive scan is needed to find them all.
 *
 * @param {string} tenantId
 * @param {string} clientId
 * @param {string} clientSecret
 * @param {string} driveId
 * @param {string} [path] - Defaults to 'root'. Scans recursively from here.
 * @returns {Promise<Array>}
 */
export async function listLoopFilesInDrive(
  tenantId,
  clientId,
  clientSecret,
  driveId,
  path
) {
  try {
    const accessToken = await getAccessToken(tenantId, clientId, clientSecret);

    const loopFiles = [];

    async function scanFolder(folderUrl) {
      let nextLink = folderUrl;
      while (nextLink) {
        const response = await axios.get(nextLink, {
          headers: { Authorization: `Bearer ${accessToken}` },
        });
        const items = response.data.value || [];

        for (const item of items) {
          if (item.folder) {
            // Recurse into subfolders
            const childUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/children`;
            await scanFolder(childUrl);
          } else if (item.file) {
            const name = (item.name || '').toLowerCase();
            if (name.endsWith('.loop') || name.endsWith('.fluid')) {
              loopFiles.push({
                id: item.id,
                name: item.name,
                driveId,
                size: item.size,
                webUrl: item.webUrl,
                lastModified: item.lastModifiedDateTime,
                mimeType: item.file?.mimeType,
              });
            }
          }
        }

        nextLink = response.data['@odata.nextLink'] || null;
      }
    }

    let rootUrl;
    if (!path || path === "/" || path.toLowerCase() === "root") {
      rootUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
    } else {
      const cleanPath = path.replace(/^\/|\/$/g, '');
      rootUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${cleanPath}:/children`;
    }

    await scanFolder(rootUrl);
    return loopFiles;
  } catch (error) {
    const detail = JSON.stringify(error.response?.data ?? error.message);
    const status = error.response?.status ?? 'N/A';
    const msg = `listLoopFilesInDrive failed (HTTP ${status}): ${detail}`;
    console.error(msg);
    throw new Error(msg);
  }
}
