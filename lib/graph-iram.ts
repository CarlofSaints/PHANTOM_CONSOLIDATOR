/**
 * iRAM SharePoint Graph client
 * Used for: creating folders and uploading per-store XLSX reports
 */

const TENANT_ID = process.env.IRAM_TENANT_ID!;
const CLIENT_ID = process.env.IRAM_CLIENT_ID!;
const CLIENT_SECRET = process.env.IRAM_CLIENT_SECRET!;
const SP_HOST = process.env.IRAM_SP_HOST ?? 'iramsa.sharepoint.com';
const LIBRARY_NAME = process.env.IRAM_SP_LIBRARY ?? 'Instore';
const BASE_FOLDER = process.env.IRAM_BASE_FOLDER ?? 'PHANTOM REPORTS (OJ)';

// ── Auth ─────────────────────────────────────────────────────────────────────

async function getToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) {
    throw new Error(`iRAM auth failed: ${data.error_description ?? JSON.stringify(data)}`);
  }
  return data.access_token as string;
}

function encodePath(path: string): string {
  return path.split('/').map((seg) => encodeURIComponent(seg)).join('/');
}

type DriveContext = { token: string; driveId: string };

async function getDriveContext(): Promise<DriveContext> {
  const token = await getToken();

  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!siteRes.ok) throw new Error(`iRAM: could not get site: ${await siteRes.text()}`);
  const site = await siteRes.json();

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drives = await drivesRes.json();
  const drive = drives.value?.find((d: { name: string }) => d.name === LIBRARY_NAME);
  if (!drive) {
    const names = drives.value?.map((d: { name: string }) => d.name).join(', ');
    throw new Error(`iRAM: library "${LIBRARY_NAME}" not found. Available: ${names}`);
  }
  return { token, driveId: drive.id as string };
}

// ── Ensure folder path exists ────────────────────────────────────────────────

async function ensureFolderExists(
  token: string,
  driveId: string,
  folderPath: string // e.g. "PHANTOM REPORTS (OJ)/Wayne July/2026-02-28"
): Promise<void> {
  const segments = folderPath.split('/');
  let currentPath = '';

  for (const segment of segments) {
    const parentPath = currentPath
      ? encodePath(currentPath)
      : undefined;

    const url = parentPath
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${parentPath}:/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

    await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: segment,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      }),
    });
    // Ignore errors — 409 conflict means already exists, which is fine

    currentPath = currentPath ? `${currentPath}/${segment}` : segment;
  }
}

// ── Upload file ──────────────────────────────────────────────────────────────

export interface UploadResult {
  webUrl: string;
  fileId: string;
}

export async function uploadReport(
  buffer: Buffer,
  l1Name: string,
  reportDate: string,
  fileName: string
): Promise<UploadResult> {
  const { token, driveId } = await getDriveContext();

  const folderPath = `${BASE_FOLDER}/${l1Name}/${reportDate}`;
  await ensureFolderExists(token, driveId, folderPath);

  const filePath = encodePath(`${folderPath}/${fileName}`);

  const uploadRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: new Uint8Array(buffer),
    }
  );

  if (!uploadRes.ok) {
    throw new Error(`iRAM: upload failed (${uploadRes.status}): ${await uploadRes.text()}`);
  }

  const uploaded = await uploadRes.json();
  return {
    webUrl: uploaded.webUrl as string,
    fileId: uploaded.id as string,
  };
}
