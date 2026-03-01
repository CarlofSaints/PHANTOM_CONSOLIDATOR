/**
 * OuterJoin SharePoint Graph client
 * Used for: reading the email control file + sending mail via Graph
 */

const TENANT_ID = process.env.OJ_TENANT_ID!;
const CLIENT_ID = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST = process.env.OJ_SP_HOST ?? 'exceler8xl.sharepoint.com';
const LIBRARY_NAME = process.env.OJ_SP_LIBRARY ?? 'Clients';
const CONTROL_FILE_FOLDER = process.env.OJ_CONTROL_FILE_FOLDER ?? '';
const CONTROL_FILE_NAME = process.env.OJ_CONTROL_FILE_NAME ?? 'email-control.xlsx';
const EMAIL_FROM = process.env.EMAIL_FROM ?? '';

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
    throw new Error(`OJ auth failed: ${data.error_description ?? JSON.stringify(data)}`);
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
  if (!siteRes.ok) throw new Error(`OJ: could not get site: ${await siteRes.text()}`);
  const site = await siteRes.json();

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drives = await drivesRes.json();
  const drive = drives.value?.find((d: { name: string }) => d.name === LIBRARY_NAME);
  if (!drive) {
    const names = drives.value?.map((d: { name: string }) => d.name).join(', ');
    throw new Error(`OJ: library "${LIBRARY_NAME}" not found. Available: ${names}`);
  }
  return { token, driveId: drive.id as string };
}

// ── Read control file ────────────────────────────────────────────────────────

export async function readControlFileBuffer(): Promise<Buffer> {
  const { token, driveId } = await getDriveContext();
  const filePath = encodePath(`${CONTROL_FILE_FOLDER}/${CONTROL_FILE_NAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`OJ: could not read control file: ${await res.text()}`);
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
}

// ── Send email via Graph sendMail ────────────────────────────────────────────

export interface Attachment {
  name: string;
  contentBytes: string; // base64
}

export interface EmailPayload {
  to: string;
  subject: string;
  htmlBody: string;
  attachments?: Attachment[];
}

export async function sendEmail(payload: EmailPayload): Promise<void> {
  const token = await getToken();

  const message: Record<string, unknown> = {
    subject: payload.subject,
    body: {
      contentType: 'HTML',
      content: payload.htmlBody,
    },
    toRecipients: [
      { emailAddress: { address: payload.to } },
    ],
  };

  if (payload.attachments && payload.attachments.length > 0) {
    message.attachments = payload.attachments.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentBytes: a.contentBytes,
    }));
  }

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(EMAIL_FROM)}/sendMail`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message, saveToSentItems: false }),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OJ: sendMail failed (${res.status}): ${text}`);
  }
}
