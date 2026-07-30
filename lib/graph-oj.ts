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
const CHANNELS_CONFIG_NAME = 'phantom-channels.json';
const DEFAULT_CHANNELS = ['PnP', 'Builders Warehouse', 'Makro', 'Game', 'Checkers', 'Dis-Chem', 'Clicks'];
const EMAIL_FROM = (process.env.OJ_EMAIL_FROM ?? process.env.EMAIL_FROM ?? '')
  .replace(/\\n|\\r|\n|\r/g, '') // strip literal \n or real newlines (common env var corruption)
  .trim()
  .replace(/^["']|["']$/g, ''); // strip accidental surrounding quotes

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

export type DriveContext = { token: string; driveId: string };

export async function getDriveContext(): Promise<DriveContext> {
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

/**
 * Reads the control file for a specific channel.
 * File naming convention: "{channelName} - User Control File - Phantom Consolidator.xlsx"
 * Pass a shared DriveContext to avoid re-authenticating per channel.
 */
export async function readControlFileBuffer(channelName: string, ctx?: DriveContext): Promise<Buffer> {
  const { token, driveId } = ctx ?? await getDriveContext();
  const fileName = `${channelName} - User Control File - Phantom Consolidator.xlsx`;
  const filePath = encodePath(CONTROL_FILE_FOLDER ? `${CONTROL_FILE_FOLDER}/${fileName}` : fileName);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`OJ: could not read control file for "${channelName}": ${await res.text()}`);
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
}

// ── Channel config (phantom-channels.json on OJ SP) ──────────────────────────

export async function readChannelsConfig(ctx?: DriveContext): Promise<string[]> {
  try {
    const { token, driveId } = ctx ?? await getDriveContext();
    const filePath = encodePath(
      CONTROL_FILE_FOLDER ? `${CONTROL_FILE_FOLDER}/${CHANNELS_CONFIG_NAME}` : CHANNELS_CONFIG_NAME
    );
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return [...DEFAULT_CHANNELS];
    const data = await res.json() as { channels?: string[] };
    return Array.isArray(data.channels) && data.channels.length > 0 ? data.channels : [...DEFAULT_CHANNELS];
  } catch {
    return [...DEFAULT_CHANNELS];
  }
}

export async function writeChannelsConfig(channels: string[], ctx?: DriveContext): Promise<void> {
  const { token, driveId } = ctx ?? await getDriveContext();
  const filePath = encodePath(
    CONTROL_FILE_FOLDER ? `${CONTROL_FILE_FOLDER}/${CHANNELS_CONFIG_NAME}` : CHANNELS_CONFIG_NAME
  );
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ channels }),
    }
  );
  if (!res.ok) throw new Error(`OJ: could not write channels config: ${await res.text()}`);
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
  if (!EMAIL_FROM) {
    throw new Error('OJ: EMAIL_FROM is not configured — set OJ_EMAIL_FROM in Vercel env vars');
  }
  console.log(`[sendEmail] FROM=${EMAIL_FROM} TO=${payload.to} SUBJECT=${payload.subject}`);
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
