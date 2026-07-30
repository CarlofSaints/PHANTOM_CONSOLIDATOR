import { NextResponse, after } from 'next/server';
import { buildStoreReport, sanitizeFilename } from '@/lib/report-builder';
import { uploadReport, getDriveContext } from '@/lib/graph-iram';
import { sendEmail, getDriveContext as getOjDriveContext, readControlFileBuffer } from '@/lib/graph-oj';
import { parseControlBuffer } from '@/lib/parse-control-file';
import {
  buildL2StoreEmail,
  buildL1RepEmail,
  buildL1SummaryEmail,
} from '@/lib/email-builder';
import type {
  RawRow,
  ControlMap,
  ProcessSummary,
  StoreResult,
} from '@/types';

export const maxDuration = 300;
export const runtime = 'nodejs';

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

function isPhantom(row: RawRow, includeNegative: boolean): boolean {
  const val = row.Phantom_Indicator.trim().toUpperCase();
  if (val === 'TRUE') return true;
  if (includeNegative && val === 'NEGATIVE') return true;
  return false;
}

async function fetchControlMap(channels: string[]): Promise<ControlMap> {
  const ctx = await getOjDriveContext();
  const results = await Promise.allSettled(
    channels.map((ch) => readControlFileBuffer(ch, ctx))
  );

  const controlMap: ControlMap = {};
  for (let i = 0; i < results.length; i++) {
    const result = results[i];
    if (result.status === 'rejected') {
      console.warn(`[process] Could not load control file for channel "${channels[i]}":`, result.reason);
    } else {
      parseControlBuffer(result.value, controlMap);
    }
  }
  return controlMap;
}

interface ProcessRequest {
  rowHeaders: string[];
  rowData: string[][];
  reportDate: string;
  mostRecentDateCol: string;
  includeNegative: boolean;
  recipientMode: 'l1' | 'l2' | 'both';
  actionMode: 'both' | 'sharepoint' | 'email';
  selectedProvinces?: Record<string, string[]>; // clientName → allowed provinces
  selectedChannels: string[];                   // channel names for control file + SP folder
}

export async function POST(req: Request) {
  const {
    rowHeaders, rowData, reportDate, mostRecentDateCol,
    includeNegative, recipientMode, actionMode = 'both', selectedProvinces,
    selectedChannels = [],
  } = await req.json() as ProcessRequest;

  // Reconstruct rows from compact array format
  const rows: RawRow[] = rowData.map((values) => {
    const row: Record<string, string> = {};
    rowHeaders.forEach((h, i) => { row[h] = values[i] ?? ''; });
    return row as RawRow;
  });

  // Fetch controlMap from SharePoint (one file per selected channel, merged)
  const controlMap = await fetchControlMap(selectedChannels);

  const summary: ProcessSummary = {
    stores: 0,
    reps: 0,
    emailsSent: 0,
    errors: [],
    storeResults: [],
  };

  try {
    // 1. Filter to phantom rows + province filter
    const phantomRows = rows.filter((r) => {
      if (!isPhantom(r, includeNegative)) return false;
      if (selectedProvinces) {
        const allowed = selectedProvinces[r.CLIENT];
        if (allowed && !allowed.includes(r.Province)) return false;
      }
      return true;
    });

    if (phantomRows.length === 0) {
      return NextResponse.json({
        success: true,
        summary: { ...summary, errors: ['No phantom rows found after filtering'] },
      });
    }

    // 2. Group by Store_Name
    const storeMap = new Map<string, RawRow[]>();
    for (const row of phantomRows) {
      const key = row.Store_Name || 'UNKNOWN STORE';
      if (!storeMap.has(key)) storeMap.set(key, []);
      storeMap.get(key)!.push(row);
    }

    summary.stores = storeMap.size;

    // Helper: resolve which channel folder a store's rows belong to.
    // Match the row's Channel value (case-insensitive) against selectedChannels;
    // fall back to the first selected channel if no match found.
    const resolveChannel = (rowChannel: string): string => {
      const norm = rowChannel.trim().toLowerCase();
      const match = selectedChannels.find((c) => c.toLowerCase() === norm);
      return match ?? selectedChannels[0] ?? 'UNKNOWN';
    };

    // 3. Build XLSX buffers synchronously (no network calls)
    const storeBuffers = new Map<string, Buffer>();
    const storeInfos = Array.from(storeMap.entries()).map(([storeName, storeRows]) => {
      const firstRow = storeRows[0];
      const l2Name = firstRow.Personnel_Level_2 || 'Unknown Rep';
      const repInfo = controlMap[l2Name.toLowerCase()];
      const l1Name = repInfo?.l1Name || 'Unknown Manager';
      const channel = resolveChannel(firstRow.Channel ?? '');
      const safeStore = sanitizeFilename(storeName);
      const safeL2 = sanitizeFilename(l2Name);
      const fileName = `${safeStore}_${safeL2}_${reportDate}.xlsx`;
      const buffer = buildStoreReport(storeRows, mostRecentDateCol);
      storeBuffers.set(storeName, buffer);
      return { storeName, storeRows, l2Name, l1Name, channel, fileName, buffer };
    });

    const storeResults: StoreResult[] = storeInfos.map(({ storeName, storeRows, l2Name, l1Name, fileName }) => ({
      storeName, l2Name, l1Name, rowCount: storeRows.length, webUrl: '', fileName,
    }));

    summary.storeResults = storeResults;

    // 4. Group by L2
    const byL2 = new Map<string, StoreResult[]>();
    for (const sr of storeResults) {
      if (!byL2.has(sr.l2Name)) byL2.set(sr.l2Name, []);
      byL2.get(sr.l2Name)!.push(sr);
    }

    summary.reps = new Set(storeResults.map((r) => r.l2Name)).size;

    // ── SP uploads — fire-and-forget in after() ───────────────────────────────
    after(async () => {
      if (actionMode === 'email') return;

      let ctx;
      try {
        ctx = await getDriveContext();
      } catch (e) {
        console.error('[process/sp] Failed to get drive context:', e);
        return;
      }

      const BATCH_SIZE = 3;
      for (let i = 0; i < storeInfos.length; i += BATCH_SIZE) {
        const batch = storeInfos.slice(i, i + BATCH_SIZE);
        await Promise.all(
          batch.map(async ({ storeName, l1Name, channel, fileName, buffer }) => {
            try {
              await uploadReport(buffer, l1Name, reportDate, fileName, channel, ctx);
            } catch (e) {
              console.error(`[process/sp] Upload failed for ${storeName}:`, e);
            }
          })
        );
        if (i + BATCH_SIZE < storeInfos.length) await sleep(300);
      }
    });

    // ── Emails — synchronous so results are visible in the response ───────────
    if (actionMode !== 'sharepoint') {
      const byL1 = new Map<string, { repInfo: { l1Email: string; l2Name: string }; stores: StoreResult[] }[]>();

      // L2 emails: one per store
      if (recipientMode === 'l2' || recipientMode === 'both') {
        for (const [storeName, storeRows] of storeMap.entries()) {
          const l2Name = storeRows[0].Personnel_Level_2 || '';
          const repInfo = controlMap[l2Name.toLowerCase()];

          if (!repInfo?.l2Email) {
            summary.errors.push(`No L2 email for rep "${l2Name}" (store: ${storeName})`);
            continue;
          }

          const storeBuffer = storeBuffers.get(storeName);
          const sr = storeResults.find((r) => r.storeName === storeName);

          try {
            await sendEmail({
              to: repInfo.l2Email,
              subject: `Phantom Stock Report – ${storeName} – ${reportDate}`,
              htmlBody: buildL2StoreEmail(storeName, l2Name, storeRows, reportDate, mostRecentDateCol),
              attachments: storeBuffer
                ? [{ name: sr?.fileName ?? `${sanitizeFilename(storeName)}_${reportDate}.xlsx`, contentBytes: storeBuffer.toString('base64') }]
                : [],
            });
            summary.emailsSent++;
          } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            summary.errors.push(`L2 email failed for ${storeName} → ${repInfo.l2Email}: ${msg}`);
          }
        }
      }

      // L1 emails: one per L2 rep
      if (recipientMode === 'l1' || recipientMode === 'both') {
        for (const [l2Name, l2Stores] of byL2.entries()) {
          const repInfo = controlMap[l2Name.toLowerCase()];
          if (!repInfo?.l1Email) {
            summary.errors.push(`No L1 email for manager of "${l2Name}"`);
            continue;
          }

          const attachments = l2Stores
            .map((sr) => {
              const buf = storeBuffers.get(sr.storeName);
              return buf ? { name: sr.fileName, contentBytes: buf.toString('base64') } : null;
            })
            .filter((a): a is { name: string; contentBytes: string } => a !== null);

          try {
            await sendEmail({
              to: repInfo.l1Email,
              subject: `Phantom Reports for ${l2Name} – ${reportDate}`,
              htmlBody: buildL1RepEmail(
                repInfo.l1Name || 'Manager',
                { l2Name, stores: l2Stores.map((s) => ({ storeName: s.storeName, rowCount: s.rowCount })) },
                reportDate
              ),
              attachments,
            });
            summary.emailsSent++;

            if (!byL1.has(repInfo.l1Email)) byL1.set(repInfo.l1Email, []);
            byL1.get(repInfo.l1Email)!.push({ repInfo: { l1Email: repInfo.l1Email, l2Name }, stores: l2Stores });
          } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            summary.errors.push(`L1 email failed for "${l2Name}" → ${repInfo.l1Email}: ${msg}`);
          }
        }

        // L1 summary email: one per unique L1
        for (const [l1Email, l2Groups] of byL1.entries()) {
          const l1Name = controlMap[l2Groups[0].repInfo.l2Name.toLowerCase()]?.l1Name || 'Manager';
          const summaryRows = l2Groups.map((g) => ({
            l2Name: g.repInfo.l2Name,
            storeCount: g.stores.length,
            reportsSent: g.stores.length,
          }));

          try {
            await sendEmail({
              to: l1Email,
              subject: `Phantom Report Summary – ${reportDate}`,
              htmlBody: buildL1SummaryEmail(l1Name, summaryRows, reportDate),
            });
            summary.emailsSent++;
          } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            summary.errors.push(`L1 summary email failed → ${l1Email}: ${msg}`);
          }
        }
      }
    }

    return NextResponse.json({ success: true, summary });
  } catch (e) {
    console.error('[process]', e);
    const msg = e instanceof Error ? e.message : 'Processing failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
