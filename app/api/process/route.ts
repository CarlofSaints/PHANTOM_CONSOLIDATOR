import { NextResponse, after } from 'next/server';
import { buildStoreReport, sanitizeFilename } from '@/lib/report-builder';
import { uploadReport } from '@/lib/graph-iram';
import { sendEmail } from '@/lib/graph-oj';
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

export const maxDuration = 60;

interface ProcessRequest {
  rows: RawRow[];
  controlMap: ControlMap;
  reportDate: string;
  mostRecentDateCol: string;
  includeNegative: boolean;
  recipientMode: 'l1' | 'l2' | 'both';
}

function isPhantom(row: RawRow, includeNegative: boolean): boolean {
  const val = row.Phantom_Indicator.trim().toUpperCase();
  if (val === 'TRUE') return true;
  if (includeNegative && val === 'NEGATIVE') return true;
  return false;
}

export async function POST(req: Request) {
  const { rows, controlMap, reportDate, mostRecentDateCol, includeNegative, recipientMode } =
    await req.json() as ProcessRequest;

  // Return an early acknowledgement and do heavy work in after()
  // But for simplicity + Vercel timeout handling we keep a streaming approach —
  // process is synchronous, after() handles any fire-and-forget cleanup.

  const summary: ProcessSummary = {
    stores: 0,
    reps: 0,
    emailsSent: 0,
    errors: [],
    storeResults: [],
  };

  try {
    // 1. Filter rows to phantom only
    const phantomRows = rows.filter((r) => isPhantom(r, includeNegative));

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

    // 3. Process each store: build XLSX, upload to iRAM
    const storeResults: StoreResult[] = [];
    const storeBuffers = new Map<string, Buffer>(); // storeName → xlsx buffer

    for (const [storeName, storeRows] of storeMap.entries()) {
      const firstRow = storeRows[0];
      const l2Name = firstRow.Personnel_Level_2 || 'Unknown Rep';
      const repInfo = controlMap[l2Name];
      const l1Name = repInfo?.l1Name || 'Unknown Manager';

      const safeStore = sanitizeFilename(storeName);
      const safeL2 = sanitizeFilename(l2Name);
      const fileName = `${safeStore}_${safeL2}_${reportDate}.xlsx`;

      try {
        const buffer = buildStoreReport(storeRows, mostRecentDateCol);
        storeBuffers.set(storeName, buffer);

        const { webUrl } = await uploadReport(buffer, l1Name, reportDate, fileName);

        storeResults.push({
          storeName,
          l2Name,
          l1Name,
          rowCount: storeRows.length,
          webUrl,
          fileName,
        });
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        summary.errors.push(`Upload failed for ${storeName}: ${msg}`);
        storeResults.push({
          storeName,
          l2Name,
          l1Name,
          rowCount: storeRows.length,
          webUrl: '',
          fileName,
          error: msg,
        });
      }
    }

    summary.storeResults = storeResults;

    // 4. Group results by L2 rep
    const byL2 = new Map<string, StoreResult[]>();
    for (const sr of storeResults) {
      if (!byL2.has(sr.l2Name)) byL2.set(sr.l2Name, []);
      byL2.get(sr.l2Name)!.push(sr);
    }

    // Group results by L1 (for L1 emails)
    const byL1 = new Map<string, { repInfo: { l1Email: string; l2Name: string }; stores: StoreResult[] }[]>();

    // 5. Send emails via after() to avoid Vercel timeout on slow SMTP
    after(async () => {
      const emailErrors: string[] = [];

      // ── Level 2 emails: ONE PER STORE ─────────────────────────────────────
      if (recipientMode === 'l2' || recipientMode === 'both') {
        for (const [storeName, storeRows] of storeMap.entries()) {
          const firstRow = storeRows[0];
          const l2Name = firstRow.Personnel_Level_2 || '';
          const repInfo = controlMap[l2Name];

          if (!repInfo?.l2Email) {
            emailErrors.push(`No L2 email for rep "${l2Name}" (store: ${storeName})`);
            continue;
          }

          const storeBuffer = storeBuffers.get(storeName);
          const sr = storeResults.find((r) => r.storeName === storeName);

          try {
            await sendEmail({
              to: repInfo.l2Email,
              subject: `Phantom Stock Report – ${storeName} – ${reportDate}`,
              htmlBody: buildL2StoreEmail(
                storeName,
                l2Name,
                storeRows,
                reportDate,
                mostRecentDateCol
              ),
              attachments: storeBuffer
                ? [
                    {
                      name: sr?.fileName ?? `${sanitizeFilename(storeName)}_${reportDate}.xlsx`,
                      contentBytes: storeBuffer.toString('base64'),
                    },
                  ]
                : [],
            });
          } catch (e) {
            emailErrors.push(
              `L2 email failed for ${storeName} to ${repInfo.l2Email}: ${e instanceof Error ? e.message : String(e)}`
            );
          }
        }
      }

      // ── Level 1 emails: ONE PER L2 (covering that L2's stores) ───────────
      if (recipientMode === 'l1' || recipientMode === 'both') {
        for (const [l2Name, l2Stores] of byL2.entries()) {
          const repInfo = controlMap[l2Name];
          if (!repInfo?.l1Email) {
            emailErrors.push(`No L1 email for L1 of rep "${l2Name}"`);
            continue;
          }

          // Build attachments for all this L2's stores
          const attachments = l2Stores
            .map((sr) => {
              const buf = storeBuffers.get(sr.storeName);
              return buf
                ? { name: sr.fileName, contentBytes: buf.toString('base64') }
                : null;
            })
            .filter((a): a is { name: string; contentBytes: string } => a !== null);

          const l1Name = repInfo.l1Name || 'Manager';

          try {
            await sendEmail({
              to: repInfo.l1Email,
              subject: `Phantom Reports for ${l2Name} – ${reportDate}`,
              htmlBody: buildL1RepEmail(
                l1Name,
                {
                  l2Name,
                  stores: l2Stores.map((s) => ({
                    storeName: s.storeName,
                    rowCount: s.rowCount,
                  })),
                },
                reportDate
              ),
              attachments,
            });

            // Track for summary email
            if (!byL1.has(repInfo.l1Email)) byL1.set(repInfo.l1Email, []);
            byL1.get(repInfo.l1Email)!.push({
              repInfo: { l1Email: repInfo.l1Email, l2Name },
              stores: l2Stores,
            });
          } catch (e) {
            emailErrors.push(
              `L1 email failed for ${l2Name} to ${repInfo.l1Email}: ${e instanceof Error ? e.message : String(e)}`
            );
          }
        }

        // ── Level 1 SUMMARY email: one per unique L1 ──────────────────────
        for (const [l1Email, l2Groups] of byL1.entries()) {
          const l1Name =
            controlMap[l2Groups[0].repInfo.l2Name]?.l1Name || 'Manager';

          const summaryRows = l2Groups.map((g) => ({
            l2Name: g.repInfo.l2Name,
            storeCount: g.stores.length,
            reportsSent: g.stores.filter((s) => !s.error).length,
          }));

          try {
            await sendEmail({
              to: l1Email,
              subject: `Phantom Report Summary – ${reportDate}`,
              htmlBody: buildL1SummaryEmail(l1Name, summaryRows, reportDate),
            });
          } catch (e) {
            emailErrors.push(
              `L1 summary email failed to ${l1Email}: ${e instanceof Error ? e.message : String(e)}`
            );
          }
        }
      }

      if (emailErrors.length > 0) {
        console.error('[process] Email errors:', emailErrors);
      }
    });

    const uniqueReps = new Set(storeResults.map((r) => r.l2Name)).size;
    summary.reps = uniqueReps;

    return NextResponse.json({ success: true, summary });
  } catch (e) {
    console.error('[process]', e);
    const msg = e instanceof Error ? e.message : 'Processing failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
