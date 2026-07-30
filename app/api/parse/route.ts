import { NextResponse } from 'next/server';
import { parseExcelBuffer } from '@/lib/excel-parser';

export const maxDuration = 60;

// Fields sent in compact row format (header names sent once, not repeated per row)
const NEEDED_FIELDS = [
  'CLIENT', 'Product_Principle', 'Channel', 'Sub_Channel', 'SiteCode', 'Store_Name',
  'Store_Status', 'Product_Brand', 'Product_Sub_Category', 'Channel_ArticleCode',
  'Client_Product_ID', 'Product_Description', 'Product_Status', 'Range_Indicator',
  'Personnel_Level_1', 'Personnel_Level_2', 'Phantom_Indicator', 'Province',
];

const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5 MB

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    if (!files || files.length === 0) {
      return NextResponse.json({ error: 'No files provided' }, { status: 400 });
    }

    // ── Per-file size check ───────────────────────────────────────────────────
    const tooLarge = files.filter((f) => f.size > MAX_FILE_SIZE);
    if (tooLarge.length > 0) {
      return NextResponse.json(
        { error: `File(s) exceed 5 MB limit: ${tooLarge.map((f) => f.name).join(', ')}. Remove all FALSE phantom rows and reload.` },
        { status: 400 }
      );
    }

    const results = [];
    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const parsed = parseExcelBuffer(buffer, file.name);
      results.push(parsed);
    }

    // ── Duplicate-data detection via fingerprint ──────────────────────────────
    const fingerprintMap = new Map<string, string>(); // fingerprint → first fileName
    const duplicateOf: Record<string, string | null> = {};
    for (const r of results) {
      if (fingerprintMap.has(r.fingerprint)) {
        duplicateOf[r.fileName] = fingerprintMap.get(r.fingerprint)!;
      } else {
        fingerprintMap.set(r.fingerprint, r.fileName);
        duplicateOf[r.fileName] = null;
      }
    }

    const allRows = results.flatMap((r) => r.rows);
    const allDateCols = [...new Set(results.flatMap((r) => r.dateColumns))].sort();
    const allChannelCols = [...new Set(results.flatMap((r) => r.channels))].sort();
    const mostRecentDateCol = allDateCols.length > 0 ? allDateCols[allDateCols.length - 1] : null;

    // Filter to phantom rows only
    const phantomRows = allRows.filter((r) => {
      const val = r.Phantom_Indicator.trim().toUpperCase();
      return val === 'TRUE' || val === 'NEGATIVE';
    });

    // Compact row format
    const rowHeaders = mostRecentDateCol
      ? [...NEEDED_FIELDS, mostRecentDateCol]
      : NEEDED_FIELDS;

    const rowData = phantomRows.map((row) =>
      rowHeaders.map((h) => row[h] ?? '')
    );

    return NextResponse.json({
      files: results.map((r) => ({
        fileName: r.fileName,
        clientName: r.clientName,
        rowCount: r.rows.length,
        dateColumns: r.dateColumns,
        provinces: r.provinces,
        channels: r.channels,
        missingFields: r.missingFields,
        duplicateOf: duplicateOf[r.fileName] ?? null,
        fingerprint: r.fingerprint,
      })),
      totalRows: allRows.length,
      phantomCount: phantomRows.length,
      allDateColumns: allDateCols,
      allChannels: allChannelCols,
      mostRecentDateCol,
      rowHeaders,
      rowData,
    });
  } catch (e) {
    console.error('[parse]', e);
    const msg = e instanceof Error ? e.message : 'Failed to parse files';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
