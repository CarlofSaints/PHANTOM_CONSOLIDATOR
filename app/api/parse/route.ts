import { NextResponse } from 'next/server';
import { parseExcelBuffer } from '@/lib/excel-parser';
import type { ParsedFile } from '@/types';

export const maxDuration = 60;

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    if (!files || files.length === 0) {
      return NextResponse.json({ error: 'No files provided' }, { status: 400 });
    }

    const results: ParsedFile[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const parsed = parseExcelBuffer(buffer, file.name);
      results.push(parsed);
    }

    // Combine all rows and merge date columns across all files
    const allRows = results.flatMap((r) => r.rows);
    const allDateCols = [...new Set(results.flatMap((r) => r.dateColumns))].sort();

    // Only keep phantom rows (TRUE or NEGATIVE) — discards the bulk of non-phantom data
    // Process route will apply the includeNegative filter at send time
    const phantomRows = allRows.filter((r) => {
      const val = r.Phantom_Indicator.trim().toUpperCase();
      return val === 'TRUE' || val === 'NEGATIVE';
    });

    return NextResponse.json({
      files: results.map((r) => ({
        fileName: r.fileName,
        clientName: r.clientName,
        rowCount: r.rows.length,
        dateColumns: r.dateColumns,
      })),
      totalRows: allRows.length,
      phantomCount: phantomRows.length,
      allDateColumns: allDateCols,
      mostRecentDateCol: allDateCols.length > 0 ? allDateCols[allDateCols.length - 1] : null,
      rows: phantomRows,
    });
  } catch (e) {
    console.error('[parse]', e);
    const msg = e instanceof Error ? e.message : 'Failed to parse files';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
