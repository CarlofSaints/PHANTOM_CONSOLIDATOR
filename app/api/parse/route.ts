import { NextResponse } from 'next/server';
import { parseExcelBuffer } from '@/lib/excel-parser';

export const maxDuration = 60;

// Fields needed by the process route — sent once as headers, not repeated per row
const NEEDED_FIELDS = [
  'CLIENT', 'Product_Principle', 'Sub_Channel', 'SiteCode', 'Store_Name',
  'Store_Status', 'Product_Brand', 'Product_Sub_Category', 'Channel_ArticleCode',
  'Client_Product_ID', 'Product_Description', 'Product_Status', 'Range_Indicator',
  'Personnel_Level_1', 'Personnel_Level_2', 'Phantom_Indicator',
];

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    if (!files || files.length === 0) {
      return NextResponse.json({ error: 'No files provided' }, { status: 400 });
    }

    const results = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const parsed = parseExcelBuffer(buffer, file.name);
      results.push(parsed);
    }

    const allRows = results.flatMap((r) => r.rows);
    const allDateCols = [...new Set(results.flatMap((r) => r.dateColumns))].sort();
    const mostRecentDateCol = allDateCols.length > 0 ? allDateCols[allDateCols.length - 1] : null;

    // Filter to phantom rows only
    const phantomRows = allRows.filter((r) => {
      const val = r.Phantom_Indicator.trim().toUpperCase();
      return val === 'TRUE' || val === 'NEGATIVE';
    });

    // Compact array format: field names sent once, rows as value arrays only
    // This roughly halves payload size vs sending full JSON objects
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
      })),
      totalRows: allRows.length,
      phantomCount: phantomRows.length,
      allDateColumns: allDateCols,
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
