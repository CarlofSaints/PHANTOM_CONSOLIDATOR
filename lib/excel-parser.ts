/**
 * Server-side Excel parser
 * Detects columns by header name (not position).
 * Only import in API routes — never in client components.
 */
import * as XLSX from 'xlsx';
import type { RawRow, ParsedFile } from '@/types';

const DATE_COL_REGEX = /^\d{4}\/\d{2}\/\d{2}$/;

// Map of plan field names → possible header variations (case-insensitive)
const HEADER_MAP: Record<string, string[]> = {
  CLIENT: ['client'],
  Product_Principle: ['product principle', 'product_principle'],
  Channel: ['channel'],
  Sub_Channel: ['sub_channel', 'sub channel'],
  Province: ['province'],
  Personnel_Level_1: ['personnel_level_1', 'personnel level 1', 'personnel_level1'],
  Personnel_Level_2: ['personnel_level_2', 'personnel level 2', 'personnel_level2'],
  SiteCode: ['sitecode', 'site code', 'site_code'],
  Store_Name: ['store name', 'store_name'],
  Store_Status: ['store status', 'store_status'],
  Product_Brand: ['product brand', 'product_brand'],
  Product_Sub_Category: ['product sub category', 'product_sub_category', 'product sub_category'],
  Channel_ArticleCode: ['channel articlecode', 'channel_articlecode', 'channel article code'],
  Client_Product_ID: ['client product id', 'client_product_id'],
  Product_Description: ['product description', 'product_description'],
  Product_Status: ['product status', 'product_status'],
  Range_Indicator: ['range indicator', 'range_indicator'],
  Phantom_Indicator: ['phantom indicator', 'phantom_indicator', 'phantom stock indicator'],
};

function normalise(s: string): string {
  return s.trim().toLowerCase();
}

function buildHeaderIndex(headers: string[]): Record<string, number> {
  return Object.fromEntries(headers.map((h, i) => [normalise(h), i]));
}

function findColIndex(
  headerIndex: Record<string, number>,
  aliases: string[]
): number {
  for (const alias of aliases) {
    if (alias in headerIndex) return headerIndex[alias];
  }
  return -1;
}

export function parseExcelBuffer(buffer: Buffer, fileName: string): ParsedFile {
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  // Use first sheet
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as string[][];

  if (raw.length === 0) throw new Error(`${fileName}: empty sheet`);

  const headerRow = raw[0].map(String);
  const headerIndex = buildHeaderIndex(headerRow);

  // Detect date columns
  const dateColumns = headerRow.filter((h) => DATE_COL_REGEX.test(h.trim()));

  // Build column index map for typed fields
  const colIdx: Record<string, number> = {};
  for (const [field, aliases] of Object.entries(HEADER_MAP)) {
    colIdx[field] = findColIndex(headerIndex, aliases);
  }

  // Date column indices
  const dateColIndices: Record<string, number> = {};
  for (const dc of dateColumns) {
    dateColIndices[dc] = headerIndex[normalise(dc)] ?? -1;
  }

  const rows: RawRow[] = [];

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (row.every((cell) => cell === '' || cell === null || cell === undefined)) continue;

    const get = (field: string): string => {
      const idx = colIdx[field];
      return idx >= 0 ? String(row[idx] ?? '').trim() : '';
    };

    const built: RawRow = {
      CLIENT: get('CLIENT'),
      Product_Principle: get('Product_Principle'),
      Channel: get('Channel'),
      Sub_Channel: get('Sub_Channel'),
      Province: get('Province'),
      Personnel_Level_1: get('Personnel_Level_1'),
      Personnel_Level_2: get('Personnel_Level_2'),
      SiteCode: get('SiteCode'),
      Store_Name: get('Store_Name'),
      Store_Status: get('Store_Status'),
      Product_Brand: get('Product_Brand'),
      Product_Sub_Category: get('Product_Sub_Category'),
      Channel_ArticleCode: get('Channel_ArticleCode'),
      Client_Product_ID: get('Client_Product_ID'),
      Product_Description: get('Product_Description'),
      Product_Status: get('Product_Status'),
      Range_Indicator: get('Range_Indicator'),
      Phantom_Indicator: get('Phantom_Indicator'),
    };

    // Attach date column values
    for (const [dc, idx] of Object.entries(dateColIndices)) {
      built[dc] = idx >= 0 ? String(row[idx] ?? '').trim() : '';
    }

    rows.push(built);
  }

  // Determine client name: use first non-empty CLIENT value in rows, or filename stem
  const clientName =
    rows.find((r) => r.CLIENT)?.CLIENT ??
    fileName.replace(/\.[^.]+$/, '');

  return { fileName, clientName, rows, dateColumns };
}
