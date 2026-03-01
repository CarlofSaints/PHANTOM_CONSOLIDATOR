/**
 * Builds per-store XLSX report from filtered rows.
 * Server-side only.
 */
import * as XLSX from 'xlsx';
import type { RawRow } from '@/types';

const REPORT_FIELDS: (keyof RawRow)[] = [
  'CLIENT',
  'Product_Principle',
  'Sub_Channel',
  'SiteCode',
  'Store_Name',
  'Store_Status',
  'Product_Brand',
  'Product_Sub_Category',
  'Channel_ArticleCode',
  'Client_Product_ID',
  'Product_Description',
  'Product_Status',
  'Range_Indicator',
];

// Display names for the headers in the output Excel
const FIELD_LABELS: Record<string, string> = {
  CLIENT: 'Client',
  Product_Principle: 'Product Principle',
  Sub_Channel: 'Sub Channel',
  SiteCode: 'Site Code',
  Store_Name: 'Store Name',
  Store_Status: 'Store Status',
  Product_Brand: 'Product Brand',
  Product_Sub_Category: 'Product Sub Category',
  Channel_ArticleCode: 'Channel ArticleCode',
  Client_Product_ID: 'Client Product ID',
  Product_Description: 'Product Description',
  Product_Status: 'Product Status',
  Range_Indicator: 'Range Indicator',
};

export function sanitizeFilename(name: string): string {
  return name.replace(/[/\\:*?"<>|]/g, '-').trim();
}

export function buildStoreReport(
  rows: RawRow[],
  mostRecentDateCol: string
): Buffer {
  const ws: XLSX.WorkSheet = {};

  const outputHeaders = [
    ...REPORT_FIELDS.map((f) => FIELD_LABELS[f] ?? f),
    mostRecentDateCol,
  ];

  // Write header row
  outputHeaders.forEach((h, c) => {
    ws[XLSX.utils.encode_cell({ r: 0, c })] = {
      v: h,
      t: 's',
      s: { font: { bold: true } },
    };
  });

  // Write data rows
  rows.forEach((row, rowIdx) => {
    const r = rowIdx + 1;
    REPORT_FIELDS.forEach((field, c) => {
      ws[XLSX.utils.encode_cell({ r, c })] = { v: row[field] ?? '', t: 's' };
    });
    // Most recent date column
    const dateVal = row[mostRecentDateCol] ?? '';
    ws[XLSX.utils.encode_cell({ r, c: REPORT_FIELDS.length })] = {
      v: dateVal,
      t: 's',
    };
  });

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rows.length, c: outputHeaders.length - 1 },
  });

  ws['!cols'] = outputHeaders.map((h) => ({
    wch: Math.max(h.length + 2, 18),
  }));

  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft' };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Phantom Stock');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}
