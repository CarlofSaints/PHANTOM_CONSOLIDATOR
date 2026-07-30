import * as XLSX from 'xlsx';
import type { ControlMap, RepInfo } from '@/types';

/**
 * Parses an Excel control file buffer into a ControlMap.
 * Merges into an existing map if provided (later entries overwrite earlier on conflict).
 */
export function parseControlBuffer(buffer: Buffer, into: ControlMap = {}): ControlMap {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName =
    workbook.SheetNames.find((n) => n.trim().toLowerCase() === 'sheet1') ??
    workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const xlsxRows = XLSX.utils.sheet_to_json(ws, { defval: '' }) as Record<string, string>[];

  for (const row of xlsxRows) {
    let l1Name = '', l2Name = '', l1Email = '', l2Email = '', l2Alias = '';

    for (const [k, v] of Object.entries(row)) {
      const nk = k.trim().toLowerCase();
      if (nk === 'personnel_level_1' || nk === 'personnel level 1')
        l1Name = String(v).trim().replace(/\s+\S+@\S+\.\S+$/, '').trim();
      if (nk === 'personnel_level_2' || nk === 'personnel level 2')
        l2Name = String(v).trim().replace(/\s+\S+@\S+\.\S+$/, '').trim();
      if (nk.includes('level_1') && nk.includes('email')) l1Email = String(v).trim();
      if (nk.includes('level_2') && nk.includes('email')) l2Email = String(v).trim();
      if ((nk.includes('level_2') || nk.includes('level 2')) && (nk.includes('alias') || nk.includes('alt'))) {
        l2Alias = String(v).trim();
      }
    }

    if (l2Name) {
      const info: RepInfo = { l1Name, l1Email, l2Email };
      into[l2Name.toLowerCase()] = info;
      if (l2Alias) into[l2Alias.toLowerCase()] = info;
    }
  }

  return into;
}
