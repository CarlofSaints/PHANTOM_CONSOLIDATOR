import { NextResponse } from 'next/server';
import { readControlFileBuffer } from '@/lib/graph-oj';
import * as XLSX from 'xlsx';
import type { ControlMap, RepInfo } from '@/types';

export async function GET() {
  try {
    const buffer = await readControlFileBuffer();
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' }) as Record<string, string>[];

    // Expected headers (case-insensitive):
    // Personnel_Level_1 | Personnel_Level_2 | Personnel_Level_1 - EMAIL | Personnel_Level_2 - EMAIL

    const controlMap: ControlMap = {};

    for (const row of rows) {
      // Normalise keys
      const norm = (k: string) => k.trim().toLowerCase();
      const entries = Object.entries(row);

      let l1Name = '';
      let l2Name = '';
      let l1Email = '';
      let l2Email = '';

      for (const [k, v] of entries) {
        const nk = norm(k);
        if (nk === 'personnel_level_1' || nk === 'personnel level 1') l1Name = String(v).trim();
        if (nk === 'personnel_level_2' || nk === 'personnel level 2') l2Name = String(v).trim();
        if (nk === 'personnel_level_1 - email' || nk === 'personnel level 1 - email' || nk === 'personnel_level_1 email') l1Email = String(v).trim();
        if (nk === 'personnel_level_2 - email' || nk === 'personnel level 2 - email' || nk === 'personnel_level_2 email') l2Email = String(v).trim();
      }

      if (l2Name) {
        const info: RepInfo = { l1Name, l1Email, l2Email };
        controlMap[l2Name] = info;
      }
    }

    return NextResponse.json({ controlMap });
  } catch (e) {
    console.error('[control-file]', e);
    const msg = e instanceof Error ? e.message : 'Failed to load control file';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
