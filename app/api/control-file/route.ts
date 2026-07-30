import { NextResponse } from 'next/server';
import { getDriveContext, readControlFileBuffer } from '@/lib/graph-oj';
import { parseControlBuffer } from '@/lib/parse-control-file';
import type { ControlMap } from '@/types';

export async function GET(req: Request) {
  try {
    const url = new URL(req.url);
    const channelsParam = url.searchParams.get('channels') ?? '';
    const channels = channelsParam.split(',').map((s) => s.trim()).filter(Boolean);

    if (channels.length === 0) {
      return NextResponse.json({ controlMap: {} });
    }

    // Fetch a shared drive context (one auth call) then read all channel files in parallel
    const ctx = await getDriveContext();
    const results = await Promise.allSettled(
      channels.map((ch) => readControlFileBuffer(ch, ctx))
    );

    const controlMap: ControlMap = {};
    const errors: string[] = [];

    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const channelName = channels[i];
      if (result.status === 'rejected') {
        errors.push(`"${channelName}": ${result.reason instanceof Error ? result.reason.message : String(result.reason)}`);
      } else {
        parseControlBuffer(result.value, controlMap);
      }
    }

    return NextResponse.json({ controlMap, errors });
  } catch (e) {
    console.error('[control-file]', e);
    const msg = e instanceof Error ? e.message : 'Failed to load control file';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
