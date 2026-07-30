import { NextResponse } from 'next/server';
import { getDriveContext, readChannelsConfig, writeChannelsConfig } from '@/lib/graph-oj';

export const runtime = 'nodejs';

export async function GET() {
  try {
    const ctx = await getDriveContext();
    const channels = await readChannelsConfig(ctx);
    return NextResponse.json({ channels });
  } catch (e) {
    console.error('[admin/channels GET]', e);
    const msg = e instanceof Error ? e.message : 'Failed to load channels';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

export async function PUT(req: Request) {
  try {
    const body = await req.json() as { channels?: unknown };
    if (!Array.isArray(body.channels)) {
      return NextResponse.json({ error: 'channels must be an array' }, { status: 400 });
    }
    const channels = (body.channels as unknown[])
      .map((s) => String(s).trim())
      .filter(Boolean);

    const ctx = await getDriveContext();
    await writeChannelsConfig(channels, ctx);
    return NextResponse.json({ channels });
  } catch (e) {
    console.error('[admin/channels PUT]', e);
    const msg = e instanceof Error ? e.message : 'Failed to save channels';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
