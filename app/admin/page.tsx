'use client';

import { useState, useEffect } from 'react';
import Link from 'next/link';

type SaveState = 'idle' | 'saving' | 'saved' | 'error';

export default function AdminPage() {
  const [channels, setChannels] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [newChannel, setNewChannel] = useState('');
  const [saveState, setSaveState] = useState<SaveState>('idle');
  const [saveError, setSaveError] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);

  useEffect(() => {
    void (async () => {
      try {
        const res = await fetch('/api/admin/channels');
        const data = await res.json() as { channels?: string[]; error?: string };
        if (!res.ok) throw new Error(data.error ?? 'Failed to load channels');
        setChannels(data.channels ?? []);
      } catch (e) {
        setLoadError(e instanceof Error ? e.message : 'Failed to load channels');
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const update = (next: string[]) => { setChannels(next); setDirty(true); setSaveState('idle'); };

  const handleEdit = (index: number, value: string) => {
    const next = [...channels];
    next[index] = value;
    update(next);
  };

  const handleDelete = (index: number) => {
    update(channels.filter((_, i) => i !== index));
  };

  const handleMoveUp = (index: number) => {
    if (index === 0) return;
    const next = [...channels];
    [next[index - 1], next[index]] = [next[index], next[index - 1]];
    update(next);
  };

  const handleMoveDown = (index: number) => {
    if (index === channels.length - 1) return;
    const next = [...channels];
    [next[index], next[index + 1]] = [next[index + 1], next[index]];
    update(next);
  };

  const handleAdd = () => {
    const name = newChannel.trim();
    if (!name) return;
    if (channels.some((c) => c.toLowerCase() === name.toLowerCase())) {
      setSaveError(`"${name}" already exists.`);
      return;
    }
    update([...channels, name]);
    setNewChannel('');
    setSaveError(null);
  };

  const handleSave = async () => {
    setSaveState('saving');
    setSaveError(null);
    try {
      const res = await fetch('/api/admin/channels', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ channels }),
      });
      const data = await res.json() as { channels?: string[]; error?: string };
      if (!res.ok) throw new Error(data.error ?? 'Save failed');
      setChannels(data.channels ?? channels);
      setSaveState('saved');
      setDirty(false);
    } catch (e) {
      setSaveError(e instanceof Error ? e.message : 'Save failed');
      setSaveState('error');
    }
  };

  return (
    <div className="min-h-screen bg-background text-foreground">
      {/* Header */}
      <header className="border-b border-border px-6 py-3 flex items-center justify-between sticky top-0 bg-card z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-1.5 h-8 bg-accent rounded" />
          <div>
            <h1 className="text-xl font-bold text-foreground">Admin</h1>
            <p className="text-muted text-xs">Channel management &mdash; Phantom Consolidator</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <Link
            href="/"
            className="text-sm text-muted hover:text-foreground transition-colors flex items-center gap-1"
          >
            &#8592; Back to app
          </Link>
          {/* iRam logo */}
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/iram-logo.png" alt="iRam" className="h-9 w-auto object-contain" />
        </div>
      </header>

      <main className="max-w-2xl mx-auto px-6 py-8">
        <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
          <h2 className="text-lg font-bold mb-1" style={{ color: '#79BE43' }}>Channel List</h2>
          <p className="text-muted text-xs mb-6">
            These channels appear in the main app&apos;s channel selector. Names must match exactly what appears in the <strong>Channel</strong> column of your uploaded data files.
          </p>

          {loading && (
            <p className="text-muted text-sm animate-pulse">Loading channels from SharePoint...</p>
          )}

          {loadError && (
            <div className="bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm mb-4">
              &#9888; {loadError}
            </div>
          )}

          {!loading && !loadError && (
            <>
              {channels.length === 0 && (
                <p className="text-muted text-sm italic mb-4">No channels defined. Add one below.</p>
              )}

              <div className="space-y-2 mb-6">
                {channels.map((ch, i) => (
                  <div key={i} className="flex items-center gap-2 bg-background border border-border rounded-lg px-3 py-2">
                    {/* Reorder */}
                    <div className="flex flex-col gap-0.5">
                      <button
                        onClick={() => handleMoveUp(i)}
                        disabled={i === 0}
                        className="text-muted hover:text-foreground disabled:opacity-20 text-xs leading-none px-1"
                        title="Move up"
                      >
                        ▲
                      </button>
                      <button
                        onClick={() => handleMoveDown(i)}
                        disabled={i === channels.length - 1}
                        className="text-muted hover:text-foreground disabled:opacity-20 text-xs leading-none px-1"
                        title="Move down"
                      >
                        ▼
                      </button>
                    </div>

                    {/* Inline edit */}
                    <input
                      type="text"
                      value={ch}
                      onChange={(e) => handleEdit(i, e.target.value)}
                      className="flex-1 bg-transparent text-foreground text-sm focus:outline-none focus:ring-1 focus:ring-accent rounded px-1 py-0.5"
                      style={{ '--tw-ring-color': '#79BE43' } as React.CSSProperties}
                    />

                    {/* Delete */}
                    <button
                      onClick={() => handleDelete(i)}
                      className="text-danger/60 hover:text-danger text-xs px-2 py-1 rounded transition-colors flex-shrink-0"
                      title="Remove channel"
                    >
                      &#10005;
                    </button>
                  </div>
                ))}
              </div>

              {/* Add new */}
              <div className="flex gap-2 mb-6">
                <input
                  type="text"
                  value={newChannel}
                  onChange={(e) => { setNewChannel(e.target.value); setSaveError(null); }}
                  onKeyDown={(e) => { if (e.key === 'Enter') handleAdd(); }}
                  placeholder="New channel name..."
                  className="flex-1 bg-background border border-border rounded-lg px-3 py-2 text-sm text-foreground placeholder:text-muted focus:outline-none focus:ring-1 focus:ring-accent"
                  style={{ '--tw-ring-color': '#79BE43' } as React.CSSProperties}
                />
                <button
                  onClick={handleAdd}
                  disabled={!newChannel.trim()}
                  className="px-4 py-2 text-sm font-bold rounded-lg text-white disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  style={{ background: newChannel.trim() ? '#79BE43' : undefined }}
                  onMouseEnter={(e) => { if (newChannel.trim()) e.currentTarget.style.background = '#69a938'; }}
                  onMouseLeave={(e) => { if (newChannel.trim()) e.currentTarget.style.background = '#79BE43'; }}
                >
                  Add
                </button>
              </div>

              {saveError && (
                <div className="bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm mb-4">
                  &#9888; {saveError}
                </div>
              )}

              {/* Save */}
              <div className="flex items-center gap-4">
                <button
                  onClick={() => void handleSave()}
                  disabled={!dirty || saveState === 'saving'}
                  className="px-6 py-2.5 text-sm font-bold rounded-lg text-white disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  style={{ background: dirty && saveState !== 'saving' ? '#79BE43' : undefined }}
                  onMouseEnter={(e) => { if (dirty && saveState !== 'saving') e.currentTarget.style.background = '#69a938'; }}
                  onMouseLeave={(e) => { if (dirty && saveState !== 'saving') e.currentTarget.style.background = '#79BE43'; }}
                >
                  {saveState === 'saving' ? 'Saving...' : 'Save changes'}
                </button>

                {saveState === 'saved' && (
                  <span className="text-success text-sm">&#10003; Saved to SharePoint</span>
                )}

                {!dirty && saveState !== 'saved' && (
                  <span className="text-muted text-xs">No unsaved changes</span>
                )}
              </div>
            </>
          )}
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t border-border px-6 py-4 flex justify-end items-center gap-3">
        <span className="text-muted text-xs">Powered by</span>
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/oj-logo.png" alt="OuterJoin" className="h-5 w-auto object-contain opacity-75" />
      </footer>
    </div>
  );
}
