'use client';

import { useState, useRef, useCallback } from 'react';
import type { RawRow, ControlMap, ProcessSummary } from '@/types';

// ── Types for client-side state ──────────────────────────────────────────────

interface ParseResponse {
  files: { fileName: string; clientName: string; rowCount: number; dateColumns: string[] }[];
  totalRows: number;
  allDateColumns: string[];
  mostRecentDateCol: string | null;
  rows: RawRow[];
}

type Stage = 'idle' | 'parsed' | 'processing' | 'done' | 'error';

// ── Tiny UI components ───────────────────────────────────────────────────────

function Badge({ label, value, color = 'accent' }: { label: string; value: string | number; color?: string }) {
  const colorClass =
    color === 'accent' ? 'text-accent' : color === 'success' ? 'text-success' : 'text-warning';
  return (
    <div className="bg-card border border-border rounded-lg p-4 text-center">
      <div className={`text-2xl font-bold ${colorClass}`}>{value}</div>
      <div className="text-muted text-sm mt-1">{label}</div>
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-card border border-border rounded-xl p-6 mb-6">
      <h2 className="text-lg font-bold text-foreground mb-4 border-b border-border pb-3">{title}</h2>
      {children}
    </div>
  );
}

// ── Main Page ────────────────────────────────────────────────────────────────

export default function Home() {
  const [stage, setStage] = useState<Stage>('idle');
  const [parseResult, setParseResult] = useState<ParseResponse | null>(null);
  const [controlMap, setControlMap] = useState<ControlMap | null>(null);
  const [controlError, setControlError] = useState<string | null>(null);
  const [processSummary, setProcessSummary] = useState<ProcessSummary | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [isFetchingControl, setIsFetchingControl] = useState(false);
  const [includeNegative, setIncludeNegative] = useState(false);
  const [recipientMode, setRecipientMode] = useState<'l1' | 'l2' | 'both'>('both');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── File upload + parse ──────────────────────────────────────────────────

  const handleFiles = useCallback(async (files: FileList | File[]) => {
    const fileArr = Array.from(files).filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (fileArr.length === 0) return;

    setIsUploading(true);
    setErrorMsg(null);

    try {
      const fd = new FormData();
      for (const f of fileArr) fd.append('files', f);

      const res = await fetch('/api/parse', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error((data as { error?: string }).error ?? 'Parse failed');

      setParseResult(data as ParseResponse);
      setStage('parsed');

      // Fetch control file
      setIsFetchingControl(true);
      setControlError(null);
      try {
        const ctrl = await fetch('/api/control-file');
        const ctrlData = await ctrl.json();
        if (!ctrl.ok) throw new Error((ctrlData as { error?: string }).error ?? 'Control file fetch failed');
        setControlMap((ctrlData as { controlMap: ControlMap }).controlMap);
      } catch (e) {
        setControlError(e instanceof Error ? e.message : 'Could not load control file');
      } finally {
        setIsFetchingControl(false);
      }
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Upload failed');
      setStage('error');
    } finally {
      setIsUploading(false);
    }
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      void handleFiles(e.dataTransfer.files);
    },
    [handleFiles]
  );

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = () => setIsDragging(false);

  // ── Process ──────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (!parseResult || !controlMap) return;

    setStage('processing');
    setErrorMsg(null);

    try {
      const res = await fetch('/api/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          rows: parseResult.rows,
          controlMap,
          reportDate: parseResult.mostRecentDateCol
            ? parseResult.mostRecentDateCol.replace(/\//g, '-')
            : new Date().toISOString().split('T')[0],
          mostRecentDateCol: parseResult.mostRecentDateCol ?? '',
          includeNegative,
          recipientMode,
        }),
      });

      const data = await res.json();
      if (!res.ok) throw new Error((data as { error?: string }).error ?? 'Processing failed');

      setProcessSummary((data as { summary: ProcessSummary }).summary);
      setStage('done');
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Processing failed');
      setStage('error');
    }
  };

  // ── Derived data for preview ──────────────────────────────────────────────

  const uniqueStores = parseResult
    ? new Set(parseResult.rows.map((r) => r.Store_Name)).size
    : 0;

  const uniqueL2s = parseResult
    ? new Set(parseResult.rows.map((r) => r.Personnel_Level_2).filter(Boolean))
    : new Set<string>();

  const missingReps = controlMap && parseResult
    ? [...uniqueL2s].filter((name) => !controlMap[name])
    : [];

  const foundReps = controlMap && parseResult
    ? [...uniqueL2s].filter((name) => !!controlMap[name])
    : [];

  const reportDate = parseResult?.mostRecentDateCol
    ? parseResult.mostRecentDateCol.replace(/\//g, '-')
    : '—';

  // ── Render ───────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-background text-foreground">
      {/* Header */}
      <header className="border-b border-border px-6 py-4 flex items-center justify-between sticky top-0 bg-background z-10">
        <div className="flex items-center gap-3">
          <div className="w-2 h-8 bg-accent rounded" />
          <div>
            <h1 className="text-xl font-bold text-foreground">Phantom Consolidator</h1>
            <p className="text-muted text-xs">Multi-vendor phantom stock reporting &mdash; OuterJoin</p>
          </div>
        </div>
        <div className="flex gap-2">
          {stage === 'parsed' && (
            <span className="bg-accent/10 text-accent border border-accent/30 text-xs px-3 py-1 rounded-full">
              Files Loaded
            </span>
          )}
          {stage === 'processing' && (
            <span className="bg-warning/10 text-warning border border-warning/30 text-xs px-3 py-1 rounded-full animate-pulse">
              Processing...
            </span>
          )}
          {stage === 'done' && (
            <span className="bg-success/10 text-success border border-success/30 text-xs px-3 py-1 rounded-full">
              Complete
            </span>
          )}
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8">

        {/* Section 1: Upload */}
        <Section title="1 — Upload Files">
          <div
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-lg p-10 text-center cursor-pointer transition-colors ${
              isDragging
                ? 'border-accent bg-accent/5'
                : 'border-border hover:border-accent/50'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              multiple
              className="hidden"
              onChange={(e) => { if (e.target.files) void handleFiles(e.target.files); }}
            />
            {isUploading ? (
              <p className="text-accent animate-pulse">Parsing files...</p>
            ) : (
              <>
                <p className="text-foreground font-medium">Drop Excel files here or click to browse</p>
                <p className="text-muted text-sm mt-1">Accepts multiple .xlsx / .xls files (one per vendor)</p>
              </>
            )}
          </div>

          {/* File list */}
          {parseResult && parseResult.files.length > 0 && (
            <div className="mt-4 space-y-2">
              {parseResult.files.map((f, i) => (
                <div
                  key={i}
                  className="flex items-center justify-between bg-background border border-border rounded-lg px-4 py-3"
                >
                  <div>
                    <span className="text-foreground font-medium text-sm">{f.fileName}</span>
                    <span className="ml-3 text-accent text-xs font-mono">{f.clientName}</span>
                  </div>
                  <div className="text-right text-xs text-muted">
                    <div>{f.rowCount.toLocaleString()} rows</div>
                    <div>{f.dateColumns.length > 0 ? f.dateColumns.join(', ') : 'No date cols'}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Section>

        {/* Section 2: Settings */}
        <Section title="2 — Settings">
          <div className="space-y-4">
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={includeNegative}
                onChange={(e) => setIncludeNegative(e.target.checked)}
                className="w-4 h-4"
                style={{ accentColor: '#f97316' }}
              />
              <span className="text-foreground">Include NEGATIVE phantom rows</span>
            </label>

            <div>
              <p className="text-foreground text-sm mb-2 font-medium">Send reports to:</p>
              <div className="flex gap-6">
                {(['l1', 'l2', 'both'] as const).map((mode) => (
                  <label key={mode} className="flex items-center gap-2 cursor-pointer">
                    <input
                      type="radio"
                      name="recipientMode"
                      value={mode}
                      checked={recipientMode === mode}
                      onChange={() => setRecipientMode(mode)}
                      style={{ accentColor: '#f97316' }}
                    />
                    <span className="text-foreground text-sm">
                      {mode === 'l1' ? 'Level 1 (managers)' : mode === 'l2' ? 'Level 2 (reps)' : 'Both'}
                    </span>
                  </label>
                ))}
              </div>
            </div>
          </div>
        </Section>

        {/* Section 3: Preview */}
        {stage === 'parsed' && parseResult && (
          <Section title="3 — Preview">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
              <Badge label="Total Rows" value={parseResult.totalRows.toLocaleString()} />
              <Badge label="Unique Stores" value={uniqueStores} />
              <Badge label="L2 Reps" value={uniqueL2s.size} />
              <Badge label="Report Date" value={reportDate} />
            </div>

            {isFetchingControl && (
              <p className="text-muted text-sm animate-pulse mb-3">Loading control file from SharePoint...</p>
            )}

            {controlError && (
              <div className="bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm mb-3">
                &#9888; Control file error: {controlError}
              </div>
            )}

            {controlMap && (
              <div className="space-y-3">
                <div className="flex items-center gap-2 text-success text-sm">
                  <span>&#10003;</span>
                  <span>Control file loaded &mdash; {Object.keys(controlMap).length} reps mapped</span>
                </div>

                {foundReps.length > 0 && (
                  <div className="text-sm text-muted">
                    <span className="text-success font-medium">Found: </span>
                    {foundReps.join(', ')}
                  </div>
                )}

                {missingReps.length > 0 && (
                  <div className="bg-warning/10 border border-warning/30 rounded-lg px-4 py-3">
                    <p className="text-warning text-sm font-medium">
                      &#9888; {missingReps.length} rep(s) not in control file — emails skipped:
                    </p>
                    <p className="text-warning/80 text-xs mt-1">{missingReps.join(', ')}</p>
                  </div>
                )}
              </div>
            )}
          </Section>
        )}

        {/* Section 4: Process & Send */}
        {(stage === 'parsed' || stage === 'processing' || stage === 'done' || stage === 'error') && (
          <Section title="4 — Process &amp; Send">
            {stage !== 'done' && (
              <button
                onClick={() => { void handleProcess(); }}
                disabled={stage === 'processing' || !parseResult || !controlMap}
                className="w-full font-bold py-3 px-6 rounded-lg transition-colors text-sm text-white disabled:opacity-50 disabled:cursor-not-allowed"
                style={{ background: stage === 'processing' || !parseResult || !controlMap ? undefined : '#f97316' }}
                onMouseEnter={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#ea6c0a'; }}
                onMouseLeave={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#f97316'; }}
              >
                {stage === 'processing' ? 'Processing...' : 'Process & Send Reports'}
              </button>
            )}

            {stage === 'processing' && (
              <div className="mt-4 text-center">
                <p className="text-muted text-sm animate-pulse">
                  Building XLSX reports, uploading to iRAM SharePoint, queuing emails via Graph API...
                </p>
                <p className="text-muted text-xs mt-1">This may take a minute for large batches.</p>
              </div>
            )}

            {errorMsg && (
              <div className="mt-4 bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm">
                &#10007; {errorMsg}
              </div>
            )}

            {stage === 'done' && processSummary && (
              <div className="mt-4 space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                  <Badge label="Reports Saved" value={processSummary.stores} color="success" />
                  <Badge label="Reps Covered" value={processSummary.reps} color="success" />
                  <Badge
                    label="Errors"
                    value={processSummary.errors.length}
                    color={processSummary.errors.length > 0 ? 'warning' : 'success'}
                  />
                </div>

                <div className="text-success text-sm">
                  &#10003; Reports saved to iRAM SharePoint. Emails dispatched in background.
                </div>

                {processSummary.errors.length > 0 && (
                  <div className="bg-warning/10 border border-warning/30 rounded-lg px-4 py-3">
                    <p className="text-warning text-sm font-medium mb-2">Errors encountered:</p>
                    <ul className="text-warning/80 text-xs space-y-1">
                      {processSummary.errors.map((err, i) => (
                        <li key={i}>&#8226; {err}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {processSummary.storeResults.length > 0 && (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs border-collapse">
                      <thead>
                        <tr className="border-b border-border">
                          <th className="text-left py-2 px-3 text-muted font-medium">Store</th>
                          <th className="text-left py-2 px-3 text-muted font-medium">L2 Rep</th>
                          <th className="text-left py-2 px-3 text-muted font-medium">L1 Manager</th>
                          <th className="text-right py-2 px-3 text-muted font-medium">Lines</th>
                          <th className="text-center py-2 px-3 text-muted font-medium">Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {processSummary.storeResults.map((sr, i) => (
                          <tr key={i} className="border-b border-border/50">
                            <td className="py-2 px-3 text-foreground">
                              {sr.webUrl ? (
                                <a
                                  href={sr.webUrl}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="text-accent hover:underline"
                                >
                                  {sr.storeName}
                                </a>
                              ) : (
                                sr.storeName
                              )}
                            </td>
                            <td className="py-2 px-3 text-foreground">{sr.l2Name}</td>
                            <td className="py-2 px-3 text-muted">{sr.l1Name}</td>
                            <td className="py-2 px-3 text-right text-foreground">{sr.rowCount}</td>
                            <td className="py-2 px-3 text-center">
                              {sr.error ? (
                                <span className="text-danger" title={sr.error}>&#10007;</span>
                              ) : (
                                <span className="text-success">&#10003;</span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                <button
                  onClick={() => {
                    setStage('idle');
                    setParseResult(null);
                    setControlMap(null);
                    setControlError(null);
                    setProcessSummary(null);
                    setErrorMsg(null);
                  }}
                  className="text-muted text-sm hover:text-accent underline"
                >
                  Start a new batch
                </button>
              </div>
            )}
          </Section>
        )}
      </main>
    </div>
  );
}
