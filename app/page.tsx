'use client';

import { useState, useRef, useCallback, useEffect } from 'react';
import Link from 'next/link';
import type { ControlMap, ProcessSummary } from '@/types';

const MAX_FILE_BYTES = 5 * 1024 * 1024; // 5 MB client-side guard

// Mirror of NEEDED_FIELDS from parse/route.ts — used for client-side row combining
const NEEDED_FIELDS_CLIENT = [
  'CLIENT', 'Product_Principle', 'Channel', 'Sub_Channel', 'SiteCode', 'Store_Name',
  'Store_Status', 'Product_Brand', 'Product_Sub_Category', 'Channel_ArticleCode',
  'Client_Product_ID', 'Product_Description', 'Product_Status', 'Range_Indicator',
  'Personnel_Level_1', 'Personnel_Level_2', 'Phantom_Indicator', 'Province',
];

// ── Confirmation Modal ────────────────────────────────────────────────────────

function ConfirmUploadModal({ onConfirm, onCancel }: { onConfirm: () => void; onCancel: () => void }) {
  const [showTooltip, setShowTooltip] = useState(false);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60">
      <div className="bg-card border border-border rounded-xl p-6 max-w-md w-full mx-4 shadow-2xl">
        <h2 className="text-foreground font-bold text-base mb-3">Confirm Upload</h2>

        <div className="flex items-start gap-2 mb-6">
          <p className="text-foreground text-sm leading-relaxed">
            Are you sure the 60 days sales data matches the dates of the two SOH dates?
          </p>
          <div className="relative flex-shrink-0">
            <button
              onMouseEnter={() => setShowTooltip(true)}
              onMouseLeave={() => setShowTooltip(false)}
              onFocus={() => setShowTooltip(true)}
              onBlur={() => setShowTooltip(false)}
              className="w-5 h-5 rounded-full border border-border text-muted text-xs font-bold flex items-center justify-center hover:border-accent hover:text-accent transition-colors mt-0.5"
              aria-label="More information"
            >
              i
            </button>
            {showTooltip && (
              <div className="absolute right-0 top-7 w-72 bg-background border border-border rounded-lg p-3 text-xs text-muted leading-relaxed shadow-xl z-10">
                In the raw file, there are two tabs: one for sales and one for SOH. The SOH tab has SOH for two dates 60 days apart. The sales tab needs to be filtered to those same dates so that the sales period matches exactly the two dates selected in the SOH tab.
                <br /><br />
                <span className="text-foreground">For example:</span> if the start SOH date is 1 Jan and the SOH end date is 2 March, then the sales needs to cover those same 60 days.
              </div>
            )}
          </div>
        </div>

        <div className="flex gap-3 justify-end">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-sm rounded-lg border border-border text-muted hover:text-foreground hover:border-foreground/30 transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            className="px-4 py-2 text-sm rounded-lg font-bold text-white transition-colors"
            style={{ background: '#79BE43' }}
            onMouseEnter={(e) => { e.currentTarget.style.background = '#69a938'; }}
            onMouseLeave={(e) => { e.currentTarget.style.background = '#79BE43'; }}
          >
            Yes, proceed
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Types ─────────────────────────────────────────────────────────────────────

interface FileInfo {
  fileName: string;
  clientName: string;
  rowCount: number;
  dateColumns: string[];
  provinces: string[];
  channels: string[];
  missingFields: string[];
  duplicateOf: string | null;
  fingerprint?: string;
}

interface ParseResponse {
  files: FileInfo[];
  totalRows: number;
  phantomCount: number;
  allDateColumns: string[];
  mostRecentDateCol: string | null;
  rowHeaders: string[];
  rowData: string[][];
  allChannels: string[];
}

// Merges per-file parse results into one unified ParseResponse.
// Files are uploaded one-at-a-time to stay under Vercel's 4.5 MB body limit,
// so cross-file duplicate detection and row-header normalisation happen here.
function combineParseResults(results: ParseResponse[]): ParseResponse {
  const globalDateCols = [...new Set(results.flatMap((r) => r.allDateColumns))].sort();
  const globalMostRecent = globalDateCols.length > 0 ? globalDateCols[globalDateCols.length - 1] : null;
  const globalRowHeaders = globalMostRecent
    ? [...NEEDED_FIELDS_CLIENT, globalMostRecent]
    : [...NEEDED_FIELDS_CLIENT];
  const targetLen = globalRowHeaders.length;

  // Normalise each row to targetLen (pad with '' if the file had no date col)
  const combinedRowData: string[][] = results.flatMap((r) =>
    r.rowData.map((row) =>
      row.length >= targetLen ? row : [...row, ...Array(targetLen - row.length).fill('')]
    )
  );

  // Cross-file fingerprint duplicate detection
  const fpMap = new Map<string, string>(); // fingerprint → first fileName
  const combinedFiles: FileInfo[] = results.flatMap((r) => r.files).map((f) => {
    const fp = f.fingerprint ?? '';
    if (fp && fpMap.has(fp)) return { ...f, duplicateOf: fpMap.get(fp)! };
    if (fp) fpMap.set(fp, f.fileName);
    return { ...f };
  });

  const allChannels = [...new Set(results.flatMap((r) => r.allChannels ?? []))].sort();

  return {
    files: combinedFiles,
    totalRows: results.reduce((s, r) => s + r.totalRows, 0),
    phantomCount: results.reduce((s, r) => s + r.phantomCount, 0),
    allDateColumns: globalDateCols,
    mostRecentDateCol: globalMostRecent,
    rowHeaders: globalRowHeaders,
    rowData: combinedRowData,
    allChannels,
  };
}

type Stage = 'idle' | 'parsed' | 'processing' | 'done' | 'error';

// ── Tiny UI helpers ───────────────────────────────────────────────────────────

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
    <div className="bg-card border border-border rounded-xl p-6 mb-6 shadow-sm">
      <h2 className="text-lg font-bold mb-4 border-b border-border pb-3" style={{ color: '#79BE43' }}>{title}</h2>
      {children}
    </div>
  );
}

// ── Main Page ─────────────────────────────────────────────────────────────────

export default function Home() {
  const [stage, setStage] = useState<Stage>('idle');
  const [parseResult, setParseResult] = useState<ParseResponse | null>(null);
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [controlMap, setControlMap] = useState<ControlMap | null>(null);
  const [controlError, setControlError] = useState<string | null>(null);
  const [processSummary, setProcessSummary] = useState<ProcessSummary | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<string | null>(null);
  const [pendingFiles, setPendingFiles] = useState<File[] | null>(null);
  const [isFetchingControl, setIsFetchingControl] = useState(false);
  const [includeNegative, setIncludeNegative] = useState(false);
  const [recipientMode, setRecipientMode] = useState<'l1' | 'l2' | 'both'>('both');
  const [actionMode, setActionMode] = useState<'both' | 'sharepoint' | 'email'>('both');
  const [selectedProvinces, setSelectedProvinces] = useState<Record<string, string[]>>({});
  const [selectedChannels, setSelectedChannels] = useState<string[]>([]);
  const [channels, setChannels] = useState<string[]>([]);
  const [channelsLoading, setChannelsLoading] = useState(true);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Load channel list from SharePoint config on mount ────────────────────

  useEffect(() => {
    void fetch('/api/admin/channels')
      .then((r) => r.json())
      .then((d: { channels?: string[] }) => { if (d.channels?.length) setChannels(d.channels); })
      .catch(() => {/* silently fall through — channels stays [] */})
      .finally(() => setChannelsLoading(false));
  }, []);

  // ── Auto-fetch control file when channels or uploaded data changes ────────

  useEffect(() => {
    if (!parseResult) return;

    if (selectedChannels.length === 0) {
      setControlMap(null);
      setControlError('Select at least one channel above to load the control file.');
      return;
    }

    let cancelled = false;
    setIsFetchingControl(true);
    setControlError(null);
    setControlMap(null);

    void (async () => {
      try {
        const params = selectedChannels.map(encodeURIComponent).join(',');
        const ctrl = await fetch(`/api/control-file?channels=${params}`);
        const ctrlData = await ctrl.json() as { controlMap?: ControlMap; errors?: string[]; error?: string };
        if (cancelled) return;
        if (!ctrl.ok) throw new Error(ctrlData.error ?? 'Control file fetch failed');
        setControlMap(ctrlData.controlMap ?? {});
        if (ctrlData.errors && ctrlData.errors.length > 0) {
          setControlError(`Some control files could not be loaded:\n${ctrlData.errors.join('\n')}`);
        }
      } catch (e) {
        if (!cancelled) setControlError(e instanceof Error ? e.message : 'Could not load control file');
      } finally {
        if (!cancelled) setIsFetchingControl(false);
      }
    })();

    return () => { cancelled = true; };
  }, [selectedChannels, parseResult]);

  // ── File upload + parse ──────────────────────────────────────────────────
  // Files are sent one at a time to stay under Vercel's 4.5 MB request body limit.
  // Results are combined client-side via combineParseResults().

  const handleFiles = useCallback(async (files: FileList | File[]) => {
    const fileArr = Array.from(files).filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (fileArr.length === 0) return;

    setIsUploading(true);
    setUploadProgress(null);
    setErrorMsg(null);
    setUploadError(null);

    try {
      const perFileResults: ParseResponse[] = [];

      for (let i = 0; i < fileArr.length; i++) {
        setUploadProgress(`Parsing file ${i + 1} of ${fileArr.length}: ${fileArr[i].name}`);
        const fd = new FormData();
        fd.append('files', fileArr[i]);

        const res = await fetch('/api/parse', { method: 'POST', body: fd });
        if (!res.ok) {
          const text = await res.text();
          throw new Error(`Error parsing "${fileArr[i].name}" (${res.status}): ${text.slice(0, 200)}`);
        }
        perFileResults.push(await res.json() as ParseResponse);
      }

      const combined = combineParseResults(perFileResults);
      setParseResult(combined);
      setUploadedFiles(fileArr);
      setStage('parsed');

      // Initialise province selection — all provinces selected by default
      const provInit: Record<string, string[]> = {};
      for (const f of combined.files) {
        provInit[f.clientName] = [...f.provinces];
      }
      setSelectedProvinces(provInit);
      // Control file is fetched by the useEffect watching [selectedChannels, parseResult]
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Upload failed');
      setStage('error');
    } finally {
      setIsUploading(false);
      setUploadProgress(null);
    }
  }, []);

  // ── Stage files with client-side validation ──────────────────────────────

  const stageFiles = useCallback((files: FileList | File[]) => {
    const fileArr = Array.from(files).filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (fileArr.length === 0) return;

    setUploadError(null);

    // Issue 4: file too large (>5MB)
    const tooLarge = fileArr.filter((f) => f.size > MAX_FILE_BYTES);
    if (tooLarge.length > 0) {
      setUploadError(
        `File${tooLarge.length > 1 ? 's' : ''} too large (max 5 MB): ${tooLarge.map((f) => f.name).join(', ')}. ` +
        `Open the file, delete all rows where Phantom Indicator = FALSE, save, then reload.`
      );
      return;
    }

    // Issue 1: duplicate filename within the batch
    const seen = new Set<string>();
    const dupNames: string[] = [];
    for (const f of fileArr) {
      if (seen.has(f.name)) dupNames.push(f.name);
      seen.add(f.name);
    }
    if (dupNames.length > 0) {
      setUploadError(`Duplicate filename${dupNames.length > 1 ? 's' : ''} in selection: ${dupNames.join(', ')}. Each file must be unique.`);
      return;
    }

    setPendingFiles(fileArr);
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      stageFiles(e.dataTransfer.files);
    },
    [stageFiles]
  );

  const onDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); };
  const onDragLeave = () => setIsDragging(false);

  // ── Toggle province ──────────────────────────────────────────────────────

  const toggleProvince = (clientName: string, province: string) => {
    setSelectedProvinces((prev) => {
      const curr = prev[clientName] ?? [];
      return {
        ...prev,
        [clientName]: curr.includes(province) ? curr.filter((p) => p !== province) : [...curr, province],
      };
    });
  };

  // ── Process ──────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (!parseResult || !controlMap || uploadedFiles.length === 0) return;

    setStage('processing');
    setErrorMsg(null);

    try {
      const res = await fetch('/api/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          rowHeaders: parseResult.rowHeaders,
          rowData: parseResult.rowData,
          reportDate: parseResult.mostRecentDateCol
            ? parseResult.mostRecentDateCol.replace(/\//g, '-')
            : new Date().toISOString().split('T')[0],
          mostRecentDateCol: parseResult.mostRecentDateCol ?? '',
          includeNegative,
          recipientMode,
          actionMode,
          selectedProvinces,
          selectedChannels,
        }),
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Server error ${res.status}: ${text.slice(0, 300)}`);
      }
      const data = await res.json();

      setProcessSummary((data as { summary: ProcessSummary }).summary);
      setStage('done');
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Processing failed');
      setStage('error');
    }
  };

  // ── Derived data ─────────────────────────────────────────────────────────

  const storeNameIdx = parseResult?.rowHeaders.indexOf('Store_Name') ?? -1;
  const l2Idx = parseResult?.rowHeaders.indexOf('Personnel_Level_2') ?? -1;

  const uniqueStores = parseResult && storeNameIdx >= 0
    ? new Set(parseResult.rowData.map((r) => r[storeNameIdx])).size
    : 0;

  const uniqueL2s = parseResult && l2Idx >= 0
    ? new Set(parseResult.rowData.map((r) => r[l2Idx]).filter(Boolean))
    : new Set<string>();

  const missingReps = controlMap && parseResult
    ? [...uniqueL2s].filter((name) => !controlMap[name.toLowerCase()])
    : [];

  const foundReps = controlMap && parseResult
    ? [...uniqueL2s].filter((name) => !!controlMap[name.toLowerCase()])
    : [];

  const reportDate = parseResult?.mostRecentDateCol
    ? parseResult.mostRecentDateCol.replace(/\//g, '-')
    : '—';

  const hasValidationIssues = parseResult?.files.some(
    (f) => f.missingFields.length > 0 || f.duplicateOf !== null
  );

  const totalSelectedProvinces = Object.values(selectedProvinces).reduce((s, a) => s + a.length, 0);

  // Channel mismatch: per-file, check if any Channel values in the data match selected channels
  const channelMismatches = parseResult && selectedChannels.length > 0
    ? parseResult.files.filter((f) => {
        if (f.channels.length === 0) return false; // no channel data — can't validate
        return !f.channels.some((c) => selectedChannels.some((sc) => sc.toLowerCase() === c.toLowerCase()));
      })
    : [];

  // ── Render ───────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-background text-foreground">
      {pendingFiles && (
        <ConfirmUploadModal
          onConfirm={() => { void handleFiles(pendingFiles); setPendingFiles(null); }}
          onCancel={() => setPendingFiles(null)}
        />
      )}

      {/* Header */}
      <header className="border-b border-border px-6 py-3 flex items-center justify-between sticky top-0 bg-card z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-1.5 h-8 bg-accent rounded" />
          <div>
            <h1 className="text-xl font-bold text-foreground">Phantom Consolidator</h1>
            <p className="text-muted text-xs">Multi-vendor phantom stock reporting &mdash; iRam</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
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
          {/* Admin link */}
          <Link href="/admin" className="text-xs text-muted hover:text-accent transition-colors">
            Admin
          </Link>
          {/* iRam logo */}
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/iram-logo.png" alt="iRam" className="h-9 w-auto object-contain" />
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8">

        {/* ── Section 1: Channel Selection ── */}
        <Section title="1 — Channel">
          <div className="flex items-center justify-between mb-3">
            <p className="text-muted text-xs">
              Select the channel(s) for this run. The app loads the matching control file(s) and routes uploads to the correct SharePoint folder.
            </p>
            <Link
              href="/admin"
              className="text-xs text-accent hover:underline flex-shrink-0 ml-4"
            >
              Manage channels &#8250;
            </Link>
          </div>

          {channelsLoading ? (
            <p className="text-muted text-xs animate-pulse">Loading channels...</p>
          ) : channels.length === 0 ? (
            <div className="bg-warning/10 border border-warning/30 text-warning rounded-lg px-4 py-3 text-sm">
              &#9888; No channels configured.{' '}
              <Link href="/admin" className="underline">Set them up in Admin</Link>.
            </div>
          ) : (
            <div className="flex flex-wrap gap-2">
              {channels.map((ch) => {
                const selected = selectedChannels.includes(ch);
                return (
                  <button
                    key={ch}
                    onClick={() =>
                      setSelectedChannels((prev) =>
                        prev.includes(ch) ? prev.filter((c) => c !== ch) : [...prev, ch]
                      )
                    }
                    className={`px-3 py-1.5 rounded-full text-sm font-medium border transition-colors ${
                      selected
                        ? 'border-accent text-accent'
                        : 'border-border text-muted hover:border-foreground/30 hover:text-foreground'
                    }`}
                    style={selected ? { background: 'rgba(121,190,67,0.12)' } : {}}
                  >
                    {selected ? '✓ ' : ''}{ch}
                  </button>
                );
              })}
            </div>
          )}

          {!channelsLoading && channels.length > 0 && selectedChannels.length === 0 && (
            <p className="text-warning text-xs mt-3">&#9888; Select at least one channel before uploading files.</p>
          )}
          {selectedChannels.length > 0 && (
            <p className="text-muted text-xs mt-3">
              Control file{selectedChannels.length > 1 ? 's' : ''}: {selectedChannels.map((c) => `"${c} - User Control File - Phantom Consolidator.xlsx"`).join(', ')}
            </p>
          )}
        </Section>

        {/* ── Section 2: Upload ── */}
        <Section title="2 — Upload Files">
          <div
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-lg p-10 text-center cursor-pointer transition-colors ${
              isDragging ? 'border-accent bg-accent/5' : 'border-border hover:border-accent/50'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              multiple
              className="hidden"
              onChange={(e) => { if (e.target.files) stageFiles(e.target.files); e.target.value = ''; }}
            />
            {isUploading ? (
              <p className="text-accent animate-pulse">{uploadProgress ?? 'Parsing files...'}</p>
            ) : (
              <>
                <p className="text-foreground font-medium">Drop Excel files here or click to browse</p>
                <p className="text-muted text-sm mt-1">Accepts multiple .xlsx / .xls files (one per vendor) &mdash; max 5 MB each</p>
              </>
            )}
          </div>

          {/* Pre-upload validation error */}
          {uploadError && (
            <div className="mt-3 bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm">
              &#9888; {uploadError}
            </div>
          )}

          {/* File list with per-file validation */}
          {parseResult && parseResult.files.length > 0 && (
            <div className="mt-4 space-y-2">
              {parseResult.files.map((f, i) => (
                <div key={i} className="bg-background border border-border rounded-lg px-4 py-3">
                  <div className="flex items-start justify-between gap-4">
                    <div className="min-w-0">
                      <span className="text-foreground font-medium text-sm block truncate">{f.fileName}</span>
                      <span className="text-accent text-xs font-mono">{f.clientName}</span>
                    </div>
                    <div className="text-right text-xs text-muted flex-shrink-0">
                      <div>{f.rowCount.toLocaleString()} rows</div>
                      <div>{f.dateColumns.length > 0 ? f.dateColumns.join(', ') : 'No date cols'}</div>
                    </div>
                  </div>
                  {/* Validation badges */}
                  {(f.missingFields.length > 0 || f.duplicateOf ||
                    (selectedChannels.length > 0 && f.channels.length > 0 &&
                     !f.channels.some((c) => selectedChannels.some((sc) => sc.toLowerCase() === c.toLowerCase())))
                  ) && (
                    <div className="mt-2 flex flex-wrap gap-2">
                      {f.missingFields.length > 0 && (
                        <span className="text-xs bg-danger/10 text-danger border border-danger/30 rounded px-2 py-0.5">
                          &#9888; Missing fields: {f.missingFields.join(', ')}
                        </span>
                      )}
                      {f.duplicateOf && (
                        <span className="text-xs bg-warning/10 text-warning border border-warning/30 rounded px-2 py-0.5">
                          &#9888; Same data as &ldquo;{f.duplicateOf}&rdquo; &mdash; possible duplicate upload
                        </span>
                      )}
                      {selectedChannels.length > 0 && f.channels.length > 0 &&
                       !f.channels.some((c) => selectedChannels.some((sc) => sc.toLowerCase() === c.toLowerCase())) && (
                        <span className="text-xs bg-danger/10 text-danger border border-danger/30 rounded px-2 py-0.5">
                          &#9888; Channel mismatch — file contains &ldquo;{f.channels.join(', ')}&rdquo; but selected: {selectedChannels.join(', ')}
                        </span>
                      )}
                    </div>
                  )}
                </div>
              ))}

              {/* Global duplicate-data warning */}
              {hasValidationIssues && (
                <div className="bg-warning/10 border border-warning/30 text-warning rounded-lg px-4 py-3 text-sm mt-2">
                  &#9888; One or more files have validation issues. Review before processing.
                </div>
              )}
            </div>
          )}
        </Section>

        {/* ── Section 3: Settings ── */}
        <Section title="3 — Settings">
          <div className="space-y-4">
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={includeNegative}
                onChange={(e) => setIncludeNegative(e.target.checked)}
                className="w-4 h-4"
                style={{ accentColor: '#79BE43' }}
              />
              <span className="text-foreground">Include NEGATIVE phantom rows</span>
            </label>

            <div>
              <p className="text-foreground text-sm mb-2 font-medium">Action:</p>
              <select
                value={actionMode}
                onChange={(e) => setActionMode(e.target.value as 'both' | 'sharepoint' | 'email')}
                className="bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
              >
                <option value="both">Email &amp; Save to SharePoint</option>
                <option value="sharepoint">Save to SharePoint only</option>
                <option value="email">Email only</option>
              </select>
            </div>

            <div className={actionMode === 'sharepoint' ? 'opacity-40 pointer-events-none' : ''}>
              <p className="text-foreground text-sm mb-2 font-medium">Send emails to:</p>
              <div className="flex gap-6">
                {(['l1', 'l2', 'both'] as const).map((mode) => (
                  <label key={mode} className="flex items-center gap-2 cursor-pointer">
                    <input
                      type="radio"
                      name="recipientMode"
                      value={mode}
                      checked={recipientMode === mode}
                      onChange={() => setRecipientMode(mode)}
                      style={{ accentColor: '#79BE43' }}
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

        {/* ── Section 4: Province Filter ── */}
        {stage === 'parsed' && parseResult && parseResult.files.length > 0 && (
          <Section title="4 — Province Filter">
            <p className="text-muted text-xs mb-4">
              Deselect provinces to exclude them from reports and emails. Defaults to all selected.
            </p>
            <div className="space-y-4">
              {parseResult.files.map((f, fi) => (
                <div key={fi} className="border border-border rounded-lg p-4">
                  <div className="flex items-center justify-between mb-3">
                    <span className="text-foreground font-medium text-sm">{f.clientName}</span>
                    <div className="flex items-center gap-3">
                      <button
                        onClick={() => setSelectedProvinces((prev) => ({ ...prev, [f.clientName]: [...f.provinces] }))}
                        className="text-xs text-accent hover:underline"
                      >
                        Select All
                      </button>
                      <span className="text-muted text-xs">&middot;</span>
                      <button
                        onClick={() => setSelectedProvinces((prev) => ({ ...prev, [f.clientName]: [] }))}
                        className="text-xs text-muted hover:text-foreground hover:underline"
                      >
                        Clear All
                      </button>
                    </div>
                  </div>
                  {f.provinces.length === 0 ? (
                    <p className="text-muted text-xs italic">No province data found in this file.</p>
                  ) : (
                    <div className="flex flex-wrap gap-2">
                      {f.provinces.map((prov) => {
                        const selected = (selectedProvinces[f.clientName] ?? []).includes(prov);
                        return (
                          <button
                            key={prov}
                            onClick={() => toggleProvince(f.clientName, prov)}
                            className={`px-3 py-1 rounded-full text-xs font-medium border transition-colors ${
                              selected
                                ? 'border-accent text-accent'
                                : 'border-border text-muted hover:border-foreground/30 hover:text-foreground'
                            }`}
                            style={selected ? { background: 'rgba(249,115,22,0.12)' } : {}}
                          >
                            {selected ? '✓ ' : ''}{prov}
                          </button>
                        );
                      })}
                    </div>
                  )}
                  {f.provinces.length > 0 && (selectedProvinces[f.clientName] ?? []).length === 0 && (
                    <p className="text-warning text-xs mt-2">&#9888; No provinces selected — this file will be skipped entirely.</p>
                  )}
                </div>
              ))}
            </div>
            {totalSelectedProvinces === 0 && (
              <div className="mt-3 bg-warning/10 border border-warning/30 text-warning rounded-lg px-4 py-3 text-sm">
                &#9888; No provinces selected across any file. Nothing will be processed.
              </div>
            )}
          </Section>
        )}

        {/* ── Section 5: Preview ── */}
        {stage === 'parsed' && parseResult && (
          <Section title="5 — Preview">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
              <Badge label="Phantom Rows ✓" value={parseResult.phantomCount.toLocaleString()} />
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
                      &#9888; {missingReps.length} rep(s) not in control file &mdash; emails skipped:
                    </p>
                    <p className="text-warning/80 text-xs mt-1">{missingReps.join(', ')}</p>
                  </div>
                )}
              </div>
            )}
          </Section>
        )}

        {/* ── Section 6: Process & Send ── */}
        {(stage === 'parsed' || stage === 'processing' || stage === 'done' || stage === 'error') && (
          <Section title="6 — Process &amp; Send">
            {stage !== 'done' && (
              <button
                onClick={() => { void handleProcess(); }}
                disabled={stage === 'processing' || !parseResult || !controlMap || totalSelectedProvinces === 0 || selectedChannels.length === 0}
                className="w-full font-bold py-3 px-6 rounded-lg transition-colors text-sm text-white disabled:opacity-50 disabled:cursor-not-allowed"
                style={{ background: stage === 'processing' || !parseResult || !controlMap || totalSelectedProvinces === 0 || selectedChannels.length === 0 ? undefined : '#79BE43' }}
                onMouseEnter={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#69a938'; }}
                onMouseLeave={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#79BE43'; }}
              >
                {stage === 'processing'
                  ? 'Processing...'
                  : actionMode === 'sharepoint'
                  ? 'Upload to SharePoint Only'
                  : actionMode === 'email'
                  ? 'Send Emails Only'
                  : 'Process & Send Reports'}
              </button>
            )}

            {stage === 'processing' && (
              <div className="mt-4 text-center">
                <p className="text-muted text-sm animate-pulse">
                  Building XLSX reports, queuing uploads and emails in background...
                </p>
                <p className="text-muted text-xs mt-1">This may take a moment for large batches.</p>
              </div>
            )}

            {errorMsg && (
              <div className="mt-4 bg-danger/10 border border-danger/30 text-danger rounded-lg px-4 py-3 text-sm">
                &#10007; {errorMsg}
              </div>
            )}

            {stage === 'done' && processSummary && (
              <div className="mt-4 space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <Badge label="Reports Queued" value={processSummary.stores} color="success" />
                  <Badge label="Reps Covered" value={processSummary.reps} color="success" />
                  <Badge label="Emails Sent" value={processSummary.emailsSent} color={processSummary.emailsSent > 0 ? 'success' : 'warning'} />
                  <Badge
                    label="Errors"
                    value={processSummary.errors.length}
                    color={processSummary.errors.length > 0 ? 'warning' : 'success'}
                  />
                </div>

                <div className="text-success text-sm">
                  &#10003;{' '}
                  {actionMode === 'sharepoint'
                    ? 'Reports uploading to iRAM SharePoint in background.'
                    : actionMode === 'email'
                    ? 'Emails dispatching in background.'
                    : 'Reports uploading to iRAM SharePoint + emails dispatching — both running independently in background.'}
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
                        </tr>
                      </thead>
                      <tbody>
                        {processSummary.storeResults.map((sr, i) => (
                          <tr key={i} className="border-b border-border/50">
                            <td className="py-2 px-3 text-foreground">{sr.storeName}</td>
                            <td className="py-2 px-3 text-foreground">{sr.l2Name}</td>
                            <td className="py-2 px-3 text-muted">{sr.l1Name}</td>
                            <td className="py-2 px-3 text-right text-foreground">{sr.rowCount}</td>
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
                    setUploadedFiles([]);
                    setControlMap(null);
                    setControlError(null);
                    setProcessSummary(null);
                    setErrorMsg(null);
                    setUploadError(null);
                    setSelectedProvinces({});
                    setSelectedChannels([]);
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

      {/* Footer */}
      <footer className="border-t border-border px-6 py-4 flex justify-end items-center gap-3">
        <span className="text-muted text-xs">Powered by</span>
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/oj-logo.png" alt="OuterJoin" className="h-5 w-auto object-contain opacity-75" />
      </footer>
    </div>
  );
}
