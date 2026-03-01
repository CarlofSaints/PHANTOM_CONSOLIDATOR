// ── Raw parsed row from uploaded Excel files ────────────────────────────────

export interface RawRow {
  CLIENT: string;
  Product_Principle: string;
  Channel: string;
  Sub_Channel: string;
  Province: string;
  Personnel_Level_1: string;
  Personnel_Level_2: string;
  SiteCode: string;
  Store_Name: string;
  Store_Status: string;
  Product_Brand: string;
  Product_Sub_Category: string;
  Channel_ArticleCode: string;
  Client_Product_ID: string;
  Product_Description: string;
  Product_Status: string;
  Range_Indicator: string;
  Phantom_Indicator: string;
  // dynamic date columns come through as extra keys
  [key: string]: string;
}

// ── Per-file parse result ───────────────────────────────────────────────────

export interface ParsedFile {
  fileName: string;
  clientName: string;
  rows: RawRow[];
  dateColumns: string[];
}

// ── Control file lookup ─────────────────────────────────────────────────────

export interface RepInfo {
  l1Name: string;
  l1Email: string;
  l2Email: string;
}

export type ControlMap = Record<string, RepInfo>; // key = Personnel_Level_2 name

// ── Process request / response ──────────────────────────────────────────────

export interface ProcessSettings {
  rows: RawRow[];
  controlMap: ControlMap;
  reportDate: string;
  includeNegative: boolean;
  recipientMode: 'l1' | 'l2' | 'both';
}

export interface StoreResult {
  storeName: string;
  l2Name: string;
  l1Name: string;
  rowCount: number;
  webUrl: string;
  fileName: string;
  error?: string;
}

export interface ProcessSummary {
  stores: number;
  reps: number;
  emailsSent: number;
  errors: string[];
  storeResults: StoreResult[];
}
