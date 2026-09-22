export type Status = "red" | "amber" | "green";
export type RevisionType = "Dokumen diterima" | "Hasil Review dikembalikan ke Process Owner";

export interface RevisionEntry {
  type: RevisionType;
  planStart: string;
  planDuration: number;      // working days
  actualStart: string | null;
  actualDuration: number | null; // working days
  note: string;
}

export type BentukDokumen = "HC" | "SC" | "";
export type KeteranganDRFI = "DR" | "FI" | "";

export interface Document {
  id: string;
  refCode: string;
  title: string;
  pic: string;
  picInitials: string;
  division: string;
  status: Status;
  progress: number;
  revision: string;
  lastUpdated: string;
  startDate: string;
  endDate: string;
  actualStartDate: string | null;
  actualEndDate: string | null;
  description: string;
  referenceDocuments: string[];
  revisionHistory: RevisionEntry[];
  // Additional metadata
  bentukDokumen: BentukDokumen;
  kategori: string;
  keteranganDRFI: KeteranganDRFI;
}

// ─── Calendar utilities ───────────────────────────────────────────────────────

export function getRealTodayIso(): string {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export const TODAY = getRealTodayIso();
export const YEAR = 2026;
export const YEAR_START = new Date(`${YEAR}-01-01`);

export const MONTHS_2026 = [
  { label: "Jan", days: 31 },
  { label: "Feb", days: 28 },
  { label: "Mar", days: 31 },
  { label: "Apr", days: 30 },
  { label: "May", days: 31 },
  { label: "Jun", days: 30 },
  { label: "Jul", days: 31 },
  { label: "Aug", days: 31 },
  { label: "Sep", days: 30 },
  { label: "Oct", days: 31 },
  { label: "Nov", days: 30 },
  { label: "Dec", days: 31 },
];

export const YEAR_DAYS = 365;

export function dayOffset(iso: string): number {
  const d = new Date(iso);
  return Math.round((d.getTime() - YEAR_START.getTime()) / 86400000);
}

export function quarter(iso: string): 1 | 2 | 3 | 4 {
  const m = new Date(iso).getMonth();
  if (m <= 2) return 1;
  if (m <= 5) return 2;
  if (m <= 8) return 3;
  return 4;
}

// ─── Working days ─────────────────────────────────────────────────────────────

function buildWorkingDays(year: number): string[] {
  const list: string[] = [];
  const d = new Date(`${year}-01-01`);
  while (d.getFullYear() === year) {
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) {
      list.push(d.toISOString().slice(0, 10));
    }
    d.setDate(d.getDate() + 1);
  }
  return list;
}

export const WORKING_DAYS = buildWorkingDays(YEAR);
const WD_MAP = new Map(WORKING_DAYS.map((d, i) => [d, i]));

/** Working day index (0-based). Weekends snap forward to next Monday. */
export function wdOffset(iso: string): number {
  const direct = WD_MAP.get(iso);
  if (direct !== undefined) return direct;
  // Clamp dates outside 2026
  if (iso < `${YEAR}-01-01`) return 0;
  if (iso > `${YEAR}-12-31`) return WORKING_DAYS.length - 1;
  // Weekend: walk forward up to 3 days
  const d = new Date(iso);
  for (let i = 1; i <= 3; i++) {
    d.setDate(d.getDate() + 1);
    const nxt = WD_MAP.get(d.toISOString().slice(0, 10));
    if (nxt !== undefined) return nxt;
  }
  return 0;
}

export const TODAY_WD = wdOffset(TODAY);
export const TODAY_DAY = dayOffset(TODAY);
export const TOTAL_WORKING_DAYS = WORKING_DAYS.length;

/** Add N working days to a start date, returns ISO string of the resulting working day. */
export function addWorkingDays(startIso: string, days: number): string {
  const startIdx = wdOffset(startIso);
  const endIdx = Math.min(startIdx + Math.max(0, days - 1), WORKING_DAYS.length - 1);
  return WORKING_DAYS[Math.max(0, endIdx)] ?? startIso;
}

/** Working days grouped by month index (0–11). */
export const WD_BY_MONTH: string[][] = MONTHS_2026.map((_, mi) => {
  const prefix = `${YEAR}-${String(mi + 1).padStart(2, "0")}-`;
  return WORKING_DAYS.filter((d) => d.startsWith(prefix));
});

/** Calendar week index (0-based), simple floor(dayOffset/7). */
export function weekOf(iso: string): number {
  return Math.max(0, Math.floor(dayOffset(iso) / 7));
}

export const TOTAL_WEEKS = Math.ceil(YEAR_DAYS / 7); // 53

// ─── Sample data ──────────────────────────────────────────────────────────────

export const DOCUMENTS: Document[] = [];

