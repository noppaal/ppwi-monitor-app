import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  DOCUMENTS, MONTHS_2026, TODAY, getRealTodayIso, YEAR_DAYS, TOTAL_WORKING_DAYS,
  WORKING_DAYS, WD_BY_MONTH, TOTAL_WEEKS,
  dayOffset, wdOffset, weekOf, quarter, TODAY_WD, TODAY_DAY, addWorkingDays,
  type Document, type Status, type RevisionType, type RevisionEntry,
  type BentukDokumen, type KeteranganDRFI,
} from "./data";
import { useLocalStorage } from "./useLocalStorage";

// ─── Constants ────────────────────────────────────────────────────────────────

const WD_W   = 14;   // px per working day
const WK_W   = 50;   // px per calendar week
const MO_W   = 80;   // px per month
const LABEL_W = 130;

export type TimelineMode = "day" | "week" | "month";

// ─── Helpers ──────────────────────────────────────────────────────────────────

function statusColor(s: Status) {
  if (s === "red")   return { bg: "var(--accent-red-bg)",   text: "var(--accent-red)",   dot: "#C8102E" };
  if (s === "amber") return { bg: "var(--accent-amber-bg)", text: "var(--accent-amber)", dot: "#D97706" };
  return               { bg: "var(--accent-green-bg)", text: "var(--accent-green)", dot: "#059669" };
}
function statusLabel(s: Status) {
  if (s === "red")   return "Perlu Tindakan";
  if (s === "amber") return "Dalam Review";
  return "Selesai";
}
function progressColor(p: number) {
  if (p === 100) return "#059669";
  if (p >= 50)   return "#D97706";
  return "#C8102E";
}
function fmt(iso: string) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return "—";
  return new Date(y, m - 1, d).toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
}
function fmtShort(iso: string) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return "—";
  return new Date(y, m - 1, d).toLocaleDateString("id-ID", { day: "2-digit", month: "short" });
}
function fmtHeaderDate(iso: string) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return "—";
  return new Date(y, m - 1, d).toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "short", year: "numeric" });
}
function isOverdue(doc: Document) {
  return doc.progress < 100 && doc.endDate < TODAY;
}
function calDays(start: string, end: string) {
  if (!start || !end) return 1;
  return Math.max(1, dayOffset(end) - dayOffset(start) + 1);
}

// Gantt pixel positions for each mode
function safeMonth(iso: string): number {
  const m = new Date(iso).getMonth();
  return isNaN(m) ? 0 : m;
}
function ganttPos(iso: string, mode: TimelineMode): number {
  if (!iso) return 0;
  if (mode === "day")   return (wdOffset(iso) || 0) * WD_W;
  if (mode === "week")  return (weekOf(iso)   || 0) * WK_W;
  return safeMonth(iso) * MO_W;
}
function ganttWidth(startIso: string, endIso: string, mode: TimelineMode): number {
  const minW = mode === "day" ? WD_W : mode === "week" ? WK_W : MO_W;
  if (!startIso || !endIso) return minW;
  if (mode === "day")  return Math.max(minW, ((wdOffset(endIso) || 0) - (wdOffset(startIso) || 0) + 1) * WD_W);
  if (mode === "week") return Math.max(minW, ((weekOf(endIso)   || 0) - (weekOf(startIso)   || 0) + 1) * WK_W);
  const ms = safeMonth(startIso), me = safeMonth(endIso);
  return Math.max(minW, (me - ms + 1) * MO_W);
}
function ganttTotal(mode: TimelineMode): number {
  if (mode === "day")  return TOTAL_WORKING_DAYS * WD_W;
  if (mode === "week") return TOTAL_WEEKS * WK_W;
  return 12 * MO_W;
}
function todayPos(mode: TimelineMode): number {
  if (mode === "day")  return TODAY_WD  * WD_W  + WD_W  / 2;
  if (mode === "week") return weekOf(TODAY) * WK_W + WK_W / 2;
  return new Date(TODAY).getMonth() * MO_W + MO_W / 2;
}

// ─── Badge colors ─────────────────────────────────────────────────────────────

const BADGE_COLORS: Record<string, string> = {
  AR: "#1D4ED8", NH: "#7C3AED", KS: "#059669", FM: "#C8102E", RI: "#D97706",
};
function getInitials(name: string) {
  return name.trim().split(/\s+/).slice(0, 2).map((w) => w[0]?.toUpperCase() ?? "").join("");
}

// ─── Shared atoms ─────────────────────────────────────────────────────────────

function PicBadge({ initials, size = "sm" }: { initials: string; size?: "sm" | "md" }) {
  const color = BADGE_COLORS[initials] ?? "#6B6860";
  const dim = size === "sm" ? 28 : 36;
  return (
    <div style={{
      width: dim, height: dim, borderRadius: "50%", background: color, color: "#fff",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: size === "sm" ? 11 : 13, fontWeight: 600,
      fontFamily: "var(--font-mono)", flexShrink: 0, letterSpacing: "0.04em",
    }}>{initials}</div>
  );
}

function StatusPill({ status }: { status: Status }) {
  const c = statusColor(status);
  return (
    <span style={{
      background: c.bg, color: c.text,
      fontSize: 11, fontWeight: 600, fontFamily: "var(--font-mono)",
      padding: "2px 8px", borderRadius: 3, letterSpacing: "0.05em",
      textTransform: "uppercase", whiteSpace: "nowrap",
    }}>{statusLabel(status)}</span>
  );
}

function ProgressBar({ value, showLabel = false }: { value: number; showLabel?: boolean }) {
  const color = progressColor(value);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div className="progress-bar" style={{ flex: 1 }}>
        <div className="progress-fill" style={{ width: `${value}%`, background: color }} />
      </div>
      {showLabel && (
        <span style={{ fontSize: 12, fontWeight: 700, color, fontFamily: "var(--font-mono)", minWidth: 32, textAlign: "right" }}>
          {value}%
        </span>
      )}
    </div>
  );
}

function MetricCard({ label, value, accent }: { label: string; value: number; accent?: string }) {
  return (
    <div style={{
      background: "var(--surface)", border: "1px solid var(--border)",
      borderRadius: 8, padding: "12px 14px", flex: 1, minWidth: 0,
    }}>
      <div style={{ fontSize: 22, fontWeight: 800, color: accent ?? "var(--text-primary)", lineHeight: 1, letterSpacing: "-0.03em" }}>
        {value}
      </div>
      <div style={{ fontSize: 11, color: "var(--text-secondary)", marginTop: 4, fontWeight: 500, lineHeight: 1.3 }}>
        {label}
      </div>
    </div>
  );
}

// ─── Document Card ────────────────────────────────────────────────────────────

function DocumentCard({ doc, onClick }: { doc: Document; onClick: () => void }) {
  const c = statusColor(doc.status);
  const overdue = isOverdue(doc);
  return (
    <button onClick={onClick} className="card-hover" style={{
      width: "100%", textAlign: "left",
      background: "var(--surface)", border: `1px solid ${overdue ? "var(--accent-red)" : "var(--border)"}`,
      borderRadius: 8, padding: "14px 16px",
      display: "flex", flexDirection: "column", gap: 10, cursor: "pointer",
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text-primary)", lineHeight: 1.35, marginBottom: 3 }}>
            {doc.title}
          </div>
          <div style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", letterSpacing: "0.04em" }}>
            {doc.refCode}
          </div>
        </div>
        <div style={{ width: 8, height: 8, borderRadius: "50%", background: c.dot, flexShrink: 0, marginTop: 4 }} />
      </div>
      <ProgressBar value={doc.progress} showLabel />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <PicBadge initials={doc.picInitials} />
          <div>
            <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-primary)" }}>{doc.pic}</div>
            <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>{doc.division}</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <StatusPill status={doc.status} />
          <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>{doc.revision}</span>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 10, color: overdue ? "var(--accent-red)" : "var(--text-tertiary)", fontFamily: "var(--font-mono)", fontWeight: overdue ? 700 : 400 }}>
          {overdue ? "⚠ Overdue · " : ""}Due {fmtShort(doc.endDate)} · {calDays(doc.startDate, doc.endDate)} hari
        </span>
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
          <path d="M5 3l4 4-4 4" stroke="var(--text-tertiary)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      </div>
    </button>
  );
}

// ─── Gantt Row ────────────────────────────────────────────────────────────────

function GanttRow({ doc, mode, expanded, onToggle }: {
  doc: Document; mode: TimelineMode; expanded: boolean; onToggle: () => void;
}) {
  const c = statusColor(doc.status);
  const total = ganttTotal(mode);
  const pLeft  = ganttPos(doc.startDate, mode);
  const pWidth = ganttWidth(doc.startDate, doc.endDate, mode);
  const aLeft  = doc.actualStartDate ? ganttPos(doc.actualStartDate, mode) : null;
  const aEnd   = doc.actualEndDate ?? TODAY;
  const aWidth = aLeft !== null ? ganttWidth(doc.actualStartDate!, aEnd, mode) : null;
  const tPos   = todayPos(mode);

  return (
    <>
      {/* Document row */}
      <div
        onClick={onToggle}
        className="card-hover"
        style={{
          display: "flex", alignItems: "center", minHeight: 52,
          borderBottom: expanded ? "none" : "1px solid var(--border)",
          cursor: "pointer",
          background: expanded ? "var(--surface-2)" : "var(--surface)",
        }}
      >
        <div style={{
          width: LABEL_W, flexShrink: 0, padding: "8px 10px",
          borderRight: "1px solid var(--border)",
          background: expanded ? "var(--surface-2)" : "var(--surface)",
          position: "sticky", left: 0, zIndex: 2,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            {/* Expand chevron */}
            <svg
              width="10" height="10" viewBox="0 0 10 10" fill="none"
              style={{ flexShrink: 0, transform: expanded ? "rotate(90deg)" : "rotate(0deg)", transition: "transform 0.18s ease" }}
            >
              <path d="M3 2l4 3-4 3" stroke="var(--text-tertiary)" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: c.dot, flexShrink: 0 }} />
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text-primary)", lineHeight: 1.3, wordBreak: "break-word" }}>
                {doc.title}
              </div>
              <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", marginTop: 1 }}>
                {doc.refCode.slice(5)}
              </div>
            </div>
          </div>
        </div>
        <div style={{ position: "relative", height: 52, width: total, flexShrink: 0 }}>
          <div style={{
            position: "absolute", top: 11, left: pLeft, width: pWidth, height: 11,
            borderRadius: 3, background: "#E0DED9", border: "1px solid var(--border-strong)",
          }} />
          {aLeft !== null && aWidth !== null && (
            <div style={{
              position: "absolute", top: 28, left: aLeft, width: aWidth, height: 11,
              borderRadius: 3, background: c.dot, opacity: 0.82,
            }} />
          )}
          <div style={{
            position: "absolute", top: 0, bottom: 0, left: tPos,
            width: 1.5, background: "#C8102E", opacity: 0.65, zIndex: 1, pointerEvents: "none",
          }} />
        </div>
      </div>

      {/* Inline history expansion */}
      {expanded && (
        <div style={{ borderBottom: "2px solid var(--border)" }}>
          {doc.revisionHistory.length === 0 ? (
            <div style={{ display: "flex" }}>
              <div style={{ width: LABEL_W, flexShrink: 0, padding: "10px 10px", borderRight: "1px solid var(--border)", background: "var(--surface-2)", position: "sticky", left: 0, zIndex: 2 }}>
                <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>Belum ada riwayat</div>
              </div>
              <div style={{ width: total, flexShrink: 0 }} />
            </div>
          ) : (
            doc.revisionHistory.map((r, ri) => {
              const isDiterima = r.type === "Dokumen diterima";
              const hAccent = isDiterima ? "#1D4ED8" : "#D97706";
              const planEnd = addWorkingDays(r.planStart, r.planDuration);
              const actualEnd = r.actualStart && r.actualDuration
                ? addWorkingDays(r.actualStart, r.actualDuration)
                : (r.actualStart ? TODAY : null);

              const hPLeft  = ganttPos(r.planStart, mode);
              const hPWidth = ganttWidth(r.planStart, planEnd, mode);
              const hALeft  = r.actualStart ? ganttPos(r.actualStart, mode) : null;
              const hAWidth = hALeft !== null && actualEnd ? ganttWidth(r.actualStart!, actualEnd, mode) : null;

              return (
                <div key={ri} style={{ display: "flex", alignItems: "center", minHeight: 46, borderTop: "1px solid var(--border)" }}>
                  <div style={{
                    width: LABEL_W, flexShrink: 0, padding: "6px 10px 6px 22px",
                    borderRight: "1px solid var(--border)", background: "var(--surface-2)",
                    position: "sticky", left: 0, zIndex: 2,
                  }}>
                    <div style={{
                      display: "inline-block", padding: "1px 6px", borderRadius: 3,
                      background: isDiterima ? "var(--accent-blue-bg)" : "var(--accent-amber-bg)",
                      marginBottom: 2,
                    }}>
                      <span style={{ fontSize: 8, fontWeight: 700, color: hAccent, fontFamily: "var(--font-mono)", letterSpacing: "0.04em" }}>
                        {isDiterima ? "DITERIMA" : "DIKEMBALIKAN"}
                      </span>
                    </div>
                    <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>
                      P: {r.planDuration}hk
                      {r.actualDuration != null ? ` · A: ${r.actualDuration}hk` : ""}
                    </div>
                  </div>
                  <div style={{ position: "relative", height: 46, width: total, flexShrink: 0 }}>
                    {/* Plan bar */}
                    <div style={{
                      position: "absolute", top: 9, left: hPLeft, width: hPWidth, height: 10,
                      borderRadius: 3, background: "#E0DED9", border: "1px solid var(--border-strong)",
                    }} />
                    {/* Actual bar */}
                    {hALeft !== null && hAWidth !== null && (
                      <div style={{
                        position: "absolute", top: 25, left: hALeft, width: hAWidth, height: 10,
                        borderRadius: 3, background: hAccent, opacity: 0.75,
                      }} />
                    )}
                    {/* Today line */}
                    <div style={{
                      position: "absolute", top: 0, bottom: 0, left: tPos,
                      width: 1.5, background: "#C8102E", opacity: 0.5, zIndex: 1, pointerEvents: "none",
                    }} />
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}
    </>
  );
}

// ─── Gantt Header ─────────────────────────────────────────────────────────────

function GanttHeader({ mode }: { mode: TimelineMode }) {
  if (mode === "day") {
    return (
      <>
        {/* Month row */}
        <div style={{ display: "flex", borderBottom: "1px solid var(--border)", position: "sticky", top: 0, zIndex: 10, background: "var(--surface)" }}>
          <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 12 }}>
            <div style={{ padding: "5px 10px" }}>
              <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase" }}>Dokumen</span>
            </div>
          </div>
          {WD_BY_MONTH.map((days, mi) => {
            const isCurrentMonth = new Date(TODAY).getMonth() === mi;
            return (
              <div key={mi} style={{
                width: days.length * WD_W, flexShrink: 0,
                borderRight: "1px solid var(--border)", padding: "5px 6px",
                background: isCurrentMonth ? "var(--accent-red-bg)" : "transparent",
              }}>
                <span style={{
                  fontSize: 10, fontWeight: 700, fontFamily: "var(--font-mono)",
                  color: isCurrentMonth ? "var(--accent-red)" : "var(--text-primary)",
                }}>
                  {MONTHS_2026[mi].label}
                </span>
                <span style={{ fontSize: 9, color: "var(--text-tertiary)", marginLeft: 4, fontFamily: "var(--font-mono)" }}>
                  {days.length}h
                </span>
              </div>
            );
          })}
        </div>
        {/* Working day date row */}
        <div style={{ display: "flex", borderBottom: "1px solid var(--border)", background: "var(--bg)", position: "sticky", top: 27, zIndex: 9 }}>
          <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 11 }} />
          {WORKING_DAYS.map((iso, i) => {
            const isToday = iso === TODAY;
            const dayNum = parseInt(iso.slice(8));
            return (
              <div key={iso} style={{
                width: WD_W, flexShrink: 0, textAlign: "center", padding: "2px 0",
                background: isToday ? "var(--accent-red-bg)" : "transparent",
                borderRight: (i + 1) % 5 === 0 ? "1px solid var(--border)" : "none",
              }}>
                <span style={{
                  fontSize: 7, display: "block",
                  color: isToday ? "var(--accent-red)" : "var(--text-tertiary)",
                  fontFamily: "var(--font-mono)", fontWeight: isToday ? 700 : 400,
                  lineHeight: 1.4,
                }}>
                  {isToday ? "▼" : dayNum}
                </span>
              </div>
            );
          })}
        </div>
        {/* Sequential working day number row */}
        <div style={{ display: "flex", borderBottom: "2px solid var(--border)", background: "var(--surface-2)", position: "sticky", top: 47, zIndex: 8 }}>
          <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 11,
            display: "flex", alignItems: "center", padding: "0 10px" }}>
            <span style={{ fontSize: 8, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase", letterSpacing: "0.05em" }}>HK</span>
          </div>
          {WORKING_DAYS.map((iso, i) => {
            const isToday = iso === TODAY;
            const seqNum = i + 1;
            return (
              <div key={iso} style={{
                width: WD_W, flexShrink: 0, textAlign: "center", padding: "2px 0",
                background: isToday ? "var(--accent-red-bg)" : "transparent",
                borderRight: (i + 1) % 5 === 0 ? "1px solid var(--border)" : "none",
              }}>
                <span style={{
                  fontSize: 7, display: "block",
                  color: isToday ? "var(--accent-red)" : "var(--text-secondary)",
                  fontFamily: "var(--font-mono)", fontWeight: isToday ? 800 : 500,
                  lineHeight: 1.4,
                }}>
                  {seqNum}
                </span>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  if (mode === "week") {
    // Build month spans over weeks
    const monthSpans: { label: string; startWeek: number; endWeek: number }[] = [];
    MONTHS_2026.forEach((m, mi) => {
      const prefix = `${2026}-${String(mi + 1).padStart(2, "0")}-`;
      const firstDay = prefix + "01";
      const lastDayNum = m.days;
      const lastDay = prefix + String(lastDayNum).padStart(2, "0");
      const sw = weekOf(firstDay), ew = weekOf(lastDay);
      const last = monthSpans[monthSpans.length - 1];
      if (last && last.endWeek === sw) {
        monthSpans.push({ label: m.label, startWeek: sw + 1, endWeek: ew });
      } else {
        monthSpans.push({ label: m.label, startWeek: sw, endWeek: ew });
      }
    });

    return (
      <>
        <div style={{ display: "flex", borderBottom: "1px solid var(--border)", position: "sticky", top: 0, zIndex: 10, background: "var(--surface)" }}>
          <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 12 }}>
            <div style={{ padding: "5px 10px" }}>
              <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase" }}>Dokumen</span>
            </div>
          </div>
          {monthSpans.map((ms, i) => {
            const spanWeeks = ms.endWeek - ms.startWeek + 1;
            const isCurrentMonth = new Date(TODAY).getMonth() === i;
            return (
              <div key={i} style={{
                width: spanWeeks * WK_W, flexShrink: 0,
                borderRight: "1px solid var(--border)", padding: "5px 6px",
                background: isCurrentMonth ? "var(--accent-red-bg)" : "transparent",
              }}>
                <span style={{
                  fontSize: 10, fontWeight: 700, fontFamily: "var(--font-mono)",
                  color: isCurrentMonth ? "var(--accent-red)" : "var(--text-primary)",
                }}>
                  {ms.label}
                </span>
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", borderBottom: "2px solid var(--border)", background: "var(--bg)", position: "sticky", top: 27, zIndex: 9 }}>
          <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 11 }} />
          {Array.from({ length: TOTAL_WEEKS }, (_, i) => {
            const isCurrent = weekOf(TODAY) === i;
            return (
              <div key={i} style={{
                width: WK_W, flexShrink: 0, textAlign: "center", padding: "2px 0",
                background: isCurrent ? "#FFF0F2" : "transparent",
                borderRight: "1px solid var(--border)",
              }}>
                <span style={{
                  fontSize: 8, color: isCurrent ? "#C8102E" : "var(--text-tertiary)",
                  fontFamily: "var(--font-mono)", fontWeight: isCurrent ? 700 : 400,
                }}>
                  {isCurrent ? "▼" : `W${i + 1}`}
                </span>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // Month mode
  return (
    <div style={{ display: "flex", borderBottom: "2px solid var(--border)", position: "sticky", top: 0, zIndex: 10, background: "var(--surface)" }}>
      <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid var(--border)", background: "var(--surface)", position: "sticky", left: 0, zIndex: 12 }}>
        <div style={{ padding: "8px 10px" }}>
          <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase" }}>Dokumen</span>
        </div>
      </div>
      {MONTHS_2026.map((m, mi) => {
        const isCurrentMonth = new Date(TODAY).getMonth() === mi;
        return (
          <div key={mi} style={{
            width: MO_W, flexShrink: 0, padding: "8px 6px",
            borderRight: "1px solid var(--border)",
            background: isCurrentMonth ? "var(--accent-red-bg)" : "transparent",
            display: "flex", flexDirection: "column", justifyContent: "center",
          }}>
            <span style={{
              fontSize: 11, fontWeight: 700, fontFamily: "var(--font-mono)",
              color: isCurrentMonth ? "var(--accent-red)" : "var(--text-primary)",
              display: "block",
            }}>
              {isCurrentMonth ? `▼ ${m.label}` : m.label}
            </span>
          </div>
        );
      })}
    </div>
  );
}

// ─── Detail Sheet ─────────────────────────────────────────────────────────────

const FIELD: React.CSSProperties = {
  width: "100%", padding: "10px 12px", borderRadius: 8,
  border: "1px solid var(--border)", background: "var(--surface)",
  fontSize: 13, color: "var(--text-primary)", fontFamily: "var(--font-sans)",
  outline: "none", boxSizing: "border-box",
};
const LBL: React.CSSProperties = {
  fontSize: 11, fontWeight: 700, color: "var(--text-secondary)",
  letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5, display: "block",
};

function DetailSheet({ doc, onClose, onSave }: {
  doc: Document;
  onClose: () => void;
  onSave: (updated: Document) => void;
}) {
  const [editMode, setEditMode] = useState(false);
  const [showAddHistory, setShowAddHistory] = useState(false);

  // Editable fields
  const [title, setTitle]           = useState(doc.title);
  const [refCode, setRefCode]       = useState(doc.refCode);
  const [pic, setPic]               = useState(doc.pic);
  const [division, setDivision]     = useState(doc.division);
  const [revision, setRevision]     = useState(doc.revision);
  const [description, setDescription] = useState(doc.description);
  const [status, setStatus]         = useState<Status>(doc.status);
  const [progress, setProgress]     = useState(doc.progress);
  const [startDate, setStartDate]   = useState(doc.startDate);
  const [endDate, setEndDate]       = useState(doc.endDate);
  const [actualStart, setActualStart] = useState(doc.actualStartDate ?? "");
  const [actualEnd, setActualEnd]   = useState(doc.actualEndDate ?? "");
  const [refDocs, setRefDocs]       = useState(doc.referenceDocuments.join(", "));

  // New / edit history entry
  const [hType, setHType]           = useState<RevisionType>("Dokumen diterima");
  const [hPlanStart, setHPlanStart] = useState(TODAY);
  const [hPlanDur, setHPlanDur]     = useState("");
  const [hActualStart, setHActualStart] = useState("");
  const [hActualDur, setHActualDur] = useState("");
  const [hNote, setHNote]           = useState("");

  const [editingIdx, setEditingIdx]     = useState<number | null>(null);
  const [deleteIdx, setDeleteIdx]       = useState<number | null>(null);

  // New metadata fields
  const [bentukDokumen, setBentukDokumen] = useState<BentukDokumen>(doc.bentukDokumen ?? "");
  const [kategori, setKategori]           = useState(doc.kategori ?? "");
  const [keteranganDRFI, setKeteranganDRFI] = useState<KeteranganDRFI>(doc.keteranganDRFI ?? "");

  const overdue = isOverdue(doc);

  function saveEdit() {
    const updated: Document = {
      ...doc,
      title: title.trim() || doc.title,
      refCode: refCode.trim() || doc.refCode,
      pic: pic.trim() || doc.pic,
      picInitials: getInitials(pic.trim() || doc.pic),
      division: division.trim() || doc.division,
      revision: revision.trim() || doc.revision,
      description: description.trim(),
      status,
      progress,
      startDate,
      endDate,
      actualStartDate: actualStart || null,
      actualEndDate: actualEnd || null,
      referenceDocuments: refDocs.split(",").map((r) => r.trim()).filter(Boolean),
      bentukDokumen,
      kategori: kategori.trim(),
      keteranganDRFI,
      lastUpdated: TODAY,
    };
    onSave(updated);
    setEditMode(false);
  }

  function resetHistoryForm() {
    setHType("Dokumen diterima");
    setHPlanStart(TODAY); setHPlanDur(""); setHActualStart(""); setHActualDur(""); setHNote("");
  }

  function addHistory() {
    if (!hPlanStart || !hPlanDur) return;
    const entry: RevisionEntry = {
      type: hType,
      planStart: hPlanStart,
      planDuration: Number(hPlanDur),
      actualStart: hActualStart || null,
      actualDuration: hActualDur ? Number(hActualDur) : null,
      note: hNote.trim(),
    };
    onSave({ ...doc, revisionHistory: [entry, ...doc.revisionHistory], lastUpdated: TODAY });
    resetHistoryForm();
    setShowAddHistory(false);
  }

  function startEditHistory(i: number) {
    const r = doc.revisionHistory[i];
    setHType(r.type);
    setHPlanStart(r.planStart);
    setHPlanDur(String(r.planDuration));
    setHActualStart(r.actualStart ?? "");
    setHActualDur(r.actualDuration != null ? String(r.actualDuration) : "");
    setHNote(r.note);
    setEditingIdx(i);
    setShowAddHistory(false);
  }

  function saveEditHistory() {
    if (editingIdx === null || !hPlanStart || !hPlanDur) return;
    const updated = doc.revisionHistory.map((r, i) =>
      i === editingIdx
        ? { type: hType, planStart: hPlanStart, planDuration: Number(hPlanDur), actualStart: hActualStart || null, actualDuration: hActualDur ? Number(hActualDur) : null, note: hNote.trim() }
        : r
    );
    onSave({ ...doc, revisionHistory: updated, lastUpdated: TODAY });
    resetHistoryForm();
    setEditingIdx(null);
  }

  function confirmDelete(i: number) { setDeleteIdx(i); }

  function doDelete() {
    if (deleteIdx === null) return;
    const updated = doc.revisionHistory.filter((_, i) => i !== deleteIdx);
    onSave({ ...doc, revisionHistory: updated, lastUpdated: TODAY });
    setDeleteIdx(null);
  }

  return (
    <>
      <div className="sheet-overlay" onClick={onClose} />
      <div className="sheet-panel scrollbar-hide">
        {/* Handle */}
        <div style={{ display: "flex", justifyContent: "center", padding: "12px 0 4px" }}>
          <div style={{ width: 36, height: 4, borderRadius: 2, background: "var(--border-strong)" }} />
        </div>

        {/* Header */}
        <div style={{ padding: "12px 20px 0", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div style={{ flex: 1, paddingRight: 12 }}>
            {editMode
              ? <input value={title} onChange={(e) => setTitle(e.target.value)} style={{ ...FIELD, fontSize: 16, fontWeight: 800 }} />
              : <>
                  <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 4 }}>
                    {doc.refCode}
                  </div>
                  <h2 style={{ fontSize: 18, fontWeight: 800, color: "var(--text-primary)", lineHeight: 1.25, letterSpacing: "-0.02em", margin: 0 }}>
                    {doc.title}
                  </h2>
                </>
            }
          </div>
          <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
            <button onClick={() => { if (editMode) saveEdit(); else setEditMode(true); }} style={{
              padding: "6px 14px", borderRadius: 20,
              background: editMode ? "var(--text-primary)" : "var(--surface-2)",
              border: `1px solid ${editMode ? "var(--text-primary)" : "var(--border)"}`,
              fontSize: 11, fontWeight: 700, color: editMode ? "var(--text-inverse)" : "var(--text-secondary)",
              cursor: "pointer",
            }}>
              {editMode ? "Simpan" : "Edit"}
            </button>
            <button onClick={() => { if (editMode) setEditMode(false); else onClose(); }} style={{
              background: "var(--surface-2)", border: "none", borderRadius: "50%",
              width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
            }}>
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                <path d="M2 2l10 10M12 2L2 12" stroke="var(--text-secondary)" strokeWidth="1.8" strokeLinecap="round" />
              </svg>
            </button>
          </div>
        </div>

        {!editMode && (
          <div style={{ padding: "10px 20px 0", display: "flex", gap: 8, flexWrap: "wrap" }}>
            <StatusPill status={doc.status} />
            {overdue && (
              <span style={{ fontSize: 11, color: "var(--accent-red)", fontFamily: "var(--font-mono)", background: "var(--accent-red-bg)", padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>
                OVERDUE
              </span>
            )}
            <span style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", background: "var(--surface-2)", padding: "2px 8px", borderRadius: 3 }}>
              {doc.division}
            </span>
            <span style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", background: "var(--surface-2)", padding: "2px 8px", borderRadius: 3 }}>
              {doc.revision}
            </span>
          </div>
        )}

        <div style={{ height: 1, background: "var(--border)", margin: "14px 0 0" }} />

        {/* ── EDIT MODE ── */}
        {editMode && (
          <div style={{ padding: "16px 20px", display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text-secondary)", letterSpacing: "0.08em", textTransform: "uppercase" }}>
              Edit Dokumen
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div><label style={LBL}>Kode Referensi</label>
                <input value={refCode} onChange={(e) => setRefCode(e.target.value)} style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
              <div><label style={LBL}>Revisi</label>
                <input value={revision} onChange={(e) => setRevision(e.target.value)} style={{ ...FIELD, fontSize: 12 }} /></div>
            </div>

            <div><label style={LBL}>PIC</label>
              <input value={pic} onChange={(e) => setPic(e.target.value)} style={FIELD} /></div>

            <div><label style={LBL}>Divisi</label>
              <input value={division} onChange={(e) => setDivision(e.target.value)} style={FIELD} /></div>

            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <label style={{ ...LBL, margin: 0 }}>Progress</label>
                <span style={{ fontSize: 16, fontWeight: 800, color: progressColor(progress), fontFamily: "var(--font-mono)" }}>{progress}%</span>
              </div>
              <input type="range" min={0} max={100} step={1} value={progress}
                onChange={(e) => setProgress(Number(e.target.value))}
                style={{ accentColor: progressColor(progress) }} />
            </div>

            <div>
              <label style={LBL}>Status</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                {(["red", "amber", "green"] as Status[]).map((s) => {
                  const c = statusColor(s); const active = status === s;
                  return (
                    <button key={s} onClick={() => setStatus(s)} style={{
                      padding: "8px", borderRadius: 6, cursor: "pointer",
                      border: `1.5px solid ${active ? c.dot : "var(--border)"}`,
                      background: active ? c.bg : "var(--surface)", transition: "all 0.12s",
                    }}>
                      <div style={{ width: 7, height: 7, borderRadius: "50%", background: c.dot, margin: "0 auto 3px" }} />
                      <div style={{ fontSize: 10, fontWeight: 600, color: active ? c.text : "var(--text-secondary)", fontFamily: "var(--font-mono)" }}>
                        {s === "red" ? "Action" : s === "amber" ? "Review" : "Done"}
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div><label style={LBL}>Tgl Mulai (Rencana)</label>
                <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
              <div><label style={LBL}>Tgl Selesai (Rencana)</label>
                <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
              <div><label style={LBL}>Tgl Mulai (Aktual)</label>
                <input type="date" value={actualStart} onChange={(e) => setActualStart(e.target.value)} style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
              <div><label style={LBL}>Tgl Selesai (Aktual)</label>
                <input type="date" value={actualEnd} onChange={(e) => setActualEnd(e.target.value)} style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
            </div>

            <div><label style={LBL}>Deskripsi</label>
              <textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={3}
                style={{ ...FIELD, resize: "none", lineHeight: 1.6 }} /></div>

            <div><label style={LBL}>Dokumen Referensi</label>
              <input value={refDocs} onChange={(e) => setRefDocs(e.target.value)}
                placeholder="Pisahkan dengan koma" style={FIELD} /></div>

            {/* New metadata fields */}
            <div>
              <label style={LBL}>Bentuk Dokumen</label>
              <div style={{ display: "flex", gap: 8 }}>
                {(["HC", "SC", ""] as BentukDokumen[]).map((v) => (
                  <button key={v} onClick={() => setBentukDokumen(v)} style={{
                    flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer",
                    border: `1.5px solid ${bentukDokumen === v ? "var(--text-primary)" : "var(--border)"}`,
                    background: bentukDokumen === v ? "var(--text-primary)" : "var(--surface)",
                    color: bentukDokumen === v ? "var(--text-inverse)" : "var(--text-secondary)",
                    fontSize: 12, fontWeight: 700, fontFamily: "var(--font-mono)",
                    transition: "all 0.12s",
                  }}>{v === "" ? "—" : v}</button>
                ))}
              </div>
            </div>

            <div><label style={LBL}>Kategori (PPWI)</label>
              <input value={kategori} onChange={(e) => setKategori(e.target.value)}
                placeholder="Contoh: PPWI-ENG" style={{ ...FIELD, fontFamily: "var(--font-mono)" }} /></div>

            <div>
              <label style={LBL}>Keterangan DR / FI</label>
              <div style={{ display: "flex", gap: 8 }}>
                {(["DR", "FI", ""] as KeteranganDRFI[]).map((v) => (
                  <button key={v} onClick={() => setKeteranganDRFI(v)} style={{
                    flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer",
                    border: `1.5px solid ${keteranganDRFI === v ? "var(--text-primary)" : "var(--border)"}`,
                    background: keteranganDRFI === v ? "var(--text-primary)" : "var(--surface)",
                    color: keteranganDRFI === v ? "var(--text-inverse)" : "var(--text-secondary)",
                    fontSize: 12, fontWeight: 700, fontFamily: "var(--font-mono)",
                    transition: "all 0.12s",
                  }}>{v === "" ? "—" : v}</button>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── VIEW MODE ── */}
        {!editMode && (
          <>
            <div style={{ padding: "16px 20px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text-secondary)", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 14 }}>
                Detail Dokumen
              </div>
              {doc.description && (
                <div style={{ fontSize: 13, color: "var(--text-secondary)", lineHeight: 1.6, marginBottom: 16 }}>
                  {doc.description}
                </div>
              )}
              {/* Total Day banner */}
              {(() => {
                const totalPlanDays = doc.revisionHistory.reduce((s, r) => s + r.planDuration, 0);
                const totalActualDays = doc.revisionHistory.reduce((s, r) => s + (r.actualDuration ?? 0), 0);
                const hasSomeActual = doc.revisionHistory.some((r) => r.actualDuration != null);
                return (
                  <div style={{
                    display: "flex", gap: 8, marginBottom: 16,
                    background: "var(--surface-2)", borderRadius: 10, padding: "12px 14px",
                    border: "1px solid var(--border)",
                  }}>
                    <div style={{ flex: 1, textAlign: "center" }}>
                      <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>
                        Total Hari (Rencana)
                      </div>
                      <div style={{ fontSize: 20, fontWeight: 800, color: "var(--text-primary)", letterSpacing: "-0.03em" }}>
                        {totalPlanDays}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", marginTop: 2 }}>hari kerja</div>
                    </div>
                    {hasSomeActual && (
                      <>
                        <div style={{ width: 1, background: "var(--border)" }} />
                        <div style={{ flex: 1, textAlign: "center" }}>
                          <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>
                            Total Hari (Aktual)
                          </div>
                          <div style={{ fontSize: 20, fontWeight: 800, color: progressColor(doc.progress), letterSpacing: "-0.03em" }}>
                            {totalActualDays}
                          </div>
                          <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", marginTop: 2 }}>hari kerja</div>
                        </div>
                      </>
                    )}
                    <div style={{ width: 1, background: "var(--border)" }} />
                    <div style={{ flex: 1, textAlign: "center" }}>
                      <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>
                        Durasi Kalender
                      </div>
                      <div style={{ fontSize: 20, fontWeight: 800, color: "var(--text-primary)", letterSpacing: "-0.03em" }}>
                        {calDays(doc.startDate, doc.endDate)}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", marginTop: 2 }}>hari</div>
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
                {([
                  ["PIC", doc.pic],
                  ["Divisi", doc.division],
                  ["Tgl Mulai", fmt(doc.startDate)],
                  ["Tenggat", fmt(doc.endDate)],
                  ["Diperbarui", fmt(doc.lastUpdated)],
                  ...(doc.actualStartDate ? [["Aktual Mulai", fmt(doc.actualStartDate)]] : []),
                  ...(doc.actualEndDate   ? [["Aktual Selesai", fmt(doc.actualEndDate)]] : []),
                  ...(doc.bentukDokumen   ? [["Bentuk Dokumen", doc.bentukDokumen]] : []),
                  ...(doc.kategori        ? [["Kategori", doc.kategori]] : []),
                  ...(doc.keteranganDRFI  ? [["DR / FI", doc.keteranganDRFI]] : []),
                ] as [string, string][]).map(([k, v]) => (
                  <div key={k}>
                    <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 2 }}>{k}</div>
                    <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-primary)" }}>{v}</div>
                  </div>
                ))}
              </div>

              {doc.referenceDocuments.length > 0 && (
                <div>
                  <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>
                    Dokumen Referensi
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    {doc.referenceDocuments.map((ref) => (
                      <div key={ref} style={{
                        display: "flex", alignItems: "center", gap: 8,
                        padding: "8px 10px", borderRadius: 6, background: "var(--surface-2)",
                      }}>
                        <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                          <rect x="1" y="1" width="10" height="10" rx="1.5" stroke="var(--text-tertiary)" strokeWidth="1.2" />
                          <path d="M3.5 4.5h5M3.5 6.5h5M3.5 8.5h3" stroke="var(--text-tertiary)" strokeWidth="1" strokeLinecap="round" />
                        </svg>
                        <span style={{ fontSize: 11, color: "var(--text-secondary)", fontFamily: "var(--font-mono)" }}>{ref}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div style={{ height: 1, background: "var(--border)" }} />

            {/* ── RIWAYAT ── */}
            <div style={{ padding: "16px 20px 32px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text-secondary)", letterSpacing: "0.08em", textTransform: "uppercase" }}>
                  Riwayat
                </div>
                <button onClick={() => setShowAddHistory(!showAddHistory)} style={{
                  display: "flex", alignItems: "center", gap: 5,
                  padding: "5px 12px", borderRadius: 20,
                  background: showAddHistory ? "var(--text-primary)" : "var(--surface-2)",
                  border: `1px solid ${showAddHistory ? "var(--text-primary)" : "var(--border)"}`,
                  fontSize: 11, fontWeight: 700,
                  color: showAddHistory ? "var(--text-inverse)" : "var(--text-secondary)",
                  cursor: "pointer",
                }}>
                  <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
                    <path d="M5.5 1v9M1 5.5h9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
                  </svg>
                  Tambah
                </button>
              </div>

              {/* Add history form */}
              {showAddHistory && (
                <div style={{
                  background: "var(--surface-2)", borderRadius: 10,
                  padding: 14, marginBottom: 20,
                  border: "1px solid var(--border)",
                  display: "flex", flexDirection: "column", gap: 12,
                }}>
                  <div>
                    <label style={LBL}>Jenis</label>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {(["Dokumen diterima", "Hasil Review dikembalikan ke Process Owner"] as RevisionType[]).map((t) => (
                        <button key={t} onClick={() => setHType(t)} style={{
                          padding: "9px 12px", borderRadius: 8, cursor: "pointer", textAlign: "left",
                          border: `1.5px solid ${hType === t ? "var(--text-primary)" : "var(--border)"}`,
                          background: hType === t ? "var(--text-primary)" : "var(--surface)",
                          color: hType === t ? "var(--text-inverse)" : "var(--text-secondary)",
                          fontSize: 12, fontWeight: 600, transition: "all 0.12s",
                        }}>
                          {t}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <div>
                      <label style={LBL}>Plan Start</label>
                      <input type="date" value={hPlanStart} onChange={(e) => setHPlanStart(e.target.value)}
                        style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} />
                    </div>
                    <div>
                      <label style={LBL}>Plan Duration (hari)</label>
                      <input type="number" min={1} value={hPlanDur} onChange={(e) => setHPlanDur(e.target.value)}
                        placeholder="0"
                        style={{ ...FIELD, fontFamily: "var(--font-mono)" }} />
                    </div>
                    <div>
                      <label style={LBL}>Actual Start</label>
                      <input type="date" value={hActualStart} onChange={(e) => setHActualStart(e.target.value)}
                        style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} />
                    </div>
                    <div>
                      <label style={LBL}>Actual Duration (hari)</label>
                      <input type="number" min={1} value={hActualDur} onChange={(e) => setHActualDur(e.target.value)}
                        placeholder="0"
                        style={{ ...FIELD, fontFamily: "var(--font-mono)" }} />
                    </div>
                  </div>

                  <div>
                    <label style={LBL}>Catatan</label>
                    <textarea value={hNote} onChange={(e) => setHNote(e.target.value)}
                      placeholder="Keterangan tambahan…" rows={2}
                      style={{ ...FIELD, resize: "none", lineHeight: 1.6 }} />
                  </div>

                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => setShowAddHistory(false)} style={{
                      flex: 1, padding: 11, borderRadius: 8, border: "1px solid var(--border)",
                      background: "var(--surface)", fontSize: 13, fontWeight: 600,
                      color: "var(--text-secondary)", cursor: "pointer",
                    }}>Batal</button>
                    <button onClick={addHistory} disabled={!hPlanStart || !hPlanDur} style={{
                      flex: 2, padding: 11, borderRadius: 8, border: "none",
                      background: hPlanStart && hPlanDur ? "var(--text-primary)" : "var(--surface-2)",
                      fontSize: 13, fontWeight: 700,
                      color: hPlanStart && hPlanDur ? "var(--text-inverse)" : "var(--text-tertiary)",
                      cursor: hPlanStart && hPlanDur ? "pointer" : "default",
                    }}>Simpan Riwayat</button>
                  </div>
                </div>
              )}

              {/* History list */}
              {doc.revisionHistory.length === 0 && (
                <div style={{ fontSize: 13, color: "var(--text-tertiary)", textAlign: "center", padding: "24px 0" }}>
                  Belum ada riwayat.
                </div>
              )}
              {doc.revisionHistory.map((r, i) => {
                const isDiterima = r.type === "Dokumen diterima";
                const accent   = isDiterima ? "#1D4ED8" : "#D97706";
                const accentBg = isDiterima ? "var(--accent-blue-bg)" : "var(--accent-amber-bg)";
                const isEditing = editingIdx === i;

                return (
                  <div key={i} style={{ display: "flex", gap: 12, paddingBottom: i < doc.revisionHistory.length - 1 ? 18 : 0 }}>
                    {/* Timeline dot + line */}
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                      <div style={{ width: 8, height: 8, borderRadius: "50%", background: i === 0 ? accent : "var(--border-strong)", marginTop: 3, flexShrink: 0 }} />
                      {i < doc.revisionHistory.length - 1 && (
                        <div style={{ width: 1, flex: 1, background: "var(--border)", marginTop: 4 }} />
                      )}
                    </div>

                    <div style={{ flex: 1, paddingBottom: 4 }}>
                      {/* Type badge + action buttons */}
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <div style={{ display: "inline-block", padding: "2px 8px", borderRadius: 3, background: accentBg }}>
                          <span style={{ fontSize: 10, fontWeight: 700, color: accent, fontFamily: "var(--font-mono)", letterSpacing: "0.03em" }}>
                            {r.type}
                          </span>
                        </div>
                        {!isEditing && (
                          <div style={{ display: "flex", gap: 4 }}>
                            <button
                              onClick={() => startEditHistory(i)}
                              style={{
                                width: 28, height: 28, borderRadius: 6, border: "1px solid var(--border)",
                                background: "var(--surface-2)", cursor: "pointer",
                                display: "flex", alignItems: "center", justifyContent: "center",
                              }}
                            >
                              <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                                <path d="M8.5 1.5l2 2L4 10H2v-2L8.5 1.5z" stroke="var(--text-secondary)" strokeWidth="1.2" strokeLinejoin="round"/>
                              </svg>
                            </button>
                            <button
                              onClick={() => confirmDelete(i)}
                              style={{
                                width: 28, height: 28, borderRadius: 6, border: "1px solid var(--border)",
                                background: "var(--surface-2)", cursor: "pointer",
                                display: "flex", alignItems: "center", justifyContent: "center",
                              }}
                            >
                              <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                                <path d="M2 3h8M5 3V2h2v1M4.5 3v6.5M7.5 3v6.5M3 3l.5 7h5L9 3" stroke="var(--accent-red)" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
                              </svg>
                            </button>
                          </div>
                        )}
                      </div>

                      {/* Inline edit form */}
                      {isEditing ? (
                        <div style={{ background: "var(--surface-2)", borderRadius: 10, padding: 12, border: "1px solid var(--border)", display: "flex", flexDirection: "column", gap: 10 }}>
                          <div>
                            <label style={LBL}>Jenis</label>
                            {(["Dokumen diterima", "Hasil Review dikembalikan ke Process Owner"] as RevisionType[]).map((t) => (
                              <button key={t} onClick={() => setHType(t)} style={{
                                display: "block", width: "100%", padding: "8px 12px", borderRadius: 8, cursor: "pointer", textAlign: "left", marginBottom: 4,
                                border: `1.5px solid ${hType === t ? "var(--text-primary)" : "var(--border)"}`,
                                background: hType === t ? "var(--text-primary)" : "var(--surface)",
                                color: hType === t ? "var(--text-inverse)" : "var(--text-secondary)",
                                fontSize: 12, fontWeight: 600, transition: "all 0.12s",
                              }}>{t}</button>
                            ))}
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                            <div><label style={LBL}>Plan Start</label>
                              <input type="date" value={hPlanStart} onChange={(e) => setHPlanStart(e.target.value)}
                                style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
                            <div><label style={LBL}>Plan Duration (hari)</label>
                              <input type="number" min={1} value={hPlanDur} onChange={(e) => setHPlanDur(e.target.value)}
                                placeholder="0" style={{ ...FIELD, fontFamily: "var(--font-mono)" }} /></div>
                            <div><label style={LBL}>Actual Start</label>
                              <input type="date" value={hActualStart} onChange={(e) => setHActualStart(e.target.value)}
                                style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} /></div>
                            <div><label style={LBL}>Actual Duration (hari)</label>
                              <input type="number" min={1} value={hActualDur} onChange={(e) => setHActualDur(e.target.value)}
                                placeholder="0" style={{ ...FIELD, fontFamily: "var(--font-mono)" }} /></div>
                          </div>
                          <div><label style={LBL}>Catatan</label>
                            <textarea value={hNote} onChange={(e) => setHNote(e.target.value)} rows={2}
                              style={{ ...FIELD, resize: "none", lineHeight: 1.6 }} /></div>
                          <div style={{ display: "flex", gap: 8 }}>
                            <button onClick={() => { setEditingIdx(null); resetHistoryForm(); }} style={{
                              flex: 1, padding: 10, borderRadius: 8, border: "1px solid var(--border)",
                              background: "var(--surface)", fontSize: 12, fontWeight: 600,
                              color: "var(--text-secondary)", cursor: "pointer",
                            }}>Batal</button>
                            <button onClick={saveEditHistory} disabled={!hPlanStart || !hPlanDur} style={{
                              flex: 2, padding: 10, borderRadius: 8, border: "none",
                              background: hPlanStart && hPlanDur ? "var(--text-primary)" : "var(--surface-2)",
                              fontSize: 12, fontWeight: 700,
                              color: hPlanStart && hPlanDur ? "var(--text-inverse)" : "var(--text-tertiary)",
                              cursor: hPlanStart && hPlanDur ? "pointer" : "default",
                            }}>Simpan</button>
                          </div>
                        </div>
                      ) : (
                        <>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: r.note ? 8 : 0 }}>
                            {[
                              ["Plan Start",      r.planStart ? fmt(r.planStart) : "—"],
                              ["Plan Duration",   `${r.planDuration} hari kerja`],
                              ["Actual Start",    r.actualStart ? fmt(r.actualStart) : "—"],
                              ["Actual Duration", r.actualDuration ? `${r.actualDuration} hari kerja` : "—"],
                            ].map(([label, val]) => (
                              <div key={label} style={{ background: "var(--surface-2)", borderRadius: 6, padding: "7px 10px" }}>
                                <div style={{ fontSize: 9, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 2 }}>{label}</div>
                                <div style={{ fontSize: 12, fontWeight: 600, color: val === "—" ? "var(--text-tertiary)" : "var(--text-primary)" }}>{val}</div>
                              </div>
                            ))}
                          </div>
                          {r.note && (
                            <div style={{ fontSize: 11, color: "var(--text-secondary)", lineHeight: 1.5 }}>{r.note}</div>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                );
              })}

              {/* Delete confirmation inline */}
              {deleteIdx !== null && (
                <div style={{
                  position: "fixed", inset: 0, zIndex: 60,
                  display: "flex", alignItems: "flex-end", justifyContent: "center",
                }}>
                  <div style={{ position: "absolute", inset: 0, background: "rgba(20,20,18,0.4)" }} onClick={() => setDeleteIdx(null)} />
                  <div style={{
                    position: "relative", background: "var(--surface)", borderRadius: "16px 16px 0 0",
                    padding: "20px 20px 32px", width: "100%",
                    animation: "slideUp 0.22s cubic-bezier(0.32,0.72,0,1)",
                  }}>
                    <div style={{ width: 36, height: 4, borderRadius: 2, background: "var(--border-strong)", margin: "0 auto 16px" }} />
                    <h3 style={{ fontSize: 16, fontWeight: 800, color: "var(--text-primary)", margin: "0 0 6px", letterSpacing: "-0.02em" }}>
                      Hapus Riwayat?
                    </h3>
                    <p style={{ fontSize: 13, color: "var(--text-secondary)", margin: "0 0 20px", lineHeight: 1.6 }}>
                      Entri riwayat ini akan dihapus permanen dan tidak bisa dikembalikan.
                    </p>
                    <div style={{ display: "flex", gap: 10 }}>
                      <button onClick={() => setDeleteIdx(null)} style={{
                        flex: 1, padding: 13, borderRadius: 8,
                        background: "var(--surface-2)", border: "1px solid var(--border)",
                        fontSize: 13, fontWeight: 600, color: "var(--text-secondary)", cursor: "pointer",
                      }}>Batal</button>
                      <button onClick={doDelete} style={{
                        flex: 1, padding: 13, borderRadius: 8, border: "none",
                        background: "var(--accent-red)", fontSize: 13, fontWeight: 700, color: "#fff", cursor: "pointer",
                      }}>Ya, Hapus</button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </>
        )}
      </div>
    </>
  );
}

// ─── Add Document Sheet ───────────────────────────────────────────────────────

function AddDocSheet({ onClose, onAdd }: { onClose: () => void; onAdd: (doc: Document) => void }) {
  const [title, setTitle]           = useState("");
  const [refCode, setRefCode]       = useState("");
  const [pic, setPic]               = useState("");
  const [division, setDivision]     = useState("");
  const [status, setStatus]         = useState<Status>("amber");
  const [progress, setProgress]     = useState(0);
  const [description, setDescription] = useState("");
  const [startDate, setStartDate]   = useState(TODAY);
  const [endDate, setEndDate]       = useState("");
  const [refDocs, setRefDocs]       = useState("");
  const [bentukDokumen, setBentukDokumen] = useState<BentukDokumen>("");
  const [kategori, setKategori]     = useState("");
  const [keteranganDRFI, setKeteranganDRFI] = useState<KeteranganDRFI>("");
  const [errors, setErrors]         = useState<Record<string, string>>({});

  function validate() {
    const e: Record<string, string> = {};
    if (!title.trim())  e.title   = "Judul dokumen wajib diisi";
    if (!refCode.trim()) e.refCode = "Kode referensi wajib diisi";
    if (!pic.trim())    e.pic     = "Nama PIC wajib diisi";
    if (!endDate)       e.endDate = "Tanggal selesai wajib diisi";
    if (endDate && endDate < startDate) e.endDate = "Tanggal selesai harus setelah tanggal mulai";
    return e;
  }

  function handleSubmit() {
    const e = validate();
    if (Object.keys(e).length) { setErrors(e); return; }
    onAdd({
      id: `D${Date.now()}`,
      refCode: refCode.trim(),
      title: title.trim(),
      pic: pic.trim(),
      picInitials: getInitials(pic),
      division: division.trim(),
      status,
      progress,
      revision: "Rev. 1",
      lastUpdated: TODAY,
      startDate,
      endDate,
      actualStartDate: null,
      actualEndDate: null,
      description: description.trim(),
      referenceDocuments: refDocs.split(",").map((r) => r.trim()).filter(Boolean),
      bentukDokumen,
      kategori: kategori.trim(),
      keteranganDRFI,
      revisionHistory: [],
    });
    onClose();
  }

  const dur = startDate && endDate && endDate >= startDate ? calDays(startDate, endDate) : null;

  return (
    <>
      <div className="sheet-overlay" onClick={onClose} />
      <div className="sheet-panel scrollbar-hide">
        <div style={{ display: "flex", justifyContent: "center", padding: "12px 0 4px" }}>
          <div style={{ width: 36, height: 4, borderRadius: 2, background: "var(--border-strong)" }} />
        </div>
        <div style={{ padding: "12px 20px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2 style={{ fontSize: 18, fontWeight: 800, color: "var(--text-primary)", margin: 0, letterSpacing: "-0.02em" }}>
            Tambah Dokumen
          </h2>
          <button onClick={onClose} style={{
            background: "var(--surface-2)", border: "none", borderRadius: "50%",
            width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
          }}>
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <path d="M2 2l10 10M12 2L2 12" stroke="var(--text-secondary)" strokeWidth="1.8" strokeLinecap="round" />
            </svg>
          </button>
        </div>
        <div style={{ height: 1, background: "var(--border)", margin: "14px 0 0" }} />

        <div style={{ padding: "16px 20px 32px", display: "flex", flexDirection: "column", gap: 14 }}>
          <div>
            <label style={LBL}>Judul Dokumen *</label>
            <input value={title} onChange={(e) => setTitle(e.target.value)}
              placeholder="Contoh: Wellbore Integrity Assessment Procedure"
              style={{ ...FIELD, borderColor: errors.title ? "var(--accent-red)" : "var(--border)" }} />
            {errors.title && <div style={{ fontSize: 11, color: "var(--accent-red)", marginTop: 4, fontFamily: "var(--font-mono)" }}>{errors.title}</div>}
          </div>
          <div>
            <label style={LBL}>Kode Referensi *</label>
            <input value={refCode} onChange={(e) => setRefCode(e.target.value)}
              placeholder="Contoh: PPWI-ENG-010"
              style={{ ...FIELD, fontFamily: "var(--font-mono)", borderColor: errors.refCode ? "var(--accent-red)" : "var(--border)" }} />
            {errors.refCode && <div style={{ fontSize: 11, color: "var(--accent-red)", marginTop: 4, fontFamily: "var(--font-mono)" }}>{errors.refCode}</div>}
          </div>
          <div>
            <label style={LBL}>Person in Charge (PIC) *</label>
            <input value={pic} onChange={(e) => setPic(e.target.value)}
              placeholder="Contoh: Ahmad Rizwan"
              style={{ ...FIELD, borderColor: errors.pic ? "var(--accent-red)" : "var(--border)" }} />
            {errors.pic && <div style={{ fontSize: 11, color: "var(--accent-red)", marginTop: 4, fontFamily: "var(--font-mono)" }}>{errors.pic}</div>}
            {pic.trim() && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6 }}>
                <PicBadge initials={getInitials(pic)} />
                <span style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>Badge: {getInitials(pic)}</span>
              </div>
            )}
          </div>
          <div>
            <label style={LBL}>Divisi</label>
            <input value={division} onChange={(e) => setDivision(e.target.value)}
              placeholder="Contoh: Engineering, HSE, Legal…"
              style={FIELD} />
          </div>
          <div>
            <label style={LBL}>Status Awal</label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
              {(["red", "amber", "green"] as Status[]).map((s) => {
                const c = statusColor(s); const active = status === s;
                return (
                  <button key={s} onClick={() => setStatus(s)} style={{
                    padding: "10px 8px", borderRadius: 6, cursor: "pointer",
                    border: `1.5px solid ${active ? c.dot : "var(--border)"}`,
                    background: active ? c.bg : "var(--surface)", transition: "all 0.12s",
                  }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: c.dot, margin: "0 auto 4px" }} />
                    <div style={{ fontSize: 10, fontWeight: 600, color: active ? c.text : "var(--text-secondary)", fontFamily: "var(--font-mono)" }}>
                      {s === "red" ? "Action" : s === "amber" ? "Review" : "Done"}
                    </div>
                  </button>
                );
              })}
            </div>
          </div>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <label style={{ ...LBL, margin: 0 }}>Progress Awal</label>
              <span style={{ fontSize: 16, fontWeight: 800, color: progressColor(progress), fontFamily: "var(--font-mono)" }}>{progress}%</span>
            </div>
            <input type="range" min={0} max={100} step={1} value={progress}
              onChange={(e) => setProgress(Number(e.target.value))}
              style={{ accentColor: progressColor(progress) }} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div>
              <label style={LBL}>Tanggal Mulai</label>
              <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)}
                style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12 }} />
            </div>
            <div>
              <label style={LBL}>Tanggal Selesai *</label>
              <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)}
                style={{ ...FIELD, fontFamily: "var(--font-mono)", fontSize: 12, borderColor: errors.endDate ? "var(--accent-red)" : "var(--border)" }} />
              {errors.endDate && <div style={{ fontSize: 11, color: "var(--accent-red)", marginTop: 4, fontFamily: "var(--font-mono)" }}>{errors.endDate}</div>}
            </div>
          </div>
          {dur !== null && (
            <div style={{ marginTop: -6, display: "flex", alignItems: "center", gap: 6 }}>
              <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                <circle cx="6" cy="6" r="5" stroke="var(--text-tertiary)" strokeWidth="1.2"/>
                <path d="M6 3.5V6l1.5 1.5" stroke="var(--text-tertiary)" strokeWidth="1.2" strokeLinecap="round"/>
              </svg>
              <span style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>
                Durasi: <strong style={{ color: "var(--text-primary)" }}>{dur} hari</strong>
              </span>
            </div>
          )}
          <div>
            <label style={LBL}>Deskripsi</label>
            <textarea value={description} onChange={(e) => setDescription(e.target.value)}
              placeholder="Deskripsi singkat…" rows={3}
              style={{ ...FIELD, resize: "none", lineHeight: 1.6 }} />
          </div>
          <div>
            <label style={LBL}>Dokumen Referensi</label>
            <input value={refDocs} onChange={(e) => setRefDocs(e.target.value)}
              placeholder="API RP 90, ISO 16530-1" style={FIELD} />
            <div style={{ fontSize: 10, color: "var(--text-tertiary)", marginTop: 4, fontFamily: "var(--font-mono)" }}>Opsional · pisahkan dengan koma</div>
          </div>

          <div>
            <label style={LBL}>Bentuk Dokumen</label>
            <div style={{ display: "flex", gap: 8 }}>
              {(["HC", "SC", ""] as BentukDokumen[]).map((v) => (
                <button key={v} onClick={() => setBentukDokumen(v)} style={{
                  flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer",
                  border: `1.5px solid ${bentukDokumen === v ? "var(--text-primary)" : "var(--border)"}`,
                  background: bentukDokumen === v ? "var(--text-primary)" : "var(--surface)",
                  color: bentukDokumen === v ? "var(--text-inverse)" : "var(--text-secondary)",
                  fontSize: 12, fontWeight: 700, fontFamily: "var(--font-mono)",
                  transition: "all 0.12s",
                }}>{v === "" ? "—" : v}</button>
              ))}
            </div>
          </div>

          <div>
            <label style={LBL}>Kategori (PPWI)</label>
            <input value={kategori} onChange={(e) => setKategori(e.target.value)}
              placeholder="Contoh: PPWI-ENG"
              style={{ ...FIELD, fontFamily: "var(--font-mono)" }} />
          </div>

          <div>
            <label style={LBL}>Keterangan DR / FI</label>
            <div style={{ display: "flex", gap: 8 }}>
              {(["DR", "FI", ""] as KeteranganDRFI[]).map((v) => (
                <button key={v} onClick={() => setKeteranganDRFI(v)} style={{
                  flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer",
                  border: `1.5px solid ${keteranganDRFI === v ? "var(--text-primary)" : "var(--border)"}`,
                  background: keteranganDRFI === v ? "var(--text-primary)" : "var(--surface)",
                  color: keteranganDRFI === v ? "var(--text-inverse)" : "var(--text-secondary)",
                  fontSize: 12, fontWeight: 700, fontFamily: "var(--font-mono)",
                  transition: "all 0.12s",
                }}>{v === "" ? "—" : v}</button>
              ))}
            </div>
          </div>

          <button onClick={handleSubmit} style={{
            width: "100%", padding: 14, borderRadius: 8, marginTop: 4,
            background: "var(--text-primary)", color: "var(--text-inverse)",
            border: "none", fontWeight: 700, fontSize: 14, cursor: "pointer",
          }}>Simpan Dokumen</button>
          <button onClick={onClose} style={{
            width: "100%", padding: 13, borderRadius: 8,
            background: "transparent", color: "var(--text-secondary)",
            border: "1px solid var(--border)", fontWeight: 600, fontSize: 13, cursor: "pointer",
          }}>Batal</button>
        </div>
      </div>
    </>
  );
}

// ─── App Main Component ───────────────────────────────────────────────────────

type Tab = "dashboard" | "timeline" | "detail";

export default function App() {
  const [docs, setDocs] = useLocalStorage<Document[]>("ppwi_clean", DOCUMENTS);
  const [activeTab, setActiveTab] = useState<Tab>("dashboard");
  const [selectedDoc, setSelectedDoc] = useState<Document | null>(null);
  const [darkMode, setDarkMode] = useLocalStorage<boolean>("ppwi_dark", false);
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [resetPhrase, setResetPhrase] = useState("");
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const RESET_PHRASE = "SAYA YAKIN UNTUK RESET DATA";

  const [currentToday, setCurrentToday] = useState(getRealTodayIso());

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", darkMode ? "dark" : "light");
  }, [darkMode]);

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentToday(getRealTodayIso());
    }, 60000);
    return () => clearInterval(timer);
  }, []);

  const [showSettings, setShowSettings] = useState(false);
  const [showAddDoc, setShowAddDoc] = useState(false);
  const [timelineSearch, setTimelineSearch] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [filterStatus, setFilterStatus] = useState<Status | null>(null);
  const [filterDivision, setFilterDivision] = useState<string | null>(null);
  const [filterQuarter, setFilterQuarter] = useState<number | null>(null);
  const [timelineMode, setTimelineMode] = useState<TimelineMode>("week");
  const [expandedDocId, setExpandedDocId] = useState<string | null>(null);
  const ganttRef = useRef<HTMLDivElement>(null);

  const total     = docs.length;
  const inReview  = docs.filter((d) => d.status === "amber").length;
  const completed = docs.filter((d) => d.progress === 100).length;
  const overdue   = docs.filter((d) => isOverdue(d)).length;

  const uniqueDivisions = [...new Set(docs.map((d) => d.division).filter(Boolean))].sort();

  const filtered = docs.filter((d) => {
    if (filterStatus   && d.status !== filterStatus)           return false;
    if (filterDivision && d.division !== filterDivision)       return false;
    if (filterQuarter  && quarter(d.endDate) !== filterQuarter) return false;
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      return d.title.toLowerCase().includes(q) || d.refCode.toLowerCase().includes(q) || d.pic.toLowerCase().includes(q);
    }
    return true;
  });

  useEffect(() => {
    if (activeTab !== "timeline" || !ganttRef.current) return;
    let scrollX = 0;
    if (timelineMode === "day")   scrollX = Math.max(0, (TODAY_WD  - 15) * WD_W);
    if (timelineMode === "week")  scrollX = Math.max(0, (weekOf(TODAY) - 4) * WK_W);
    if (timelineMode === "month") scrollX = Math.max(0, (new Date(TODAY).getMonth() - 2) * MO_W);
    ganttRef.current.scrollLeft = scrollX;
  }, [activeTab, timelineMode]);

  function showToast(msg: string) {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(null), 3000);
  }

  function handleSave(updated: Document) {
    setDocs((prev) => prev.map((d) => d.id === updated.id ? updated : d));
    setSelectedDoc(updated);
  }
  function handleAdd(doc: Document) {
    setDocs((prev) => [doc, ...prev]);
    showToast("Dokumen berhasil ditambahkan!");
  }
  function handleReset() {
    setDocs([]);
    setShowResetConfirm(false);
    setResetPhrase("");
    setShowSettings(false);
    showToast("Seluruh data telah di-reset!");
  }

  // Export ke Excel
  function handleExportExcel() {
    type Row = Record<string, string | number>;
    const rows: Row[] = docs.flatMap((doc): Row[] => {
      if (doc.revisionHistory.length === 0) {
        return [{
          "Ref Code": doc.refCode, "Judul": doc.title, "PIC": doc.pic,
          "Divisi": doc.division, "Status": doc.status, "Progress (%)": doc.progress,
          "Revisi": doc.revision, "Bentuk": doc.bentukDokumen, "Kategori": doc.kategori,
          "DR/FI": doc.keteranganDRFI,
          "Tgl Mulai Rencana": doc.startDate, "Tgl Selesai Rencana": doc.endDate,
          "Tgl Mulai Aktual": doc.actualStartDate ?? "", "Tgl Selesai Aktual": doc.actualEndDate ?? "",
          "Riwayat #": 0, "Tipe Riwayat": "", "Plan Start": "", "Plan Durasi (HK)": 0,
          "Aktual Start": "", "Aktual Durasi (HK)": 0, "Catatan": "",
        }];
      }
      return doc.revisionHistory.map((r, i) => ({
        "Ref Code": i === 0 ? doc.refCode : "", "Judul": i === 0 ? doc.title : "", "PIC": i === 0 ? doc.pic : "",
        "Divisi": i === 0 ? doc.division : "", "Status": i === 0 ? doc.status : "", "Progress (%)": i === 0 ? doc.progress : "",
        "Revisi": i === 0 ? doc.revision : "", "Bentuk": i === 0 ? doc.bentukDokumen : "", "Kategori": i === 0 ? doc.kategori : "",
        "DR/FI": i === 0 ? doc.keteranganDRFI : "",
        "Tgl Mulai Rencana": i === 0 ? doc.startDate : "", "Tgl Selesai Rencana": i === 0 ? doc.endDate : "",
        "Tgl Mulai Aktual": i === 0 ? (doc.actualStartDate ?? "") : "", "Tgl Selesai Aktual": i === 0 ? (doc.actualEndDate ?? "") : "",
        "Riwayat #": i + 1, "Tipe Riwayat": r.type, "Plan Start": r.planStart, "Plan Durasi (HK)": r.planDuration,
        "Aktual Start": r.actualStart ?? "", "Aktual Durasi (HK)": r.actualDuration ?? "", "Catatan": r.note,
      }));
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PPWI Monitor");
    XLSX.writeFile(wb, `PPWI_Monitor_${TODAY}.xlsx`);
    showToast("File Excel berhasil diunduh!");
  }

  // Export Backup JSON
  function handleExportJSON() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(docs, null, 2));
    const downloadAnchor = document.createElement("a");
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `PPWI_Backup_${TODAY}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showToast("Backup JSON berhasil diunduh!");
  }

  // Restore/Import JSON
  function handleImportJSON(e: React.ChangeEvent<HTMLInputElement>) {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], "UTF-8");
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            setDocs(parsed);
            showToast(`Berhasil merestore ${parsed.length} dokumen dari JSON!`);
          } else {
            alert("Format JSON tidak valid. Data harus berupa array dokumen.");
          }
        } catch (err) {
          alert("Gagal membaca file JSON. Pastikan file valid.");
        }
      };
    }
  }

  return (
    <div style={{
      background: "var(--bg)", minHeight: "100%", width: "100%",
      display: "flex", flexDirection: "column",
      fontFamily: "var(--font-sans)", position: "relative",
    }}>

      {/* Toast Notification */}
      {toastMsg && (
        <div style={{
          position: "fixed", top: 20, left: "50%", transform: "translateX(-50%)",
          zIndex: 100, background: "var(--text-primary)", color: "var(--text-inverse)",
          padding: "10px 18px", borderRadius: 20, fontSize: 12, fontWeight: 600,
          boxShadow: "0 4px 12px rgba(0,0,0,0.2)", pointerEvents: "none",
          animation: "fadeIn 0.2s ease",
        }}>
          {toastMsg}
        </div>
      )}

      {/* ── Top Bar ── */}
      <div style={{
        background: "var(--surface)", borderBottom: "1px solid var(--border)",
        padding: "16px 20px 14px", position: "sticky", top: 0, zIndex: 20,
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)", letterSpacing: "0.1em", textTransform: "uppercase" }}>
              {fmtHeaderDate(currentToday)}
            </div>
            <h1 style={{ fontSize: 20, fontWeight: 800, color: "var(--text-primary)", margin: "2px 0 0", letterSpacing: "-0.03em", lineHeight: 1 }}>
              PPWI Monitor
            </h1>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setShowSettings(true)} style={{
              width: 36, height: 36, borderRadius: "50%",
              background: "var(--surface-2)", border: "1px solid var(--border)",
              display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
            }}>
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M8 10a2 2 0 100-4 2 2 0 000 4z" stroke="var(--text-secondary)" strokeWidth="1.5" />
                <path d="M6.8 1.5h2.4l.4 1.7a5.2 5.2 0 011.3.7l1.7-.5 1.2 2.1-1.2 1.3c.1.4.1.9 0 1.3l1.2 1.3-1.2 2.1-1.7-.5c-.4.3-.8.6-1.3.7l-.4 1.7H6.8l-.4-1.7c-.5-.1-.9-.4-1.3-.7l-1.7.5-1.2-2.1 1.2-1.3a5.2 5.2 0 010-1.3L2.2 5.4l1.2-2.1 1.7.5c.4-.3.8-.6 1.3-.7l.4-1.6z" stroke="var(--text-secondary)" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            <button onClick={() => setShowAddDoc(true)} style={{
              width: 36, height: 36, borderRadius: "50%",
              background: "var(--text-primary)", border: "none",
              display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
            }}>
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M8 3v10M3 8h10" stroke="var(--bg)" strokeWidth="2" strokeLinecap="round"/>
              </svg>
            </button>
          </div>
        </div>
      </div>


      {/* ─── DASHBOARD ─── */}
      {activeTab === "dashboard" && (
        <div style={{ flex: 1, overflowY: "auto" }} className="scrollbar-hide">
          <div style={{ padding: "16px 20px 0" }}>
            <div style={{ display: "flex", gap: 8 }}>
              <MetricCard label="Total Dokumen" value={total} />
              <MetricCard label="Dalam Review" value={inReview} accent="var(--accent-amber)" />
              <MetricCard label="Selesai" value={completed} accent="var(--accent-green)" />
              <MetricCard label="Overdue" value={overdue} accent="var(--accent-red)" />
            </div>
          </div>

          <div style={{ padding: "12px 20px 0" }}>
            <div style={{
              display: "flex", alignItems: "center", gap: 10,
              background: "var(--surface)", border: "1px solid var(--border)",
              borderRadius: 8, padding: "10px 14px",
            }}>
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                <circle cx="6" cy="6" r="4.5" stroke="var(--text-tertiary)" strokeWidth="1.4" />
                <path d="M9.5 9.5L12 12" stroke="var(--text-tertiary)" strokeWidth="1.4" strokeLinecap="round" />
              </svg>
              <input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Cari dokumen, kode ref, PIC…"
                style={{ flex: 1, border: "none", outline: "none", background: "none", fontSize: 13, color: "var(--text-primary)", fontFamily: "var(--font-sans)" }} />
              {searchQuery && (
                <button onClick={() => setSearchQuery("")} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
                  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                    <circle cx="7" cy="7" r="6" fill="var(--surface-2)" />
                    <path d="M5 5l4 4M9 5l-4 4" stroke="var(--text-secondary)" strokeWidth="1.3" strokeLinecap="round" />
                  </svg>
                </button>
              )}
            </div>
          </div>

          <div style={{ padding: "10px 20px 0", overflowX: "auto" }} className="scrollbar-hide">
            <div style={{ display: "flex", gap: 6, minWidth: "max-content" }}>
              {(["red", "amber", "green"] as Status[]).map((s) => {
                const c = statusColor(s); const active = filterStatus === s;
                return (
                  <button key={s} onClick={() => setFilterStatus(active ? null : s)} style={{
                    padding: "5px 12px", borderRadius: 20,
                    border: `1px solid ${active ? c.dot : "var(--border)"}`,
                    background: active ? c.bg : "var(--surface)",
                    color: active ? c.text : "var(--text-secondary)",
                    fontSize: 11, fontWeight: 600, cursor: "pointer",
                    fontFamily: "var(--font-mono)", display: "flex", alignItems: "center", gap: 5,
                    transition: "all 0.12s",
                  }}>
                    <div style={{ width: 6, height: 6, borderRadius: "50%", background: c.dot }} />
                    {s === "red" ? "Action" : s === "amber" ? "Review" : "Done"}
                  </button>
                );
              })}
              <div style={{ width: 1, background: "var(--border)", margin: "0 2px" }} />
              {uniqueDivisions.map((div) => {
                const active = filterDivision === div;
                return (
                  <button key={div} onClick={() => setFilterDivision(active ? null : div)} style={{
                    padding: "5px 12px", borderRadius: 20,
                    border: `1px solid ${active ? "var(--text-primary)" : "var(--border)"}`,
                    background: active ? "var(--text-primary)" : "var(--surface)",
                    color: active ? "var(--text-inverse)" : "var(--text-secondary)",
                    fontSize: 11, fontWeight: 600, cursor: "pointer", transition: "all 0.12s",
                  }}>{div}</button>
                );
              })}
              <div style={{ width: 1, background: "var(--border)", margin: "0 2px" }} />
              {[1, 2, 3, 4].map((q) => {
                const active = filterQuarter === q;
                return (
                  <button key={q} onClick={() => setFilterQuarter(active ? null : q)} style={{
                    padding: "5px 12px", borderRadius: 20,
                    border: `1px solid ${active ? "var(--accent-blue)" : "var(--border)"}`,
                    background: active ? "var(--accent-blue-bg)" : "var(--surface)",
                    color: active ? "var(--accent-blue)" : "var(--text-secondary)",
                    fontSize: 11, fontWeight: 600, cursor: "pointer",
                    fontFamily: "var(--font-mono)", transition: "all 0.12s",
                  }}>Q{q}</button>
                );
              })}
            </div>
          </div>

          <div style={{ padding: "12px 20px 4px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 11, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>
              {filtered.length} dokumen{(filterStatus || filterDivision || filterQuarter || searchQuery) ? " · filtered" : ""}
            </span>
            {(filterStatus || filterDivision || filterQuarter || searchQuery) && (
              <button onClick={() => { setFilterStatus(null); setFilterDivision(null); setFilterQuarter(null); setSearchQuery(""); }}
                style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "var(--accent-blue)", fontWeight: 600 }}>
                Hapus filter
              </button>
            )}
          </div>

          <div style={{ padding: "0 20px 100px", display: "flex", flexDirection: "column", gap: 8 }}>
            {filtered.length === 0
              ? <div style={{ textAlign: "center", padding: "48px 0", fontSize: 13, color: "var(--text-tertiary)" }}>Tidak ada dokumen yang sesuai.</div>
              : filtered.map((doc) => (
                  <DocumentCard key={doc.id} doc={doc}
                    onClick={() => { setSelectedDoc(doc); setActiveTab("detail"); }} />
                ))
            }
          </div>
        </div>
      )}

      {/* ─── TIMELINE ─── */}
      {activeTab === "timeline" && (
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {/* Search */}
          <div style={{ padding: "10px 16px 8px", background: "var(--surface)", borderBottom: "1px solid var(--border)" }}>
            <div style={{
              display: "flex", alignItems: "center", gap: 8,
              background: "var(--surface-2)", border: "1px solid var(--border)",
              borderRadius: 8, padding: "8px 12px",
            }}>
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                <circle cx="6" cy="6" r="4.5" stroke="var(--text-tertiary)" strokeWidth="1.4"/>
                <path d="M9.5 9.5L12 12" stroke="var(--text-tertiary)" strokeWidth="1.4" strokeLinecap="round"/>
              </svg>
              <input
                value={timelineSearch}
                onChange={(e) => setTimelineSearch(e.target.value)}
                placeholder="Cari dokumen atau kode ref…"
                style={{ flex: 1, border: "none", outline: "none", background: "none", fontSize: 12, color: "var(--text-primary)", fontFamily: "var(--font-sans)" }}
              />
              {timelineSearch && (
                <button onClick={() => setTimelineSearch("")} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
                  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                    <circle cx="7" cy="7" r="6" fill="var(--border-strong)"/>
                    <path d="M5 5l4 4M9 5l-4 4" stroke="#fff" strokeWidth="1.3" strokeLinecap="round"/>
                  </svg>
                </button>
              )}
            </div>
          </div>
          {/* Controls */}
          <div style={{ padding: "8px 16px", background: "var(--surface)", borderBottom: "1px solid var(--border)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ display: "flex", gap: 5 }}>
              {(["day", "week", "month"] as TimelineMode[]).map((m) => (
                <button key={m} onClick={() => setTimelineMode(m)} style={{
                  padding: "5px 12px", borderRadius: 20,
                  border: `1px solid ${timelineMode === m ? "var(--text-primary)" : "var(--border)"}`,
                  background: timelineMode === m ? "var(--text-primary)" : "var(--surface)",
                  color: timelineMode === m ? "var(--text-inverse)" : "var(--text-secondary)",
                  fontSize: 11, fontWeight: 700, cursor: "pointer",
                  fontFamily: "var(--font-mono)", transition: "all 0.12s",
                }}>
                  {m === "day" ? "Hari" : m === "week" ? "Minggu" : "Bulan"}
                </button>
              ))}
            </div>
            <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <div style={{ width: 16, height: 7, background: "#E0DED9", borderRadius: 2, border: "1px solid var(--border-strong)" }} />
                <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>Rencana</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <div style={{ width: 16, height: 7, background: "#059669", borderRadius: 2, opacity: 0.82 }} />
                <span style={{ fontSize: 10, color: "var(--text-tertiary)", fontFamily: "var(--font-mono)" }}>Aktual</span>
              </div>
            </div>
          </div>

          <div ref={ganttRef} style={{ flex: 1, overflowX: "auto", overflowY: "auto" }} className="scrollbar-hide">
            <div style={{ minWidth: LABEL_W + ganttTotal(timelineMode), position: "relative" }}>
              <GanttHeader mode={timelineMode} />
              {docs.filter((d) => {
                if (!timelineSearch) return true;
                const q = timelineSearch.toLowerCase();
                return d.title.toLowerCase().includes(q) || d.refCode.toLowerCase().includes(q) || d.pic.toLowerCase().includes(q);
              }).map((doc) => (
                <GanttRow
                  key={doc.id}
                  doc={doc}
                  mode={timelineMode}
                  expanded={expandedDocId === doc.id}
                  onToggle={() => setExpandedDocId(expandedDocId === doc.id ? null : doc.id)}
                />
              ))}
              <div style={{ height: 80 }} />
            </div>
          </div>
        </div>
      )}

      {/* ─── BOTTOM NAV ─── */}
      <div style={{
        position: "fixed", bottom: 0, left: 0, right: 0,
        width: "100%",
        background: "var(--surface)", borderTop: "1px solid var(--border)",
        display: "flex", padding: "10px 0 max(16px, env(safe-area-inset-bottom, 16px))", zIndex: 30,
      }}>
        {[
          { id: "dashboard", label: "Dashboard", icon: (a: boolean) => (
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <rect x="2" y="2" width="7" height="7" rx="1.5" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.6" fill={a ? "var(--text-primary)" : "none"} />
              <rect x="11" y="2" width="7" height="7" rx="1.5" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.6" />
              <rect x="2" y="11" width="7" height="7" rx="1.5" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.6" />
              <rect x="11" y="11" width="7" height="7" rx="1.5" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.6" />
            </svg>
          )},
          { id: "timeline", label: "Timeline", icon: (a: boolean) => (
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <path d="M2 5h16M2 10h10M2 15h6" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.8" strokeLinecap="round" />
              <circle cx="16" cy="10" r="2.5" stroke={a ? "var(--text-primary)" : "var(--text-tertiary)"} strokeWidth="1.5" fill={a ? "var(--text-primary)" : "none"} />
            </svg>
          )},
        ].map((tab) => {
          const active = activeTab === tab.id;
          return (
            <button key={tab.id} onClick={() => setActiveTab(tab.id as Tab)} style={{
              flex: 1, display: "flex", flexDirection: "column", alignItems: "center",
              gap: 4, background: "none", border: "none", cursor: "pointer", padding: "2px 0",
            }}>
              {tab.icon(active)}
              <span style={{ fontSize: 10, fontWeight: active ? 700 : 500, color: active ? "var(--text-primary)" : "var(--text-tertiary)" }}>
                {tab.label}
              </span>
            </button>
          );
        })}
      </div>

      {/* ─── SETTINGS & DATA MANAGEMENT ─── */}
      {showSettings && (
        <>
          <div className="sheet-overlay" onClick={() => { setShowSettings(false); setShowResetConfirm(false); setResetPhrase(""); }} />
          <div style={{
            position: "fixed", bottom: 0, left: 0, right: 0,
            width: "100%",
            background: "var(--surface)", borderRadius: "16px 16px 0 0",
            padding: "24px 20px max(36px, env(safe-area-inset-bottom, 36px))", zIndex: 50,
            animation: "slideUp 0.25s cubic-bezier(0.32,0.72,0,1)",
            maxHeight: "90vh", maxHeight: "90dvh", overflowY: "auto",
          }} className="scrollbar-hide">
            <div style={{ width: 36, height: 4, borderRadius: 2, background: "var(--border-strong)", margin: "0 auto 20px" }} />
            <h3 style={{ fontSize: 17, fontWeight: 800, color: "var(--text-primary)", margin: "0 0 20px", letterSpacing: "-0.02em" }}>
              Settings & Data
            </h3>

            {/* Appearance section */}
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 10, fontFamily: "var(--font-mono)", letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--text-tertiary)", marginBottom: 10 }}>
                Tampilan
              </div>
              <div style={{
                display: "flex", alignItems: "center", justifyContent: "space-between",
                padding: "13px 16px", borderRadius: 10,
                background: "var(--surface-2)", border: "1px solid var(--border)",
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <div style={{ width: 36, height: 36, borderRadius: 8, background: darkMode ? "var(--text-primary)" : "var(--border)", display: "flex", alignItems: "center", justifyContent: "center", transition: "background 0.2s" }}>
                    {darkMode ? (
                      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                        <path d="M13.5 9.5A6 6 0 016.5 2.5a6 6 0 100 11 6 6 0 007-4z" stroke="var(--bg)" strokeWidth="1.4" strokeLinejoin="round"/>
                      </svg>
                    ) : (
                      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                        <circle cx="8" cy="8" r="3" stroke="var(--text-secondary)" strokeWidth="1.4"/>
                        <path d="M8 1.5v1M8 13.5v1M1.5 8h1M13.5 8h1M3.4 3.4l.7.7M11.9 11.9l.7.7M3.4 12.6l.7-.7M11.9 4.1l.7-.7" stroke="var(--text-secondary)" strokeWidth="1.4" strokeLinecap="round"/>
                      </svg>
                    )}
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text-primary)" }}>
                      {darkMode ? "Dark Mode" : "Light Mode"}
                    </div>
                    <div style={{ fontSize: 11, color: "var(--text-secondary)" }}>
                      {darkMode ? "Tampilan gelap aktif" : "Tampilan terang aktif"}
                    </div>
                  </div>
                </div>
                {/* Toggle switch */}
                <button
                  onClick={() => setDarkMode(!darkMode)}
                  style={{
                    width: 44, height: 26, borderRadius: 13, border: "none", cursor: "pointer",
                    background: darkMode ? "var(--text-primary)" : "var(--border-strong)",
                    position: "relative", flexShrink: 0, transition: "background 0.2s",
                  }}
                >
                  <div style={{
                    position: "absolute", top: 3, left: darkMode ? 21 : 3,
                    width: 20, height: 20, borderRadius: "50%",
                    background: darkMode ? "var(--bg)" : "var(--surface)",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.25)",
                    transition: "left 0.2s cubic-bezier(0.4,0,0.2,1)",
                  }} />
                </button>
              </div>
            </div>

            {/* Backup & Export Section */}
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 10, fontFamily: "var(--font-mono)", letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--text-tertiary)", marginBottom: 10 }}>
                Status Penyimpanan HP & Export
              </div>

              {/* Cache Indicator */}
              <div style={{
                padding: "12px 14px", borderRadius: 10,
                background: "var(--surface-2)", border: "1px solid var(--border)",
                display: "flex", alignItems: "center", gap: 12, marginBottom: 12,
              }}>
                <div style={{ width: 34, height: 34, borderRadius: 8, background: "#059669", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                    <path d="M4 6.5A4.5 4.5 0 0113 5a3.5 3.5 0 012 6.5H4a3 3 0 010-6z" stroke="#fff" strokeWidth="1.5" strokeLinejoin="round"/>
                    <path d="M6 14l3 3 3-3" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text-primary)", marginBottom: 2 }}>Cache & Penyimpanan HP</div>
                  <div style={{ fontSize: 11, color: "#059669", fontWeight: 600, fontFamily: "var(--font-mono)" }}>
                    ✓ Tersimpan Otomatis di HP (Offline Ready)
                  </div>
                </div>
              </div>
              
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {/* Excel Export */}
                <button onClick={handleExportExcel} style={{
                  width: "100%", padding: "12px 14px", borderRadius: 10,
                  background: "var(--accent-green-bg)", border: "1px solid #A7F3D0",
                  display: "flex", alignItems: "center", gap: 12, cursor: "pointer", textAlign: "left",
                }}>
                  <div style={{ width: 34, height: 34, borderRadius: 8, background: "var(--accent-green)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                      <rect x="2" y="2" width="14" height="14" rx="2" stroke="#fff" strokeWidth="1.5"/>
                      <path d="M5 6h3M5 9h3M5 12h3M10 6h3M10 9h3M10 12h3" stroke="#fff" strokeWidth="1.3" strokeLinecap="round"/>
                    </svg>
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--accent-green)", marginBottom: 2 }}>Export ke Excel (.xlsx)</div>
                    <div style={{ fontSize: 11, color: "var(--text-secondary)" }}>Unduh laporan spreadsheet </div>
                  </div>
                </button>

                {/* Backup JSON */}
                <button onClick={handleExportJSON} style={{
                  width: "100%", padding: "12px 14px", borderRadius: 10,
                  background: "var(--accent-blue-bg)", border: "1px solid #BFDBFE",
                  display: "flex", alignItems: "center", gap: 12, cursor: "pointer", textAlign: "left",
                }}>
                  <div style={{ width: 34, height: 34, borderRadius: 8, background: "var(--accent-blue)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                      <path d="M3 13.5v1.5a1.5 1.5 0 001.5 1.5h9a1.5 1.5 0 001.5-1.5v-1.5M9 1.5v10.5M5.5 8.5L9 12l3.5-3.5" stroke="#fff" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--accent-blue)", marginBottom: 2 }}>Backup Data (JSON)</div>
                    <div style={{ fontSize: 11, color: "var(--text-secondary)" }}>Download file cadangan JSON</div>
                  </div>
                </button>

                {/* Restore JSON */}
                <button onClick={() => fileInputRef.current?.click()} style={{
                  width: "100%", padding: "12px 14px", borderRadius: 10,
                  background: "var(--surface-2)", border: "1px solid var(--border)",
                  display: "flex", alignItems: "center", gap: 12, cursor: "pointer", textAlign: "left",
                }}>
                  <div style={{ width: 34, height: 34, borderRadius: 8, background: "var(--text-primary)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                      <path d="M3 13.5v1.5a1.5 1.5 0 001.5 1.5h9a1.5 1.5 0 001.5-1.5v-1.5M9 12V1.5M5.5 5L9 1.5 12.5 5" stroke="var(--bg)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text-primary)", marginBottom: 2 }}>Restore Data (JSON)</div>
                    <div style={{ fontSize: 11, color: "var(--text-secondary)" }}>Import file cadangan untuk memulihkan</div>
                  </div>
                </button>
                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleImportJSON}
                  accept=".json"
                  style={{ display: "none" }}
                />
              </div>
            </div>

            {/* Danger zone */}
            <div style={{ marginTop: 20 }}>
              <div style={{ fontSize: 10, fontFamily: "var(--font-mono)", letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--text-tertiary)", marginBottom: 10 }}>
                Danger Zone
              </div>
              {!showResetConfirm ? (
                <button onClick={() => setShowResetConfirm(true)} style={{
                  width: "100%", padding: "13px 16px", borderRadius: 10,
                  background: "var(--accent-red-bg)", border: "1px solid #FECDD3",
                  display: "flex", alignItems: "center", gap: 12, cursor: "pointer", textAlign: "left",
                }}>
                  <div style={{ width: 36, height: 36, borderRadius: 8, background: "var(--accent-red)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <path d="M2 4h12M6 4V2.5h4V4M5 4v8.5h6V4" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  </div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--accent-red)", marginBottom: 2 }}>Reset Semua Data</div>
                    <div style={{ fontSize: 11, color: "var(--text-secondary)" }}>Hapus seluruh data. Tidak dapat diurungkan.</div>
                  </div>
                </button>
              ) : (
                <div style={{ borderRadius: 12, background: "var(--accent-red-bg)", border: "1.5px solid #FECDD3", overflow: "hidden" }}>
                  <div style={{ padding: "14px 16px 12px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                      <svg width="15" height="15" viewBox="0 0 15 15" fill="none">
                        <path d="M7.5 2L13.5 12.5H1.5L7.5 2Z" stroke="var(--accent-red)" strokeWidth="1.4" strokeLinejoin="round"/>
                        <path d="M7.5 6v3.5M7.5 11v.5" stroke="var(--accent-red)" strokeWidth="1.4" strokeLinecap="round"/>
                      </svg>
                      <span style={{ fontSize: 12, fontWeight: 700, color: "var(--accent-red)" }}>Konfirmasi Reset</span>
                    </div>
                    <p style={{ fontSize: 12, color: "var(--text-secondary)", margin: "0 0 12px", lineHeight: 1.6 }}>
                      Tindakan ini akan <strong style={{ color: "var(--text-primary)" }}>menghapus seluruh data</strong> secara permanen. Untuk melanjutkan, ketik frasa berikut persis:
                    </p>
                    <div style={{
                      padding: "8px 12px", borderRadius: 6,
                      background: "var(--surface)", border: "1px solid #FECDD3",
                      fontFamily: "var(--font-mono)", fontSize: 11, fontWeight: 600,
                      color: "var(--accent-red)", letterSpacing: "0.02em", marginBottom: 12,
                      userSelect: "all",
                    }}>
                      {RESET_PHRASE}
                    </div>
                    <input
                      value={resetPhrase}
                      onChange={(e) => setResetPhrase(e.target.value)}
                      placeholder="Ketik frasa di atas…"
                      autoCapitalize="characters"
                      style={{
                        width: "100%", padding: "10px 12px", borderRadius: 8,
                        border: `1.5px solid ${resetPhrase === RESET_PHRASE ? "var(--accent-green)" : resetPhrase.length > 0 ? "#FECDD3" : "var(--border)"}`,
                        background: "var(--surface)", fontSize: 12,
                        fontFamily: "var(--font-mono)", color: "var(--text-primary)",
                        outline: "none", boxSizing: "border-box",
                        transition: "border-color 0.15s",
                      }}
                    />
                  </div>
                  <div style={{ display: "flex", borderTop: "1px solid #FECDD3" }}>
                    <button onClick={() => { setShowResetConfirm(false); setResetPhrase(""); }} style={{
                      flex: 1, padding: "12px 0",
                      background: "none", border: "none", borderRight: "1px solid #FECDD3",
                      fontSize: 13, fontWeight: 600, color: "var(--text-secondary)", cursor: "pointer",
                    }}>Batal</button>
                    <button
                      onClick={handleReset}
                      disabled={resetPhrase !== RESET_PHRASE}
                      style={{
                        flex: 1, padding: "12px 0",
                        background: resetPhrase === RESET_PHRASE ? "var(--accent-red)" : "transparent",
                        border: "none", fontSize: 13, fontWeight: 700,
                        color: resetPhrase === RESET_PHRASE ? "#fff" : "#FECDD3",
                        cursor: resetPhrase === RESET_PHRASE ? "pointer" : "not-allowed",
                        transition: "all 0.15s",
                      }}
                    >Hapus Semua Data</button>
                  </div>
                </div>
              )}
            </div>

            <button onClick={() => { setShowSettings(false); setShowResetConfirm(false); setResetPhrase(""); }} style={{
              width: "100%", padding: 13, borderRadius: 8, marginTop: 16,
              background: "transparent", color: "var(--text-secondary)",
              border: "1px solid var(--border)", fontWeight: 600, fontSize: 13, cursor: "pointer",
            }}>Tutup</button>
          </div>
        </>
      )}

      {/* ─── DETAIL SHEET ─── */}
      {selectedDoc && activeTab === "detail" && (
        <DetailSheet
          doc={docs.find((d) => d.id === selectedDoc.id) ?? selectedDoc}
          onClose={() => { setSelectedDoc(null); setActiveTab("dashboard"); }}
          onSave={handleSave}
        />
      )}

      {/* ─── ADD DOC ─── */}
      {showAddDoc && (
        <AddDocSheet onClose={() => setShowAddDoc(false)} onAdd={handleAdd} />
      )}

      {/* FAB */}
      {activeTab === "dashboard" && !showAddDoc && !showSettings && !selectedDoc && (
        <button onClick={() => setShowAddDoc(true)} style={{
          position: "fixed", bottom: "calc(env(safe-area-inset-bottom, 0px) + 76px)", right: 20,
          width: 52, height: 52, borderRadius: "50%",
          background: "var(--text-primary)", border: "none",
          display: "flex", alignItems: "center", justifyContent: "center",
          cursor: "pointer", zIndex: 25,
          boxShadow: "0 4px 16px rgba(20,20,18,0.18)",
        }}>
          <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
            <path d="M10 4v12M4 10h12" stroke="var(--bg)" strokeWidth="2.2" strokeLinecap="round"/>
          </svg>
        </button>
      )}
    </div>
  );
}
