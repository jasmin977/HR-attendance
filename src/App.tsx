import React from "react";
import { useState, useCallback, useRef } from "react";
import type { ChangeEvent, KeyboardEvent } from "react";
import * as XLSX from "xlsx";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Employee {
  id: string | number;
  name: string;
  dept: string;
}

interface RawRecord {
  date: string; // DD/MM/YYYY
  times: string[]; // HH:MM
}

interface Period {
  from: string; // DD/MM/YYYY
  to: string;
  hours: string;
}

type DayStatus = "working" | "absent" | "dayoff" | "dayoff_worked" | "error";
type Override = "full" | "half";

interface AnalyzedDay extends RawRecord {
  deduped: string[];
  isDayOff: boolean;
  status: DayStatus;
  worked: number | null;
  extra?: number; // positive = heures sup, negative = manque
  isError?: boolean;
}

interface Summary {
  totalWorked: number; // minutes travaillées réelles
  workDays: number;
  reqDays: number;
  absent: number;
  overtime: number; // minutes de dépassement brut
  undertime: number; // minutes de manque brut (absent + sous-temps)
  netBalance: number; // overtime - undertime (positif = crédit, négatif = débit)
  errors: number;
  dayEquiv: number; // totalWorked en jours
  avgReqPerDay: number; // exact average scheduled minutes per worked day
}

interface StatusConfig {
  label: string;
  bg: string;
  border: string;
  text: string;
}
interface SummaryCard {
  icon: string;
  label: string;
  val: string | number;
  sub: string;
  accent: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

const toMins = (t: string): number => {
  const [h, m] = t.trim().split(":").map(Number);
  return h * 60 + m;
};

// Round a minute value to the nearest quarter-hour (±10 min).
// Applied to worked duration and extra/under so that minor clock

const roundQ10 = (mins: number): number => Math.round(mins / 5) * 5;

const fmtHHMM = (m: number | null | undefined): string => {
  if (m == null || isNaN(m)) return "—";
  const abs = Math.abs(Math.round(m));
  return `${m < 0 ? "-" : ""}${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
};

const fmtAbs = (m: number | null | undefined): string => {
  if (m == null || isNaN(m)) return "—";
  const abs = Math.abs(Math.round(m));
  return `${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
};

// Collapse times that are within CLUSTER_GAP minutes of each other into one
// (keeps the earliest of the cluster). Also removes exact duplicates first.
const CLUSTER_GAP = 5; // minutes

const clusterTimes = (arr: string[]): string[] => {
  // 1. Exact dedup, then sort chronologically
  const sorted = [...new Set(arr)].sort((a, b) => toMins(a) - toMins(b));
  if (sorted.length === 0) return [];
  // 2. Walk sorted list; start a new cluster whenever gap > CLUSTER_GAP
  const clusters: string[] = [sorted[0]];
  for (let i = 1; i < sorted.length; i++) {
    const prev = toMins(clusters[clusters.length - 1]);
    const cur = toMins(sorted[i]);
    if (cur - prev > CLUSTER_GAP) {
      clusters.push(sorted[i]); // far enough — new cluster representative
    }
    // else: within 5 min of the cluster representative → discard
  }
  return clusters;
};

interface SessionResult {
  worked: number | null;
  workedRounded?: number;
  isError: boolean;
  isEmpty: boolean;
  deduped: string[]; // after clustering
}

const computeSession = (times: string[]): SessionResult => {
  if (!times || times.length === 0)
    return { worked: null, isError: false, isEmpty: true, deduped: [] };
  const d = clusterTimes(times);
  if (d.length === 1)
    return { worked: null, isError: true, isEmpty: false, deduped: d };
  if (d.length % 2 !== 0)
    return { worked: null, isError: true, isEmpty: false, deduped: d };
  let total = 0;
  for (let i = 0; i < d.length; i += 2)
    total += toMins(d[i + 1]) - toMins(d[i]);
  return {
    // worked: total,
    worked: roundQ10(total),
    isError: false,
    isEmpty: false,
    deduped: d,
  };
};

const JR_FR = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];

const getDayName = (dateStr: string): string => {
  const [dd, mm, yyyy] = dateStr.split("/").map(Number);
  return JR_FR[new Date(yyyy, mm - 1, dd).getDay()];
};

// ─── Parse Excel ──────────────────────────────────────────────────────────────

const parseFile = (
  buffer: Uint8Array,
): { emp: Employee | null; records: RawRecord[] } => {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
  let emp: Employee | null = null;
  const records: RawRecord[] = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];
    if (!row || !row[3]) continue;
    if (!emp && row[0])
      emp = {
        id: row[0] as string | number,
        name: row[1] as string,
        dept: row[2] as string,
      };
    const raw = row[3];
    let dateStr: string;
    if (typeof raw === "number") {
      const d = (XLSX as any).SSF.parse_date_code(raw);
      dateStr = `${String(d.d).padStart(2, "0")}/${String(d.m).padStart(2, "0")}/${d.y}`;
    } else {
      dateStr = String(raw).trim();
    }
    const pt = row[4] ? String(row[4]).trim() : "";
    const times = pt
      ? pt.split(/\s+/).filter((t) => /^\d{2}:\d{2}$/.test(t))
      : [];
    records.push({ date: dateStr, times });
  }
  return { emp, records };
};

// ─── Analyze ──────────────────────────────────────────────────────────────────

const analyze = (
  records: RawRecord[],
  daysOff: Set<string>,
  getReq: (d: string) => number,
  overrides: Record<string, Override>,
): AnalyzedDay[] =>
  records.map((r): AnalyzedDay => {
    const isDayOff = daysOff.has(r.date);
    const req = getReq(r.date);
    const ov = overrides[r.date];
    const { worked, isError, isEmpty, deduped } = computeSession(r.times);

    if (ov === "full")
      return {
        ...r,
        deduped,
        isDayOff,
        status: "working",
        worked: req,
        extra: 0,
      };
    if (ov === "half")
      return {
        ...r,
        deduped,
        isDayOff,
        status: "working",
        worked: req / 2,
        extra: -req / 2,
      };

    if (isDayOff) {
      if (isEmpty)
        return { ...r, deduped, isDayOff, status: "dayoff", worked: 0 };
      return {
        ...r,
        deduped,
        isDayOff,
        status: "dayoff_worked",
        worked,
        isError,
      };
    }
    if (isError)
      return { ...r, deduped, isDayOff, status: "error", worked: null };
    if (isEmpty)
      return { ...r, deduped, isDayOff, status: "absent", worked: 0 };
    return {
      ...r,
      deduped,
      isDayOff,
      status: "working",
      worked,
      extra: worked! - req,
    };
  });

// ─── Status labels (FR) ───────────────────────────────────────────────────────

const STATUS_CFG: Record<DayStatus, StatusConfig> = {
  working: {
    label: "Travaillé",
    bg: "#f0fdf4",
    border: "#86efac",
    text: "#166534",
  },
  absent: {
    label: "Absent",
    bg: "#fff7ed",
    border: "#fdba74",
    text: "#9a3412",
  },
  dayoff: { label: "Repos", bg: "#eff6ff", border: "#93c5fd", text: "#1d4ed8" },
  dayoff_worked: {
    label: "Repos + Trav",
    bg: "#fdf4ff",
    border: "#d8b4fe",
    text: "#7e22ce",
  },
  error: {
    label: "⚠ Erreur",
    bg: "#fff1f2",
    border: "#fca5a5",
    text: "#9f1239",
  },
};

const Chip = ({ status }: { status: DayStatus }): React.ReactNode => {
  const c = STATUS_CFG[status] ?? STATUS_CFG.working;
  return (
    <span
      style={{
        background: c.bg,
        border: `1.5px solid ${c.border}`,
        color: c.text,
        borderRadius: 20,
        padding: "2px 9px",
        fontSize: 10.5,
        fontWeight: 700,
        whiteSpace: "nowrap",
      }}
    >
      {c.label}
    </span>
  );
};

// ─── Constants ────────────────────────────────────────────────────────────────

const JOURS_SEMAINE: string[] = [
  "Dim",
  "Lun",
  "Mar",
  "Mer",
  "Jeu",
  "Ven",
  "Sam",
];

const getDatesForWeekdays = (
  records: RawRecord[],
  indices: Set<number>,
): Set<string> => {
  const out = new Set<string>();
  records.forEach(({ date }) => {
    const [dd, mm, yyyy] = date.split("/").map(Number);
    if (indices.has(new Date(yyyy, mm - 1, dd).getDay())) out.add(date);
  });
  return out;
};

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App(): React.ReactNode {
  const [data, setData] = useState<{
    emp: Employee | null;
    records: RawRecord[];
  } | null>(null);
  const [restWeekdays, setRestWeekdays] = useState<Set<number>>(new Set());
  const [extraDaysOff, setExtraDaysOff] = useState<Set<string>>(new Set());
  const [daysOffInput, setDaysOffInput] = useState<string>();
  const [periods, setPeriods] = useState<Period[]>([]);
  const [defaultHrs, setDefaultHrs] = useState<string>("9");
  const [overrides, setOverrides] = useState<Record<string, Override>>({});
  const [result, setResult] = useState<AnalyzedDay[] | null>(null);
  const printRef = useRef<HTMLDivElement>(null);

  const daysOff: Set<string> = data
    ? new Set([
        ...getDatesForWeekdays(data.records, restWeekdays),
        ...extraDaysOff,
      ])
    : extraDaysOff;

  // ── File loading ────────────────────────────────────────────────────────────
  const load = useCallback((e: ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const { emp, records } = parseFile(
        new Uint8Array(ev.target!.result as ArrayBuffer),
      );
      setData({ emp, records });
      setResult(null);
      setOverrides({});
      setRestWeekdays(new Set());
      setExtraDaysOff(new Set());
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // ── Weekday toggle ──────────────────────────────────────────────────────────
  const toggleWeekday = (idx: number): void => {
    const s = new Set(restWeekdays);
    s.has(idx) ? s.delete(idx) : s.add(idx);
    setRestWeekdays(s);
  };

  // ── Extra days off ──────────────────────────────────────────────────────────
  const addExtraDaysOff = (): void => {
    const vals =
      daysOffInput
        ?.split(",")
        .map((v) => v.trim())
        .filter(Boolean) ?? [];

    setExtraDaysOff(new Set([...extraDaysOff, ...vals]));
    setDaysOffInput("");
  };

  // ── Required minutes for a date ─────────────────────────────────────────────
  // Parse a DD/MM/YYYY string into a Date (midnight)
  const parseDMY = (s: string): Date | null => {
    const parts = s.trim().split("/");
    if (parts.length !== 3) return null;
    const [dd, mm, yyyy] = parts.map(Number);
    if (!dd || !mm || !yyyy || yyyy < 100) return null;
    return new Date(yyyy, mm - 1, dd);
  };

  const getReq = useCallback(
    (dateStr: string): number => {
      const [dd, mm, yyyy] = dateStr.split("/").map(Number);
      const d = new Date(yyyy, mm - 1, dd);
      for (const p of periods) {
        if (!p.from || !p.to) continue;
        try {
          const from = parseDMY(p.from);
          const to = parseDMY(p.to);
          if (from && to && d >= from && d <= to)
            return parseFloat(p.hours) * 60;
        } catch {}
      }
      return parseFloat(defaultHrs) * 60;
    },
    [periods, defaultHrs],
  );

  // ── Analyze ─────────────────────────────────────────────────────────────────
  const runAnalyze = (): void => {
    if (!data) return;
    setResult(analyze(data.records, daysOff, getReq, overrides));
  };

  // ── Overrides ────────────────────────────────────────────────────────────────
  const setOv = (date: string, val: Override): void => {
    const newOv = { ...overrides, [date]: val };
    setOverrides(newOv);
    if (result && data)
      setResult(analyze(data.records, daysOff, getReq, newOv));
  };

  const resetOv = (date: string): void => {
    const newOv = { ...overrides };
    delete newOv[date];
    setOverrides(newOv);
    if (result && data)
      setResult(analyze(data.records, daysOff, getReq, newOv));
  };

  // ── Summary with net balance ─────────────────────────────────────────────────
  // Logic:
  //   overtime  = sum of positive extras + dayoff_worked minutes
  //   undertime = sum of |negative extras| + absent days × required hours
  //   netBalance = overtime - undertime  (+ = crédit, - = débit)
  const summary: Summary | null = result
    ? (() => {
        let totalWorked = 0,
          workDays = 0,
          absent = 0,
          overtime = 0,
          undertime = 0,
          errors = 0;

        // Sum of required minutes for each actually-worked day — used to compute day-equivalent
        // accurately when different periods have different working hours.
        let totalReqMinsWorked = 0;
        const reqDays = result.filter((d) => !d.isDayOff).length;

        result.forEach((d) => {
          const req = getReq(d.date);
          if (d.status === "error") {
            errors++;
            undertime += req;
            return;
          }
          if (d.status === "absent") {
            absent++;
            undertime += req; // full day deducted
          }
          if (d.status === "working") {
            workDays++;
            totalWorked += d.worked ?? 0;
            totalReqMinsWorked += req; // accumulate this day's scheduled hours
            const e = d.extra ?? 0;
            if (e > 0) overtime += e;
            if (e < 0) undertime += Math.abs(e);
          }
          if (d.status === "dayoff_worked" && d.worked) {
            workDays++;
            totalReqMinsWorked += req; // count this day's scheduled hours toward the avg
            totalWorked += d.worked;
            overtime += d.worked;
          }
        });

        const netBalance = overtime - undertime;
        // Use raw integer totalReqMinsWorked/workDays — never derive from dayEquiv
        // to avoid floating-point cascade when computing breakdown.
        const avgReqPerDay =
          workDays > 0
            ? totalReqMinsWorked / workDays
            : (parseFloat(defaultHrs) || 9) * 60;
        const dayEquiv = totalWorked / avgReqPerDay;
        const fullDays = Math.floor(totalWorked / avgReqPerDay);
        const remainingMinutes = totalWorked % avgReqPerDay;
        const hours = Math.floor(remainingMinutes / 60);
        return {
          totalWorked,
          workDays,
          reqDays,
          absent,
          overtime,
          undertime,
          netBalance,
          errors,
          dayEquiv,
          avgReqPerDay, // expose so consumers never re-derive it
          fullDays,
          hours,
        };
      })()
    : null;

  // ── PDF Export ───────────────────────────────────────────────────────────────
  const handleExportPDF = async (): Promise<void> => {
    const jsPDFModule = await import(
      "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js" as any
    );
    const jsPDF: any =
      (jsPDFModule as any)?.jsPDF ??
      (jsPDFModule as any)?.default?.jsPDF ??
      (window as any).jspdf?.jsPDF ??
      (window as any).jsPDF;
    if (!jsPDF) {
      alert("Impossible de charger jsPDF — vérifiez votre connexion.");
      return;
    }

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
    });
    const W: number = doc.internal.pageSize.getWidth();
    const H: number = doc.internal.pageSize.getHeight();
    const mg = 12;
    let y = mg;

    const month = "mai";

    // Header
    doc.setFillColor(248, 250, 252);
    doc.rect(0, 0, W, 18, "F");
    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setFillColor(100, 116, 139);
    doc.text("RAPPORT DE POINTAGE", mg, 12);
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.setFillColor(100, 116, 139);
    doc.text(
      `Généré le ${new Date().toLocaleDateString("fr-FR")}`,
      W - mg,
      12,
      { align: "right" },
    );
    y = 26;

    // Employee
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text(data?.emp?.name ?? "", mg, y);
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(100, 116, 139);
    doc.text(
      `Matricule #${data?.emp?.id}  ·  ${data?.emp?.dept}  ·  ${month} 2026`,
      mg,
      y + 5,
    );

    const infoLines: string[] = [];
    if (restWeekdays.size > 0)
      infoLines.push(
        "Repos : " +
          JOURS_SEMAINE.filter((_, i) => restWeekdays.has(i)).join(", "),
      );

    if (infoLines.length) {
      doc.setTextColor(99, 102, 241);
      doc.setFontSize(9);
      doc.text(infoLines.join("  ·  "), W - mg, y, { align: "right" });
    }
    y += 14;

    // Summary boxes
    if (summary) {
      const avgReqPerDay = summary.avgReqPerDay;

      const boxes: Array<{
        label: string;
        val: string;
        sub: string;
        color: [number, number, number];
      }> = [
        {
          label: "Heures travaillées",
          val: fmtAbs(summary.totalWorked),
          sub: (() => {
            const fullDays = Math.floor(summary.totalWorked / avgReqPerDay);
            const remainingMinutes = summary.totalWorked % avgReqPerDay;
            const hours = Math.floor(remainingMinutes / 60);
            return `${fullDays}J ${hours}h`;
          })(),
          color: [99, 102, 241],
        },
        {
          label: "Jours travaillés",
          val: `${summary.workDays} / ${summary.reqDays}`,
          sub: "requis",
          color: [14, 165, 233],
        },
        {
          label: "Absences",
          val: `${summary.absent}`,
          sub: "jours",
          color: [249, 115, 22],
        },
        {
          label: "Heures sup.",
          val: fmtAbs(summary.overtime),
          sub: "crédit brut",
          color: [16, 185, 129],
        },
        {
          label: "Heures manq.",
          val: fmtAbs(summary.undertime),
          sub: "débit brut",
          color: [239, 68, 68],
        },
        (() => {
          const absNet = Math.abs(summary.netBalance);
          const fullDays = Math.floor(absNet / avgReqPerDay);
          const remainingMinutes = absNet % avgReqPerDay;
          const hours = Math.floor(remainingMinutes / 60);
          const sub =
            absNet >= 60
              ? `${fullDays}J ${hours}h`
              : summary.netBalance >= 0
                ? "crédit"
                : "débit";
          return {
            label: "Solde net",
            val:
              (summary.netBalance >= 0 ? "+" : "") +
              fmtHHMM(summary.netBalance),
            sub,
            color: (summary.netBalance >= 0
              ? [16, 185, 129]
              : [239, 68, 68]) as [number, number, number],
          };
        })(),
        ...(summary.errors
          ? [
              {
                label: "Erreurs",
                val: `${summary.errors}`,
                sub: "à vérifier",
                color: [244, 63, 94] as [number, number, number],
              },
            ]
          : []),
      ];
      const bw = (W - mg * 2 - (boxes.length - 1) * 2.5) / boxes.length;
      boxes.forEach((b, i) => {
        const bx = mg + i * (bw + 2.5);
        doc.setFillColor(248, 250, 252);
        doc.roundedRect(bx, y, bw, 19, 2, 2, "F");
        doc.setFillColor(...b.color);
        doc.rect(bx, y, 2.5, 19, "F");
        doc.setTextColor(100, 116, 139);
        doc.setFontSize(6);
        doc.setFont("helvetica", "bold");
        doc.text(b.label.toUpperCase(), bx + 4, y + 5.5);
        doc.setTextColor(...b.color);
        doc.setFontSize(10);
        doc.setFont("helvetica", "bold");
        doc.text(b.val, bx + 4, y + 12.5);
        doc.setTextColor(148, 163, 184);
        doc.setFontSize(7);
        doc.setFont("helvetica", "normal");
        doc.text(b.sub, bx + 4, y + 17);
      });
      y += 25;
    }

    // Table
    const cols: Array<{ header: string; w: number }> = [
      { header: "DATE", w: 20 },
      { header: "JOUR", w: 9 },
      { header: "POINTAGES", w: 48 },
      { header: "TRAVAILLÉ", w: 17 },
      { header: "REQUIS", w: 17 },
      { header: "SUP/MANQ", w: 20 },
      { header: "STATUT", w: 22 },
    ];
    const totalW = cols.reduce((s, c) => s + c.w, 0);
    const scale = (W - mg * 2) / totalW;
    const sw = cols.map((c) => c.w * scale);
    let cx = mg;
    const cx_arr = sw.map((w) => {
      const x = cx;
      cx += w;
      return x;
    });
    const rh = 5.8;

    doc.setFillColor(100, 116, 139);
    doc.rect(mg, y, W - mg * 2, rh, "F");
    doc.setTextColor(226, 232, 240);
    doc.setFontSize(6);
    doc.setFont("helvetica", "bold");
    cols.forEach((c, i) => doc.text(c.header, cx_arr[i] + 2, y + 3.8));
    y += rh;

    const rowBg: Record<DayStatus, [number, number, number]> = {
      working: [255, 255, 255],
      absent: [255, 247, 237],
      dayoff: [239, 246, 255],
      dayoff_worked: [253, 244, 255],
      error: [255, 241, 242],
    };
    const stLabel: Record<DayStatus, string> = {
      working: "Travaillé",
      absent: "Absent",
      dayoff: "Repos",
      dayoff_worked: "Repos+Trav",
      error: "! Erreur",
    };
    const stColor: Record<DayStatus, [number, number, number]> = {
      working: [22, 101, 52],
      absent: [154, 52, 18],
      dayoff: [29, 78, 216],
      dayoff_worked: [126, 34, 206],
      error: [159, 18, 57],
    };

    if (result)
      result.forEach((d) => {
        if (y + rh > H - 10) {
          doc.addPage();
          y = mg;
        }
        const req = getReq(d.date);
        const bg = rowBg[d.status];
        doc.setFillColor(...bg);
        doc.rect(mg, y, W - mg * 2, rh, "F");
        doc.setDrawColor(226, 232, 240);
        doc.setLineWidth(0.15);
        doc.line(mg, y + rh, W - mg, y + rh);
        const ty = y + 3.8;

        doc.setTextColor(30, 41, 59);
        doc.setFontSize(7);
        doc.setFont("helvetica", "bold");
        doc.text(d.date, cx_arr[0] + 2, ty);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100, 116, 139);
        doc.text(getDayName(d.date), cx_arr[1] + 2, ty);
        doc.setTextColor(71, 85, 105);
        doc.text(
          doc.splitTextToSize(
            d.deduped?.length ? d.deduped.join("  ") : "—",
            sw[2] - 3,
          )[0],
          cx_arr[2] + 2,
          ty,
        );
        doc.setTextColor(30, 41, 59);
        doc.setFont("helvetica", "bold");
        doc.text(
          d.worked != null && d.status !== "absent" ? fmtAbs(d.worked) : "—",
          cx_arr[3] + 2,
          ty,
        );
        doc.setFont("helvetica", "normal");
        doc.setTextColor(148, 163, 184);
        doc.text(d.isDayOff ? "—" : fmtAbs(req), cx_arr[4] + 2, ty);

        if (d.status === "working" && d.extra != null && d.extra !== 0) {
          const pos = d.extra > 0;
          doc.setTextColor(pos ? 5 : 220, pos ? 150 : 38, pos ? 105 : 38);
          doc.setFont("helvetica", "bold");
          doc.text(
            (d.extra > 0 ? "+" : "") + fmtHHMM(d.extra),
            cx_arr[5] + 2,
            ty,
          );
        } else if (d.status === "absent") {
          doc.setTextColor(220, 38, 38);
          doc.setFont("helvetica", "bold");
          doc.text("-" + fmtAbs(req), cx_arr[5] + 2, ty);
        } else if (d.status === "dayoff_worked" && d.worked) {
          doc.setTextColor(124, 58, 237);
          doc.setFont("helvetica", "bold");
          doc.text("+" + fmtAbs(d.worked), cx_arr[5] + 2, ty);
        } else {
          doc.setTextColor(200, 200, 200);
          doc.setFont("helvetica", "normal");
          doc.text("—", cx_arr[5] + 2, ty);
        }

        doc.setTextColor(...stColor[d.status]);
        doc.setFontSize(6);
        doc.setFont("helvetica", "bold");
        doc.text(stLabel[d.status], cx_arr[6] + 3, ty);
        y += rh;
      });

    // Footer
    doc.setFillColor(248, 250, 252);
    doc.rect(0, H - 8, W, 8, "F");
    doc.setDrawColor(226, 232, 240);
    doc.setLineWidth(0.3);
    doc.line(0, H - 8, W, H - 8);
    doc.setTextColor(148, 163, 184);
    doc.setFontSize(7.5);
    doc.setFont("helvetica", "normal");
    doc.text(
      `Rapport de Pointage · ${data?.emp?.name} · ${data?.emp?.dept} · ${month} 2026`,
      mg,
      H - 3,
    );
    doc.text(`${doc.internal.getNumberOfPages()} page(s)`, W - mg, H - 3, {
      align: "right",
    });
    doc.save(
      `Pointage_${(data?.emp?.name ?? "rapport").replace(/\s+/g, "_")}_${month}_2026.pdf`,
    );
  };

  // ─── Render ───────────────────────────────────────────────────────────────────

  return (
    <div
      ref={printRef}
      style={{
        fontFamily: "'DM Sans', system-ui, sans-serif",
        background: "#f8fafc",
        minHeight: "100vh",
        minWidth: "100vw",
        color: "#0f172a",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; }
        .mono { font-family: 'DM Mono', monospace; }
        .card { background: white; border-radius: 16px; border: 1px solid #e2e8f0; padding: 20px 24px; }
        .btn  { cursor: pointer; border: none; border-radius: 9px; font-weight: 600; font-size: 13px; padding: 8px 16px; transition: .15s; font-family: inherit; }
        .btn:hover { filter: brightness(.93); }
        input[type=text], input[type=number], input[type=date] {
          border: 1.5px solid #e2e8f0; border-radius: 8px; padding: 6px 10px;
          font-size: 13px; outline: none; font-family: inherit; background: white;
        }
        input[type=date] { cursor: pointer; }
        input:focus { border-color: #6366f1; }
        table { width: 100%; border-collapse: collapse; }
        thead th { background: #0f172a; color: #e2e8f0; font-size: 11px; font-weight: 600;
          text-transform: uppercase; letter-spacing: .5px; padding: 10px; text-align: left;
          position: sticky; top: 0; z-index: 2; }
        tbody tr { border-bottom: 1px solid #f1f5f9; }
        tbody tr:hover td { background: rgba(99,102,241,.04) !important; }
        td { padding: 7px 10px; font-size: 12.5px; vertical-align: middle; }
        .row-error        td { background: #fff1f2; }
        .row-dayoff       td { background: #eff6ff; }
        .row-absent       td { background: #fff7ed; }
        .row-dayoff_worked td { background: #fdf4ff; }
        .row-working      td { background: white; }
      `}</style>

      {/* ── Barre de navigation ─────────────────────────────────────────────── */}
      <div
        style={{
          background: "#0f172a",
          padding: "16px 32px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <div>
          <div
            className="mono"
            style={{
              fontSize: 10,
              color: "#475569",
              letterSpacing: 2,
              textTransform: "uppercase",
            }}
          >
            RH · Présences
          </div>
          <div style={{ color: "white", fontWeight: 700, fontSize: 20 }}>
            Analyseur de Pointage
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {data && (
            <label
              className="btn"
              style={{ background: "#1e293b", color: "#94a3b8", fontSize: 12 }}
            >
              ↩ Nouveau fichier
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e: ChangeEvent<HTMLInputElement>) => {
                  load(e);
                  e.target.value = "";
                }}
                style={{ display: "none" }}
              />
            </label>
          )}
          {result && (
            <button
              className="btn"
              onClick={handleExportPDF}
              style={{
                background: "#6366f1",
                color: "white",
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <span style={{ fontSize: 15 }}>⬇</span> Exporter PDF
            </button>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1160, margin: "0 auto", padding: "24px 16px" }}>
        {/* ── Zone d'import ──────────────────────────────────────────────────── */}
        {!data ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              paddingTop: 60,
            }}
          >
            <label
              style={{
                border: "2px dashed #cbd5e1",
                borderRadius: 20,
                padding: "60px 80px",
                textAlign: "center",
                cursor: "pointer",
                background: "white",
                display: "block",
                transition: ".2s",
              }}
              onMouseOver={(e) =>
                (e.currentTarget.style.borderColor = "#6366f1")
              }
              onMouseOut={(e) =>
                (e.currentTarget.style.borderColor = "#cbd5e1")
              }
            >
              <div style={{ fontSize: 48, marginBottom: 8 }}>📋</div>
              <div style={{ fontWeight: 700, fontSize: 18 }}>
                Importer un fichier Excel de pointage
              </div>
              <div style={{ fontSize: 13, color: "#94a3b8", marginTop: 4 }}>
                Cliquez ou glissez-déposez votre fichier .xlsx
              </div>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={load}
                style={{ display: "none" }}
              />
            </label>
          </div>
        ) : (
          <>
            {/* ── Infos employé ────────────────────────────────────────────── */}
            <div
              className="card"
              style={{
                marginBottom: 16,
                display: "flex",
                gap: 20,
                alignItems: "center",
              }}
            >
              <div
                style={{
                  width: 44,
                  height: 44,
                  borderRadius: 12,
                  background: "#e0e7ff",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 20,
                }}
              >
                👤
              </div>
              <div>
                <div style={{ fontWeight: 700, fontSize: 17 }}>
                  {data.emp?.name}
                  <span
                    className="mono"
                    style={{ color: "#94a3b8", fontSize: 12 }}
                  >
                    {" "}
                    #{data.emp?.id}
                  </span>
                </div>
                <div style={{ fontSize: 13, color: "#64748b" }}>
                  {data.emp?.dept} · Mars 2026 · {data.records.length}{" "}
                  enregistrements
                </div>
              </div>
            </div>

            {/* ── Configuration ────────────────────────────────────────────── */}
            <div className="card" style={{ marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}>
                <span
                  style={{
                    background: "#f1f5f9",
                    borderRadius: 8,
                    padding: "4px 10px",
                    fontSize: 12,
                    color: "#475569",
                  }}
                >
                  ⚙ Configuration
                </span>
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 24,
                }}
              >
                {/* Jours de repos */}
                <div>
                  <div
                    style={{
                      fontWeight: 600,
                      fontSize: 12,
                      color: "#475569",
                      marginBottom: 10,
                      textTransform: "uppercase",
                      letterSpacing: 0.5,
                    }}
                  >
                    Jours de repos
                  </div>

                  <div
                    style={{
                      fontSize: 12,
                      color: "#64748b",
                      marginBottom: 6,
                      fontWeight: 500,
                    }}
                  >
                    Jour(s) de repos hebdomadaire
                  </div>
                  <div
                    style={{
                      display: "flex",
                      gap: 6,
                      marginBottom: 14,
                      flexWrap: "wrap",
                    }}
                  >
                    {JOURS_SEMAINE.map((jour, idx) => {
                      const active = restWeekdays.has(idx);
                      return (
                        <button
                          key={jour}
                          onClick={() => toggleWeekday(idx)}
                          className="btn"
                          style={{
                            padding: "6px 12px",
                            fontSize: 12,
                            fontWeight: 700,
                            background: active ? "#0f172a" : "#f1f5f9",
                            color: active ? "white" : "#64748b",
                            border: active
                              ? "2px solid #0f172a"
                              : "2px solid #e2e8f0",
                            borderRadius: 10,
                            boxShadow: active
                              ? "0 2px 8px rgba(15,23,42,.25)"
                              : "none",
                          }}
                        >
                          {jour}
                        </button>
                      );
                    })}
                  </div>
                  {restWeekdays.size > 0 && (
                    <div
                      style={{
                        fontSize: 11,
                        color: "#6366f1",
                        marginBottom: 12,
                        background: "#eef2ff",
                        borderRadius: 8,
                        padding: "6px 10px",
                      }}
                    >
                      {JOURS_SEMAINE.filter((_, i) => restWeekdays.has(i)).join(
                        " & ",
                      )}{" "}
                      marqués comme repos →{" "}
                      <strong>
                        {getDatesForWeekdays(data.records, restWeekdays).size}{" "}
                        jours
                      </strong>{" "}
                      ce mois
                    </div>
                  )}

                  <div
                    style={{
                      fontSize: 12,
                      color: "#64748b",
                      marginBottom: 6,
                      fontWeight: 500,
                    }}
                  >
                    Dates spécifiques supplémentaires
                  </div>
                  <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                    <input
                      type="text"
                      placeholder="JJ/MM/AAAA, JJ/MM/AAAA…"
                      value={daysOffInput}
                      onChange={(e: ChangeEvent<HTMLInputElement>) =>
                        setDaysOffInput(e.target.value)
                      }
                      onKeyDown={(e: KeyboardEvent<HTMLInputElement>) =>
                        e.key === "Enter" && addExtraDaysOff()
                      }
                      style={{ flex: 1 }}
                    />
                    <button
                      className="btn"
                      style={{ background: "#6366f1", color: "white" }}
                      onClick={addExtraDaysOff}
                    >
                      Ajouter
                    </button>
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {[...extraDaysOff].sort().map((d) => (
                      <span
                        key={d}
                        style={{
                          background: "#dbeafe",
                          color: "#1d4ed8",
                          borderRadius: 20,
                          padding: "3px 10px",
                          fontSize: 11.5,
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 5,
                        }}
                      >
                        {d}
                        <span
                          onClick={() => {
                            const s = new Set(extraDaysOff);
                            s.delete(d);
                            setExtraDaysOff(s);
                          }}
                          style={{
                            cursor: "pointer",
                            color: "#93c5fd",
                            fontWeight: 700,
                            fontSize: 14,
                            lineHeight: 1,
                          }}
                        >
                          ×
                        </span>
                      </span>
                    ))}
                    {extraDaysOff.size === 0 && (
                      <span style={{ color: "#cbd5e1", fontSize: 11 }}>
                        Aucune date spécifique ajoutée
                      </span>
                    )}
                  </div>
                </div>

                {/* Horaires */}
                <div>
                  <div
                    style={{
                      fontWeight: 600,
                      fontSize: 12,
                      color: "#475569",
                      marginBottom: 8,
                      textTransform: "uppercase",
                      letterSpacing: 0.5,
                    }}
                  >
                    Horaires
                  </div>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      marginBottom: 12,
                    }}
                  >
                    <span style={{ fontSize: 13, color: "#64748b" }}>
                      Heures par défaut / jour :
                    </span>
                    <input
                      type="number"
                      value={defaultHrs}
                      min={1}
                      max={24}
                      step={0.5}
                      onChange={(e: ChangeEvent<HTMLInputElement>) =>
                        setDefaultHrs(e.target.value)
                      }
                      style={{ width: 64 }}
                    />
                    <span style={{ fontSize: 12, color: "#94a3b8" }}>h</span>
                  </div>

                  <div
                    style={{ fontSize: 12, color: "#94a3b8", marginBottom: 8 }}
                  >
                    Périodes d'exception (Ramadan, etc.)
                  </div>
                  {periods.map((p, i) => (
                    <div
                      key={i}
                      style={{
                        display: "flex",
                        gap: 6,
                        alignItems: "center",
                        marginBottom: 8,
                        flexWrap: "wrap",
                        background: "#f8fafc",
                        borderRadius: 10,
                        padding: "8px 10px",
                        border: "1px solid #e2e8f0",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: 2,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 10,
                            color: "#94a3b8",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: 0.5,
                          }}
                        >
                          Du
                        </span>
                        <input
                          type="text"
                          value={p.from}
                          placeholder="JJ/MM/AAAA"
                          maxLength={10}
                          onChange={(e: ChangeEvent<HTMLInputElement>) => {
                            let v = e.target.value.replace(/[^\d/]/g, "");
                            // Auto-insert slashes at positions 2 and 5
                            if (v.length === 2 && !v.includes("/")) v = v + "/";
                            else if (
                              v.length === 5 &&
                              v.split("/").length === 2
                            )
                              v = v + "/";
                            const s = [...periods];
                            s[i] = { ...s[i], from: v };
                            setPeriods(s);
                          }}
                          style={{ width: 120, letterSpacing: 1 }}
                        />
                      </div>
                      <span
                        style={{
                          color: "#cbd5e1",
                          fontSize: 16,
                          marginTop: 14,
                        }}
                      >
                        →
                      </span>
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: 2,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 10,
                            color: "#94a3b8",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: 0.5,
                          }}
                        >
                          Au
                        </span>
                        <input
                          type="text"
                          value={p.to}
                          placeholder="JJ/MM/AAAA"
                          maxLength={10}
                          onChange={(e: ChangeEvent<HTMLInputElement>) => {
                            let v = e.target.value.replace(/[^\d/]/g, "");
                            if (v.length === 2 && !v.includes("/")) v = v + "/";
                            else if (
                              v.length === 5 &&
                              v.split("/").length === 2
                            )
                              v = v + "/";
                            const s = [...periods];
                            s[i] = { ...s[i], to: v };
                            setPeriods(s);
                          }}
                          style={{ width: 120, letterSpacing: 1 }}
                        />
                      </div>
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: 2,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 10,
                            color: "#94a3b8",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: 0.5,
                          }}
                        >
                          H / jour
                        </span>
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 4,
                          }}
                        >
                          <input
                            type="number"
                            value={p.hours}
                            min={1}
                            max={24}
                            step={0.5}
                            onChange={(e: ChangeEvent<HTMLInputElement>) => {
                              const s = [...periods];
                              s[i] = { ...s[i], hours: e.target.value };
                              setPeriods(s);
                            }}
                            style={{ width: 58 }}
                          />
                          <span style={{ fontSize: 12, color: "#94a3b8" }}>
                            h
                          </span>
                        </div>
                      </div>
                      {/^\d{2}\/\d{2}\/\d{4}$/.test(p.from) &&
                        /^\d{2}\/\d{2}\/\d{4}$/.test(p.to) && (
                          <div
                            style={{
                              marginTop: 14,
                              fontSize: 11,
                              color: "#6366f1",
                              fontWeight: 600,
                              background: "#eef2ff",
                              borderRadius: 6,
                              padding: "3px 8px",
                            }}
                          >
                            {p.hours}h/j
                          </div>
                        )}
                      <button
                        className="btn"
                        style={{
                          background: "#fee2e2",
                          color: "#b91c1c",
                          padding: "4px 9px",
                          fontSize: 12,
                          marginTop: 14,
                          marginLeft: "auto",
                        }}
                        onClick={() =>
                          setPeriods(periods.filter((_, j) => j !== i))
                        }
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                  <button
                    className="btn"
                    style={{
                      background: "#f1f5f9",
                      color: "#475569",
                      fontSize: 12,
                      marginTop: 2,
                    }}
                    onClick={() =>
                      setPeriods([...periods, { from: "", to: "", hours: "8" }])
                    }
                  >
                    + Ajouter une période
                  </button>
                </div>
              </div>

              <div
                style={{
                  borderTop: "1px solid #f1f5f9",
                  marginTop: 16,
                  paddingTop: 14,
                }}
              >
                <button
                  className="btn"
                  style={{
                    background: "#0f172a",
                    color: "white",
                    fontSize: 14,
                    padding: "10px 24px",
                  }}
                  onClick={runAnalyze}
                >
                  🔍 Analyser le pointage
                </button>
              </div>
            </div>

            {/* ── Cartes récapitulatives ────────────────────────────────────── */}
            {summary && (
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))",
                  gap: 12,
                  marginBottom: 20,
                }}
              >
                {(
                  [
                    {
                      icon: "⏱",
                      label: "Heures travaillées",
                      val: fmtAbs(summary.totalWorked),
                      sub: (() => {
                        const avg = summary.avgReqPerDay;
                        const fullDays = Math.floor(summary.totalWorked / avg);
                        const remainMins = summary.totalWorked % avg;
                        const hours = Math.floor(remainMins / 60);
                        return `${fullDays}J ${hours}h`;
                      })(),
                      accent: "#6366f1",
                    },
                    {
                      icon: "📅",
                      label: "Jours travaillés",
                      val: `${summary.workDays}`,
                      sub: `sur ${summary.reqDays} requis`,
                      accent: "#0ea5e9",
                    },
                    {
                      icon: "🔴",
                      label: "Absences",
                      val: summary.absent,
                      sub: "jours",
                      accent: "#f97316",
                    },
                    {
                      icon: "⬆",
                      label: "Heures sup.",
                      val: fmtAbs(summary.overtime),
                      sub: "crédit brut",
                      accent: "#10b981",
                    },
                    {
                      icon: "⬇",
                      label: "Heures manq.",
                      val: fmtAbs(summary.undertime),
                      sub: "débit brut",
                      accent: "#ef4444",
                    },
                    {
                      icon: summary.netBalance >= 0 ? "✅" : "❌",
                      label: "Solde net",
                      val:
                        (summary.netBalance >= 0 ? "+" : "") +
                        fmtHHMM(summary.netBalance),
                      sub: (() => {
                        const absNet = Math.abs(summary.netBalance);
                        if (absNet >= 60) {
                          const avg = summary.avgReqPerDay;
                          const fullDays = Math.floor(absNet / avg);
                          const remainMins = absNet % avg;
                          const hours = Math.floor(remainMins / 60);
                          return `${fullDays}J ${hours}h`;
                        }
                        return summary.netBalance >= 0 ? "crédit" : "débit";
                      })(),
                      accent: summary.netBalance >= 0 ? "#10b981" : "#ef4444",
                    },
                    ...(summary.errors
                      ? [
                          {
                            icon: "⚠",
                            label: "Erreurs",
                            val: summary.errors,
                            sub: "à vérifier",
                            accent: "#f43f5e",
                          },
                        ]
                      : []),
                  ] as SummaryCard[]
                ).map((s) => (
                  <div
                    key={s.label}
                    className="card"
                    style={{
                      padding: "14px 16px",
                      borderLeft: `4px solid ${s.accent}`,
                    }}
                  >
                    <div
                      style={{
                        fontSize: 10,
                        color: "#94a3b8",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: 0.5,
                      }}
                    >
                      {s.icon} {s.label}
                    </div>
                    <div
                      className="mono"
                      style={{
                        fontSize: 22,
                        fontWeight: 700,
                        color: s.accent,
                        marginTop: 2,
                      }}
                    >
                      {s.val}
                    </div>
                    <div style={{ fontSize: 11, color: "#94a3b8" }}>
                      {s.sub}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* ── Légende ──────────────────────────────────────────────────── */}
            {result && (
              <div
                style={{
                  display: "flex",
                  gap: 10,
                  marginBottom: 12,
                  flexWrap: "wrap",
                }}
              >
                {(
                  Object.entries(STATUS_CFG) as [DayStatus, StatusConfig][]
                ).map(([k, v]) => (
                  <span
                    key={k}
                    style={{
                      background: v.bg,
                      border: `1.5px solid ${v.border}`,
                      color: v.text,
                      borderRadius: 20,
                      padding: "3px 12px",
                      fontSize: 11,
                      fontWeight: 600,
                    }}
                  >
                    {v.label}
                  </span>
                ))}
              </div>
            )}

            {/* ── Tableau ───────────────────────────────────────────────────── */}
            {result && (
              <div
                className="card"
                style={{ padding: 0, overflow: "hidden", borderRadius: 16 }}
              >
                <div style={{ overflowX: "auto" }}>
                  <table>
                    <thead>
                      <tr>
                        {[
                          "Date",
                          "Jour",
                          "Pointages",
                          "Heures trav.",
                          "Requis",
                          "Sup / Manq",
                          "Statut",
                          "Action",
                        ].map((h) => (
                          <th key={h}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {result.map((d) => {
                        const req = getReq(d.date);
                        return (
                          <tr key={d.date} className={`row-${d.status}`}>
                            <td
                              className="mono"
                              style={{ fontWeight: 700, color: "#1e293b" }}
                            >
                              {d.date}
                            </td>
                            <td style={{ color: "#64748b", fontSize: 11.5 }}>
                              {getDayName(d.date)}
                            </td>
                            <td
                              className="mono"
                              style={{ fontSize: 11, color: "#334155" }}
                            >
                              {d.deduped?.length ? (
                                d.deduped.join(" · ")
                              ) : (
                                <span style={{ color: "#d1d5db" }}>—</span>
                              )}
                            </td>
                            <td className="mono" style={{ fontWeight: 700 }}>
                              {d.worked != null && d.status !== "absent" ? (
                                fmtAbs(d.worked)
                              ) : (
                                <span style={{ color: "#cbd5e1" }}>—</span>
                              )}
                            </td>
                            <td
                              className="mono"
                              style={{ color: "#94a3b8", fontSize: 11.5 }}
                            >
                              {d.isDayOff ? (
                                <span style={{ color: "#bfdbfe" }}>—</span>
                              ) : (
                                fmtAbs(req)
                              )}
                            </td>
                            <td className="mono" style={{ fontWeight: 700 }}>
                              {d.status === "working" &&
                              d.extra != null &&
                              d.extra !== 0 ? (
                                <span
                                  style={{
                                    color: d.extra > 0 ? "#059669" : "#dc2626",
                                  }}
                                >
                                  {d.extra > 0 ? "+" : ""}
                                  {fmtHHMM(d.extra)}
                                </span>
                              ) : d.status === "absent" ? (
                                <span style={{ color: "#dc2626" }}>
                                  -{fmtAbs(req)}
                                </span>
                              ) : d.status === "dayoff_worked" && d.worked ? (
                                <span style={{ color: "#7c3aed" }}>
                                  +{fmtAbs(d.worked)}
                                </span>
                              ) : (
                                <span style={{ color: "#e2e8f0" }}>—</span>
                              )}
                            </td>
                            <td>
                              <Chip status={d.status} />
                            </td>
                            <td>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button
                                  className="btn"
                                  style={{
                                    background: "#d1fae5",
                                    color: "#065f46",
                                    padding: "3px 10px",
                                    fontSize: 11,
                                  }}
                                  onClick={() => setOv(d.date, "full")}
                                >
                                  Journée
                                </button>
                                <button
                                  className="btn"
                                  style={{
                                    background: "#fef9c3",
                                    color: "#854d0e",
                                    padding: "3px 10px",
                                    fontSize: 11,
                                  }}
                                  onClick={() => setOv(d.date, "half")}
                                >
                                  Demi-j.
                                </button>
                                {overrides[d.date] && (
                                  <button
                                    className="btn"
                                    style={{
                                      background: "#fee2e2",
                                      color: "#b91c1c",
                                      padding: "3px 8px",
                                      fontSize: 10,
                                      marginTop: 2,
                                    }}
                                    onClick={() => resetOv(d.date)}
                                  >
                                    Réinitialiser
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {result && (
              <div
                style={{
                  marginTop: 16,
                  fontSize: 11,
                  color: "#94a3b8",
                  textAlign: "center",
                }}
              >
                Rapport généré pour {data.emp?.name} · {data.emp?.dept} · 2026
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
