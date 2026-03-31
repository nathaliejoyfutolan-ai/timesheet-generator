/**
 * NOC Sprout Timesheet Generator (FULL SCRIPT)
 *
 * UPDATE:
 * A) Raw_Invoice is treated as PURELY RAW INPUT (no workshift fields required or used there).
 * B) Workshift Start and End are read from TBRegularOT on "Regular vs Overage".
 *    Manual columns are preserved on every run:
 *      - Col A Workshift Start (manual if specified)
 *      - Col B Workshift End (manual if specified)
 *
 * NEW DEFAULT BEHAVIOR:
 * If Workshift Start / End is blank for an employee, assume:
 *   - Workshift Start = 11:00 PM
 *   - Workshift End   = 8:00 AM
 *
 * TBRegularOT behavior (Regular vs Overage table):
 * Manual if specified:
 *  - Col A Workshift Start
 *  - Col B Workshift End
 *
 * Autodetect from Raw_Sprout each run (employee list varies):
 *  - Col C Employee key (ONLY the numeric ID)
 *  - Col D Employee name
 *
 * Auto-calc each run:
 *  - Remaining columns in TBRegularOT (totals, costs, etc)
 *
 * ALSO INCLUDED:
 * 1) LOCAL Excel date parsing (no day shifting).
 * 2) OT logic: Paid OT comes ONLY from Raw_Approval Center Overtime approvals AND is limited by Raw_Invoice otHours pools.
 * 3) Outside-shift hours from Sprout logs remain UNPAID by default (unpaidH).
 * 4) Overnight shift-day attribution with 3-hour post-end grace for grouping.
 * 5) Overage day rollups split Paid vs Unpaid so "Paid" rows show ONLY paidOverH.
 * 6) Approval tags (SIL, COA, Official Business, Overtime) grouped under project code KCG-CORP,
 *    with different Project Name labels.
 * 7) Populate Dashboard (2) flat tables, starting Row 4:
 *    - Employee table A-F: Employee ID, Employee Name, Project Code, Project Name, Hours, Cost
 *    - Project table  H-M: Project Code, Project Name, Employee ID, Employee Name, Hours, Cost
 *
 * IMPORTANT:
 * TBRegularOT Workshift Start and Workshift End are now WRITTEN AS TEXT
 * like "11:00 PM" and "8:00 AM" so they link smoothly with Power Apps.
 */

const MAX_SEGMENT_HOURS = 12;

// Default shift fallback
const DEFAULT_SHIFT_START_MIN = 23 * 60; // 11:00 PM
const DEFAULT_SHIFT_END_MIN = 8 * 60;    // 8:00 AM

// Shift-day grace after shift end (overnight)
const SHIFT_POST_END_GRACE_MIN = 180; // 3 hours

// Corporate project code for approval tags
const CORP_PROJECT_CODE = "KCG-CORP";

// Night diff window: 10:00 PM to 6:00 AM
const NIGHT_START_MIN = 22 * 60;
const NIGHT_END_MIN = 6 * 60;

// TBRegularOT table name (in "Regular vs Overage")
const REGOT_TABLE = "TBRegularOT";

// Dashboard upload signature storage
const DASH_META_SIGNATURE_CELL = "Z1";

const SHEETS = {
  RAW_SPROUT: "Raw_Sprout",
  RAW_INVOICE: "Raw_Invoice",
  RAW_APPROVAL: "Raw_Approval Center",
  DASH: "Dashboard",
  DASH2: "Dashboard (2)",
  EMP: "Employees",
  PROJ: "Projects",
  REGVSOV: "Regular vs Overage",
  ADJUST: "Adjust Hours",
  RATES: "Rates"
};

// Dashboard blocks (original Dashboard)
const DASH_EMP = { headerRow: 11, startCol: 1, cols: 8 };    // A-H
const DASH_PROJ = { headerRow: 11, startCol: 10, cols: 8 };  // J-Q
const DASH_CLEAR_ROWS = 2500;

// Employees tab blocks
const EMP_SUM = { headerRow: 3, startCol: 1, cols: 8, clearRows: 6000 };
const EMP_REG = { headerRow: 3, startCol: 11, cols: 5, clearRows: 6000 }; // K-O
const EMP_OV = { headerRow: 3, startCol: 17, cols: 6, clearRows: 6000 };  // Q-V

// Projects tab blocks
const PROJ_SUM = { headerRow: 3, startCol: 1, cols: 8, clearRows: 6000 };
const PROJ_REG = { headerRow: 3, startCol: 11, cols: 5, clearRows: 6000 }; // K-O
const PROJ_OV = { headerRow: 3, startCol: 17, cols: 6, clearRows: 6000 };  // Q-V

// Adjust Hours
const ADJ = { headerRow: 1, startCol: 1, cols: 12, clearRows: 6000 };

// Dashboard (2) flat tables: headers row 3, data starts row 4
const DASH2_EMP = { headerRow: 3, startCol: 1, cols: 6, clearRows: 12000 }; // A-F
const DASH2_PROJ = { headerRow: 3, startCol: 8, cols: 6, clearRows: 12000 }; // H-M

type SproutRow = {
  empId: string;
  empName: string;
  dt: Date;
  inOut: "IN" | "OUT" | "OTHER";
  projectCode: string;
  projectName: string;
  rawRowIndex1: number;
};

type InvoiceRow = {
  empName: string;
  monthlyCostUsd: number;
  otHours1st: number;
  otHours2nd: number;
};

type ShiftRow = {
  empId: string;
  empName: string;
  workshiftStartMin: number;
  workshiftEndMin: number;
};

type Segment = {
  empId: string;
  empName: string;
  projectCode: string;
  projectName: string;
  start: Date;
  end: Date;

  regularH: number;
  unpaidH: number;
  paidOverH: number;

  startRow1: number;
  endRow1: number;
};

type EmpAgg = {
  empId: string;
  empName: string;

  regularHours: number;
  paidOverHours: number;
  unpaidHours: number;

  totalPaidHours: number;
  hourlyRate: number;
  totalCost: number;
};

type EmpProjAgg = {
  empId: string;
  empName: string;
  projectCode: string;
  projectName: string;
  hoursPaid: number;
  costPaid: number;
};

type ProjAgg = {
  projectCode: string;
  projectName: string;
  totalHours: number;
  totalCost: number;
};

type ProjEmpAgg = {
  projectCode: string;
  projectName: string;
  empId: string;
  empName: string;
  hoursPaid: number;
  costPaid: number;
};

type DayRow = {
  who: string;
  date: Date;
  count: number;
  hours: number;
  cost: number;
  paidOrUnpaid: "Paid" | "Unpaid";
};

type ApprovalRow = {
  empId: string;
  empName: string;
  appType: string;
  dateFrom: Date;
  dateTo: Date;
  status: string;
};

type CellOut = string | number | boolean;
type PreservedShift = { wsStart: unknown; wsEnd: unknown };

// -------------------- Helpers --------------------

function norm(s: string): string {
  return (s || "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}

function toNumber(v: unknown): number {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const t = String(v).replace(/[$,]/g, "").trim();
  const n = Number(t);
  return isFinite(n) ? n : 0;
}

/**
 * LOCAL epoch to avoid date shifting.
 * Excel day 0 = 1899-12-30.
 */
function excelDateToJSDate(excelSerial: number): Date {
  const epoch = new Date(1899, 11, 30);
  const ms = Math.round(excelSerial * 24 * 60 * 60 * 1000);
  return new Date(epoch.getTime() + ms);
}

function jsDateToExcelSerial(d: Date): number {
  const epoch = new Date(1899, 11, 30);
  const ms = d.getTime() - epoch.getTime();
  return ms / (24 * 60 * 60 * 1000);
}

function parseDateCell(v: unknown): Date | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return excelDateToJSDate(v);
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function parseTimeToMinutes(v: unknown): number | null {
  if (v === null || v === undefined || v === "") return null;

  if (typeof v === "number") {
    const totalMinutes = Math.round(v * 24 * 60);
    return (totalMinutes % (24 * 60) + 24 * 60) % (24 * 60);
  }

  const s = String(v).trim();
  const d = new Date(`1970-01-01 ${s}`);
  if (isNaN(d.getTime())) return null;
  return d.getHours() * 60 + d.getMinutes();
}

function minutesToTimeText(totalMinutes: number): string {
  const mins = ((Math.round(totalMinutes) % (24 * 60)) + (24 * 60)) % (24 * 60);
  const hh24 = Math.floor(mins / 60);
  const mm = mins % 60;
  const ampm = hh24 >= 12 ? "PM" : "AM";
  const h12 = hh24 % 12 === 0 ? 12 : hh24 % 12;
  return `${h12}:${String(mm).padStart(2, "0")} ${ampm}`;
}

function combineDateAndTime(dateVal: unknown, timeVal: unknown): Date | null {
  if (
    typeof dateVal === "string" &&
    String(dateVal).includes(":") &&
    (timeVal === null || timeVal === undefined || timeVal === "")
  ) {
    const dd = new Date(String(dateVal));
    if (!isNaN(dd.getTime())) return dd;
  }

  const d = parseDateCell(dateVal);
  if (!d) return null;

  if (timeVal === null || timeVal === undefined || timeVal === "") return d;

  if (typeof timeVal === "number") {
    const minutes = Math.round(timeVal * 24 * 60);
    const out = new Date(d.getTime());
    out.setHours(0, 0, 0, 0);
    out.setMinutes(minutes);
    return out;
  }

  const t = new Date(`1970-01-01 ${String(timeVal).trim()}`);
  if (isNaN(t.getTime())) return d;

  const out = new Date(d.getTime());
  out.setHours(t.getHours(), t.getMinutes(), t.getSeconds(), 0);
  return out;
}

function startOfDay(d: Date): Date {
  const x = new Date(d.getTime());
  x.setHours(0, 0, 0, 0);
  return x;
}

function addDays(d: Date, days: number): Date {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + days);
  return x;
}

function isSunday(d: Date): boolean {
  return d.getDay() === 0;
}

function inRange(dt: Date, start: Date, end: Date): boolean {
  return dt.getTime() >= start.getTime() && dt.getTime() <= end.getTime();
}

function hoursBetween(a: Date, b: Date): number {
  const ms = b.getTime() - a.getTime();
  return ms > 0 ? ms / 36e5 : 0;
}

function overlapHours(aStart: Date, aEnd: Date, bStart: Date, bEnd: Date): number {
  const s = Math.max(aStart.getTime(), bStart.getTime());
  const e = Math.min(aEnd.getTime(), bEnd.getTime());
  return e > s ? (e - s) / 36e5 : 0;
}

function formatMDY(d: Date): string {
  const m = d.getMonth() + 1;
  const dd = d.getDate();
  const y = d.getFullYear();
  return `${m}/${dd}/${y}`;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function timeHM(d: Date): string {
  const hh = d.getHours();
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ampm = hh >= 12 ? "PM" : "AM";
  const h12 = hh % 12 === 0 ? 12 : hh % 12;
  return `${h12}:${mm} ${ampm}`;
}

function safeSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const ws = workbook.getWorksheet(name);
  if (!ws) throw new Error(`Missing sheet: ${name}`);
  return ws;
}

function findHeaderMap(headerRow: unknown[]): Map<string, number> {
  const m = new Map<string, number>();
  for (let c = 0; c < headerRow.length; c++) {
    const key = norm(String(headerRow[c] ?? ""));
    if (key) m.set(key, c);
  }
  return m;
}

function getCol(map: Map<string, number>, aliases: string[]): number {
  for (let i = 0; i < aliases.length; i++) {
    const idx = map.get(norm(aliases[i]));
    if (idx !== undefined) return idx;
  }
  return -1;
}

function detectHeaderRow(values: unknown[][], requiredAliases: string[][], scanRows: number): number {
  const maxR = Math.min(values.length, scanRows);
  let bestRow = 0;
  let bestScore = -1;

  for (let r = 0; r < maxR; r++) {
    const hmap = findHeaderMap(values[r]);
    let score = 0;
    for (let k = 0; k < requiredAliases.length; k++) {
      if (getCol(hmap, requiredAliases[k]) >= 0) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  const minNeed = Math.min(3, requiredAliases.length);
  if (bestScore < minNeed) return 0;
  return bestRow;
}

function shiftLengthHours(wsStartMin: number, wsEndMin: number): number {
  if (wsStartMin === wsEndMin) return 24;
  if (wsEndMin > wsStartMin) return (wsEndMin - wsStartMin) / 60;
  return ((24 * 60 - wsStartMin) + wsEndMin) / 60;
}

function getCutoff(d: Date): 1 | 2 {
  return d.getDate() <= 15 ? 1 : 2;
}

function getDefaultShift(empId: string, empName: string): ShiftRow {
  return {
    empId,
    empName,
    workshiftStartMin: DEFAULT_SHIFT_START_MIN,
    workshiftEndMin: DEFAULT_SHIFT_END_MIN
  };
}

function getSheetUsedRowCount(workbook: ExcelScript.Workbook, sheetName: string): number {
  const ws = safeSheet(workbook, sheetName);
  const used = ws.getUsedRange();
  if (!used) return 0;
  return used.getRowCount();
}

function nightDiffOverlapForSegment(segStart: Date, segEnd: Date): number {
  let total = 0;
  let cursor = new Date(segStart.getTime());

  while (cursor.getTime() < segEnd.getTime()) {
    const day0 = startOfDay(cursor);
    const nextDay0 = addDays(day0, 1);
    const pieceEnd = segEnd.getTime() < nextDay0.getTime() ? segEnd : nextDay0;

    const w1s = new Date(day0.getTime());
    const w1e = new Date(day0.getTime());
    w1e.setMinutes(NIGHT_END_MIN);
    total += overlapHours(cursor, pieceEnd, w1s, w1e);

    const w2s = new Date(day0.getTime());
    w2s.setMinutes(NIGHT_START_MIN);
    const w2e = new Date(nextDay0.getTime());
    total += overlapHours(cursor, pieceEnd, w2s, w2e);

    cursor = new Date(pieceEnd.getTime());
  }

  return total;
}

// -------------------- Dashboard date range --------------------

function detectSproutMinMaxDates(workbook: ExcelScript.Workbook): { min: Date; max: Date } {
  const ws = safeSheet(workbook, SHEETS.RAW_SPROUT);
  const used = ws.getUsedRange();
  if (!used) throw new Error("Raw_Sprout is empty.");

  const values = used.getValues();
  if (values.length < 2) throw new Error("Raw_Sprout has no data rows.");

  const required: string[][] = [["logdate", "log date", "date"]];
  const headerRowIdx = detectHeaderRow(values, required, 30);
  const headers = values[headerRowIdx];
  const hmap = findHeaderMap(headers);
  const cLogDate = getCol(hmap, ["logdate", "log date", "date"]);
  if (cLogDate < 0) throw new Error("Raw_Sprout must contain LogDate.");

  let minD: Date | null = null;
  let maxD: Date | null = null;

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const d = parseDateCell(row[cLogDate]);
    if (!d) continue;

    const sd = startOfDay(d);
    if (!minD || sd.getTime() < minD.getTime()) minD = sd;
    if (!maxD || sd.getTime() > maxD.getTime()) maxD = sd;
  }

  if (!minD || !maxD) throw new Error("Could not detect min/max dates from Raw_Sprout.");
  return { min: minD, max: maxD };
}

function getCurrentUploadSignature(workbook: ExcelScript.Workbook): string {
  const sproutRows = getSheetUsedRowCount(workbook, SHEETS.RAW_SPROUT);
  const invoiceRows = getSheetUsedRowCount(workbook, SHEETS.RAW_INVOICE);
  const approvalRows = getSheetUsedRowCount(workbook, SHEETS.RAW_APPROVAL);
  const mm = detectSproutMinMaxDates(workbook);

  return [
    `spr:${sproutRows}`,
    `inv:${invoiceRows}`,
    `app:${approvalRows}`,
    `min:${formatMDY(mm.min)}`,
    `max:${formatMDY(mm.max)}`
  ].join("|");
}

function setDashboardDateRange(
  workbook: ExcelScript.Workbook,
  startDay: Date,
  endDay: Date
): { start: Date; end: Date } {
  const ws = safeSheet(workbook, SHEETS.DASH);

  let s = startOfDay(startDay);
  let eDay = startOfDay(endDay);

  if (s.getTime() > eDay.getTime()) {
    const tmp = s;
    s = eDay;
    eDay = tmp;
  }

  ws.getRange("B2").setValue(jsDateToExcelSerial(s));
  ws.getRange("B3").setValue(jsDateToExcelSerial(eDay));
  ws.getRange("B2:B3").setNumberFormatLocal("m/d/yyyy");

  const e = addDays(eDay, 1);
  e.setMilliseconds(e.getMilliseconds() - 1);

  return { start: s, end: e };
}

function readOrInitDashboardDateRange(workbook: ExcelScript.Workbook): { start: Date; end: Date } {
  const ws = safeSheet(workbook, SHEETS.DASH);

  const currentSignature = getCurrentUploadSignature(workbook);
  const savedSignature = String(ws.getRange(DASH_META_SIGNATURE_CELL).getValue() ?? "").trim();

  const vStart = ws.getRange("B2").getValue();
  const vEnd = ws.getRange("B3").getValue();

  const dStart = parseDateCell(vStart);
  const dEnd = parseDateCell(vEnd);

  const mm = detectSproutMinMaxDates(workbook);

  if (!savedSignature || savedSignature !== currentSignature) {
    const dr = setDashboardDateRange(workbook, mm.min, mm.max);
    ws.getRange(DASH_META_SIGNATURE_CELL).setValue(currentSignature);
    return dr;
  }

  if (dStart && dEnd) {
    const dr = setDashboardDateRange(workbook, dStart, dEnd);
    ws.getRange(DASH_META_SIGNATURE_CELL).setValue(currentSignature);
    return dr;
  }

  const dr = setDashboardDateRange(workbook, mm.min, mm.max);
  ws.getRange(DASH_META_SIGNATURE_CELL).setValue(currentSignature);
  return dr;
}

// -------------------- Employee key parser --------------------

function parseEmployeeKey(cellVal: unknown): { empId: string; empName: string } {
  const s = String(cellVal ?? "").trim();
  if (!s) return { empId: "", empName: "" };

  const m = s.match(/^(\d+)\s+(.*)$/);
  if (m) return { empId: m[1], empName: m[2].trim() };

  if (/^\d+$/.test(s)) return { empId: s, empName: "" };

  return { empId: "", empName: s };
}

// -------------------- Shift source = TBRegularOT --------------------

function loadWorkshiftsFromTBRegularOT(workbook: ExcelScript.Workbook): { byEmpId: Map<string, ShiftRow>; byEmpName: Map<string, ShiftRow> } {
  const ws = safeSheet(workbook, SHEETS.REGVSOV);
  const byEmpId = new Map<string, ShiftRow>();
  const byEmpName = new Map<string, ShiftRow>();

  let tbl: ExcelScript.Table;
  try {
    tbl = ws.getTable(REGOT_TABLE);
  } catch {
    return { byEmpId, byEmpName };
  }

  const headers = tbl.getHeaderRowRange().getValues()[0] as unknown[];
  const hmap = findHeaderMap(headers);

  const cWSStart = getCol(hmap, ["workshift start"]);
  const cWSEnd = getCol(hmap, ["workshift end"]);
  const cEmpKey = getCol(hmap, ["employee key"]);
  const cEmpName = getCol(hmap, ["employee name"]);

  if (cWSStart < 0 || cWSEnd < 0) return { byEmpId, byEmpName };

  const vals = tbl.getRangeBetweenHeaderAndTotal().getValues();
  if (!vals || vals.length === 0) return { byEmpId, byEmpName };

  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];

    const wsStartMin = parseTimeToMinutes(row[cWSStart]);
    const wsEndMin = parseTimeToMinutes(row[cWSEnd]);
    if (wsStartMin === null || wsEndMin === null) continue;

    let empId = "";
    let empName = "";

    if (cEmpKey >= 0) {
      const parsed = parseEmployeeKey(row[cEmpKey]);
      empId = parsed.empId;
      empName = parsed.empName;
    }
    if (!empName && cEmpName >= 0) empName = String(row[cEmpName] ?? "").trim();

    if (!empId && !empName) continue;

    const sr: ShiftRow = {
      empId,
      empName,
      workshiftStartMin: wsStartMin,
      workshiftEndMin: wsEndMin
    };
    if (empId) byEmpId.set(empId, sr);
    if (empName) byEmpName.set(norm(empName), sr);
  }

  return { byEmpId, byEmpName };
}

function getShiftWindowForMoment(t: Date, wsStartMin: number, wsEndMin: number): { s: Date; e: Date } {
  const day0 = startOfDay(t);
  const mins = t.getHours() * 60 + t.getMinutes();
  const overnight = wsEndMin <= wsStartMin;

  if (!overnight) {
    const s = new Date(day0.getTime());
    s.setMinutes(wsStartMin);
    const e = new Date(day0.getTime());
    e.setMinutes(wsEndMin);
    return { s, e };
  }

  if (mins >= wsStartMin) {
    const s = new Date(day0.getTime());
    s.setMinutes(wsStartMin);
    const e = addDays(new Date(day0.getTime()), 1);
    e.setMinutes(wsEndMin);
    return { s, e };
  }

  const s2 = addDays(new Date(day0.getTime()), -1);
  s2.setMinutes(wsStartMin);
  const e2 = new Date(day0.getTime());
  e2.setMinutes(wsEndMin);
  return { s: s2, e: e2 };
}

function getShiftDayForMoment(t: Date, wsStartMin: number, wsEndMin: number): Date {
  const day0 = startOfDay(t);
  const mins = t.getHours() * 60 + t.getMinutes();
  const overnight = wsEndMin <= wsStartMin;

  if (!overnight) return day0;

  const cutoff = wsEndMin + SHIFT_POST_END_GRACE_MIN;
  if (mins <= cutoff) return addDays(day0, -1);
  return day0;
}

function shiftDayForEmp(
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>,
  empId: string,
  empName: string,
  t: Date
): Date {
  const sr =
    (empId && shiftByEmpId.get(empId)) ||
    shiftByEmpName.get(norm(empName)) ||
    getDefaultShift(empId, empName);

  return getShiftDayForMoment(t, sr.workshiftStartMin, sr.workshiftEndMin);
}

function getShiftForEmp(
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>,
  empId: string,
  empName: string
): ShiftRow {
  return (
    (empId && shiftByEmpId.get(empId)) ||
    shiftByEmpName.get(norm(empName)) ||
    getDefaultShift(empId, empName)
  );
}

// -------------------- Load Raw_Invoice --------------------

function loadInvoice(workbook: ExcelScript.Workbook): Map<string, InvoiceRow> {
  const ws = safeSheet(workbook, SHEETS.RAW_INVOICE);
  const used = ws.getUsedRange();
  const out = new Map<string, InvoiceRow>();
  if (!used) return out;

  const values = used.getValues();
  if (values.length < 2) return out;

  const required: string[][] = [["name", "employee", "employee name"]];
  const headerRowIdx = detectHeaderRow(values, required, 30);
  const headers = values[headerRowIdx];
  const hmap = findHeaderMap(headers);

  const cName = getCol(hmap, ["employee", "employee name", "name"]);
  const cTotalUsd = getCol(hmap, ["totalusd", "total usd", "total"]);
  const cResource = getCol(hmap, ["resourcecostusd", "resource cost usd"]);
  const cFeeEx = getCol(hmap, ["feeexproration", "fee ex proration"]);
  const cSubtotal = getCol(hmap, ["subtotal"]);
  const cOtHours1st = getCol(hmap, ["othours1st"]);
  const cOtHours2nd = getCol(hmap, ["othours2nd"]);

  if (cName < 0) throw new Error("Raw_Invoice must contain a Name or Employee column.");

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const name = String(row[cName] ?? "").trim();
    if (!name) continue;

    let monthly = 0;
    if (cTotalUsd >= 0) monthly = toNumber(row[cTotalUsd]);
    if (!monthly && cResource >= 0) monthly = toNumber(row[cResource]);
    if (!monthly && cFeeEx >= 0) monthly = toNumber(row[cFeeEx]);
    if (!monthly && cSubtotal >= 0) monthly = toNumber(row[cSubtotal]);

    const ot1 = cOtHours1st >= 0 ? toNumber(row[cOtHours1st]) : 0;
    const ot2 = cOtHours2nd >= 0 ? toNumber(row[cOtHours2nd]) : 0;

    out.set(norm(name), {
      empName: name,
      monthlyCostUsd: monthly,
      otHours1st: ot1,
      otHours2nd: ot2
    });
  }

  return out;
}

// -------------------- Load Raw_Sprout --------------------

function loadSprout(workbook: ExcelScript.Workbook, start: Date, end: Date): SproutRow[] {
  const ws = safeSheet(workbook, SHEETS.RAW_SPROUT);
  const used = ws.getUsedRange();
  if (!used) return [];

  const baseRowIndex1 = used.getRowIndex() + 1;
  const values = used.getValues();
  if (values.length < 2) return [];

  const required: string[][] = [
    ["biometric id", "biometricid", "biometric"],
    ["full name", "fullname", "name"],
    ["logdate", "log date", "date"],
    ["logtime", "log time", "log (time)", "log"],
    ["in/out mode", "in out mode", "in/out", "inout", "mode"]
  ];

  const headerRowIdx = detectHeaderRow(values, required, 30);
  const headers = values[headerRowIdx];
  const hmap = findHeaderMap(headers);

  const cEmpId = getCol(hmap, ["biometric id", "biometricid", "biometric"]);
  const cName = getCol(hmap, ["full name", "fullname", "name"]);
  const cLogDate = getCol(hmap, ["logdate", "log date", "date"]);
  const cLogTime = getCol(hmap, ["logtime", "log time", "log (time)", "log"]);
  const cInOut = getCol(hmap, ["in/out mode", "in out mode", "in/out", "inout", "mode"]);
  const cProjCode = getCol(hmap, ["project code", "projectcode", "project"]);
  const cProjName = getCol(hmap, ["project name", "projectname"]);

  if (cEmpId < 0 || cName < 0 || cLogDate < 0 || cLogTime < 0 || cInOut < 0) {
    throw new Error("Raw_Sprout headers missing required columns.");
  }

  const out: SproutRow[] = [];

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const empId = String(row[cEmpId] ?? "").trim();
    const empName = String(row[cName] ?? "").trim();
    if (!empId || !empName) continue;

    const dt = combineDateAndTime(row[cLogDate], row[cLogTime]);
    if (!dt) continue;
    if (!inRange(dt, start, end)) continue;

    if (isSunday(dt)) continue;

    const ioRaw = norm(String(row[cInOut] ?? ""));
    let inOut: "IN" | "OUT" | "OTHER" = "OTHER";
    if (ioRaw.startsWith("in")) inOut = "IN";
    else if (ioRaw.startsWith("out")) inOut = "OUT";

    const projectCode = cProjCode >= 0 ? String(row[cProjCode] ?? "").trim() : "";
    const projectName = cProjName >= 0 ? String(row[cProjName] ?? "").trim() : "";

    out.push({
      empId,
      empName,
      dt,
      inOut,
      projectCode: projectCode || "No Tagged Project",
      projectName: projectName || "No Tagged Project",
      rawRowIndex1: baseRowIndex1 + r
    });
  }

  out.sort((a, b) => a.empId.localeCompare(b.empId) || a.dt.getTime() - b.dt.getTime());
  return out;
}

// -------------------- Load Raw_Approval Center --------------------

function loadApprovals(workbook: ExcelScript.Workbook): ApprovalRow[] {
  const ws = safeSheet(workbook, SHEETS.RAW_APPROVAL);
  const used = ws.getUsedRange();
  if (!used) return [];

  const values = used.getValues();
  if (values.length < 2) return [];

  const fixedBioCol0 = 8; // I
  const fixedNameCol0 = 9; // J

  const required: string[][] = [["application type"], ["date from"], ["date to"], ["status"]];
  const headerRowIdx = detectHeaderRow(values, required, 30);
  const headers = values[headerRowIdx];
  const hmap = findHeaderMap(headers);

  const cAppType = getCol(hmap, ["application type"]);
  const cDateFrom = getCol(hmap, ["date from"]);
  const cDateTo = getCol(hmap, ["date to"]);
  const cStatus = getCol(hmap, ["status"]);

  const cBioDetected = getCol(hmap, ["biometric id", "biometricid", "employee id", "employeeid"]);
  const cNameDetected = getCol(hmap, ["name", "full name", "employee name"]);

  const out: ApprovalRow[] = [];

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];

    const status = norm(String(row[cStatus] ?? ""));
    if (status !== "approved") continue;

    const df = parseDateCell(row[cDateFrom]);
    const dt = parseDateCell(row[cDateTo]);
    if (!df || !dt) continue;

    const appType = norm(String(row[cAppType] ?? "").trim());

    let bio = String(row[fixedBioCol0] ?? "").trim();
    if (!bio && cBioDetected >= 0) bio = String(row[cBioDetected] ?? "").trim();

    let nm = String(row[fixedNameCol0] ?? "").trim();
    if (!nm && cNameDetected >= 0) nm = String(row[cNameDetected] ?? "").trim();

    if (!bio && !nm) continue;

    out.push({
      empId: bio,
      empName: nm,
      appType,
      dateFrom: startOfDay(df),
      dateTo: startOfDay(dt),
      status
    });
  }

  return out;
}

// -------------------- Segments from Sprout logs --------------------

function makeSegmentFromOpenClose(
  openIn: SproutRow,
  closeEv: SproutRow,
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>
): Segment | null {
  const start = openIn.dt;
  let end = closeEv.dt;
  if (end.getTime() <= start.getTime()) return null;

  const rawDur = hoursBetween(start, end);
  if (rawDur > MAX_SEGMENT_HOURS) end = new Date(start.getTime() + MAX_SEGMENT_HOURS * 36e5);

  const shift = getShiftForEmp(shiftByEmpId, shiftByEmpName, openIn.empId, openIn.empName);

  let cursor = new Date(start.getTime());
  let regSum = 0;
  let outSum = 0;

  while (cursor.getTime() < end.getTime()) {
    const dayEnd = addDays(startOfDay(cursor), 1);
    const pieceEnd = end.getTime() < dayEnd.getTime() ? end : dayEnd;

    const pieceHours = hoursBetween(cursor, pieceEnd);
    const sh = getShiftWindowForMoment(cursor, shift.workshiftStartMin, shift.workshiftEndMin);
    const regPiece = overlapHours(cursor, pieceEnd, sh.s, sh.e);
    const outPiece = Math.max(0, pieceHours - regPiece);

    regSum += regPiece;
    outSum += outPiece;

    cursor = new Date(pieceEnd.getTime());
  }

  return {
    empId: openIn.empId,
    empName: openIn.empName,
    projectCode: openIn.projectCode,
    projectName: openIn.projectName,
    start,
    end,
    regularH: regSum,
    unpaidH: outSum,
    paidOverH: 0,
    startRow1: openIn.rawRowIndex1,
    endRow1: closeEv.rawRowIndex1
  };
}

function buildSegments(
  rows: SproutRow[],
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>
): { segments: Segment[]; adjustRows: (string | number)[][] } {
  const segs: Segment[] = [];
  const adjust: (string | number)[][] = [];

  let currentEmp = "";
  let openIn: SproutRow | null = null;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    if (r.empId !== currentEmp) {
      currentEmp = r.empId;
      openIn = null;
    }

    if (r.inOut === "OTHER") continue;

    if (r.inOut === "IN") {
      if (!openIn) {
        openIn = r;
        continue;
      }

      const sameMinute =
        openIn.dt.getFullYear() === r.dt.getFullYear() &&
        openIn.dt.getMonth() === r.dt.getMonth() &&
        openIn.dt.getDate() === r.dt.getDate() &&
        openIn.dt.getHours() === r.dt.getHours() &&
        openIn.dt.getMinutes() === r.dt.getMinutes();

      if (sameMinute) {
        openIn = r;
        continue;
      }

      const seg = makeSegmentFromOpenClose(openIn, r, shiftByEmpId, shiftByEmpName);
      if (seg) segs.push(seg);
      openIn = r;
      continue;
    }

    if (r.inOut === "OUT") {
      if (!openIn) continue;
      const seg = makeSegmentFromOpenClose(openIn, r, shiftByEmpId, shiftByEmpName);
      if (seg) segs.push(seg);
      openIn = null;
    }
  }

  const unresolvedByEmp = new Map<string, SproutRow>();
  currentEmp = "";
  openIn = null;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r.empId !== currentEmp) {
      if (openIn) unresolvedByEmp.set(currentEmp, openIn);
      currentEmp = r.empId;
      openIn = null;
    }

    if (r.inOut === "IN") openIn = r;
    else if (r.inOut === "OUT") openIn = null;
  }
  if (openIn) unresolvedByEmp.set(currentEmp, openIn);

  const unresolvedList = Array.from(unresolvedByEmp.values());
  for (let i = 0; i < unresolvedList.length; i++) {
    const u = unresolvedList[i];
    adjust.push([
      "",
      u.empId,
      u.empName,
      formatMDY(startOfDay(u.dt)),
      timeHM(u.dt),
      "",
      "",
      "",
      u.projectCode,
      u.projectName,
      "Missing clock-out (unresolved open segment)",
      u.rawRowIndex1
    ]);
  }

  return { segments: segs, adjustRows: adjust };
}

// -------------------- Approval day expansion --------------------

function approvalDaysForRow(a: ApprovalRow, dashStart: Date, dashEnd: Date): Date[] {
  const days: Date[] = [];
  const from = startOfDay(a.dateFrom);
  const to = startOfDay(a.dateTo);

  const isSingleDay = from.getTime() === to.getTime();

  let cur = new Date(from.getTime());
  while (cur.getTime() <= to.getTime()) {
    if (
      cur.getTime() >= startOfDay(dashStart).getTime() &&
      cur.getTime() <= startOfDay(dashEnd).getTime()
    ) {
      if (isSunday(cur)) {
        if (isSingleDay) days.push(new Date(cur.getTime()));
      } else {
        days.push(new Date(cur.getTime()));
      }
    }
    cur = addDays(cur, 1);
  }
  return days;
}

// -------------------- Approval synthetic segments --------------------

function buildApprovalSyntheticSegments(
  approvals: ApprovalRow[],
  dashStart: Date,
  dashEnd: Date,
  invoiceMapByName: Map<string, InvoiceRow>,
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>,
  empIdToName: Map<string, string>,
  existingRegularByEmpDay: Map<string, number>
): Segment[] {
  const segs: Segment[] = [];
  const remainingOT = new Map<string, number>();

  const empIds = Array.from(empIdToName.keys());
  for (let i = 0; i < empIds.length; i++) {
    const empId = empIds[i];
    const nm = empIdToName.get(empId) || "";
    const inv = invoiceMapByName.get(norm(nm));
    const ot1 = inv ? inv.otHours1st : 0;
    const ot2 = inv ? inv.otHours2nd : 0;
    remainingOT.set(`${empId}||1`, ot1);
    remainingOT.set(`${empId}||2`, ot2);
  }

  for (let i = 0; i < approvals.length; i++) {
    const a = approvals[i];
    if (!a.empId) continue;

    const empName = a.empName || empIdToName.get(a.empId) || "";
    const shift = getShiftForEmp(shiftByEmpId, shiftByEmpName, a.empId, empName);

    const shiftLen = shiftLengthHours(shift.workshiftStartMin, shift.workshiftEndMin);
    const days = approvalDaysForRow(a, dashStart, dashEnd);
    if (days.length === 0) continue;

    if (a.appType === "overtime") {
      for (let d = 0; d < days.length; d++) {
        const day = days[d];
        const cutoff = getCutoff(day);
        const poolKey = `${a.empId}||${cutoff}`;
        const remain = remainingOT.get(poolKey) || 0;
        if (remain <= 0) continue;

        const add = Math.min(shiftLen, remain);
        if (add <= 0) continue;

        remainingOT.set(poolKey, remain - add);

        segs.push({
          empId: a.empId,
          empName,
          projectCode: CORP_PROJECT_CODE,
          projectName: "Approved Overtime",
          start: new Date(day.getTime()),
          end: new Date(day.getTime()),
          regularH: 0,
          unpaidH: 0,
          paidOverH: add,
          startRow1: 0,
          endRow1: 0
        });
      }
      continue;
    }

    if (
      a.appType === "sil leave" ||
      a.appType === "certificate of attendance" ||
      a.appType === "official business"
    ) {
      for (let d = 0; d < days.length; d++) {
        const day = days[d];

        const dayKey = `${a.empId}||${formatMDY(day)}`;
        const existingReg = existingRegularByEmpDay.get(dayKey) || 0;
        const need = Math.max(0, shiftLen - existingReg);
        if (need <= 0) continue;

        const projName =
          a.appType === "sil leave"
            ? "Approved SIL Leave"
            : a.appType === "certificate of attendance"
              ? "Approved Certificate of Attendance"
              : "Approved Official Business";

        segs.push({
          empId: a.empId,
          empName,
          projectCode: CORP_PROJECT_CODE,
          projectName: projName,
          start: new Date(day.getTime()),
          end: new Date(day.getTime()),
          regularH: need,
          unpaidH: 0,
          paidOverH: 0,
          startRow1: 0,
          endRow1: 0
        });

        existingRegularByEmpDay.set(dayKey, existingReg + need);
      }
    }
  }

  return segs;
}

// -------------------- Aggregations --------------------

function computeEmpAgg(
  segs: Segment[],
  invoiceMapByName: Map<string, InvoiceRow>,
  empIdToName: Map<string, string>
): Map<string, EmpAgg> {
  const emp = new Map<string, EmpAgg>();

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    const resolvedName = s.empName || empIdToName.get(s.empId) || "";
    const key = `${s.empId}||${norm(resolvedName)}`;

    const reg = s.regularH;
    const paid = s.paidOverH;
    const unp = s.unpaidH;

    const cur = emp.get(key);
    if (!cur) {
      emp.set(key, {
        empId: s.empId,
        empName: resolvedName,
        regularHours: reg,
        paidOverHours: paid,
        unpaidHours: unp,
        totalPaidHours: reg + paid,
        hourlyRate: 0,
        totalCost: 0
      });
    } else {
      cur.regularHours += reg;
      cur.paidOverHours += paid;
      cur.unpaidHours += unp;
      cur.totalPaidHours += reg + paid;
    }
  }

  const keys = Array.from(emp.keys());
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const e = emp.get(k);
    if (!e) continue;

    const inv = invoiceMapByName.get(norm(e.empName));
    const monthly = inv ? inv.monthlyCostUsd : 0;

    const rate = e.totalPaidHours > 0 ? monthly / e.totalPaidHours : 0;
    e.hourlyRate = rate;
    e.totalCost = e.totalPaidHours * rate;
  }

  return emp;
}

function computeEmpProjAgg(segs: Segment[], empAgg: Map<string, EmpAgg>): EmpProjAgg[] {
  const m = new Map<string, EmpProjAgg>();
  const empVals = Array.from(empAgg.values());

  function findRateAndName(empId: string): { rate: number; name: string } {
    for (let i = 0; i < empVals.length; i++) {
      if (empVals[i].empId === empId) return { rate: empVals[i].hourlyRate, name: empVals[i].empName };
    }
    return { rate: 0, name: "" };
  }

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    const paidH = s.regularH + s.paidOverH;
    if (paidH <= 0) continue;

    const rn = findRateAndName(s.empId);
    const rate = rn.rate;
    const empName = rn.name || s.empName;

    const key = `${s.empId}||${norm(empName)}||${norm(s.projectCode)}||${norm(s.projectName)}`;
    const cur = m.get(key);
    if (!cur) {
      m.set(key, {
        empId: s.empId,
        empName,
        projectCode: s.projectCode,
        projectName: s.projectName,
        hoursPaid: paidH,
        costPaid: paidH * rate
      });
    } else {
      cur.hoursPaid += paidH;
      cur.costPaid += paidH * rate;
    }
  }

  return Array.from(m.values()).sort(
    (a, b) =>
      a.empId.localeCompare(b.empId) ||
      a.projectCode.localeCompare(b.projectCode) ||
      a.projectName.localeCompare(b.projectName)
  );
}

function aggregateProjectTotals(empProj: EmpProjAgg[]): ProjAgg[] {
  const pMap = new Map<string, ProjAgg>();
  for (let i = 0; i < empProj.length; i++) {
    const r = empProj[i];
    const pk = `${norm(r.projectCode)}||${norm(r.projectName)}`;
    const cur = pMap.get(pk);
    if (!cur) {
      pMap.set(pk, {
        projectCode: r.projectCode,
        projectName: r.projectName,
        totalHours: r.hoursPaid,
        totalCost: r.costPaid
      });
    } else {
      cur.totalHours += r.hoursPaid;
      cur.totalCost += r.costPaid;
    }
  }
  return Array.from(pMap.values()).sort(
    (a, b) => a.projectCode.localeCompare(b.projectCode) || a.projectName.localeCompare(b.projectName)
  );
}

function computeProjEmpAgg(empProj: EmpProjAgg[]): ProjEmpAgg[] {
  const m = new Map<string, ProjEmpAgg>();
  for (let i = 0; i < empProj.length; i++) {
    const r = empProj[i];
    const key = `${norm(r.projectCode)}||${norm(r.projectName)}||${r.empId}||${norm(r.empName)}`;
    const cur = m.get(key);
    if (!cur) {
      m.set(key, {
        projectCode: r.projectCode,
        projectName: r.projectName,
        empId: r.empId,
        empName: r.empName,
        hoursPaid: r.hoursPaid,
        costPaid: r.costPaid
      });
    } else {
      cur.hoursPaid += r.hoursPaid;
      cur.costPaid += r.costPaid;
    }
  }
  return Array.from(m.values()).sort(
    (a, b) =>
      a.projectCode.localeCompare(b.projectCode) ||
      a.empId.localeCompare(b.empId) ||
      a.empName.localeCompare(b.empName)
  );
}

// -------------------- Day summaries --------------------

function computeEmployeeDayRows(
  segs: Segment[],
  empAgg: Map<string, EmpAgg>,
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>
): { reg: DayRow[]; ov: DayRow[] } {
  const regMap = new Map<string, { empName: string; date: Date; projects: Map<string, boolean>; hours: number; cost: number }>();
  const ovPaidMap = new Map<string, { empName: string; date: Date; projects: Map<string, boolean>; hours: number; cost: number }>();
  const ovUnpaidMap = new Map<string, { empName: string; date: Date; projects: Map<string, boolean>; hours: number; cost: number }>();

  const empVals = Array.from(empAgg.values());
  function rateFor(empId: string): { name: string; rate: number } {
    for (let i = 0; i < empVals.length; i++) {
      if (empVals[i].empId === empId) return { name: empVals[i].empName, rate: empVals[i].hourlyRate };
    }
    return { name: "", rate: 0 };
  }

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    const d = shiftDayForEmp(shiftByEmpId, shiftByEmpName, s.empId, s.empName, s.start);
    const keyDate = formatMDY(d);
    const k = `${s.empId}||${keyDate}`;

    const rn = rateFor(s.empId);
    const empName = rn.name || s.empName;
    const rate = rn.rate;

    const projKey = `${norm(s.projectCode)}||${norm(s.projectName)}`;

    if (s.regularH > 0) {
      const cur = regMap.get(k);
      if (!cur) {
        const projects = new Map<string, boolean>();
        projects.set(projKey, true);
        regMap.set(k, { empName, date: d, projects, hours: s.regularH, cost: s.regularH * rate });
      } else {
        cur.projects.set(projKey, true);
        cur.hours += s.regularH;
        cur.cost += s.regularH * rate;
      }
    }

    if (s.paidOverH > 0) {
      const curP = ovPaidMap.get(k);
      if (!curP) {
        const projects = new Map<string, boolean>();
        projects.set(projKey, true);
        ovPaidMap.set(k, { empName, date: d, projects, hours: s.paidOverH, cost: s.paidOverH * rate });
      } else {
        curP.projects.set(projKey, true);
        curP.hours += s.paidOverH;
        curP.cost += s.paidOverH * rate;
      }
    }

    if (s.unpaidH > 0) {
      const curU = ovUnpaidMap.get(k);
      if (!curU) {
        const projects = new Map<string, boolean>();
        projects.set(projKey, true);
        ovUnpaidMap.set(k, { empName, date: d, projects, hours: s.unpaidH, cost: s.unpaidH * rate });
      } else {
        curU.projects.set(projKey, true);
        curU.hours += s.unpaidH;
        curU.cost += s.unpaidH * rate;
      }
    }
  }

  const reg: DayRow[] = [];
  const ov: DayRow[] = [];

  for (const x of Array.from(regMap.values())) {
    reg.push({ who: x.empName, date: x.date, count: x.projects.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Paid" });
  }
  for (const x of Array.from(ovPaidMap.values())) {
    ov.push({ who: x.empName, date: x.date, count: x.projects.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Paid" });
  }
  for (const x of Array.from(ovUnpaidMap.values())) {
    ov.push({ who: x.empName, date: x.date, count: x.projects.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Unpaid" });
  }

  reg.sort((a, b) => a.date.getTime() - b.date.getTime() || a.who.localeCompare(b.who));
  ov.sort((a, b) => a.date.getTime() - b.date.getTime() || a.who.localeCompare(b.who) || a.paidOrUnpaid.localeCompare(b.paidOrUnpaid));
  return { reg, ov };
}

function computeProjectDayRows(
  segs: Segment[],
  empAgg: Map<string, EmpAgg>,
  shiftByEmpId: Map<string, ShiftRow>,
  shiftByEmpName: Map<string, ShiftRow>
): { reg: DayRow[]; ov: DayRow[] } {
  const regMap = new Map<string, { projName: string; date: Date; emps: Map<string, boolean>; hours: number; cost: number }>();
  const ovPaidMap = new Map<string, { projName: string; date: Date; emps: Map<string, boolean>; hours: number; cost: number }>();
  const ovUnpaidMap = new Map<string, { projName: string; date: Date; emps: Map<string, boolean>; hours: number; cost: number }>();

  const empVals = Array.from(empAgg.values());
  function rateFor(empId: string): number {
    for (let i = 0; i < empVals.length; i++) if (empVals[i].empId === empId) return empVals[i].hourlyRate;
    return 0;
  }

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    const d = shiftDayForEmp(shiftByEmpId, shiftByEmpName, s.empId, s.empName, s.start);
    const keyDate = formatMDY(d);
    const rate = rateFor(s.empId);
    const projName = s.projectName;

    const pk = `${norm(s.projectCode)}||${norm(s.projectName)}||${keyDate}`;
    const empKey = `${s.empId}||${norm(s.empName)}`;

    if (s.regularH > 0) {
      const cur = regMap.get(pk);
      if (!cur) {
        const emps = new Map<string, boolean>();
        emps.set(empKey, true);
        regMap.set(pk, { projName, date: d, emps, hours: s.regularH, cost: s.regularH * rate });
      } else {
        cur.emps.set(empKey, true);
        cur.hours += s.regularH;
        cur.cost += s.regularH * rate;
      }
    }

    if (s.paidOverH > 0) {
      const curP = ovPaidMap.get(pk);
      if (!curP) {
        const emps = new Map<string, boolean>();
        emps.set(empKey, true);
        ovPaidMap.set(pk, { projName, date: d, emps, hours: s.paidOverH, cost: s.paidOverH * rate });
      } else {
        curP.emps.set(empKey, true);
        curP.hours += s.paidOverH;
        curP.cost += s.paidOverH * rate;
      }
    }

    if (s.unpaidH > 0) {
      const curU = ovUnpaidMap.get(pk);
      if (!curU) {
        const emps = new Map<string, boolean>();
        emps.set(empKey, true);
        ovUnpaidMap.set(pk, { projName, date: d, emps, hours: s.unpaidH, cost: s.unpaidH * rate });
      } else {
        curU.emps.set(empKey, true);
        curU.hours += s.unpaidH;
        curU.cost += s.unpaidH * rate;
      }
    }
  }

  const reg: DayRow[] = [];
  const ov: DayRow[] = [];

  for (const x of Array.from(regMap.values())) {
    reg.push({ who: x.projName, date: x.date, count: x.emps.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Paid" });
  }
  for (const x of Array.from(ovPaidMap.values())) {
    ov.push({ who: x.projName, date: x.date, count: x.emps.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Paid" });
  }
  for (const x of Array.from(ovUnpaidMap.values())) {
    ov.push({ who: x.projName, date: x.date, count: x.emps.size, hours: x.hours, cost: x.cost, paidOrUnpaid: "Unpaid" });
  }

  reg.sort((a, b) => a.date.getTime() - b.date.getTime() || a.who.localeCompare(b.who));
  ov.sort((a, b) => a.date.getTime() - b.date.getTime() || a.who.localeCompare(b.who) || a.paidOrUnpaid.localeCompare(b.paidOrUnpaid));
  return { reg, ov };
}

// -------------------- Writers helpers --------------------

function clearBlock(ws: ExcelScript.Worksheet, headerRow: number, startCol: number, cols: number, clearRows: number): void {
  const startRow = headerRow + 1;
  ws.getRangeByIndexes(startRow - 1, startCol - 1, clearRows, cols).clear(ExcelScript.ClearApplyTo.contents);
}

function clearHighlight(ws: ExcelScript.Worksheet, headerRow: number, startCol: number, cols: number, clearRows: number): void {
  const startRow = headerRow + 1;
  const r = ws.getRangeByIndexes(startRow - 1, startCol - 1, clearRows, cols);
  r.getFormat().getFill().clear();
  r.getFormat().getFont().setBold(false);
}

function writeBlock(ws: ExcelScript.Worksheet, headerRow: number, startCol: number, rows: (string | number | boolean)[][]): void {
  if (rows.length === 0) return;
  const startRow = headerRow + 1;
  ws.getRangeByIndexes(startRow - 1, startCol - 1, rows.length, rows[0].length).setValues(rows);
}

function applyRowHighlight(ws: ExcelScript.Worksheet, rowIndex1: number, startCol: number, cols: number, fillHex: string): void {
  const r = ws.getRangeByIndexes(rowIndex1 - 1, startCol - 1, 1, cols);
  r.getFormat().getFill().setColor(fillHex);
  r.getFormat().getFont().setBold(true);
}

// -------------------- Dashboard writers --------------------

function writeDashboard(
  workbook: ExcelScript.Workbook,
  empAgg: Map<string, EmpAgg>,
  empProj: EmpProjAgg[],
  projAgg: ProjAgg[],
  projEmpAgg: ProjEmpAgg[]
): void {
  const ws = safeSheet(workbook, SHEETS.DASH);

  clearBlock(ws, DASH_EMP.headerRow, DASH_EMP.startCol, DASH_EMP.cols, DASH_CLEAR_ROWS);
  clearBlock(ws, DASH_PROJ.headerRow, DASH_PROJ.startCol, DASH_PROJ.cols, DASH_CLEAR_ROWS);
  clearHighlight(ws, DASH_EMP.headerRow, DASH_EMP.startCol, DASH_EMP.cols, DASH_CLEAR_ROWS);
  clearHighlight(ws, DASH_PROJ.headerRow, DASH_PROJ.startCol, DASH_PROJ.cols, DASH_CLEAR_ROWS);

  const empRows: (string | number | boolean)[][] = [];
  const empVals = Array.from(empAgg.values()).sort((a, b) => a.empId.localeCompare(b.empId));

  const empToProjects = new Map<string, EmpProjAgg[]>();
  for (let i = 0; i < empProj.length; i++) {
    const r = empProj[i];
    const k = `${r.empId}||${norm(r.empName)}`;
    const list = empToProjects.get(k);
    if (!list) empToProjects.set(k, [r]); else list.push(r);
  }

  for (let i = 0; i < empVals.length; i++) {
    const e = empVals[i];
    const k = `${e.empId}||${norm(e.empName)}`;
    const list = empToProjects.get(k) || [];
    empRows.push([e.empId, e.empName, round2(e.totalPaidHours), round2(e.totalCost), "", "", "", ""]);
    for (let j = 0; j < list.length; j++) {
      const p = list[j];
      empRows.push(["", "", "", "", p.projectCode, p.projectName, round2(p.hoursPaid), round2(p.costPaid)]);
    }
  }

  if (empRows.length) {
    writeBlock(ws, DASH_EMP.headerRow, DASH_EMP.startCol, empRows);
    let cursor = DASH_EMP.headerRow + 1;
    for (let i = 0; i < empVals.length; i++) {
      const e = empVals[i];
      const k = `${e.empId}||${norm(e.empName)}`;
      const list = empToProjects.get(k) || [];
      applyRowHighlight(ws, cursor, DASH_EMP.startCol, DASH_EMP.cols, "#FFF7E6");
      cursor += 1 + list.length;
    }
  }

  const projRows: (string | number | boolean)[][] = [];
  const projToEmps = new Map<string, ProjEmpAgg[]>();
  for (let i = 0; i < projEmpAgg.length; i++) {
    const r = projEmpAgg[i];
    const pk = `${norm(r.projectCode)}||${norm(r.projectName)}`;
    const list = projToEmps.get(pk);
    if (!list) projToEmps.set(pk, [r]); else list.push(r);
  }

  for (let i = 0; i < projAgg.length; i++) {
    const p = projAgg[i];
    const pk = `${norm(p.projectCode)}||${norm(p.projectName)}`;
    const list = projToEmps.get(pk) || [];

    projRows.push([p.projectCode, p.projectName, round2(p.totalHours), round2(p.totalCost), "", "", "", ""]);
    for (let j = 0; j < list.length; j++) {
      const e = list[j];
      projRows.push(["", "", "", "", e.empId, e.empName, round2(e.hoursPaid), round2(e.costPaid)]);
    }
  }

  if (projRows.length) {
    writeBlock(ws, DASH_PROJ.headerRow, DASH_PROJ.startCol, projRows);
    let cursor = DASH_PROJ.headerRow + 1;
    for (let i = 0; i < projAgg.length; i++) {
      const p = projAgg[i];
      const pk = `${norm(p.projectCode)}||${norm(p.projectName)}`;
      const list = projToEmps.get(pk) || [];
      applyRowHighlight(ws, cursor, DASH_PROJ.startCol, DASH_PROJ.cols, "#E8F3FF");
      cursor += 1 + list.length;
    }
  }

  let totalH = 0;
  let totalC = 0;
  for (let i = 0; i < empVals.length; i++) {
    totalH += empVals[i].totalPaidHours;
    totalC += empVals[i].totalCost;
  }
  ws.getRange("B5").setValue(round2(totalH));
  ws.getRange("B6").setValue(round2(totalC));
  ws.getRange("B7").setValue(empVals.length);
  ws.getRange("B8").setValue(projAgg.length);
}

function writeDashboard2Flat(
  workbook: ExcelScript.Workbook,
  empProj: EmpProjAgg[],
  projEmpAgg: ProjEmpAgg[]
): void {
  const ws = safeSheet(workbook, SHEETS.DASH2);

  clearBlock(ws, DASH2_EMP.headerRow, DASH2_EMP.startCol, DASH2_EMP.cols, DASH2_EMP.clearRows);
  clearBlock(ws, DASH2_PROJ.headerRow, DASH2_PROJ.startCol, DASH2_PROJ.cols, DASH2_PROJ.clearRows);

  const empRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < empProj.length; i++) {
    const r = empProj[i];
    empRows.push([r.empId, r.empName, r.projectCode, r.projectName, round2(r.hoursPaid), round2(r.costPaid)]);
  }
  if (empRows.length) {
    writeBlock(ws, DASH2_EMP.headerRow, DASH2_EMP.startCol, empRows);
    const startRow0 = DASH2_EMP.headerRow;
    const dataRowCount = empRows.length;
    ws.getRangeByIndexes(startRow0, 4, dataRowCount, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(startRow0, 5, dataRowCount, 1).setNumberFormatLocal("$#,##0.00");
  }

  const projRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < projEmpAgg.length; i++) {
    const r = projEmpAgg[i];
    projRows.push([r.projectCode, r.projectName, r.empId, r.empName, round2(r.hoursPaid), round2(r.costPaid)]);
  }
  if (projRows.length) {
    writeBlock(ws, DASH2_PROJ.headerRow, DASH2_PROJ.startCol, projRows);
    const startRow0 = DASH2_PROJ.headerRow;
    const dataRowCount = projRows.length;
    ws.getRangeByIndexes(startRow0, DASH2_PROJ.startCol - 1 + 4, dataRowCount, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(startRow0, DASH2_PROJ.startCol - 1 + 5, dataRowCount, 1).setNumberFormatLocal("$#,##0.00");
  }
}

function writeEmployeesTab(
  workbook: ExcelScript.Workbook,
  empAgg: Map<string, EmpAgg>,
  empProj: EmpProjAgg[],
  empDayReg: DayRow[],
  empDayOv: DayRow[]
): void {
  const ws = safeSheet(workbook, SHEETS.EMP);

  clearBlock(ws, EMP_SUM.headerRow, EMP_SUM.startCol, EMP_SUM.cols, EMP_SUM.clearRows);
  clearBlock(ws, EMP_REG.headerRow, EMP_REG.startCol, EMP_REG.cols, EMP_REG.clearRows);
  clearBlock(ws, EMP_OV.headerRow, EMP_OV.startCol, EMP_OV.cols, EMP_OV.clearRows);

  const empVals = Array.from(empAgg.values()).sort((a, b) => a.empId.localeCompare(b.empId));
  const empToProjects = new Map<string, EmpProjAgg[]>();
  for (let i = 0; i < empProj.length; i++) {
    const r = empProj[i];
    const k = `${r.empId}||${norm(r.empName)}`;
    const list = empToProjects.get(k);
    if (!list) empToProjects.set(k, [r]); else list.push(r);
  }

  const rows: (string | number | boolean)[][] = [];
  for (let i = 0; i < empVals.length; i++) {
    const e = empVals[i];
    const k = `${e.empId}||${norm(e.empName)}`;
    const list = empToProjects.get(k) || [];
    rows.push([e.empId, e.empName, round2(e.totalPaidHours), round2(e.totalCost), "", "", "", ""]);
    for (let j = 0; j < list.length; j++) {
      const p = list[j];
      rows.push(["", "", "", "", p.projectCode, p.projectName, round2(p.hoursPaid), round2(p.costPaid)]);
    }
  }
  if (rows.length) writeBlock(ws, EMP_SUM.headerRow, EMP_SUM.startCol, rows);

  const regRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < empDayReg.length; i++) {
    const d = empDayReg[i];
    regRows.push([d.who, jsDateToExcelSerial(d.date), d.count, round2(d.hours), round2(d.cost)]);
  }
  if (regRows.length) {
    writeBlock(ws, EMP_REG.headerRow, EMP_REG.startCol, regRows);
    const r0 = EMP_REG.headerRow;
    const c0 = EMP_REG.startCol - 1;
    ws.getRangeByIndexes(r0, c0 + 1, regRows.length, 1).setNumberFormatLocal("[$-en-US]dddd, mmmm d, yyyy");
    ws.getRangeByIndexes(r0, c0 + 3, regRows.length, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(r0, c0 + 4, regRows.length, 1).setNumberFormatLocal("$#,##0.00");
  }

  const ovRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < empDayOv.length; i++) {
    const d = empDayOv[i];
    ovRows.push([d.who, jsDateToExcelSerial(d.date), d.count, round2(d.hours), round2(d.cost), d.paidOrUnpaid]);
  }
  if (ovRows.length) {
    writeBlock(ws, EMP_OV.headerRow, EMP_OV.startCol, ovRows);
    const r0 = EMP_OV.headerRow;
    const c0 = EMP_OV.startCol - 1;
    ws.getRangeByIndexes(r0, c0 + 1, ovRows.length, 1).setNumberFormatLocal("[$-en-US]dddd, mmmm d, yyyy");
    ws.getRangeByIndexes(r0, c0 + 3, ovRows.length, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(r0, c0 + 4, ovRows.length, 1).setNumberFormatLocal("$#,##0.00");
  }
}

function writeProjectsTab(
  workbook: ExcelScript.Workbook,
  projAgg: ProjAgg[],
  projEmpAgg: ProjEmpAgg[],
  projDayReg: DayRow[],
  projDayOv: DayRow[]
): void {
  const ws = safeSheet(workbook, SHEETS.PROJ);

  clearBlock(ws, PROJ_SUM.headerRow, PROJ_SUM.startCol, PROJ_SUM.cols, PROJ_SUM.clearRows);
  clearBlock(ws, PROJ_REG.headerRow, PROJ_REG.startCol, PROJ_REG.cols, PROJ_REG.clearRows);
  clearBlock(ws, PROJ_OV.headerRow, PROJ_OV.startCol, PROJ_OV.cols, PROJ_OV.clearRows);

  const projToEmps = new Map<string, ProjEmpAgg[]>();
  for (let i = 0; i < projEmpAgg.length; i++) {
    const pe = projEmpAgg[i];
    const pk = `${norm(pe.projectCode)}||${norm(pe.projectName)}`;
    const list = projToEmps.get(pk);
    if (!list) projToEmps.set(pk, [pe]); else list.push(pe);
  }

  const rows: (string | number | boolean)[][] = [];
  for (let i = 0; i < projAgg.length; i++) {
    const p = projAgg[i];
    const pk = `${norm(p.projectCode)}||${norm(p.projectName)}`;
    rows.push([p.projectCode, p.projectName, round2(p.totalHours), round2(p.totalCost), "", "", "", ""]);
    const list = projToEmps.get(pk) || [];
    for (let j = 0; j < list.length; j++) {
      const pe = list[j];
      rows.push(["", "", "", "", pe.empId, pe.empName, round2(pe.hoursPaid), round2(pe.costPaid)]);
    }
  }
  if (rows.length) writeBlock(ws, PROJ_SUM.headerRow, PROJ_SUM.startCol, rows);

  const regRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < projDayReg.length; i++) {
    const d = projDayReg[i];
    regRows.push([d.who, jsDateToExcelSerial(d.date), d.count, round2(d.hours), round2(d.cost)]);
  }
  if (regRows.length) {
    writeBlock(ws, PROJ_REG.headerRow, PROJ_REG.startCol, regRows);
    const r0 = PROJ_REG.headerRow;
    const c0 = PROJ_REG.startCol - 1;
    ws.getRangeByIndexes(r0, c0 + 1, regRows.length, 1).setNumberFormatLocal("[$-en-US]dddd, mmmm d, yyyy");
    ws.getRangeByIndexes(r0, c0 + 3, regRows.length, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(r0, c0 + 4, regRows.length, 1).setNumberFormatLocal("$#,##0.00");
  }

  const ovRows: (string | number | boolean)[][] = [];
  for (let i = 0; i < projDayOv.length; i++) {
    const d = projDayOv[i];
    ovRows.push([d.who, jsDateToExcelSerial(d.date), d.count, round2(d.hours), round2(d.cost), d.paidOrUnpaid]);
  }
  if (ovRows.length) {
    writeBlock(ws, PROJ_OV.headerRow, PROJ_OV.startCol, ovRows);
    const r0 = PROJ_OV.headerRow;
    const c0 = PROJ_OV.startCol - 1;
    ws.getRangeByIndexes(r0, c0 + 1, ovRows.length, 1).setNumberFormatLocal("[$-en-US]dddd, mmmm d, yyyy");
    ws.getRangeByIndexes(r0, c0 + 3, ovRows.length, 1).setNumberFormatLocal("0.00");
    ws.getRangeByIndexes(r0, c0 + 4, ovRows.length, 1).setNumberFormatLocal("$#,##0.00");
  }
}

// -------------------- TBRegularOT writer --------------------

function writeTBRegularOT(
  workbook: ExcelScript.Workbook,
  sproutRows: SproutRow[],
  segmentsAll: Segment[],
  empAgg: Map<string, EmpAgg>,
  start: Date,
  end: Date
): void {
  const ws = safeSheet(workbook, SHEETS.REGVSOV);
  const tbl = ws.getTable(REGOT_TABLE);

  const headers = tbl.getHeaderRowRange().getValues()[0] as unknown[];
  const hmap = findHeaderMap(headers);

  const cWSStart = getCol(hmap, ["workshift start"]);
  const cWSEnd = getCol(hmap, ["workshift end"]);
  const cEmpKey = getCol(hmap, ["employee key"]);
  const cEmpName = getCol(hmap, ["employee name"]);
  const cDateRange = getCol(hmap, ["date range"]);
  const cRegH = getCol(hmap, ["total regular hours", "regular hours"]);
  const cRegCost = getCol(hmap, ["regular labor cost", "regular cost"]);
  const cRegRate = getCol(hmap, ["regular hourly rate", "hourly rate"]);
  const cOverH = getCol(hmap, ["total overage hours", "overage hours", "total overtime hours", "overtime hours"]);
  const cOverCost = getCol(hmap, ["overage labor cost", "overage cost", "overtime cost"]);
  const cOverRate = getCol(hmap, ["overage hourly rate", "overtime rate"]);
  const cNight = getCol(hmap, ["night diff hours", "night differential hours"]);
  const cProjCount = getCol(hmap, ["no. of projects", "no of projects", "projects"]);

  if (cWSStart < 0 || cWSEnd < 0 || cEmpKey < 0 || cEmpName < 0) {
    throw new Error("TBRegularOT headers missing required columns (Workshift Start, Workshift End, Employee key, Employee name).");
  }

  const body = tbl.getRangeBetweenHeaderAndTotal();
  const curVals = body.getValues();
  const curRowCount = curVals ? curVals.length : 0;
  const colCount = headers.length;

  const preserveByEmpId = new Map<string, PreservedShift>();
  const preserveByName = new Map<string, PreservedShift>();

  for (let r = 0; r < curRowCount; r++) {
    const row = curVals[r];
    const parsed = parseEmployeeKey(row[cEmpKey]);
    const nm = String(row[cEmpName] ?? "").trim() || parsed.empName;

    const wsStartVal: unknown = row[cWSStart];
    const wsEndVal: unknown = row[cWSEnd];

    if (parsed.empId) preserveByEmpId.set(parsed.empId, { wsStart: wsStartVal, wsEnd: wsEndVal });
    if (nm) preserveByName.set(norm(nm), { wsStart: wsStartVal, wsEnd: wsEndVal });
  }

  const empMap = new Map<string, string>();
  const empVals = Array.from(empAgg.values()).sort((a, b) => a.empId.localeCompare(b.empId));

  for (let i = 0; i < empVals.length; i++) {
    const e = empVals[i];
    if (e.empId && e.empName) empMap.set(e.empId, e.empName);
  }

  for (let i = 0; i < sproutRows.length; i++) {
    const r = sproutRows[i];
    if (r.empId && r.empName && !empMap.has(r.empId)) empMap.set(r.empId, r.empName);
  }

  const empIds = Array.from(empMap.keys()).sort((a, b) => a.localeCompare(b));

  const aggByEmpId = new Map<string, EmpAgg>();
  for (let i = 0; i < empVals.length; i++) aggByEmpId.set(empVals[i].empId, empVals[i]);

  const nightByEmp = new Map<string, number>();
  const projSetByEmp = new Map<string, Map<string, boolean>>();

  for (let i = 0; i < segmentsAll.length; i++) {
    const s = segmentsAll[i];

    if (s.startRow1 > 0) {
      const nd = nightDiffOverlapForSegment(s.start, s.end);
      nightByEmp.set(s.empId, (nightByEmp.get(s.empId) || 0) + nd);
    }

    const hasAny = s.regularH + s.paidOverH + s.unpaidH > 0;
    if (hasAny) {
      const pk = `${norm(s.projectCode)}||${norm(s.projectName)}`;
      let set = projSetByEmp.get(s.empId);
      if (!set) {
        set = new Map<string, boolean>();
        projSetByEmp.set(s.empId, set);
      }
      set.set(pk, true);
    }
  }

  const dateRangeText = `${formatMDY(start)} - ${formatMDY(end)}`;

  const outRows: CellOut[][] = [];
  for (let i = 0; i < empIds.length; i++) {
    const empId = empIds[i];
    const empName = empMap.get(empId) || "";

    const preserved = preserveByEmpId.get(empId) || preserveByName.get(norm(empName));

    const preservedStartMin = preserved ? parseTimeToMinutes(preserved.wsStart) : null;
    const preservedEndMin = preserved ? parseTimeToMinutes(preserved.wsEnd) : null;

    const finalStartMin = preservedStartMin !== null ? preservedStartMin : DEFAULT_SHIFT_START_MIN;
    const finalEndMin = preservedEndMin !== null ? preservedEndMin : DEFAULT_SHIFT_END_MIN;

    const agg = aggByEmpId.get(empId);
    const regularH = agg ? agg.regularHours : 0;
    const paidOverH = agg ? agg.paidOverHours : 0;
    const hourlyRate = agg ? agg.hourlyRate : 0;

    const regularCost = regularH * hourlyRate;
    const overCost = paidOverH * hourlyRate;
    const overRate = paidOverH > 0 ? hourlyRate : 0;

    const nightH = nightByEmp.get(empId) || 0;
    const projCount = projSetByEmp.get(empId) ? (projSetByEmp.get(empId) as Map<string, boolean>).size : 0;

    const rowOut: CellOut[] = [];
    for (let c = 0; c < colCount; c++) rowOut.push("");

    // Write as TEXT so Power Apps reads clean time strings
    rowOut[cWSStart] = minutesToTimeText(finalStartMin);
    rowOut[cWSEnd] = minutesToTimeText(finalEndMin);

    rowOut[cEmpKey] = `${empId}`;
    rowOut[cEmpName] = empName;

    if (cDateRange >= 0) rowOut[cDateRange] = dateRangeText;
    if (cRegH >= 0) rowOut[cRegH] = round2(regularH);
    if (cRegCost >= 0) rowOut[cRegCost] = round2(regularCost);
    if (cRegRate >= 0) rowOut[cRegRate] = round2(hourlyRate);
    if (cOverH >= 0) rowOut[cOverH] = round2(paidOverH);
    if (cOverCost >= 0) rowOut[cOverCost] = round2(overCost);
    if (cOverRate >= 0) rowOut[cOverRate] = round2(overRate);
    if (cNight >= 0) rowOut[cNight] = round2(nightH);
    if (cProjCount >= 0) rowOut[cProjCount] = projCount;

    outRows.push(rowOut);
  }

  const desiredRows = outRows.length;

  if (curRowCount < desiredRows) {
    const addCount = desiredRows - curRowCount;
    if (addCount > 0) {
      const emptyRows: CellOut[][] = [];
      for (let i = 0; i < addCount; i++) {
        const r: CellOut[] = [];
        for (let c = 0; c < colCount; c++) r.push("");
        emptyRows.push(r);
      }
      tbl.addRows(-1, emptyRows);
    }
  }

  const newBody = tbl.getRangeBetweenHeaderAndTotal();
  const finalRowCount = newBody.getRowCount();

  if (finalRowCount > 0) {
    newBody.clear(ExcelScript.ClearApplyTo.contents);
  }

  if (desiredRows > 0) {
    const target = tbl.getRangeBetweenHeaderAndTotal();
    const writeRows = Math.min(desiredRows, target.getRowCount());

    const writeRange = target.getResizedRange(-(target.getRowCount() - writeRows), 0);
    writeRange.setValues(outRows.slice(0, writeRows));

    // Force Workshift Start/End columns to TEXT format
    ws.getRangeByIndexes(target.getRowIndex(), cWSStart, writeRows, 1).setNumberFormatLocal("@");
    ws.getRangeByIndexes(target.getRowIndex(), cWSEnd, writeRows, 1).setNumberFormatLocal("@");
  }
}

// -------------------- Rates + Adjust Hours --------------------

function writeRatesDerived(workbook: ExcelScript.Workbook, empAgg: Map<string, EmpAgg>): void {
  const ws = safeSheet(workbook, SHEETS.RATES);

  ws.getRangeByIndexes(1, 0, 3000, 6).clear(ExcelScript.ClearApplyTo.contents);

  const vals = Array.from(empAgg.values()).sort((a, b) => a.empId.localeCompare(b.empId));
  const rows: (string | number | boolean)[][] = [];
  for (let i = 0; i < vals.length; i++) {
    rows.push([
      vals[i].empId,
      vals[i].empName,
      round2(vals[i].hourlyRate),
      round2(vals[i].regularHours),
      round2(vals[i].paidOverHours),
      round2(vals[i].unpaidHours)
    ]);
  }
  if (rows.length) ws.getRangeByIndexes(1, 0, rows.length, 6).setValues(rows);
}

function writeAdjustHoursTab(workbook: ExcelScript.Workbook, adjustRows: (string | number)[][]): void {
  const ws = safeSheet(workbook, SHEETS.ADJUST);
  clearBlock(ws, ADJ.headerRow, ADJ.startCol, ADJ.cols, ADJ.clearRows);
  if (adjustRows.length) writeBlock(ws, ADJ.headerRow, ADJ.startCol, adjustRows as (string | number | boolean)[][]);
}

// -------------------- MAIN --------------------

function main(workbook: ExcelScript.Workbook) {
  const dr = readOrInitDashboardDateRange(workbook);
  const start = dr.start;
  const end = dr.end;

  const shifts = loadWorkshiftsFromTBRegularOT(workbook);
  const shiftByEmpId = shifts.byEmpId;
  const shiftByEmpName = shifts.byEmpName;

  const invoiceMapByName = loadInvoice(workbook);
  const sproutRows = loadSprout(workbook, start, end);
  const approvals = loadApprovals(workbook);

  const empIdToName = new Map<string, string>();
  for (let i = 0; i < sproutRows.length; i++) {
    const r = sproutRows[i];
    if (r.empId && r.empName && !empIdToName.get(r.empId)) empIdToName.set(r.empId, r.empName);
  }
  for (let i = 0; i < approvals.length; i++) {
    const a = approvals[i];
    if (a.empId && a.empName && !empIdToName.get(a.empId)) empIdToName.set(a.empId, a.empName);
  }

  const built = buildSegments(sproutRows, shiftByEmpId, shiftByEmpName);
  let segments = built.segments;

  const existingRegularByEmpDay = new Map<string, number>();
  for (let i = 0; i < segments.length; i++) {
    const s = segments[i];
    const sd = shiftDayForEmp(shiftByEmpId, shiftByEmpName, s.empId, s.empName, s.start);
    const dk = `${s.empId}||${formatMDY(sd)}`;
    existingRegularByEmpDay.set(dk, (existingRegularByEmpDay.get(dk) || 0) + s.regularH);
  }

  const syn = buildApprovalSyntheticSegments(
    approvals,
    start,
    end,
    invoiceMapByName,
    shiftByEmpId,
    shiftByEmpName,
    empIdToName,
    existingRegularByEmpDay
  );
  segments = segments.concat(syn);

  const empAgg = computeEmpAgg(segments, invoiceMapByName, empIdToName);
  const empProj = computeEmpProjAgg(segments, empAgg);
  const projAgg = aggregateProjectTotals(empProj);
  const projEmpAgg = computeProjEmpAgg(empProj);

  const empDay = computeEmployeeDayRows(segments, empAgg, shiftByEmpId, shiftByEmpName);
  const projDay = computeProjectDayRows(segments, empAgg, shiftByEmpId, shiftByEmpName);

  writeDashboard(workbook, empAgg, empProj, projAgg, projEmpAgg);
  writeDashboard2Flat(workbook, empProj, projEmpAgg);
  writeEmployeesTab(workbook, empAgg, empProj, empDay.reg, empDay.ov);
  writeProjectsTab(workbook, projAgg, projEmpAgg, projDay.reg, projDay.ov);

  writeTBRegularOT(
    workbook,
    sproutRows,
    segments,
    empAgg,
    startOfDay(start),
    startOfDay(end)
  );

  writeRatesDerived(workbook, empAgg);
  writeAdjustHoursTab(workbook, built.adjustRows);
}
