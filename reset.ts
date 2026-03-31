/**
 * RESET SCRIPT (Template-safe v6)
 *
 * Fixes:
 * 1) TBRegularOT: preserves Workshift Start/End + Employee key/name by HEADER NAME,
 *    clears only the other columns (no positional assumptions).
 * 2) Date range ALWAYS refreshes from Raw_Sprout min/max and forces recalculation.
 * 3) Dashboard (2) O4/P4 are only set if they are NOT formulas (so we don't break linkages).
 */

const SHEETS = {
  RAW_SPROUT: "Raw_Sprout",
  DASH: "Dashboard",
  DASH2: "Dashboard (2)",
  EMP: "Employees",
  PROJ: "Projects",
  REGVSOV: "Regular vs Overage",
  ADJUST: "Adjust Hours",
  RATES: "Rates"
};

const TABLES = {
  REGULAR_OT: "TBRegularOT",
  DASH2_FILTERS: "TBDashboard_2_Filters"
};

// Dashboard blocks
const DASH_EMP = { headerRow: 11, startCol: 1, cols: 8, clearRows: 2500 };    // A-H
const DASH_PROJ = { headerRow: 11, startCol: 10, cols: 8, clearRows: 2500 };  // J-Q

// Dashboard (2) flat outputs
const DASH2_EMP = { headerRow: 3, startCol: 1, cols: 6, clearRows: 12000 };   // A-F
const DASH2_PROJ = { headerRow: 3, startCol: 8, cols: 6, clearRows: 12000 };  // H-M

// Employees tab blocks
const EMP_SUM = { headerRow: 3, startCol: 1, cols: 8, clearRows: 6000 };
const EMP_REG = { headerRow: 3, startCol: 11, cols: 5, clearRows: 6000 };
const EMP_OV = { headerRow: 3, startCol: 17, cols: 6, clearRows: 6000 };

// Projects tab blocks
const PROJ_SUM = { headerRow: 3, startCol: 1, cols: 8, clearRows: 6000 };
const PROJ_REG = { headerRow: 3, startCol: 11, cols: 5, clearRows: 6000 };
const PROJ_OV = { headerRow: 3, startCol: 17, cols: 6, clearRows: 6000 };

// Adjust Hours block
const ADJ = { headerRow: 1, startCol: 1, cols: 12, clearRows: 6000 };

// Rates block
const RATES_CLEAR = { startRow0: 1, startCol0: 0, rows: 3000, cols: 6 };

// Dashboard (2) filter cells (linked to Dashboard B2/B3 in your setup)
const DASH2_FILTERS_CELL_FROM = "O4";
const DASH2_FILTERS_CELL_TO = "P4";

// ---------- helpers ----------

function safeSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const ws = workbook.getWorksheet(name);
  if (!ws) throw new Error(`Missing sheet: ${name}`);
  return ws;
}

function norm(s: string): string {
  return (s || "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}

function startOfDay(d: Date): Date {
  const x = new Date(d.getTime());
  x.setHours(0, 0, 0, 0);
  return x;
}

/**
 * Excel serial day 0 = 1899-12-30 (local).
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

/**
 * Robust date parsing:
 * - excel serial numbers
 * - "Friday, January 2, 2026" (weekday prefix)
 * - normal date strings
 */
function parseDateCellRobust(v: unknown): Date | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return excelDateToJSDate(v);

  const s = String(v).trim();
  if (!s) return null;

  const cleaned = s.replace(/^[A-Za-z]+,\s*/, "");
  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? null : d;
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

  if (bestScore < 1) return 0;
  return bestRow;
}

function clearBlockContents(
  ws: ExcelScript.Worksheet,
  headerRow: number,
  startCol: number,
  cols: number,
  clearRows: number
): void {
  const startRow0 = headerRow;   // clears AFTER the header row
  const startCol0 = startCol - 1;
  ws.getRangeByIndexes(startRow0, startCol0, clearRows, cols)
    .clear(ExcelScript.ClearApplyTo.contents);
}

// ---------- Raw_Sprout min/max ----------

function detectSproutMinMaxDates(workbook: ExcelScript.Workbook): { min: Date; max: Date } {
  const ws = safeSheet(workbook, SHEETS.RAW_SPROUT);

  const used = ws.getUsedRange(true);
  if (!used) throw new Error("Raw_Sprout is empty.");

  const values = used.getValues();
  if (!values || values.length < 2) throw new Error("Raw_Sprout has no data rows.");

  const required: string[][] = [["logdate", "log date", "date"]];
  const headerRowIdx = detectHeaderRow(values, required, 30);

  const headers = values[headerRowIdx];
  const hmap = findHeaderMap(headers);

  const cLogDate = getCol(hmap, ["logdate", "log date", "date"]);
  if (cLogDate < 0) throw new Error("Raw_Sprout must contain a LogDate column.");

  let minD: Date | null = null;
  let maxD: Date | null = null;

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const d = parseDateCellRobust(values[r][cLogDate]);
    if (!d) continue;

    const sd = startOfDay(d);
    if (!minD || sd.getTime() < minD.getTime()) minD = sd;
    if (!maxD || sd.getTime() > maxD.getTime()) maxD = sd;
  }

  if (!minD || !maxD) throw new Error("Could not detect Raw_Sprout min/max dates.");
  return { min: minD, max: maxD };
}

// ---------- TBRegularOT reset (header-safe) ----------

function resetTBRegularOTOutputsOnly(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.REGVSOV);

  let tbl: ExcelScript.Table;
  try {
    tbl = ws.getTable(TABLES.REGULAR_OT);
  } catch {
    return;
  }

  const headerVals = tbl.getHeaderRowRange().getValues();
  const headerRow = (headerVals && headerVals.length) ? headerVals[0] : [];
  const hmap = findHeaderMap(headerRow);

  const body = tbl.getRangeBetweenHeaderAndTotal();
  if (!body) return;

  const rowCount = body.getRowCount();
  const colCount = body.getColumnCount();
  if (rowCount <= 0 || colCount <= 0) return;

  // Columns we MUST preserve (by header name)
  const keepCols = new Set<number>();

  const cWSStart = getCol(hmap, ["workshift start", "work shift start"]);
  const cWSEnd = getCol(hmap, ["workshift end", "work shift end"]);
  const cEmpKey = getCol(hmap, ["employee key"]);
  const cEmpName = getCol(hmap, ["employee name"]);

  if (cWSStart >= 0) keepCols.add(cWSStart);
  if (cWSEnd >= 0) keepCols.add(cWSEnd);
  if (cEmpKey >= 0) keepCols.add(cEmpKey);
  if (cEmpName >= 0) keepCols.add(cEmpName);

  // If for some reason headers aren't found, do NOTHING (safer than wiping)
  if (keepCols.size === 0) return;

  // Clear every other column (outputs) in the table body
  for (let c = 0; c < colCount; c++) {
    if (keepCols.has(c)) continue;
    body.getColumn(c).clear(ExcelScript.ClearApplyTo.contents);
  }
}

// ---------- Reset other outputs ----------

function resetDashboardOutputs(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.DASH);

  clearBlockContents(ws, DASH_EMP.headerRow, DASH_EMP.startCol, DASH_EMP.cols, DASH_EMP.clearRows);
  clearBlockContents(ws, DASH_PROJ.headerRow, DASH_PROJ.startCol, DASH_PROJ.cols, DASH_PROJ.clearRows);

  ws.getRange("B5:B8").clear(ExcelScript.ClearApplyTo.contents);

  try { ws.getRange("E6").setValue(false); } catch { /* ignore */ }
}

function resetDashboard2Outputs(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.DASH2);
  clearBlockContents(ws, DASH2_EMP.headerRow, DASH2_EMP.startCol, DASH2_EMP.cols, DASH2_EMP.clearRows);
  clearBlockContents(ws, DASH2_PROJ.headerRow, DASH2_PROJ.startCol, DASH2_PROJ.cols, DASH2_PROJ.clearRows);
}

function resetEmployees(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.EMP);
  clearBlockContents(ws, EMP_SUM.headerRow, EMP_SUM.startCol, EMP_SUM.cols, EMP_SUM.clearRows);
  clearBlockContents(ws, EMP_REG.headerRow, EMP_REG.startCol, EMP_REG.cols, EMP_REG.clearRows);
  clearBlockContents(ws, EMP_OV.headerRow, EMP_OV.startCol, EMP_OV.cols, EMP_OV.clearRows);
}

function resetProjects(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.PROJ);
  clearBlockContents(ws, PROJ_SUM.headerRow, PROJ_SUM.startCol, PROJ_SUM.cols, PROJ_SUM.clearRows);
  clearBlockContents(ws, PROJ_REG.headerRow, PROJ_REG.startCol, PROJ_REG.cols, PROJ_REG.clearRows);
  clearBlockContents(ws, PROJ_OV.headerRow, PROJ_OV.startCol, PROJ_OV.cols, PROJ_OV.clearRows);
}

function resetAdjustHours(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.ADJUST);
  clearBlockContents(ws, ADJ.headerRow, ADJ.startCol, ADJ.cols, ADJ.clearRows);
}

function resetRates(workbook: ExcelScript.Workbook): void {
  const ws = safeSheet(workbook, SHEETS.RATES);
  ws.getRangeByIndexes(RATES_CLEAR.startRow0, RATES_CLEAR.startCol0, RATES_CLEAR.rows, RATES_CLEAR.cols)
    .clear(ExcelScript.ClearApplyTo.contents);
}

// ---------- Date range reset (Raw_Sprout min/max) ----------

function isFormulaCell(r: ExcelScript.Range): boolean {
  try {
    const f = r.getFormula();
    return !!f && String(f).trim().startsWith("=");
  } catch {
    return false;
  }
}

function forceCalc(workbook: ExcelScript.Workbook): void {
  const app = workbook.getApplication();
  try { app.setCalculationMode(ExcelScript.CalculationMode.automatic); } catch { /* ignore */ }
  try { app.calculate(ExcelScript.CalculationType.full); } catch { /* ignore */ }
}

function setDateRangeToSproutMinMax(workbook: ExcelScript.Workbook): void {
  // Make sure we’re not reading stale values after paste
  forceCalc(workbook);

  const mm = detectSproutMinMaxDates(workbook);

  // Set Dashboard B2:B3 (this should drive O4:P4 via your linkage)
  const wsDash = safeSheet(workbook, SHEETS.DASH);
  wsDash.getRange("B2").setValue(jsDateToExcelSerial(mm.min));
  wsDash.getRange("B3").setValue(jsDateToExcelSerial(mm.max));
  wsDash.getRange("B2:B3").setNumberFormatLocal("m/d/yyyy");

  // If your O4/P4 are NOT formulas (or are blank), we can also set them safely
  // without breaking any linkage. If they ARE formulas, we leave them alone.
  const wsD2 = safeSheet(workbook, SHEETS.DASH2);

  const rO4 = wsD2.getRange(DASH2_FILTERS_CELL_FROM);
  const rP4 = wsD2.getRange(DASH2_FILTERS_CELL_TO);

  if (!isFormulaCell(rO4)) rO4.setValue(jsDateToExcelSerial(mm.min));
  if (!isFormulaCell(rP4)) rP4.setValue(jsDateToExcelSerial(mm.max));

  try {
    wsD2.getRange(`${DASH2_FILTERS_CELL_FROM}:${DASH2_FILTERS_CELL_TO}`).setNumberFormatLocal("m/d/yyyy");
  } catch {
    // ignore
  }

  // Final calc so all downstream formulas update
  forceCalc(workbook);
}

// ---------- MAIN ----------

function main(workbook: ExcelScript.Workbook) {
  // 1) Clear outputs but keep template inputs
  resetTBRegularOTOutputsOnly(workbook);

  resetDashboardOutputs(workbook);
  resetDashboard2Outputs(workbook);
  resetEmployees(workbook);
  resetProjects(workbook);
  resetAdjustHours(workbook);
  resetRates(workbook);

  // 2) Reset date range every run
  setDateRangeToSproutMinMax(workbook);
}
