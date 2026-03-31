/**
 * Reset Raw Inputs (STRICT RANGES ONLY)
 *
 * Clears ONLY these ranges (starting row 2, keeps headers in row 1):
 * - Raw_Sprout:          A2:J(last used row within A:J)
 * - Raw_Invoice:         A2:AU(last used row within A:AU)
 * - Raw_Approval Center: A2:H(last used row within A:H)
 *
 * IMPORTANT:
 * - Does NOT touch anything beyond those columns.
 * - Does NOT clear row 1 headers.
 * - If there are no data rows, it does nothing for that sheet.
 */

const RANGES: { sheet: string; firstCol: string; lastCol: string }[] = [
    { sheet: "Raw_Sprout", firstCol: "A", lastCol: "J" },
    { sheet: "Raw_Invoice", firstCol: "A", lastCol: "AU" },
    { sheet: "Raw_Approval Center", firstCol: "A", lastCol: "H" }
];

function safeSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
    const ws = workbook.getWorksheet(name);
    if (!ws) throw new Error(`Missing sheet: ${name}`);
    return ws;
}

// Find the last row (1-based) that has ANY content in the target columns.
// Returns 1 if only headers (row 1) or the range is empty.
function getLastDataRowInCols(ws: ExcelScript.Worksheet, lastCol: string): number {
    // Used range could extend beyond the columns we care about, so we scan only A:lastCol.
    const used = ws.getUsedRange(true);
    if (!used) return 1;

    const lastUsedRow1 = used.getRowIndex() + used.getRowCount(); // 1-based
    if (lastUsedRow1 <= 1) return 1;

    const scanRange = ws.getRange(`A1:${lastCol}${lastUsedRow1}`);
    const vals = scanRange.getValues(); // row 1..lastUsedRow1, col A..lastCol

    // Walk upward to find the last row with any non-empty cell
    for (let r = vals.length - 1; r >= 1; r--) { // start from bottom, skip header row index 0
        const row = vals[r];
        let hasData = false;
        for (let c = 0; c < row.length; c++) {
            const v = row[c];
            if (v !== null && v !== undefined && String(v).trim() !== "") {
                hasData = true;
                break;
            }
        }
        if (hasData) return r + 1; // convert 0-based array index to 1-based row number
    }

    return 1;
}

function clearStrictRange(ws: ExcelScript.Worksheet, firstCol: string, lastCol: string): void {
    const lastRow1 = getLastDataRowInCols(ws, lastCol);
    if (lastRow1 <= 1) return; // nothing to clear

    const addr = `${firstCol}2:${lastCol}${lastRow1}`;
    ws.getRange(addr).clear(ExcelScript.ClearApplyTo.contents);
}

function main(workbook: ExcelScript.Workbook) {
    for (const r of RANGES) {
        const ws = safeSheet(workbook, r.sheet);
        clearStrictRange(ws, r.firstCol, r.lastCol);
    }
}
