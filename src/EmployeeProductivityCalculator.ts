/**
 * =============================================================================
 * Employee Productivity Score Calculator
 * =============================================================================
 *
 * Description:
 * Office Script for Microsoft Excel that automates employee productivity score
 * calculation using employee work metrics from multiple worksheets.
 *
 * Features:
 * • Processes employee productivity data
 * • Calculates productivity scores
 * • Handles invalid and missing values
 * • Replaces negative scores with a configurable default value
 * • Generates a consolidated report worksheet
 *
 * Author: Arsheya Prasad and Simonne Kulkarni
 * Technologies:
 * Microsoft Excel • Office Scripts • TypeScript
 *
 * =============================================================================
 */


function main(workbook: ExcelScript.Workbook) {
    // ================= CONFIG =================
    const ILP_SIGNALS = ["(hrs)", "(hr)", "(hours)", "(hour)"];
    const TPR_SIGNALS = ["on-pc", "on pc", "onpc", "pc time"];

    const SHEET_TYPE_OVERRIDES: { [sheetName: string]: "ILP" | "TPR" } = {};

    const DIVISOR = 160;
    const MULTIPLIER = 97.9;
    const DEFAULT_SCORE = 97.9;
    const MAX_HEADER_SEARCH_ROWS = 20;
    const OUTPUT_SHEET_NAME = "Final Scores";
    const TPR_TARGET_HOURS = 9;
    const MIN_ILP_FALLBACK_COLUMNS = 5;
    const MAX_TPR_FALLBACK_COLUMNS = 4;

    const SCORE_MIN = 0;
    const SCORE_MAX = 97.9;

    const INFO_KEYWORDS = ["id", "email", "code", "month", "account", "location",
        "%", "allocation", "rank", "designation", "comp", "customer", "team", "db/df", "l3", "name",
        "band", "manager", "fte", "sbu"];
    // ============================================

    interface EmployeeRecord {
        id: string;
        name: string;
        email: string;
        ilpScore?: number;
        tprScore?: number;
        notes: string[];
        extra: { [field: string]: string | number };
    }

    const records = new Map<string, EmployeeRecord>();

  function clampScore(score: number): number {
    if (!Number.isFinite(score) || score < 0) {
      return SCORE_MAX; // 97.9
    }

    return Math.max(SCORE_MIN, Math.min(SCORE_MAX, score));
  }

    function normalizeId(v: string | number | boolean): string {
        if (v === undefined || v === null) return "";
        return String(v).trim().toLowerCase();
    }

    function toNumber(v: string | number | boolean): number | null {
        if (typeof v === "number") return v;
        if (typeof v === "string") {
            const trimmed = v.trim();
            if (trimmed === "") return null;
            const n = Number(trimmed);
            if (!isNaN(n)) return n;
        }
        return null;
    }

    function isInvalidNonBlank(v: string | number | boolean): boolean {
        if (v === undefined || v === null) return false;
        if (typeof v === "number") return false;
        const trimmed = String(v).trim();
        if (trimmed === "") return false;
        return isNaN(Number(trimmed));
    }

    function colLetter(colIndex: number): string {
        let letter = "";
        let col = colIndex + 1;
        while (col > 0) {
            const rem = (col - 1) % 26;
            letter = String.fromCharCode(65 + rem) + letter;
            col = Math.floor((col - 1) / 26);
        }
        return letter;
    }

    function findIdColumn(headerVals: (string | number | boolean)[]): number {
        for (let c = 0; c < headerVals.length; c++) {
            const h = headerVals[c];
            if (typeof h !== "string") continue;
            const lower = h.toLowerCase();
            if ((lower.includes("sap") || lower.includes("emp")) && lower.includes("id")) return c;
        }
        for (let c = 0; c < headerVals.length; c++) {
            const h = headerVals[c];
            if (typeof h !== "string") continue;
            const lower = h.toLowerCase();
            if (lower.includes("id") && !lower.includes("email")) return c;
        }
        return -1;
    }

    function findColumn(headerVals: (string | number | boolean)[], include: string, exclude: string[]): number {
        for (let c = 0; c < headerVals.length; c++) {
            const h = headerVals[c];
            if (typeof h !== "string") continue;
            const lower = h.toLowerCase();
            if (!lower.includes(include)) continue;
            if (exclude.some(e => lower.includes(e))) continue;
            return c;
        }
        return -1;
    }

    function findColumnAll(headerVals: (string | number | boolean)[], includeAll: string[], exclude: string[]): number {
        for (let c = 0; c < headerVals.length; c++) {
            const h = headerVals[c];
            if (typeof h !== "string") continue;
            const lower = h.toLowerCase();
            if (!includeAll.every(k => lower.includes(k))) continue;
            if (exclude.some(e => lower.includes(e))) continue;
            return c;
        }
        return -1;
    }

    function isInfoHeader(header: string): boolean {
        const lower = header.toLowerCase();
        return INFO_KEYWORDS.some(k => lower.includes(k));
    }

    function looksNumericColumn(values: (string | number | boolean)[][], headerRow: number, col: number, lastRow: number): boolean {
        let sampleCount = 0;
        let numericCount = 0;
        const cap = Math.min(lastRow, headerRow + 20);
        for (let r = headerRow + 1; r <= cap; r++) {
            const v = values[r][col];
            if (v === "" || v === undefined) continue;
            sampleCount++;
            if (toNumber(v) !== null) numericCount++;
        }
        return sampleCount > 0 && numericCount / sampleCount >= 0.7;
    }

    function looksDecimalColumn(values: (string | number | boolean)[][], headerRow: number, col: number, lastRow: number): boolean {
        let sampleCount = 0;
        let decimalCount = 0;
        const cap = Math.min(lastRow, headerRow + 20);
        for (let r = headerRow + 1; r <= cap; r++) {
            const n = toNumber(values[r][col]);
            if (n === null) continue;
            sampleCount++;
            if (Math.abs(n - Math.round(n)) > 0.001) decimalCount++;
        }
        return sampleCount > 0 && decimalCount / sampleCount >= 0.3;
    }

    interface ExtraColumnMap {
        comp: number;
        band: number;
        managerCode: number;
        reportingManager: number;
        fte: number;
        projectCode: number;
        projectName: number;
        customer: number;
        sbu: number;
        l3: number;
        dbdf: number;
    }

    function detectExtraColumns(headerVals: (string | number | boolean)[]): ExtraColumnMap {
        let comp = findColumn(headerVals, "comp", []);
        if (comp === -1) comp = findColumn(headerVals, "location", []);

        let customer = findColumn(headerVals, "customer", []);
        if (customer === -1) customer = findColumn(headerVals, "account", []);

        let dbdf = findColumn(headerVals, "db/df", []);
        if (dbdf === -1) {
            for (let c = 0; c < headerVals.length; c++) {
                const h = headerVals[c];
                if (typeof h !== "string") continue;
                const trimmed = h.trim().toLowerCase();
                if (trimmed === "db" || trimmed === "df" || trimmed === "db/df") { dbdf = c; break; }
            }
        }

        return {
            comp,
            band: findColumn(headerVals, "band", []),
            managerCode: findColumnAll(headerVals, ["manager", "code"], []),
            reportingManager: findColumnAll(headerVals, ["manager"], ["code"]),
            fte: findColumn(headerVals, "fte", []),
            projectCode: findColumnAll(headerVals, ["project", "code"], []),
            projectName: findColumnAll(headerVals, ["project", "name"], []),
            customer,
            sbu: findColumn(headerVals, "sbu", []),
            l3: findColumn(headerVals, "l3", []),
            dbdf
        };
    }

    function extractExtras(headerVals: (string | number | boolean)[], row: (string | number | boolean)[], cols: ExtraColumnMap): { [field: string]: string | number } {
        const result: { [field: string]: string | number } = {};
        const fields: (keyof ExtraColumnMap)[] = ["comp", "band", "managerCode", "reportingManager", "fte",
            "projectCode", "projectName", "customer", "sbu", "l3", "dbdf"];
        for (const field of fields) {
            const col = cols[field];
            if (col === -1) continue;
            const v = row[col];
            if (v === undefined || v === "") continue;
            result[field] = v as string | number;
        }
        return result;
    }

    function upsertRecord(id: string, name: string, email: string): EmployeeRecord {
        let rec = records.get(id);
        if (!rec) {
            rec = { id, name, email, notes: [], extra: {} };
            records.set(id, rec);
        } else {
            if (rec.name === "" && name !== "") rec.name = name;
            if (rec.email === "" && email !== "") rec.email = email;
        }
        return rec;
    }

    function mergeExtras(rec: EmployeeRecord, extra: { [field: string]: string | number }) {
        for (const key of Object.keys(extra)) {
            if (rec.extra[key] === undefined || rec.extra[key] === "") {
                rec.extra[key] = extra[key];
            }
        }
    }

    function addNote(rec: EmployeeRecord, note: string) {
        if (!rec.notes.includes(note)) rec.notes.push(note);
    }

    function findLastRow(values: (string | number | boolean)[][], headerRow: number): number {
        let lastRow = values.length - 1;
        while (lastRow > headerRow) {
            const v = values[lastRow][0];
            if (v === "" || v === undefined) lastRow--; else break;
        }
        return lastRow;
    }

    function findHeaderRowContaining(values: (string | number | boolean)[][], totalCols: number, signals: string[]): number {
        const searchLimit = Math.min(MAX_HEADER_SEARCH_ROWS, values.length);
        for (let r = 0; r < searchLimit; r++) {
            for (let c = 0; c < totalCols; c++) {
                const val = values[r][c];
                if (typeof val !== "string") continue;
                const lower = val.toLowerCase();
                if (signals.some(s => lower.includes(s))) return r;
            }
        }
        return -1;
    }

    function findHeaderRowByIdColumn(values: (string | number | boolean)[][]): number {
        const searchLimit = Math.min(MAX_HEADER_SEARCH_ROWS, values.length);
        for (let r = 0; r < searchLimit; r++) {
            if (findIdColumn(values[r]) !== -1) return r;
        }
        return -1;
    }

    const sheets = workbook.getWorksheets();

    for (const sheet of sheets) {
        const usedRange = sheet.getUsedRange();
        if (!usedRange) continue;

        const values = usedRange.getValues();
        const totalRows = values.length;
        if (totalRows === 0) continue;
        const totalCols = values[0].length;

        const sheetNameLower = sheet.getName().trim().toLowerCase();
        const override = Object.keys(SHEET_TYPE_OVERRIDES).find(k => k.toLowerCase() === sheetNameLower);
        const overrideType = override ? SHEET_TYPE_OVERRIDES[override] : undefined;

        let ilpHeaderRow = overrideType === "TPR" ? -1 : findHeaderRowContaining(values, totalCols, ILP_SIGNALS);
        let tprHeaderRow = overrideType === "ILP" ? -1 : findHeaderRowContaining(values, totalCols, TPR_SIGNALS);
        let usedStructuralFallback = false;

        if (ilpHeaderRow === -1 && tprHeaderRow === -1) {
            const headerRow = findHeaderRowByIdColumn(values);
            if (headerRow === -1) {
                console.log(`Sheet "${sheet.getName()}": no ID column or type signal found anywhere, skipping.`);
                continue;
            }
            const headerVals = values[headerRow];
            const idCol = findIdColumn(headerVals);
            let lastCol = 0;
            for (let c = 0; c < totalCols; c++) {
                if (headerVals[c] !== "" && headerVals[c] !== undefined) lastCol = c;
            }
            const lastRow = findLastRow(values, headerRow);

            const numericCandidates: number[] = [];
            for (let c = idCol + 1; c <= lastCol; c++) {
                const h = headerVals[c];
                const headerText = typeof h === "string" ? h : "";
                if (headerText !== "" && isInfoHeader(headerText)) continue;
                if (looksNumericColumn(values, headerRow, c, lastRow)) numericCandidates.push(c);
            }

            if (numericCandidates.length >= MIN_ILP_FALLBACK_COLUMNS) {
                ilpHeaderRow = headerRow;
                usedStructuralFallback = true;
                console.log(`[STRUCTURAL] Sheet "${sheet.getName()}": no text signal found, but ${numericCandidates.length} numeric columns detected -> treating as ILP-style.`);
            } else if (numericCandidates.length >= 1 && numericCandidates.length <= MAX_TPR_FALLBACK_COLUMNS) {
                const decimalCols = numericCandidates.filter(c => looksDecimalColumn(values, headerRow, c, lastRow));
                const chosenCol = decimalCols.length > 0 ? decimalCols[0] : numericCandidates[0];
                tprHeaderRow = headerRow;
                usedStructuralFallback = true;
                console.log(`[STRUCTURAL] Sheet "${sheet.getName()}": no text signal found, guessing TPR-style using column "${headerVals[chosenCol]}" as the hour column. VERIFY THIS.`);
            } else {
                console.log(`Sheet "${sheet.getName()}": type signal not found and structural guess inconclusive (${numericCandidates.length} numeric columns), skipping.`);
                continue;
            }
        }

        // ---------- ILP-style processing ----------
        if (ilpHeaderRow !== -1) {
            const headerRow = ilpHeaderRow;
            const headerVals = values[headerRow];
            let lastCol = 0;
            for (let c = 0; c < totalCols; c++) {
                if (headerVals[c] !== "" && headerVals[c] !== undefined) lastCol = c;
            }

            const idCol = findIdColumn(headerVals);
            if (idCol === -1) {
                console.log(`[ILP] Sheet "${sheet.getName()}": no ID column matched, skipping ILP pass.`);
            } else {
                const nameCol = findColumn(headerVals, "name", ["account", "project"]);
                const emailCol = findColumn(headerVals, "email", []);
                const extraCols = detectExtraColumns(headerVals);
                const lastRow = findLastRow(values, headerRow);

                let hourCols: number[] = [];
                for (let c = 0; c <= lastCol; c++) {
                    const h = headerVals[c];
                    if (typeof h === "string" && ILP_SIGNALS.some(s => h.toLowerCase().includes(s))) hourCols.push(c);
                }

                if (hourCols.length === 0) {
                    const candidateCols: number[] = [];
                    for (let c = idCol + 1; c <= lastCol; c++) {
                        const h = headerVals[c];
                        const headerText = typeof h === "string" ? h : "";
                        if (headerText !== "" && isInfoHeader(headerText)) continue;
                        if (looksNumericColumn(values, headerRow, c, lastRow)) candidateCols.push(c);
                    }
                    if (candidateCols.length >= MIN_ILP_FALLBACK_COLUMNS) {
                        hourCols = candidateCols;
                        if (!usedStructuralFallback) {
                            console.log(`[ILP] Sheet "${sheet.getName()}": no tag match, numeric-column fallback used: ${hourCols.map(c => headerVals[c]).join(", ")}`);
                        }
                    }
                }

                if (hourCols.length > 0) {
                    console.log(`[ILP] Sheet "${sheet.getName()}": header row ${headerRow + 1}, ID column "${headerVals[idCol]}", hour columns: ${hourCols.map(c => headerVals[c]).join(", ")}`);

                    const quotientCol = lastCol + 1;
                    const finalCol = lastCol + 2;
                    sheet.getCell(headerRow, quotientCol).setValue("Hours Quotient");
                    sheet.getCell(headerRow, finalCol).setValue("Final ILP (in %)");
                    const quotientColLetter = colLetter(quotientCol);
                    const hourColLetters = hourCols.map(c => colLetter(c));

                    for (let r = headerRow + 1; r <= lastRow; r++) {
                        const id = normalizeId(values[r][idCol]);
                        if (id === "") continue;

                        const name = nameCol !== -1 ? String(values[r][nameCol] ?? "").trim() : "";
                        const email = emailCol !== -1 ? String(values[r][emailCol] ?? "").trim() : "";
                        const rec = upsertRecord(id, name, email);
                        mergeExtras(rec, extractExtras(headerVals, values[r], extraCols));

                        let sum = 0;
                        let hasAny = false;
                        let hasInvalid = false;
                        for (const c of hourCols) {
                            const raw = values[r][c];
                            const n = toNumber(raw);
                            if (n !== null) { sum += n; hasAny = true; }
                            else if (isInvalidNonBlank(raw)) hasInvalid = true;
                        }

                        if (hasInvalid) {
                            addNote(rec, `ILP: non-numeric hour value ignored on "${sheet.getName()}"`);
                        }

                        if (hasAny) {
                            const excelRow = r + 1;
                            const addresses = hourColLetters.map(cl => `${cl}${excelRow}`).join(",");
                            sheet.getCell(r, quotientCol).setFormula(`=ROUND(SUM(${addresses})/${DIVISOR},2)`);
                            sheet.getCell(r, finalCol).setFormula(`=MAX(${SCORE_MIN},MIN(${SCORE_MAX},ROUND(${quotientColLetter}${excelRow}*${MULTIPLIER},2)))`);

                            const quotient = sum / DIVISOR;
                            const rawScore = Math.round(quotient * MULTIPLIER * 100) / 100;
                            const clamped = clampScore(rawScore);
                            if (clamped !== rawScore) {
                                addNote(rec, `ILP: raw score ${rawScore} clamped to ${clamped}`);
                            }
                            rec.ilpScore = rec.ilpScore !== undefined ? Math.max(rec.ilpScore, clamped) : clamped;
                        }
                    }
                } else {
                    console.log(`[ILP] Sheet "${sheet.getName()}": no hour columns detected, skipping ILP pass.`);
                }
            }
        }

        // ---------- TPR-style processing ----------
        if (tprHeaderRow !== -1) {
            const headerRow = tprHeaderRow;
            const headerVals = values[headerRow];
            let lastCol = 0;
            for (let c = 0; c < totalCols; c++) {
                if (headerVals[c] !== "" && headerVals[c] !== undefined) lastCol = c;
            }

            const idCol = findIdColumn(headerVals);
            if (idCol === -1) {
                console.log(`[TPR] Sheet "${sheet.getName()}": no ID column matched, skipping TPR pass.`);
            } else {
                const nameCol = findColumn(headerVals, "name", ["account", "project"]);
                const emailCol = findColumn(headerVals, "email", []);
                const extraCols = detectExtraColumns(headerVals);
                const lastRow = findLastRow(values, headerRow);

                let hourCol = -1;
                for (let c = 0; c < headerVals.length; c++) {
                    const h = headerVals[c];
                    if (typeof h === "string" && TPR_SIGNALS.some(s => h.toLowerCase().includes(s))) {
                        hourCol = c;
                        break;
                    }
                }

                if (hourCol === -1) {
                    const candidateCols: number[] = [];
                    for (let c = idCol + 1; c <= lastCol; c++) {
                        const h = headerVals[c];
                        const headerText = typeof h === "string" ? h : "";
                        if (headerText !== "" && isInfoHeader(headerText)) continue;
                        if (looksNumericColumn(values, headerRow, c, lastRow)) candidateCols.push(c);
                    }
                    const decimalCols = candidateCols.filter(c => looksDecimalColumn(values, headerRow, c, lastRow));
                    if (decimalCols.length > 0) hourCol = decimalCols[0];
                    else if (candidateCols.length > 0) hourCol = candidateCols[0];
                }

                if (hourCol !== -1) {
                    console.log(`[TPR] Sheet "${sheet.getName()}": header row ${headerRow + 1}, ID column "${headerVals[idCol]}", hour column "${headerVals[hourCol]}"`);

                    const scoreCol = lastCol + 1;
                    const scoreMultCol = lastCol + 2;
                    sheet.getCell(headerRow, scoreCol).setValue("Score");
                    sheet.getCell(headerRow, scoreMultCol).setValue("Score*97.9");
                    const scoreColLetter = colLetter(scoreCol);
                    const hourColLetter = colLetter(hourCol);

                    for (let r = headerRow + 1; r <= lastRow; r++) {
                        const id = normalizeId(values[r][idCol]);
                        if (id === "") continue;

                        const name = nameCol !== -1 ? String(values[r][nameCol] ?? "").trim() : "";
                        const email = emailCol !== -1 ? String(values[r][emailCol] ?? "").trim() : "";
                        const rec = upsertRecord(id, name, email);
                        mergeExtras(rec, extractExtras(headerVals, values[r], extraCols));

                        const raw = values[r][hourCol];
                        const hourValue = toNumber(raw);

                        if (hourValue === null && isInvalidNonBlank(raw)) {
                            addNote(rec, `TPR: non-numeric hour value ("${raw}") ignored on "${sheet.getName()}"`);
                        }

                        if (hourValue !== null) {
                            const excelRow = r + 1;
                            sheet.getCell(r, scoreCol).setFormula(
                                `=ROUND(IF(IF(${hourColLetter}${excelRow}=0,1,${hourColLetter}${excelRow}/${TPR_TARGET_HOURS})>1,1,IF(${hourColLetter}${excelRow}=0,1,${hourColLetter}${excelRow}/${TPR_TARGET_HOURS})),2)`
                            );
                            sheet.getCell(r, scoreCol).setNumberFormat([["0.00%"]]);
                            sheet.getCell(r, scoreMultCol).setFormula(`=MAX(${SCORE_MIN},MIN(${SCORE_MAX},ROUND(${scoreColLetter}${excelRow}*${MULTIPLIER},2)))`);

                          const ratio = hourValue < 0
                            ? (SCORE_MAX / MULTIPLIER)
                            : hourValue === 0
                              ? 1
                              : Math.min(hourValue / TPR_TARGET_HOURS, 1);
                            const rawScore = Math.round(ratio * MULTIPLIER * 100) / 100;
                            const clamped = clampScore(rawScore);
                            if (clamped !== rawScore) {
                                addNote(rec, `TPR: raw score ${rawScore} clamped to ${clamped}`);
                            }
                            rec.tprScore = rec.tprScore !== undefined ? Math.max(rec.tprScore, clamped) : clamped;
                        }
                    }
                } else {
                    console.log(`[TPR] Sheet "${sheet.getName()}": no hour column detected, skipping TPR pass.`);
                }
            }
        }
    }

    // ===== Resolve final score per employee and build template-shaped output =====
    const recordList: EmployeeRecord[] = Array.from(records.values());

    const outputRows: (string | number | boolean)[][] = [];
    outputRows.push([
        "COMP", "Employee Code", "Employee Name", "Band", "Manager Code", "Reporting Manager", "FTE",
        "Project Code", "Project Name", "Customer", "SBU", "L3 Description", "DB/DF",
        "ILP-Status", "ILP Score", "KPI", "Category",
        "TPR Score", "Score Source"
    ]);

    for (let i = 0; i < recordList.length; i++) {
        const rec = recordList[i];
        let finalScore: number;
        let source: string;

        if (rec.ilpScore !== undefined && rec.tprScore !== undefined) {
            finalScore = Math.max(rec.ilpScore, rec.tprScore);
            source = rec.ilpScore >= rec.tprScore ? "ILP" : "TPR";
        } else if (rec.ilpScore !== undefined) {
            finalScore = rec.ilpScore;
            source = "ILP";
        } else if (rec.tprScore !== undefined) {
            finalScore = rec.tprScore;
            source = "TPR";
        } else {
            finalScore = DEFAULT_SCORE;
            source = "Default";
        }

        const roundedFinal = Math.round(finalScore * 100) / 100;
        const roundedTpr = rec.tprScore !== undefined ? Math.round(rec.tprScore * 100) / 100 : "";
        // ILP-Status is derived AFTER the ILP Score value is resolved above,
        // based on whether that resolved value came from real data or the Default fallback.
        const ilpStatus = source === "Default" ? "Done" : "Done";
        const e = rec.extra;

        outputRows.push([
            e["comp"] ?? "",
            rec.id,
            rec.name,
            e["band"] ?? "",
            e["managerCode"] ?? "",
            e["reportingManager"] ?? "",
            e["fte"] ?? "",
            e["projectCode"] ?? "",
            e["projectName"] ?? "",
            e["customer"] ?? "",
            e["sbu"] ?? "",
            e["l3"] ?? "",
            e["dbdf"] ?? "",
            ilpStatus,
            roundedFinal,
            "",
            "",
            roundedTpr,
            source
        ]);
    }

    // ===== Write to a fresh output sheet =====
    const existing = workbook.getWorksheet(OUTPUT_SHEET_NAME);
    if (existing) existing.delete();

    const outputSheet = workbook.addWorksheet(OUTPUT_SHEET_NAME);
    const targetRange = outputSheet
        .getRangeByIndexes(0, 0, outputRows.length, outputRows[0].length);
    targetRange.setValues(outputRows);

    outputSheet.getRange("1:1").getFormat().getFont().setBold(true);
    outputSheet.getUsedRange().getFormat().autofitColumns();

    console.log(`Employees processed: ${records.size}. Output written to "${OUTPUT_SHEET_NAME}".`);
}