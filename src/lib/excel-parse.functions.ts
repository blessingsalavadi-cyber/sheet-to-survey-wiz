import { createServerFn } from "@tanstack/react-start";
import ExcelJS from "exceljs";

export type FieldType = "short_text" | "long_text" | "number" | "date" | "select" | "multi_select" | "checkbox";

export interface FormField {
  id: string;
  label: string;
  type: FieldType;
  required: boolean;
  options?: string[];
  description?: string;
}

export interface FormSection {
  title: string;
  description?: string;
  fields: FormField[];
}

export interface ParsedForm {
  title: string;
  sections: FormSection[];
  lookups: { name: string; values: string[] }[];
}

function inferType(label: string, sample?: unknown): FieldType {
  const l = label.toLowerCase();
  if (/(date|dob|when|day)/.test(l)) return "date";
  if (/(amount|count|qty|quantity|number|age|price|total|score)/.test(l)) return "number";
  if (/(comment|describe|details|notes|explain|why|how)/.test(l)) return "long_text";
  if (typeof sample === "number") return "number";
  if (sample instanceof Date) return "date";
  return "short_text";
}

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    const obj = v as { text?: string; richText?: { text: string }[]; result?: unknown; formula?: string };
    if (obj.richText) return obj.richText.map((r) => r.text).join("").trim();
    if (obj.text) return String(obj.text).trim();
    if (obj.result != null) return cellText(obj.result as ExcelJS.CellValue);
  }
  return String(v).trim();
}

function slug(s: string, i: number) {
  return (
    s
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "") || `field_${i}`
  );
}

function extractListFromValidation(
  formulae: string[] | undefined,
  lookups: Map<string, string[]>,
  workbook: ExcelJS.Workbook,
): string[] | null {
  if (!formulae || !formulae.length) return null;
  const f = formulae[0];
  if (!f) return null;
  // Inline list e.g. "Yes,No,Maybe" or '"Yes,No"'
  const stripped = f.replace(/^"|"$/g, "");
  if (!stripped.includes("!") && !stripped.includes(":")) {
    return stripped.split(",").map((x) => x.trim()).filter(Boolean);
  }
  // Range reference e.g. Sheet1!$A$1:$A$5 or Lookups!$B$2:$B$10
  const m = f.match(/^=?([^!]+)!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
  if (m) {
    const [, sheetNameRaw, c1, r1, c2, r2] = m;
    const sheetName = sheetNameRaw.replace(/^'|'$/g, "");
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) return null;
    const values: string[] = [];
    const col1 = colLetterToNum(c1);
    const col2 = colLetterToNum(c2);
    for (let r = Number(r1); r <= Number(r2); r++) {
      for (let c = col1; c <= col2; c++) {
        const t = cellText(ws.getCell(r, c).value);
        if (t) values.push(t);
      }
    }
    if (values.length) lookups.set(`${sheetName}!${c1}${r1}:${c2}${r2}`, values);
    return values;
  }
  // Named range
  return null;
}

function colLetterToNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function detectLookupSheet(ws: ExcelJS.Worksheet): { name: string; values: string[] } | null {
  // Lookup sheet heuristic: one or two columns, short repeating values, name like *list*, *lookup*, *options*
  const name = ws.name;
  if (!/lookup|list|options|values|ref/i.test(name)) return null;
  const col1: string[] = [];
  ws.eachRow((row) => {
    const v = cellText(row.getCell(1).value);
    if (v) col1.push(v);
  });
  if (col1.length < 1) return null;
  return { name, values: col1 };
}

function processSheet(ws: ExcelJS.Worksheet, workbook: ExcelJS.Workbook, lookups: Map<string, string[]>): FormSection | null {
  const rows: { rowNum: number; cells: { col: number; value: ExcelJS.CellValue; address: string }[] }[] = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    const cells: { col: number; value: ExcelJS.CellValue; address: string }[] = [];
    row.eachCell({ includeEmpty: false }, (cell, colNum) => {
      cells.push({ col: colNum, value: cell.value, address: cell.address });
    });
    if (cells.length) rows.push({ rowNum, cells });
  });
  if (!rows.length) return null;

  const fields: FormField[] = [];
  // Detect table-style: first row has 2+ string cells across contiguous columns, and subsequent rows fill same columns
  const first = rows[0];
  const firstCols = first.cells.map((c) => c.col);
  const isTable =
    first.cells.length >= 2 &&
    first.cells.every((c) => typeof c.value === "string") &&
    rows.length >= 2 &&
    rows.slice(1, 4).some((r) => r.cells.some((c) => firstCols.includes(c.col)));

  if (isTable) {
    // Each header column becomes a field. Use validations from row 2 (first data row) when present.
    first.cells.forEach((headerCell, idx) => {
      const label = cellText(headerCell.value);
      if (!label) return;
      const dataCellRow = ws.getRow(first.rowNum + 1);
      const dataCell = dataCellRow.getCell(headerCell.col);
      const validation = (dataCell as unknown as { dataValidation?: { type?: string; formulae?: string[]; allowBlank?: boolean } }).dataValidation;
      let type: FieldType = inferType(label, dataCell.value);
      let options: string[] | undefined;
      if (validation?.type === "list") {
        const opts = extractListFromValidation(validation.formulae, lookups, workbook);
        if (opts && opts.length) {
          options = opts;
          type = "select";
        }
      }
      // Collect a few example values to enrich (not required)
      fields.push({
        id: slug(label, idx),
        label,
        type,
        required: !(validation?.allowBlank ?? true),
        options,
      });
    });
  } else {
    // Two-column Q/A style. Look for column with the longest strings = question column.
    rows.forEach((row, idx) => {
      const labelCell = row.cells[0];
      const answerCell = row.cells[1];
      if (!labelCell) return;
      const label = cellText(labelCell.value);
      if (!label || label.length < 2) return;
      const validation = answerCell
        ? (ws.getCell(answerCell.address) as unknown as { dataValidation?: { type?: string; formulae?: string[]; allowBlank?: boolean } }).dataValidation
        : undefined;
      let type: FieldType = inferType(label, answerCell?.value);
      let options: string[] | undefined;
      if (validation?.type === "list") {
        const opts = extractListFromValidation(validation.formulae, lookups, workbook);
        if (opts && opts.length) {
          options = opts;
          type = "select";
        }
      }
      fields.push({
        id: slug(label, idx),
        label: label.replace(/[:?]\s*$/, ""),
        type,
        required: !(validation?.allowBlank ?? true),
        options,
      });
    });
  }

  if (!fields.length) return null;
  return { title: ws.name, fields };
}

export const parseExcel = createServerFn({ method: "POST" })
  .inputValidator((data: { fileBase64: string; fileName: string }) => data)
  .handler(async ({ data }): Promise<ParsedForm> => {
    const buf = Buffer.from(data.fileBase64, "base64");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buf);

    const lookups = new Map<string, string[]>();
    const lookupSections: { name: string; values: string[] }[] = [];
    const sections: FormSection[] = [];

    workbook.eachSheet((ws) => {
      const lookup = detectLookupSheet(ws);
      if (lookup) {
        lookupSections.push(lookup);
        return;
      }
      const section = processSheet(ws, workbook, lookups);
      if (section) sections.push(section);
    });

    return {
      title: data.fileName.replace(/\.xlsx$/i, ""),
      sections,
      lookups: [
        ...lookupSections,
        ...Array.from(lookups.entries()).map(([name, values]) => ({ name, values })),
      ],
    };
  });