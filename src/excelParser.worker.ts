import * as XLSX from "xlsx";

type InvoiceRow = {
  sellerTaxId: string;
  sellerName: string;
  supplierRaw: string;
  idLast4: string;
  invoiceDate: string;
  plateNo: string;
  totalWithTax: number;
};

const PROC_HEADER_ALIASES = {
  inboundDate: ["入库日期", "入库时间", "单据日期", "业务日期", "日期"],
  supplier: ["供应商", "供应商名称", "供货单位", "往来单位"],
  plateNo: ["车牌号", "车牌号码", "车牌", "车辆号牌"],
  totalWithTax: ["价税合计", "含税合计", "价税总计", "合计金额", "金额"],
} as const;

function normalizeHeader(v: unknown): string {
  return String(v ?? "")
    .replace(/^\ufeff/g, "")
    .replace(/\s/g, "")
    .replace(/[：:]/g, "")
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .trim();
}

function pickIndexByAliases(header: string[], aliases: readonly string[]): number {
  for (const alias of aliases) {
    const a = normalizeHeader(alias);
    const idx = header.findIndex((cell) => normalizeHeader(cell) === a);
    if (idx >= 0) return idx;
  }
  return -1;
}

function rowHasProcurementHeaders(cells: unknown[]): boolean {
  const h = cells.map(normalizeHeader);
  return (
    pickIndexByAliases(h, PROC_HEADER_ALIASES.inboundDate) >= 0 &&
    pickIndexByAliases(h, PROC_HEADER_ALIASES.supplier) >= 0 &&
    pickIndexByAliases(h, PROC_HEADER_ALIASES.totalWithTax) >= 0
  );
}

function cellAt(row: unknown[] | undefined, idx: number): unknown {
  if (!row || idx < 0) return "";
  const v = row[idx];
  return v ?? "";
}

function excelJsCellValue(v: unknown): unknown {
  if (v == null || v === "") return "";
  if (typeof v !== "object") return v;
  const o = v as Record<string, unknown>;
  if (Array.isArray(o.richText)) {
    return (o.richText as { text?: string }[]).map((x) => x.text ?? "").join("");
  }
  if (typeof o.text === "string") return o.text;
  if ("result" in o && o.result != null && o.result !== undefined) return o.result;
  return v;
}

function parseAmount(v: unknown): number {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null) return 0;
  const n = Number(String(v).replace(/,/g, "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function formatInvoiceDateTime(v: unknown): string {
  if (v == null || v === "") return "";
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())} ${pad2(v.getHours())}:${pad2(v.getMinutes())}`;
  }
  if (typeof v === "number" && Number.isFinite(v)) {
    const d = new Date((v - 25569) * 86400000);
    if (Number.isNaN(d.getTime())) return "";
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())} ${pad2(d.getUTCHours())}:${pad2(
      d.getUTCMinutes()
    )}`;
  }

  const s = String(v).trim();
  const normalized = s.replace(/[年./]/g, "-").replace(/月/g, "-").replace(/日/g, "").trim();
  const d = new Date(normalized);
  if (!Number.isNaN(d.getTime())) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }
  return s;
}

function normalizePartyName(input: string): string {
  return String(input ?? "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s\u00a0\u2000-\u200f\u2028-\u202f\u205f\u3000]/g, "")
    .replace(
      /[()（）【】\[\]<>《》「」『』〔〕〖〗.,，。:：;；!！?？'"`~～@#$%^&*+=_|\\/\-—–·•、]/g,
      ""
    )
    .trim();
}

function shouldExcludeCorporateSupplier(name: string): boolean {
  const raw = String(name ?? "").normalize("NFKC");
  const normalized = normalizePartyName(raw);
  if (!normalized) return false;
  if (normalized.includes("公司")) return true;
  if (normalized.includes("回收站")) return true;
  if (/回.{0,3}收.{0,3}站/u.test(raw)) return true;
  return /公.{0,3}司/u.test(raw);
}

function parseSupplierIdentity(rawSupplier: string): { payerName: string; identity18: string } | null {
  const src = String(rawSupplier ?? "").normalize("NFKC").trim();
  if (!src || !/[0-9]/.test(src)) return null;

  const firstDigitIdx = src.search(/[0-9]/);
  const payerName = (firstDigitIdx >= 0 ? src.slice(0, firstDigitIdx) : src).trim();
  if (!payerName) return null;

  const alnum = src.toUpperCase().replace(/[^0-9A-Z]/g, "");
  if (alnum.length < 18) return null;
  const identity18 = alnum.slice(-18);
  if (!/[0-9]/.test(identity18)) return null;
  return { payerName, identity18 };
}

function idLast4FromSupplierRule(supplierRaw: string): string {
  const s = String(supplierRaw ?? "").normalize("NFKC").trim();
  if (s.length < 2) return "";
  const penultimate = s[s.length - 2];
  if (!/[0-9]/.test(penultimate)) return "";
  return s.slice(-4);
}

function extractRowsFromAoa(aoa: (string | number | null | undefined)[][]): InvoiceRow[] | null {
  if (!aoa.length) return null;
  const headerRowIndex = aoa.findIndex((row) => Array.isArray(row) && rowHasProcurementHeaders(row));
  if (headerRowIndex < 0) return null;

  const headerCells = aoa[headerRowIndex];
  const header = headerCells.map(normalizeHeader);
  const idxDate = pickIndexByAliases(header, PROC_HEADER_ALIASES.inboundDate);
  const idxName = pickIndexByAliases(header, PROC_HEADER_ALIASES.supplier);
  const idxPlate = pickIndexByAliases(header, PROC_HEADER_ALIASES.plateNo);
  const idxTotal = pickIndexByAliases(header, PROC_HEADER_ALIASES.totalWithTax);
  if (idxDate < 0 || idxName < 0 || idxTotal < 0) {
    throw new Error("表头解析失败：未能定位「入库日期/供应商/价税合计」对应列（含常见别名）");
  }

  const rows: InvoiceRow[] = [];
  for (let i = headerRowIndex + 1; i < aoa.length; i += 1) {
    const r = aoa[i];
    if (!Array.isArray(r)) continue;
    const supplierRaw = String(cellAt(r, idxName))
      .normalize("NFKC")
      .replace(/[，。；;、,:：]+$/g, "")
      .trim();
    if (!supplierRaw) continue;
    if (shouldExcludeCorporateSupplier(supplierRaw)) continue;

    let sellerName = supplierRaw;
    let sellerTaxId = "";
    const idLast4 = idLast4FromSupplierRule(supplierRaw);
    if (/[0-9]/.test(supplierRaw)) {
      const identity = parseSupplierIdentity(supplierRaw);
      if (!identity) continue;
      sellerName = identity.payerName;
      sellerTaxId = identity.identity18;
    }

    const invoiceDate = formatInvoiceDateTime(cellAt(r, idxDate));
    const plateNo = String(cellAt(r, idxPlate)).normalize("NFKC").trim();
    const totalWithTax = parseAmount(cellAt(r, idxTotal));
    if (!sellerName) continue;
    rows.push({ sellerTaxId, sellerName, supplierRaw, idLast4, invoiceDate, plateNo, totalWithTax });
  }
  return rows;
}

function extractRows(workbook: XLSX.WorkBook): InvoiceRow[] {
  if (!workbook.SheetNames.length) return [];
  let best: InvoiceRow[] = [];
  let headerFoundAnywhere = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const aoa = XLSX.utils.sheet_to_json<(string | number | null | undefined)[]>(sheet, {
      header: 1,
      raw: true,
      defval: "",
    });
    if (!aoa.length) continue;

    const hasHeaderRow = aoa.some((row) => Array.isArray(row) && rowHasProcurementHeaders(row));
    if (hasHeaderRow) headerFoundAnywhere = true;

    const parsed = extractRowsFromAoa(aoa);
    if (parsed && parsed.length > best.length) best = parsed;
  }

  if (!headerFoundAnywhere) {
    throw new Error(
      "未在任何工作表中找到表头（需同时包含：入库日期类、供应商类、价税合计类列）。请确认未用只读/保护视图另存导致首行缺失。"
    );
  }
  return best;
}

async function extractRowsExcelJs(ab: ArrayBuffer): Promise<InvoiceRow[]> {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ab);
  let best: InvoiceRow[] = [];
  let headerFoundAnywhere = false;
  const MAX_ROWS_PER_SHEET = 250_000;

  for (const sheet of wb.worksheets) {
    const aoa: unknown[][] = [];
    sheet.eachRow({ includeEmpty: true }, (row) => {
      if (aoa.length >= MAX_ROWS_PER_SHEET) return false;
      const vals = row.values as unknown[];
      const slice = Array.isArray(vals)
        ? vals.slice(1).map((x) => (x === undefined ? "" : excelJsCellValue(x)))
        : [];
      aoa.push(slice);
    });

    const hasHeaderRow = aoa.some((r) => Array.isArray(r) && rowHasProcurementHeaders(r));
    if (hasHeaderRow) headerFoundAnywhere = true;

    let parsed: InvoiceRow[] | null;
    try {
      parsed = extractRowsFromAoa(aoa as (string | number | null | undefined)[][]);
    } catch (e) {
      throw e instanceof Error
        ? new Error(`工作表「${sheet.name}」：${e.message}`)
        : new Error(`工作表「${sheet.name}」解析失败`);
    }
    if (parsed && parsed.length > best.length) best = parsed;
  }

  if (!headerFoundAnywhere) {
    throw new Error("未在任何工作表中找到表头（需同时包含：入库日期类、供应商类、价税合计类列）。");
  }
  return best;
}

async function parseExcelArrayBuffer(ab: ArrayBuffer): Promise<InvoiceRow[]> {
  const sheetRowsCap = 250_000;
  const binaryFallbackMaxBytes = 12 * 1024 * 1024;
  try {
    const rows = extractRows(
      XLSX.read(ab, { type: "array", raw: true, cellDates: true, sheetRows: sheetRowsCap })
    );
    if (rows.length > 0) return rows;
    console.warn("SheetJS(array)仅读取到空明细，继续降级解析");
  } catch (e) {
    console.warn("SheetJS(array)解析失败，尝试 binary+936:", e);
  }

  if (ab.byteLength <= binaryFallbackMaxBytes) {
    try {
      const bytes = new Uint8Array(ab);
      let binary = "";
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
      }
      const rows = extractRows(
        XLSX.read(binary, { type: "binary", codepage: 936, raw: false, cellDates: true, sheetRows: sheetRowsCap })
      );
      if (rows.length > 0) return rows;
      console.warn("SheetJS(binary+936)仅读取到空明细，继续尝试 ExcelJS");
    } catch (e) {
      console.warn("SheetJS(binary+936)解析失败，尝试 ExcelJS:", e);
    }
  } else {
    console.warn(
      `文件较大（${Math.round(ab.byteLength / 1024 / 1024)}MB），跳过 SheetJS(binary+936) 以避免内存峰值，改用 ExcelJS`
    );
  }

  try {
    return await extractRowsExcelJs(ab);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error(`解析失败（已尝试 SheetJS 与 ExcelJS）：${msg}`);
  }
}

self.onmessage = async (event: MessageEvent<{ id: number; buffer: ArrayBuffer }>) => {
  const { id, buffer } = event.data;
  try {
    const rows = await parseExcelArrayBuffer(buffer);
    postMessage({ id, rows });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    postMessage({ id, error: message });
  }
};
