import { useMemo, useState } from "react";
import { useAccessSession } from "./accessSession";

type InvoiceRow = {
  sellerTaxId: string;
  sellerName: string;
  supplierRaw: string;
  idLast4: string;
  /** 开票时间，格式 YYYY-MM-DD HH:mm（用于展示与排序） */
  invoiceDate: string;
  plateNo: string;
  totalWithTax: number;
};

type CompanySummary = {
  id: string;
  sellerTaxId: string;
  sellerName: string;
  supplierRaw: string;
  idLast4: string;
  salesperson: string;
  monthlyTotals: Record<string, number>;
  last12MonthsTotal: number;
  isWarning: boolean;
};
type SalesMappingData = {
  salespersonByPayer: Record<string, string>;
  payerNameByKey: Record<string, string>;
  idNoByPayerKey: Record<string, string>;
};
type XlsxModule = typeof import("xlsx");

const WARNING_LINE = 4_500_000;
/** 需上传的采购入库单 Excel 个数 */
const PURCHASE_INBOUND_FILE_COUNT = 3;
const SALES_MAPPING_FILE_COUNT = 1;
const TOTAL_REQUIRED_FILE_COUNT = PURCHASE_INBOUND_FILE_COUNT + SALES_MAPPING_FILE_COUNT;
const PURCHASE_FILE_NAME_PATTERN = /^采购入库单.*\.(xlsx|xls)$/i;
const SALES_HEADER_ALIASES = {
  payerName: ["收款人", "收款人名", "收款人姓名", "姓名", "联系人"],
  salesperson: ["分管业务员", "业务员", "销售员", "负责人", "跟单员"],
  idNo: ["身份证号", "身份证号码", "证件号", "身份证"],
} as const;
let xlsxLoader: Promise<XlsxModule> | null = null;

async function loadXlsx(): Promise<XlsxModule> {
  if (!xlsxLoader) xlsxLoader = import("xlsx");
  return xlsxLoader;
}

async function parseExcelInWorker(file: File): Promise<InvoiceRow[]> {
  const buffer = await file.arrayBuffer();
  return new Promise<InvoiceRow[]>((resolve, reject) => {
    const worker = new Worker(new URL("./excelParser.worker.ts", import.meta.url), { type: "module" });
    const id = Math.floor(Math.random() * 1_000_000_000);
    const cleanup = () => {
      worker.onmessage = null;
      worker.onerror = null;
      worker.terminate();
    };
    worker.onmessage = (event: MessageEvent<{ id: number; rows?: InvoiceRow[]; error?: string }>) => {
      if (event.data.id !== id) return;
      cleanup();
      if (event.data.error) reject(new Error(event.data.error));
      else resolve(event.data.rows ?? []);
    };
    worker.onerror = (event) => {
      cleanup();
      reject(new Error(event.message || "解析 Worker 异常退出"));
    };
    worker.postMessage({ id, buffer }, [buffer]);
  });
}

function parseInvoiceDateTimeLocal(s: string): Date | null {
  const m = s.trim().match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const h = Number(m[4]);
  const min = Number(m[5]);
  const dt = new Date(y, mo - 1, d, h, min, 0, 0);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  return dt;
}

function startOfLocalDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** target 所在日历日与 now 所在日相差几天（target 更早则为负）。 */
function calendarDayDiffFromNow(now: Date, target: Date): number {
  const a = startOfLocalDay(now).getTime();
  const b = startOfLocalDay(target).getTime();
  return Math.round((b - a) / 86400000);
}

/** 当天时段（不含钟点），如「下午」「午夜」。 */
function formatDayPeriodZh(hour: number): string {
  if (hour >= 18) return "晚上";
  if (hour >= 13) return "下午";
  if (hour === 12) return "中午";
  if (hour >= 6) return "上午";
  if (hour >= 1) return "凌晨";
  return "午夜";
}

/** 统计截止时间的口头说明（不含几点）；超过24小时则用「N天前」。 */
function formatCutoffHumanZh(datetimeStr: string, now = new Date()): string | null {
  const target = parseInvoiceDateTimeLocal(datetimeStr);
  if (!target) return null;

  const MS_PER_DAY = 86_400_000;
  const elapsedMs = now.getTime() - target.getTime();

  if (elapsedMs > MS_PER_DAY) {
    const days = Math.floor(elapsedMs / MS_PER_DAY);
    return `${days}天前`;
  }

  const diff = calendarDayDiffFromNow(now, target);
  const periodStr = formatDayPeriodZh(target.getHours());

  const relativeMap: Record<number, string> = {
    0: "今天",
    [-1]: "昨天",
    [-2]: "前天",
    1: "明天",
    2: "后天",
  };

  const rel = relativeMap[diff];
  if (rel) return `${rel}${periodStr}`;

  const y = target.getFullYear();
  const mo = target.getMonth() + 1;
  const day = target.getDate();
  const datePart = y === now.getFullYear() ? `${mo}月${day}日` : `${y}年${mo}月${day}日`;
  return `${datePart}${periodStr}`;
}

function monthKey(dateStr: string): string {
  const m = dateStr.match(/^(\d{4})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}` : "未知月份";
}

function formatMonthDay(dateStr: string): string {
  const m = dateStr.match(/^\d{4}-(\d{2})-(\d{2})/);
  if (!m) return "-";
  return `${Number(m[1])}/${Number(m[2])}`;
}

function recent12Months(now = new Date()): string[] {
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const keys: string[] = [];
  for (let i = 11; i >= 0; i -= 1) {
    const total = y * 12 + (m - 1) - i;
    const yy = Math.floor(total / 12);
    const mm = (total % 12) + 1;
    keys.push(`${yy}-${String(mm).padStart(2, "0")}`);
  }
  return keys;
}

function monthHeaderParts(monthKeyValue: string): { year: string; month: string } {
  const matched = monthKeyValue.match(/^(\d{4})-(\d{2})$/);
  if (!matched) return { year: "", month: monthKeyValue };
  const yearNum = Number(matched[1]);
  const monthNum = Number(matched[2]);
  return { year: `${yearNum}年`, month: `${monthNum}月` };
}

function formatWan(v: number): string {
  const wan = v / 10000;
  if (v > WARNING_LINE) {
    return wan.toLocaleString("zh-CN", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  }
  return Math.round(wan).toLocaleString("zh-CN", { maximumFractionDigits: 0 });
}

/** 十二个月格子：整数万（不显示「万」字）；为 0 时不显示数字。 */
function MonthAmountCell({ yuan }: { yuan: number }) {
  const wan = Math.round(yuan / 10000);
  if (wan === 0) {
    return <span className="inline-flex min-h-[2.25rem] items-center justify-center" />;
  }
  return (
    <span className="inline-flex min-h-[2.25rem] items-center justify-center tabular-nums">
      {wan.toLocaleString("zh-CN", { maximumFractionDigits: 0 })}
    </span>
  );
}

function isLikelyCompanyName(name: string): boolean {
  return shouldExcludeCorporateSupplier(name);
}

function normalizePartyName(input: string): string {
  // 先做全角/兼容字符归一化，再去掉空白和常见分隔符，避免“公 司/公、司/公．司”绕过。
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

function parseSalespersonMappingAoa(aoa: (string | number | null | undefined)[][]): SalesMappingData {
  const headerRowIndex = aoa.findIndex((row) => {
    const header = row.map((cell) => normalizeHeader(cell));
    return (
      pickIndexByAliases(header, SALES_HEADER_ALIASES.payerName) >= 0 &&
      pickIndexByAliases(header, SALES_HEADER_ALIASES.salesperson) >= 0
    );
  });
  if (headerRowIndex < 0) throw new Error("未找到业务员对应表表头（需包含：收款人、分管业务员）");

  const header = aoa[headerRowIndex].map((cell) => normalizeHeader(cell));
  const payerIdx = pickIndexByAliases(header, SALES_HEADER_ALIASES.payerName);
  const salesIdx = pickIndexByAliases(header, SALES_HEADER_ALIASES.salesperson);
  const idNoIdx = pickIndexByAliases(header, SALES_HEADER_ALIASES.idNo);
  if (payerIdx < 0 || salesIdx < 0) {
    throw new Error("业务员对应表表头解析失败（需包含：收款人、分管业务员）");
  }

  const mapping: Record<string, string> = {};
  const payerNameByKey: Record<string, string> = {};
  const idNoByPayerKey: Record<string, string> = {};
  for (let i = headerRowIndex + 1; i < aoa.length; i += 1) {
    const row = aoa[i];
    if (!Array.isArray(row)) continue;
    const payer = String(row[payerIdx] ?? "").normalize("NFKC").trim();
    const salesperson = String(row[salesIdx] ?? "").normalize("NFKC").trim();
    if (!payer || !salesperson) continue;
    const normalizedPayer = normalizePartyName(payer);
    if (!normalizedPayer) continue;
    mapping[normalizedPayer] = salesperson;
    if (!payerNameByKey[normalizedPayer]) payerNameByKey[normalizedPayer] = payer;
    if (idNoIdx >= 0) {
      const idNo = String(row[idNoIdx] ?? "").normalize("NFKC").trim();
      if (idNo) idNoByPayerKey[normalizedPayer] = idNo;
    }
  }
  if (!Object.keys(mapping).length) {
    throw new Error("业务员对应表没有可用数据（收款人/分管业务员不能为空）");
  }
  return { salespersonByPayer: mapping, payerNameByKey, idNoByPayerKey };
}

async function parseSalespersonMapping(file: File): Promise<SalesMappingData> {
  const ab = await file.arrayBuffer();
  const XLSX = await loadXlsx();
  const wb = XLSX.read(ab, { type: "array", raw: false, cellDates: false, sheetRows: 250_000 });
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) continue;
    const aoa = XLSX.utils.sheet_to_json<(string | number | null | undefined)[]>(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });
    if (!aoa.length) continue;
    try {
      return parseSalespersonMappingAoa(aoa);
    } catch {
      // try next sheet
    }
  }
  throw new Error(`文件「${file.name}」中未找到业务员对应关系（需包含：收款人、分管业务员）`);
}

function shouldExcludeCorporateSupplier(name: string): boolean {
  const raw = String(name ?? "").normalize("NFKC");
  const normalized = normalizePartyName(raw);
  if (!normalized) return false;
  if (normalized.includes("公司")) return true;
  if (normalized.includes("回收站")) return true;
  if (/回.{0,3}收.{0,3}站/u.test(raw)) return true;
  // 兜底：允许“公”和“司”之间夹杂少量异常字符
  return /公.{0,3}司/u.test(raw);
}

export default function App() {
  const accessSession = useAccessSession();
  const isAdmin = accessSession?.role === "admin";
  const [rows, setRows] = useState<InvoiceRow[]>([]);
  const [salespersonByPayer, setSalespersonByPayer] = useState<Record<string, string>>({});
  const [payerNameByKey, setPayerNameByKey] = useState<Record<string, string>>({});
  const [idNoByPayerKey, setIdNoByPayerKey] = useState<Record<string, string>>({});
  const [draggingPayerKey, setDraggingPayerKey] = useState<string | null>(null);
  const [mappingFileName, setMappingFileName] = useState("废钢业务联系讯息.xlsx");
  const [error, setError] = useState("");
  const [fileNames, setFileNames] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [parsing, setParsing] = useState(false);

  const rolling12 = useMemo(() => recent12Months(), []);
  const displayMonths = useMemo(() => [...rolling12].reverse(), [rolling12]);

  const latestTransactionTime = useMemo(() => {
    let max = "";
    for (const row of rows) {
      const t = row.invoiceDate.trim();
      if (!t) continue;
      if (t > max) max = t;
    }
    return max || null;
  }, [rows]);

  const cutoffHumanZh = useMemo(
    () => (latestTransactionTime ? formatCutoffHumanZh(latestTransactionTime) : null),
    [latestTransactionTime]
  );

  const summaries = useMemo<CompanySummary[]>(() => {
    const map = new Map<string, CompanySummary>();

    for (const row of rows) {
      if (isLikelyCompanyName(row.sellerName) || shouldExcludeCorporateSupplier(row.supplierRaw)) continue;

      const id = row.sellerName;
      const month = monthKey(row.invoiceDate);

      if (!map.has(id)) {
        map.set(id, {
          id,
          sellerTaxId: row.sellerTaxId,
          sellerName: row.sellerName,
          supplierRaw: row.supplierRaw,
          idLast4: row.idLast4,
          salesperson: salespersonByPayer[normalizePartyName(row.sellerName)] ?? "-",
          monthlyTotals: {},
          last12MonthsTotal: 0,
          isWarning: false,
        });
      }

      const item = map.get(id)!;
      if (!item.idLast4 && row.idLast4) item.idLast4 = row.idLast4;
      if (item.salesperson === "-") {
        item.salesperson = salespersonByPayer[normalizePartyName(row.sellerName)] ?? "-";
      }
      item.monthlyTotals[month] = (item.monthlyTotals[month] ?? 0) + row.totalWithTax;
    }

    for (const item of map.values()) {
      item.last12MonthsTotal = rolling12.reduce((sum, month) => sum + (item.monthlyTotals[month] ?? 0), 0);
      item.isWarning = item.last12MonthsTotal > WARNING_LINE;
    }

    return [...map.values()]
      .filter((item) => !shouldExcludeCorporateSupplier(item.sellerName))
      .sort((a, b) => b.last12MonthsTotal - a.last12MonthsTotal);
  }, [rows, rolling12, salespersonByPayer]);

  const filteredCorporateCount = useMemo(() => {
    if (!rows.length) return 0;
    return rows.filter((row) => shouldExcludeCorporateSupplier(row.sellerName)).length;
  }, [rows]);
  const personalTransactionCount = rows.length - filteredCorporateCount;
  const visibleSummaries = useMemo(() => {
    if (isAdmin || !accessSession?.salesperson) return summaries;
    return summaries.filter((x) => x.salesperson === accessSession.salesperson);
  }, [summaries, isAdmin, accessSession?.salesperson]);
  const unmatchedPayers = useMemo(() => {
    const latestByKey = new Map<string, { key: string; payerName: string; invoiceDate: string; plateNo: string }>();
    for (const row of rows) {
      if (shouldExcludeCorporateSupplier(row.sellerName) || shouldExcludeCorporateSupplier(row.supplierRaw)) continue;
      const key = normalizePartyName(row.sellerName);
      if (!key) continue;
      if ((salespersonByPayer[key] ?? "").trim()) continue;
      const prev = latestByKey.get(key);
      if (!prev || row.invoiceDate > prev.invoiceDate) {
        latestByKey.set(key, {
          key,
          payerName: row.sellerName,
          invoiceDate: row.invoiceDate,
          plateNo: row.plateNo || "-",
        });
      }
    }
    return [...latestByKey.values()].sort((a, b) => a.payerName.localeCompare(b.payerName, "zh-CN"));
  }, [rows, salespersonByPayer]);
  const salespersonOptions = useMemo(() => {
    const set = new Set<string>();
    for (const name of Object.values(salespersonByPayer)) {
      const n = String(name ?? "").trim();
      if (n && n !== "-") set.add(n);
    }
    return [...set].sort((a, b) => a.localeCompare(b, "zh-CN"));
  }, [salespersonByPayer]);

  function assignSalespersonByDrop(payerKey: string | null, salesperson: string) {
    if (!payerKey) return;
    const target = unmatchedPayers.find((x) => x.key === payerKey);
    if (!target) {
      setDraggingPayerKey(null);
      return;
    }
    setError("");
    setSalespersonByPayer((prev) => ({ ...prev, [payerKey]: salesperson }));
    setPayerNameByKey((prev) => ({ ...prev, [payerKey]: prev[payerKey] || target.payerName }));
    setDraggingPayerKey(null);
  }

  async function exportUpdatedSalesMapping() {
    const keys = Array.from(new Set([...Object.keys(payerNameByKey), ...Object.keys(salespersonByPayer)]));
    const rowsToExport = keys
      .filter((k) => (salespersonByPayer[k] ?? "").trim())
      .map((k) => ({
        收款人: payerNameByKey[k] || k,
        身份证号: idNoByPayerKey[k] || "",
        分管业务员: salespersonByPayer[k],
      }));
    if (!rowsToExport.length) {
      setError("当前没有可导出的业务员对应关系");
      return;
    }
    const XLSX = await loadXlsx();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rowsToExport);
    XLSX.utils.book_append_sheet(wb, ws, "业务员对应");
    const base = mappingFileName.replace(/\.(xlsx|xls)$/i, "");
    XLSX.writeFile(wb, `${base}_补充后.xlsx`);
    setError("");
  }

  async function handleExcelFiles(files: File[]) {
    const excelFiles = files.filter((f) => /\.(xlsx|xls)$/i.test(f.name));
    if (excelFiles.length !== TOTAL_REQUIRED_FILE_COUNT) {
      setRows([]);
      setSalespersonByPayer({});
      setPayerNameByKey({});
      setIdNoByPayerKey({});
      setDraggingPayerKey(null);
      setFileNames([]);
      setError(
        `请上传恰好 ${TOTAL_REQUIRED_FILE_COUNT} 个 Excel（${PURCHASE_INBOUND_FILE_COUNT} 个采购入库单 + ${SALES_MAPPING_FILE_COUNT} 个业务员对应表），当前选中 ${excelFiles.length} 个`
      );
      return;
    }

    setError("");
    setFileNames(excelFiles.map((f) => f.name));
    setParsing(true);

    const merged: InvoiceRow[] = [];
    const invoiceByName = excelFiles.filter((f) => PURCHASE_FILE_NAME_PATTERN.test(f.name));
    const otherFiles = excelFiles.filter((f) => !PURCHASE_FILE_NAME_PATTERN.test(f.name));
    let invoiceFiles: File[] = [];
    let mappingFile: File | null = null;
    let mapping: SalesMappingData = { salespersonByPayer: {}, payerNameByKey: {}, idNoByPayerKey: {} };
    try {
      if (invoiceByName.length === PURCHASE_INBOUND_FILE_COUNT && otherFiles.length === SALES_MAPPING_FILE_COUNT) {
        invoiceFiles = invoiceByName;
        mappingFile = otherFiles[0];
        mapping = await parseSalespersonMapping(mappingFile);
      } else {
        for (let i = 0; i < excelFiles.length; i++) {
          const candidate = excelFiles[i];
          if (mappingFile) {
            invoiceFiles.push(candidate);
            continue;
          }
          try {
            mapping = await parseSalespersonMapping(candidate);
            mappingFile = candidate;
          } catch {
            invoiceFiles.push(candidate);
          }
        }
      }

      if (!mappingFile) {
        setRows([]);
        setSalespersonByPayer({});
        setPayerNameByKey({});
        setIdNoByPayerKey({});
        setDraggingPayerKey(null);
        setError("未识别到业务员对应表，请确认第4个文件包含「收款人」「分管业务员」列");
        return;
      }
      if (invoiceFiles.length !== PURCHASE_INBOUND_FILE_COUNT) {
        setRows([]);
        setSalespersonByPayer({});
        setPayerNameByKey({});
        setIdNoByPayerKey({});
        setDraggingPayerKey(null);
        setError(
          `已识别业务员对应表「${mappingFile.name}」，但采购入库单应为 ${PURCHASE_INBOUND_FILE_COUNT} 个，当前识别到 ${invoiceFiles.length} 个`
        );
        return;
      }

      try {
        const parsedChunks = await Promise.all(
          invoiceFiles.map(async (invoiceFile, idx) => {
            try {
              return await parseExcelInWorker(invoiceFile);
            } catch (e) {
              const msg = e instanceof Error ? e.message : "解析失败";
              throw new Error(`采购入库单第 ${idx + 1} 个文件「${invoiceFile.name}」：${msg}`);
            }
          })
        );
        for (const rowsChunk of parsedChunks) merged.push(...rowsChunk);
      } catch (e) {
        setRows([]);
        setSalespersonByPayer({});
        setPayerNameByKey({});
        setIdNoByPayerKey({});
        setDraggingPayerKey(null);
        setFileNames([]);
        const msg = e instanceof Error ? e.message : "解析失败";
        setError(msg);
        return;
      }

      setRows(merged);
      setSalespersonByPayer(mapping.salespersonByPayer);
      setPayerNameByKey(mapping.payerNameByKey);
      setIdNoByPayerKey(mapping.idNoByPayerKey);
      setDraggingPayerKey(null);
      setMappingFileName(mappingFile.name);
      if (!merged.length) {
        setError(
          "三个采购入库单均已解析，但未读取到有效数据；请确认每张表都包含表头：入库日期、供应商、价税合计"
        );
      }
    } finally {
      setParsing(false);
    }
  }

  const warningCount = visibleSummaries.filter((x) => x.isWarning).length;

  return (
    <div className="min-h-screen bg-gray-900 text-slate-900">
      <div className="mx-auto max-w-7xl p-4 md:p-8">
        <div>
          <div>
            <h1 className="text-2xl font-bold">税务开票预警助手</h1>
            {latestTransactionTime ? (
              <div className="mt-1 text-sm text-slate-600">
                <p>
                  <span className="text-slate-500">统计截止时间</span>
                  <span className="ml-2 font-medium tabular-nums text-slate-900">{latestTransactionTime}</span>
                </p>
                {cutoffHumanZh ? (
                  <p className="mt-1 text-slate-700">{cutoffHumanZh}</p>
                ) : null}
              </div>
            ) : null}
            <p className="mt-1 text-sm text-slate-600">按实时滚动12个月统计累计并进行450万预警。</p>
          </div>
        </div>

        {isAdmin ? (
          <div
            className={`mt-4 rounded-xl border-2 border-dashed p-8 text-center transition-colors ${
              dragging
                ? "border-sky-800 bg-sky-700 text-white"
                : "border-sky-500 bg-sky-300 text-sky-950"
            }`}
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragging(false);
              const list = e.dataTransfer.files?.length ? Array.from(e.dataTransfer.files) : [];
              if (list.length) void handleExcelFiles(list);
            }}
          >
          <p className="font-medium">拖拽 {TOTAL_REQUIRED_FILE_COUNT} 个 Excel 到这里</p>
          <p className={`mt-1 text-sm ${dragging ? "text-sky-200" : "text-sky-800"}`}>
            自动识别 {PURCHASE_INBOUND_FILE_COUNT} 个采购入库单（文件名形如「采购入库单_*.xlsx」）与 1 个业务员对应表（含收款人、分管业务员）
          </p>
          <label
            className={`mt-4 inline-block cursor-pointer rounded border-2 px-4 py-2 transition-colors ${
              dragging
                ? "border-sky-300 bg-sky-800 text-white hover:bg-sky-900 hover:border-sky-200"
                : "border-sky-700 bg-sky-300 text-sky-950 hover:bg-sky-700 hover:text-white hover:border-sky-800"
            }`}
          >
            选择文件（{TOTAL_REQUIRED_FILE_COUNT} 个）
            <input
              type="file"
              accept=".xlsx,.xls"
              multiple
              className="hidden"
              onChange={(e) => {
                const list = e.target.files?.length ? Array.from(e.target.files) : [];
                e.target.value = "";
                if (list.length) void handleExcelFiles(list);
              }}
            />
          </label>
          {fileNames.length === TOTAL_REQUIRED_FILE_COUNT ? (
            <ul
              className={`mx-auto mt-3 max-w-lg list-inside list-decimal text-left text-sm ${dragging ? "text-sky-100" : "text-sky-950"}`}
            >
              {fileNames.map((name, idx) => (
                <li key={`${idx}-${name}`} className="tabular-nums">
                  {name}
                </li>
              ))}
            </ul>
          ) : null}
          {error && (
            <p className={`mt-2 text-sm ${dragging ? "text-red-300" : "text-red-600"}`}>{error}</p>
          )}
          {parsing && (
            <p className={`mt-2 text-sm ${dragging ? "text-sky-100" : "text-sky-900"}`}>
              正在解析中，请稍候（大文件可能需要 10-60 秒）
            </p>
          )}
          </div>
        ) : (
          <div className="mt-4 rounded-xl border border-slate-200 bg-white p-4 text-sm text-slate-600 shadow-sm">
            业务员账号为只读模式，仅查看你本人负责客户的预警结果。
          </div>
        )}

        <div className="mt-4 grid gap-3 md:grid-cols-3">
          <div className="rounded-lg bg-white p-3 shadow-sm">个人交易次数：{personalTransactionCount}</div>
          <div className="rounded-lg bg-white p-3 shadow-sm">人数：{visibleSummaries.length}</div>
          <div className="rounded-lg bg-white p-3 shadow-sm">预警人数：{warningCount}</div>
        </div>
        {filteredCorporateCount > 0 ? (
          <p className="mt-2 text-sm text-slate-500">
            已在展示前自动排除疑似企业名 {filteredCorporateCount} 条（如修改数据后请刷新并重新解析）。
          </p>
        ) : null}
        {isAdmin && (fileNames.length > 0 || Object.keys(salespersonByPayer).length > 0) ? (
          <div className="mt-3 rounded-xl border-2 border-sky-200 bg-white p-4 shadow-sm">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <p className="text-sm font-semibold text-sky-800">业务员拖拽补录区</p>
              <p className="text-sm text-slate-700">
                未匹配业务员：<span className="font-semibold">{unmatchedPayers.length}</span> 人
              </p>
              <button
                type="button"
                className="rounded border border-sky-700 px-3 py-1.5 text-sm text-sky-800 hover:bg-sky-50"
                onClick={() => void exportUpdatedSalesMapping()}
              >
                导出补充后的业务员对应表
              </button>
            </div>
            {unmatchedPayers.length > 0 ? (
              <div className="mt-3 grid grid-cols-[1.6fr_1fr] gap-4 items-start">
                <div className="min-w-0">
                  <p className="mb-2 text-xs font-medium text-slate-700">左侧：未分配收款人（小方框）</p>
                  <div className="max-h-[26rem] space-y-2 overflow-y-auto pr-1">
                    {unmatchedPayers.map((item) => (
                      <button
                        key={item.key}
                        type="button"
                        draggable
                        className={`flex w-full items-center justify-between rounded border px-3 py-2 text-xs ${
                          draggingPayerKey === item.key
                            ? "border-sky-500 bg-sky-100 text-sky-900"
                            : "border-slate-300 bg-slate-50 text-slate-700"
                        }`}
                        onDragStart={(e) => {
                          e.dataTransfer.setData("text/plain", item.key);
                          e.dataTransfer.effectAllowed = "move";
                          setDraggingPayerKey(item.key);
                        }}
                        onDragEnd={() => setDraggingPayerKey(null)}
                      >
                        <span className="mr-2 truncate text-left">{item.payerName}</span>
                        <span className="shrink-0 text-[11px] text-slate-500">
                          {formatMonthDay(item.invoiceDate)} | {item.plateNo || "-"}
                        </span>
                      </button>
                    ))}
                  </div>
                </div>
                <div className="min-w-0">
                  <p className="mb-2 text-xs font-medium text-slate-700">右侧：业务员（大方框投放区）</p>
                  {salespersonOptions.length > 0 ? (
                    <div className="max-h-[26rem] space-y-2 overflow-y-auto pr-1">
                      {salespersonOptions.map((salesperson) => (
                        <div
                          key={salesperson}
                          className="min-h-16 rounded-lg border-2 border-dashed border-slate-300 bg-slate-50 p-3"
                          onDragOver={(e) => {
                            e.preventDefault();
                            e.dataTransfer.dropEffect = "move";
                          }}
                          onDrop={(e) => {
                            e.preventDefault();
                            const droppedKey = e.dataTransfer.getData("text/plain") || draggingPayerKey;
                            assignSalespersonByDrop(droppedKey || null, salesperson);
                          }}
                        >
                          <div className="text-sm font-medium text-slate-800">{salesperson}</div>
                          <div className="mt-1 text-xs text-slate-500">拖到此业务员框完成分配</div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="text-sm text-amber-700">当前对应表未识别到可用业务员名单，无法进行拖拽分配。</p>
                  )}
                </div>
              </div>
            ) : (
              <p className="mt-2 text-sm text-emerald-700">所有收款人已匹配业务员。</p>
            )}
          </div>
        ) : null}

        <div className="mt-6 overflow-x-auto rounded-xl bg-white p-4 shadow-sm">
          <h2 className="mb-3 font-semibold">个人开票金额汇总（滚动12个月 + 预警）</h2>
          <table className="min-w-full table-fixed border-separate border-spacing-0 text-sm">
            <colgroup>
              <col style={{ width: "4rem" }} />
              <col style={{ width: "6rem" }} />
              <col style={{ width: "8em" }} />
              <col style={{ width: "7rem" }} />
              <col style={{ width: "7rem" }} />
              {displayMonths.map((month) => (
                <col key={`month-col-${month}`} style={{ width: "6rem" }} />
              ))}
            </colgroup>
            <thead>
              <tr className="bg-slate-50 text-center">
                <th className="w-[4rem] min-w-[4rem] px-2 py-2 align-middle">状态</th>
                <th className="w-24 px-2 py-2 align-middle whitespace-nowrap">近12个月累计</th>
                <th className="w-[8em] max-w-[8em] min-w-[8em] px-2 py-2 align-middle whitespace-normal break-all">
                  收款人名
                </th>
                <th className="px-2 py-2 align-middle whitespace-nowrap">业务员</th>
                <th className="px-2 py-2 align-middle whitespace-nowrap">
                  <div className="mx-auto flex w-full flex-col items-center leading-tight">
                    <div className="flex h-[1.35rem] w-full shrink-0 items-end justify-center">身份证</div>
                    <div className="flex h-[1.35rem] w-full shrink-0 items-start justify-center">后四位</div>
                  </div>
                </th>
                {displayMonths.map((month, monthIndex) => {
                  const parts = monthHeaderParts(month);
                  return (
                    <th
                      key={month}
                      className={`w-24 border-t border-b border-r border-solid border-neutral-700 px-2 py-2 whitespace-nowrap align-middle bg-slate-50 ${
                        monthIndex === 0 ? "border-l" : ""
                      }`}
                    >
                      <div className="mx-auto flex w-full flex-col items-center leading-tight">
                        <div className="flex h-[1.35rem] w-full shrink-0 items-end justify-center">{parts.year}</div>
                        <div className="flex h-[1.35rem] w-full shrink-0 items-start justify-center">{parts.month}</div>
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {visibleSummaries.map((s) => (
                <tr key={s.id}>
                  <td className="w-[4rem] min-w-[4rem] px-2 py-2 align-middle whitespace-nowrap">
                    {s.isWarning ? (
                      <span className="rounded bg-red-100 px-2 py-1 text-red-700">⚠️ 已超450万预警线</span>
                    ) : (
                      <span className="rounded bg-emerald-100 px-2 py-1 text-emerald-700">正常</span>
                    )}
                  </td>
                  <td className="w-24 px-2 py-2 text-center whitespace-nowrap font-medium tabular-nums">
                    {formatWan(s.last12MonthsTotal)}万
                  </td>
                  <td className="w-[8em] max-w-[8em] min-w-[8em] px-2 py-2 align-middle whitespace-normal break-all text-left">
                    {s.sellerName || "-"}
                  </td>
                  <td className="px-2 py-2 whitespace-nowrap">{s.salesperson || "-"}</td>
                  <td className="px-2 py-2 whitespace-nowrap">{s.idLast4 || "-"}</td>
                  {displayMonths.map((month, monthIndex) => (
                    <td
                      key={`${s.id}-${month}`}
                      className={`w-24 border-t border-b border-r border-solid border-neutral-700 px-2 py-2 text-center whitespace-nowrap align-middle bg-white ${
                        monthIndex === 0 ? "border-l" : ""
                      }`}
                    >
                      <MonthAmountCell yuan={s.monthlyTotals[month] ?? 0} />
                    </td>
                  ))}
                </tr>
              ))}
              {!visibleSummaries.length && (
                <tr>
                  <td colSpan={5 + displayMonths.length} className="px-2 py-6 text-center text-slate-400">
                    暂无数据
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
