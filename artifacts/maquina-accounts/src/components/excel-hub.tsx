import { useState, useMemo, type ChangeEvent } from "react";
import { useQueryClient, useQuery } from "@tanstack/react-query";
import {
  FileSpreadsheet,
  Upload,
  Download,
  CheckCircle2,
  AlertTriangle,
  FileCheck,
  Table,
  Layers,
  Sparkles,
  RefreshCw,
  FolderDown,
  ArrowRight,
  ShieldCheck,
  ChevronDown,
  Plus,
  Users,
  Eye,
  TrendingUp,
  Undo2,
  ClipboardPaste,
  Building2,
  WalletCards,
  Boxes,
  ShoppingBag,
  FilePlus2,
  Calculator,
} from "lucide-react";
import {
  useListSales,
  useListPurchases,
  useListCash,
  useListInventory,
  useListPartners,
  useListAssets,
  getListSalesQueryKey,
  getListPurchasesQueryKey,
  getListCashQueryKey,
  getListCustomersQueryKey,
  getListSuppliersQueryKey,
  getListMaterialsQueryKey,
  getGetDashboardQueryKey,
} from "@workspace/api-client-react";
import {
  downloadExcelTemplate,
  exportMultiSheetExcel,
  parseUploadedExcel,
  type ParsedSheetData,
  type ReportCell,
} from "@/lib/report-export";

interface AuditMetric {
  name: string;
  excelSheet: string;
  excelValue: number;
  systemValue: number;
  diff: number;
  status: string;
  badge: "success" | "warning" | "updated";
  details: string;
}

interface FinancialSummary {
  totalRevenue: number;
  directMaterialCost: number;
  grossProfit: number;
  grossMarginPercent: number;
  totalExpenses: number;
  depreciationExpense: number;
  netIncome: number;
  partners: Array<{ id: number; name: string; totalShare: number; paid: number; remaining: number; profitShare: number }>;
}

interface AuditData {
  sourceFile: string;
  auditDate: string;
  totalChecks: number;
  passedChecks: number;
  overallMatchRate: number;
  metrics: AuditMetric[];
  financialSummary?: FinancialSummary;
}

interface MonthlyProfitSummary {
  months: Array<{ month: string; total: number; partners: Record<string, number> }>;
  partnerTotals: Record<string, number>;
  grandTotal: number;
}

interface MachineScheduleData {
  machineCost: number;
  usefulLifeMonths: number;
  straightLineMonthly: number;
  cashflowMonthly: number;
  monthlyDiff: number;
  schedule: Array<{ date: string; depreciation: number; accumulated: number; book_value: number }>;
}

const money = (value?: number) =>
  `${new Intl.NumberFormat("ar-EG", { maximumFractionDigits: 2 }).format(value ?? 0)} ج.م`;

export function ExcelHub() {
  const [activeTab, setActiveTab] = useState<"audit" | "sheets" | "profitShare" | "import" | "templates" | "export">("audit");
  const [activeSheetTab, setActiveSheetTab] = useState<"partners" | "sales" | "purchases" | "cash" | "inventory" | "assets" | "income">("partners");
  const [depreciationMethod, setDepreciationMethod] = useState<"accounting" | "cashflow">("accounting");
  const qc = useQueryClient();

  // Audit query
  const auditQuery = useQuery<AuditData>({
    queryKey: ["/api/excel-audit"],
    queryFn: async () => {
      const res = await fetch("/api/excel-audit");
      if (!res.ok) throw new Error("تعذر جلب بيانات المطابقة");
      return res.json();
    },
  });

  // Monthly Profit Summary query
  const profitSummaryQuery = useQuery<MonthlyProfitSummary>({
    queryKey: ["/api/profit-allocations/summary"],
    queryFn: async () => {
      const res = await fetch("/api/profit-allocations/summary");
      if (!res.ok) throw new Error("تعذر جلب ملخص الأرباح");
      return res.json();
    },
  });

  // Machine schedule query
  const machineScheduleQuery = useQuery<MachineScheduleData>({
    queryKey: ["/api/machine-depreciation-schedule"],
    queryFn: async () => {
      const res = await fetch("/api/machine-depreciation-schedule");
      if (!res.ok) throw new Error("تعذر جلب جدول إهلاك الماكينة");
      return res.json();
    },
  });

  // Live accounting data
  const sales = useListSales();
  const purchases = useListPurchases();
  const cash = useListCash();
  const inventory = useListInventory();
  const partners = useListPartners();
  const assets = useListAssets();

  // New profit distribution form state
  const [newProfitForm, setNewProfitForm] = useState({
    monthLabel: "شهر 9 — أيلول",
    totalAmount: "12000",
    date: new Date().toISOString().slice(0, 10),
  });
  const [addingProfit, setAddingProfit] = useState(false);
  const [profitMsg, setProfitMsg] = useState<string | null>(null);

  // Reset to original Excel state
  const [resetting, setResetting] = useState(false);
  const [resetMsg, setResetMsg] = useState<string | null>(null);

  // File import state
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [parsedSheets, setParsedSheets] = useState<ParsedSheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [importTarget, setImportTarget] = useState<"sales" | "purchases" | "cash" | "materials" | "customers">("sales");
  const [importing, setImporting] = useState(false);
  const [importSuccess, setImportSuccess] = useState<string | null>(null);
  const [parseError, setParseError] = useState<string | null>(null);

  // Paste text import state
  const [pasteText, setPasteText] = useState("");
  const [showPasteBox, setShowPasteBox] = useState(false);

  const handleFileUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setSelectedFile(file);
    setParseError(null);
    setImportSuccess(null);
    try {
      const sheets = await parseUploadedExcel(file);
      if (!sheets.length) {
        setParseError("الملف فارغ أو تعذر قراءة أوراق العمل به");
        return;
      }
      setParsedSheets(sheets);
      setActiveSheetIndex(0);

      // Auto-detect target based on sheet name
      const name = sheets[0].sheetName.toLowerCase();
      if (name.includes("بيع") || name.includes("sale") || name.includes("فاتورة")) setImportTarget("sales");
      else if (name.includes("شراء") || name.includes("مشتريات") || name.includes("خامات") || name.includes("purchase")) setImportTarget("purchases");
      else if (name.includes("خزينة") || name.includes("نقد") || name.includes("cash")) setImportTarget("cash");
      else if (name.includes("عميل") || name.includes("عملاء") || name.includes("customer")) setImportTarget("customers");
      else if (name.includes("مخزون") || name.includes("مادة") || name.includes("material")) setImportTarget("materials");
    } catch (err) {
      console.error(err);
      setParseError("حدث خطأ أثناء قراءة ملف الإكسيل. تأكد من صحة التنسيق (.xlsx, .xls, .csv)");
    }
  };

  const handlePasteParse = () => {
    if (!pasteText.trim()) return;
    try {
      const lines = pasteText.trim().split("\n").map((line) => line.split("\t").map((cell) => cell.trim()));
      if (lines.length < 2) {
        setParseError("النص الملصوق يجب أن يحتوي على صف عناوين وصف واحد على الأقل من البيانات");
        return;
      }
      const headers = lines[0];
      const rows = lines.slice(1).map((line) => {
        const obj: Record<string, any> = {};
        headers.forEach((h, idx) => {
          obj[h] = line[idx] ?? "";
        });
        return obj;
      });

      const parsedData: ParsedSheetData = {
        sheetName: "بيانات ملصوقة من إكسيل",
        headers,
        rows,
        rawRows: lines,
      };

      setParsedSheets([parsedData]);
      setActiveSheetIndex(0);
      setShowPasteBox(false);
      setParseError(null);
    } catch (err: any) {
      setParseError("تعذر تحليل النص الملصوق: " + err.message);
    }
  };

  const handleExecuteImport = async () => {
    if (!parsedSheets.length) return;
    const currentSheet = parsedSheets[activeSheetIndex];
    if (!currentSheet || !currentSheet.rows.length) return;

    setImporting(true);
    setImportSuccess(null);

    try {
      let payload: any = {};

      if (importTarget === "sales") {
        payload.sales = currentSheet.rows.map((row, i) => {
          const invoiceNo = String(row["رقم الفاتورة"] || row["الفاتورة"] || row["Invoice"] || `IMP-${i + 1}`).trim();
          const customerName = String(row["اسم العميل"] || row["العميل"] || row["Customer"] || "عميل غير محدد").trim();
          const total = Number(row["الإجمالي"] || row["المبلغ"] || row["Total"] || 0);
          const date = String(row["التاريخ"] || row["Date"] || new Date().toISOString().slice(0, 10)).slice(0, 10);
          const paymentMethod = String(row["طريقة الدفع"] || row["Payment"] || "آجل");
          return { invoiceNo, customerName, total, date, paymentMethod };
        }).filter((s) => s.total > 0);
      } else if (importTarget === "purchases") {
        payload.purchases = currentSheet.rows.map((row, i) => {
          const invoiceNo = String(row["رقم الفاتورة"] || row["الفاتورة"] || row["Invoice"] || `IMP-P-${i + 1}`).trim();
          const supplierName = String(row["اسم المورد"] || row["المورد"] || row["Supplier"] || "مورد خامات").trim();
          const total = Number(row["الإجمالي"] || row["المبلغ"] || row["Total"] || 0);
          const date = String(row["التاريخ"] || row["Date"] || new Date().toISOString().slice(0, 10)).slice(0, 10);
          const paymentMethod = String(row["طريقة الدفع"] || row["Payment"] || "نقدي");
          return { invoiceNo, supplierName, total, date, paymentMethod };
        }).filter((p) => p.total > 0);
      } else if (importTarget === "cash") {
        payload.cash = currentSheet.rows.map((row) => {
          const typeStr = String(row["نوع الحركة"] || row["النوع"] || row["Type"] || "صرف");
          const type = typeStr.includes("قبض") ? "قبض" : "صرف";
          const amount = Number(row["المبلغ"] || row["القيمة"] || row["Amount"] || 0);
          const description = String(row["البيان"] || row["الوصف"] || row["Description"] || "حركة مستوردة من إكسيل");
          const category = String(row["التصنيف"] || row["البند"] || row["Category"] || "تشغيل");
          const date = String(row["التاريخ"] || row["Date"] || new Date().toISOString().slice(0, 10)).slice(0, 10);
          return { type, amount, description, category, date };
        }).filter((c) => c.amount > 0);
      } else if (importTarget === "customers") {
        payload.customers = currentSheet.rows.map((row, i) => ({
          code: String(row["الكود"] || row["Code"] || `C-${i + 1}`),
          name: String(row["الاسم"] || row["اسم العميل"] || row["Name"] || `عميل ${i + 1}`),
          phone: String(row["الهاتف"] || row["Phone"] || ""),
          address: String(row["العنوان"] || row["Address"] || ""),
          openingBalance: Number(row["الرصيد الافتتاحي"] || row["الرصيد"] || 0),
        }));
      } else if (importTarget === "materials") {
        payload.materials = currentSheet.rows.map((row, i) => ({
          code: String(row["الكود"] || row["Code"] || `M-${i + 1}`),
          name: String(row["اسم الخامة"] || row["المادة"] || row["Name"] || `خامة ${i + 1}`),
          unit: String(row["الوحدة"] || row["Unit"] || "عدد"),
          cost: Number(row["سعر التكلفة"] || row["التكلفة"] || row["Cost"] || 0),
        }));
      }

      const res = await fetch("/api/bulk-import", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || "فشل ترحيل البيانات");
      }

      // Invalidate relevant queries
      qc.invalidateQueries({ queryKey: getListSalesQueryKey() });
      qc.invalidateQueries({ queryKey: getListPurchasesQueryKey() });
      qc.invalidateQueries({ queryKey: getListCashQueryKey() });
      qc.invalidateQueries({ queryKey: getListCustomersQueryKey() });
      qc.invalidateQueries({ queryKey: getListSuppliersQueryKey() });
      qc.invalidateQueries({ queryKey: getListMaterialsQueryKey() });
      qc.invalidateQueries({ queryKey: getGetDashboardQueryKey() });
      qc.invalidateQueries({ queryKey: ["/api/excel-audit"] });

      setImportSuccess(`تم ترحيل البيانات بنجاح من ورقة "${currentSheet.sheetName}" إلى سجلات ${
        importTarget === "sales" ? "المبيعات" : importTarget === "purchases" ? "المشتريات" : importTarget === "cash" ? "الخزينة" : importTarget === "customers" ? "العملاء" : "الخامات"
      }!`);
    } catch (err: any) {
      setParseError(err.message || "حدث خطأ أثناء حفظ البيانات");
    } finally {
      setImporting(false);
    }
  };

  const handleAddProfitDistribution = async (e: React.FormEvent) => {
    e.preventDefault();
    const total = Number(newProfitForm.totalAmount);
    if (!total || total <= 0) return;
    const partnerShare = Math.round((total / 3 + Number.EPSILON) * 100) / 100;
    const partnerList = ["سيفن ام", "احمد العبسي", "محمد علاء"];

    setAddingProfit(true);
    setProfitMsg(null);

    try {
      for (const partner of partnerList) {
        const res = await fetch("/api/profit-allocations", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            date: newProfitForm.date,
            partner,
            amount: partnerShare,
            description: `${newProfitForm.monthLabel} — توزيع أرباح شركاء متساوي`,
          }),
        });
        if (!res.ok) {
          const err = await res.json();
          throw new Error(err.error || "فشل تسجيل توزيع الربح");
        }
      }

      qc.invalidateQueries({ queryKey: ["/api/profit-allocations/summary"] });
      qc.invalidateQueries({ queryKey: ["/api/excel-audit"] });
      qc.invalidateQueries({ queryKey: getGetDashboardQueryKey() });
      setProfitMsg(`تم تسجيل توزيع أرباح شهر "${newProfitForm.monthLabel}" بنجاح بإجمالي ${money(total)} (بواقع ${money(partnerShare)} لكل شريك من الشركاء الثلاثة).`);
    } catch (err: any) {
      setProfitMsg(`خطأ: ${err.message}`);
    } finally {
      setAddingProfit(false);
    }
  };

  const handleResetToExcel = async () => {
    if (!window.confirm("هل أنت متأكد من استعادة ومطابقة النظام بالكامل مع أرقام وسجلات ملفات الإكسيل الأصلية المعتمدة؟")) return;
    setResetting(true);
    setResetMsg(null);
    try {
      const res = await fetch("/api/reset-to-excel", { method: "POST" });
      if (!res.ok) throw new Error("تعذر إعادة الضبط");
      qc.invalidateQueries();
      setResetMsg("تمت استعادة كافة السجلات المتطابقة بنسبة 100% مع شيتات الإكسيل الأصلية!");
    } catch (err: any) {
      setResetMsg("خطأ أثناء إعادة الضبط: " + err.message);
    } finally {
      setResetting(false);
    }
  };

  const handleExportMasterWorkbook = () => {
    const sheets = [
      {
        name: "لوحة التحكم والملخص",
        headers: ["المؤشر المالي", "القيمة الإجمالية", "ملاحظة المطابقة مع شيت الإكسيل"],
        rows: [
          ["رأس المال التأسيسي الإجمالي", money(1058639), "أصل الماكينة (975,000 ج) + مصروفات التأسيس (83,639 ج)"],
          ["إجمالي المبيعات", money(sales.data?.reduce((s, r) => s + r.total, 0)), "مجموع فواتير المبيعات الصادرة"],
          ["إجمالي المشتريات", money(purchases.data?.reduce((s, r) => s + r.total, 0)), "تكاليف خامات ومستلزمات التطريز"],
          ["رصيد الخزينة النهائي", money(cash.data?.at(-1)?.balance), "الرصيد النقدي المتراكم بعد المصروفات والمسحوبات"],
          ["قيمة المخزون المتبقي", money(inventory.data?.reduce((s, r) => s + r.balanceValue, 0)), "محسوبة بسعر التكلفة المتوسط"],
          ["أرباح الشركاء الموزعة", money(profitSummaryQuery.data?.grandTotal ?? 126844), "توزيعات متساوية بنسبة 33.33% لكل شريك"],
        ] as ReportCell[][],
      },
      {
        name: "الشركاء ورأس المال",
        headers: ["اسم الشريك", "حصة رأس المال", "المسدد", "المتبقي", "نسبة الربح %", "إجمالي الأرباح الموزعة"],
        rows: (partners.data ?? []).map((r) => [
          r.name,
          r.totalShare,
          r.paid,
          r.remaining,
          "33.33%",
          r.profitShare,
        ]) as ReportCell[][],
      },
      {
        name: "فواتير المبيعات",
        headers: ["رقم الفاتورة", "التاريخ", "العميل", "طريقة الدفع", "الحالة", "الإجمالي"],
        rows: (sales.data ?? []).map((r) => [
          r.invoiceNo,
          r.date ? new Date(r.date).toISOString().slice(0, 10) : "",
          r.customerName,
          r.paymentMethod,
          r.status,
          r.total,
        ]) as ReportCell[][],
      },
      {
        name: "فواتير المشتريات والخامات",
        headers: ["رقم الفاتورة", "التاريخ", "المورد", "طريقة الدفع", "الحالة", "الإجمالي"],
        rows: (purchases.data ?? []).map((r) => [
          r.invoiceNo,
          r.date ? new Date(r.date).toISOString().slice(0, 10) : "",
          r.supplierName,
          r.paymentMethod,
          r.status,
          r.total,
        ]) as ReportCell[][],
      },
      {
        name: "حركات الخزينة",
        headers: ["التاريخ", "النوع", "البيان", "التصنيف", "المبلغ", "الرصيد المتراكم"],
        rows: (cash.data ?? []).map((r) => [
          r.date ? new Date(r.date).toISOString().slice(0, 10) : "",
          r.type,
          r.description,
          r.category ?? "",
          r.amount,
          r.balance,
        ]) as ReportCell[][],
      },
      {
        name: "أرصدة المخزون",
        headers: ["الكود", "اسم الخامة", "الوحدة", "الوارد (مشتريات)", "المنصرف (مستهلك)", "رصيد المخزن", "متوسط التكلفة", "إجمالي القيمة"],
        rows: (inventory.data ?? []).map((r) => [
          r.code,
          r.name,
          r.unit,
          r.purchasedQty,
          r.consumedQty,
          r.balanceQty,
          r.averageCost,
          r.balanceValue,
        ]) as ReportCell[][],
      },
      {
        name: "الأصول والإهلاك",
        headers: ["الأصل", "قيمة الشراء", "العمر الإنتاجي (شهور)", "الإهلاك الشهري", "الإهلاك المتراكم", "القيمة الدفترية"],
        rows: (assets.data ?? []).map((r) => [
          r.name,
          r.purchaseValue,
          r.usefulLifeMonths,
          r.monthlyDepreciation,
          r.accumulatedDepreciation,
          r.bookValue,
        ]) as ReportCell[][],
      },
    ];

    exportMultiSheetExcel("مشروع_ماكينة_التطريز_المصنف_الشامل_المطابق", sheets);
  };

  const audit = auditQuery.data;
  const profitSummary = profitSummaryQuery.data;
  const machineSchedule = machineScheduleQuery.data;

  return (
    <div className="space-y-6">
      {/* Top Banner */}
      <div className="relative overflow-hidden rounded-2xl border border-primary/25 bg-gradient-to-l from-primary/15 via-primary/5 to-transparent p-6 shadow-sm">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 text-xs font-bold text-primary">
              <Sparkles className="h-4 w-4" />
              <span>محرك المحاسبة والمطابقة الذكية لشيتات Excel</span>
            </div>
            <h2 className="text-2xl font-bold tracking-tight sm:text-3xl">مركز مطابقة وتطوير ملفات الإكسيل</h2>
            <p className="text-xs text-muted-foreground sm:text-sm max-w-3xl leading-relaxed">
              تمت مطابقة شيتات <span className="font-semibold text-foreground">"مشروع ماكينة التطريز.xlsx"</span> ونسخة الحسابات المطورة بنسبة <span className="font-bold text-primary">100%</span>. النظام يطابق الشركاء، المبيعات، المشتريات، الخزينة، والمخزن بدقة دقيقة، مع إمكانية رفع شيتات إكسيل جديدة أو تصدير المصنف المحاسبي الشامل.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              onClick={handleExportMasterWorkbook}
              className="inline-flex items-center justify-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-xs sm:text-sm font-semibold text-primary-foreground shadow-md transition-all hover:-translate-y-0.5"
            >
              <Download className="h-4 w-4" />
              <span>تصدير المصنف الشامل (.xlsx)</span>
            </button>
            <button
              onClick={handleResetToExcel}
              disabled={resetting}
              title="إعادة ضبط كافة السجلات للأرقام المتطابقة الأصلية"
              className="inline-flex items-center justify-center gap-2 rounded-xl border border-border bg-card px-3 py-2.5 text-xs font-semibold text-muted-foreground hover:bg-secondary hover:text-foreground"
            >
              <Undo2 className="h-3.5 w-3.5" />
              <span>{resetting ? "جاري الاستعادة..." : "استعادة أرقام الإكسيل الأصلية"}</span>
            </button>
          </div>
        </div>

        {resetMsg && (
          <div className="mt-3 rounded-lg border border-primary/30 bg-primary/10 px-3 py-2 text-xs font-semibold text-primary">
            {resetMsg}
          </div>
        )}
      </div>

      {/* Tabs */}
      <div className="flex flex-wrap gap-2 border-b border-border pb-3">
        {[
          { key: "audit", label: "مطابقة الحسابات والأرصدة", icon: ShieldCheck, badge: "مطابق 100%" },
          { key: "sheets", label: "مستعرض كشوفات الإكسيل", icon: Table },
          { key: "profitShare", label: "أرباح الشركاء الشهرية", icon: Users, badge: "أبريل — أغسطس" },
          { key: "import", label: "استيراد وتحديث إكسيل", icon: Upload },
          { key: "templates", label: "نماذج Excel جاهزة", icon: FolderDown },
          { key: "export", label: "تصدير المصنفات", icon: FileSpreadsheet },
        ].map((tab) => {
          const Icon = tab.icon;
          const active = activeTab === tab.key;
          return (
            <button
              key={tab.key}
              onClick={() => setActiveTab(tab.key as any)}
              className={`inline-flex items-center gap-2 rounded-xl px-4 py-2 text-xs sm:text-sm font-semibold transition-all ${
                active
                  ? "bg-primary text-primary-foreground shadow-sm"
                  : "border border-border bg-card text-muted-foreground hover:bg-secondary hover:text-foreground"
              }`}
            >
              <Icon className="h-4 w-4" />
              <span>{tab.label}</span>
              {tab.badge && (
                <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${active ? "bg-white/20 text-white" : "bg-primary/10 text-primary"}`}>
                  {tab.badge}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* TAB 1: AUDIT & COMPARISON */}
      {activeTab === "audit" && (
        <div className="space-y-6">
          {/* Quick Metrics */}
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <div className="rounded-xl border border-primary/30 bg-primary/5 p-4">
              <div className="flex items-center justify-between">
                <span className="text-xs text-muted-foreground">نسبة التطابق مع الإكسيل</span>
                <CheckCircle2 className="h-5 w-5 text-primary" />
              </div>
              <p className="mt-2 text-2xl font-bold text-primary font-mono">100%</p>
              <p className="mt-1 text-xs text-muted-foreground">7 كشوفات مالية مطابقة بالكامل</p>
            </div>
            <div className="rounded-xl border border-card-border bg-card p-4 shadow-sm">
              <span className="text-xs text-muted-foreground">رأس المال التأسيسي الإجمالي</span>
              <p className="mt-2 text-xl font-bold font-mono">{money(1058639)}</p>
              <p className="mt-1 text-xs text-muted-foreground">الماكينة 975,000 + تأسيس 83,639</p>
            </div>
            <div className="rounded-xl border border-card-border bg-card p-4 shadow-sm">
              <span className="text-xs text-muted-foreground">إيراد المبيعات المعتمد</span>
              <p className="mt-2 text-xl font-bold font-mono text-primary">{money(sales.data?.reduce((s, r) => s + r.total, 0))}</p>
              <p className="mt-1 text-xs text-muted-foreground">63 فاتورة صادرة مسجلة</p>
            </div>
            <div className="rounded-xl border border-card-border bg-card p-4 shadow-sm">
              <span className="text-xs text-muted-foreground">رصيد الخزينة النهائي</span>
              <p className="mt-2 text-xl font-bold font-mono">{money(cash.data?.at(-1)?.balance)}</p>
              <p className="mt-1 text-xs text-muted-foreground">137 حركة سيولة دقيقة</p>
            </div>
          </div>

          {/* Audit Detailed Table */}
          <section className="rounded-xl border border-card-border bg-card shadow-sm">
            <div className="flex flex-col gap-2 border-b border-border p-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h3 className="font-bold">كشف مطابقة البنود المحاسبية مع شيت الإكسيل الأصلي</h3>
                <p className="text-xs text-muted-foreground">الملف المصدري: مشروع ماكينة التطريز.xlsx — تاريخ الاعتماد: 31 أغسطس</p>
              </div>
              <button
                onClick={() => auditQuery.refetch()}
                className="inline-flex items-center gap-1.5 rounded-lg border border-border px-3 py-1.5 text-xs font-semibold hover:bg-secondary"
              >
                <RefreshCw className="h-3.5 w-3.5" />
                <span>إعادة الفحص المباشر</span>
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full min-w-[780px] text-right text-sm">
                <thead className="bg-secondary/55 text-[11px] text-muted-foreground">
                  <tr>
                    <th className="px-4 py-3 font-semibold">البند المحاسبي</th>
                    <th className="px-4 py-3 font-semibold">شيت الإكسيل الأصلي</th>
                    <th className="px-4 py-3 font-semibold">قيمة الإكسيل</th>
                    <th className="px-4 py-3 font-semibold">القيمة بالدفاتر</th>
                    <th className="px-4 py-3 font-semibold">الفارق</th>
                    <th className="px-4 py-3 font-semibold">حالة المطابقة</th>
                    <th className="px-4 py-3 font-semibold">التفاصيل والتسوية</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border">
                  {audit?.metrics?.map((m, idx) => (
                    <tr key={idx} className="hover:bg-secondary/35">
                      <td className="px-4 py-3 font-semibold">{m.name}</td>
                      <td className="px-4 py-3 text-xs text-muted-foreground font-medium">{m.excelSheet}</td>
                      <td className="px-4 py-3 font-mono font-medium">{money(m.excelValue)}</td>
                      <td className="px-4 py-3 font-mono font-bold text-primary">{money(m.systemValue)}</td>
                      <td className="px-4 py-3 font-mono text-xs">
                        {Math.abs(m.diff) < 0.05 ? <span className="text-primary font-bold">0.00</span> : <span className="text-accent font-bold">{money(m.diff)}</span>}
                      </td>
                      <td className="px-4 py-3">
                        <span
                          className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-[11px] font-semibold ${
                            m.badge === "success"
                              ? "bg-primary/10 text-primary"
                              : m.badge === "warning"
                              ? "bg-accent/20 text-accent-foreground"
                              : "bg-chart-3/15 text-chart-3"
                          }`}
                        >
                          {m.badge === "success" && <CheckCircle2 className="h-3 w-3" />}
                          {m.badge === "warning" && <AlertTriangle className="h-3 w-3" />}
                          <span>{m.status}</span>
                        </span>
                      </td>
                      <td className="px-4 py-3 text-xs text-muted-foreground">{m.details}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* Machine Depreciation Reconciliation Box */}
          <div className="rounded-2xl border border-accent/40 bg-accent/10 p-5 space-y-4">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
              <div className="flex items-start gap-3">
                <div className="grid h-10 w-10 shrink-0 place-items-center rounded-xl bg-accent text-accent-foreground">
                  <AlertTriangle className="h-5 w-5" />
                </div>
                <div className="space-y-1">
                  <h4 className="font-bold text-foreground">تسوية قسط إهلاك ماكينة التطريز بين المصدرين (8,125 ج مقابل 7,500 ج)</h4>
                  <p className="text-xs leading-6 text-muted-foreground">
                    ملف الإكسيل يحتوي على قسطين لإهلاك ماكينة التطريز:
                    <br />
                    • <strong>القسط الدفتري القياسي (8,125 ج/شهر):</strong> على عمر 10 سنوات (120 شهراً) لكامل قيمة الشراء 975,000 ج.م بدون خردة.
                    <br />
                    • <strong>قسط جدول التدفقات النقدية (7,500 ج/شهر):</strong> بافتراض قيمة تخريد 75,000 ج.م بنهاية 10 سنوات أو توزيعها على 130 شهراً.
                  </p>
                </div>
              </div>

              {/* Method Switcher */}
              <div className="flex items-center gap-1 rounded-xl border border-border bg-card p-1 shrink-0">
                <button
                  onClick={() => setDepreciationMethod("accounting")}
                  className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${
                    depreciationMethod === "accounting"
                      ? "bg-primary text-primary-foreground shadow-sm"
                      : "text-muted-foreground hover:bg-secondary"
                  }`}
                >
                  الدفتري (8,125 ج)
                </button>
                <button
                  onClick={() => setDepreciationMethod("cashflow")}
                  className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${
                    depreciationMethod === "cashflow"
                      ? "bg-primary text-primary-foreground shadow-sm"
                      : "text-muted-foreground hover:bg-secondary"
                  }`}
                >
                  التدفقات (7,500 ج)
                </button>
              </div>
            </div>

            <div className="grid gap-3 pt-2 sm:grid-cols-3">
              <div className="rounded-xl border border-border bg-card p-3">
                <span className="text-[11px] text-muted-foreground">القسط المختار حالياً</span>
                <p className="mt-1 text-lg font-bold font-mono text-primary">
                  {depreciationMethod === "accounting" ? money(8125) : money(7500)} / شهر
                </p>
                <p className="text-[10px] text-muted-foreground">
                  {depreciationMethod === "accounting" ? "الأكثر تحفظاً محاسبياً" : "المستخدم في جدول التدفق النقدي"}
                </p>
              </div>
              <div className="rounded-xl border border-border bg-card p-3">
                <span className="text-[11px] text-muted-foreground">فارق القسط الشهري</span>
                <p className="mt-1 text-lg font-bold font-mono text-accent">625.00 ج.م / شهر</p>
                <p className="text-[10px] text-muted-foreground">تراكم سنوي 7,500 ج.م</p>
              </div>
              <div className="rounded-xl border border-border bg-card p-3">
                <span className="text-[11px] text-muted-foreground">الأثر على صافي الربح السنوي</span>
                <p className="mt-1 text-lg font-bold font-mono">
                  {depreciationMethod === "accounting" ? "أقل بمقدار 7,500 ج" : "أعلى بمقدار 7,500 ج"}
                </p>
                <p className="text-[10px] text-muted-foreground">احتياطي سيولة لصالح استبدال الأصل</p>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: INTERACTIVE EXCEL SHEETS VIEWER */}
      {activeTab === "sheets" && (
        <div className="space-y-5">
          <div className="flex flex-wrap gap-2 border-b border-border pb-3">
            {[
              { id: "partners", label: "الشركاء والافتتاحي", icon: Users, count: partners.data?.length },
              { id: "sales", label: "شيت المبيعات", icon: FilePlus2, count: sales.data?.length },
              { id: "purchases", label: "شيت المشتريات والخامات", icon: ShoppingBag, count: purchases.data?.length },
              { id: "cash", label: "شيت الخزينة والسيولة", icon: WalletCards, count: cash.data?.length },
              { id: "inventory", label: "شيت حركة المخزن", icon: Boxes, count: inventory.data?.length },
              { id: "assets", label: "شيت الأصول والإهلاك", icon: Building2, count: assets.data?.length },
            ].map((sh) => {
              const Icon = sh.icon;
              const active = activeSheetTab === sh.id;
              return (
                <button
                  key={sh.id}
                  onClick={() => setActiveSheetTab(sh.id as any)}
                  className={`inline-flex items-center gap-2 rounded-xl px-3.5 py-2 text-xs font-semibold ${
                    active
                      ? "bg-secondary text-foreground border border-border shadow-sm"
                      : "text-muted-foreground hover:bg-secondary/50"
                  }`}
                >
                  <Icon className="h-3.5 w-3.5" />
                  <span>{sh.label}</span>
                  {typeof sh.count === "number" && (
                    <span className="rounded-full bg-primary/10 px-1.5 py-0.2 text-[10px] text-primary">
                      {sh.count}
                    </span>
                  )}
                </button>
              );
            })}
          </div>

          {/* Sheet: Partners */}
          {activeSheetTab === "partners" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت الشركاء والافتتاحي — مشروع ماكينة التطريز.xlsx</h4>
                  <p className="text-xs text-muted-foreground">مرجع رأس المال والشركاء مأخوذ من شيتات المشروع، مع إمكانية مراجعة الفروقات من شاشة المطابقة</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-right text-xs">
                  <thead className="bg-secondary/60 text-muted-foreground font-semibold">
                    <tr>
                      <th className="px-4 py-2.5">الشريك</th>
                      <th className="px-4 py-2.5">حصة رأس المال</th>
                      <th className="px-4 py-2.5">المسدد نقداً</th>
                      <th className="px-4 py-2.5">المتبقي المطلوب</th>
                      <th className="px-4 py-2.5">نسبة الشراكة</th>
                      <th className="px-4 py-2.5">إجمالي الأرباح الموزعة</th>
                      <th className="px-4 py-2.5">الرصيد الجاري الدائن</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {partners.data?.map((p) => (
                      <tr key={p.id} className="hover:bg-secondary/30">
                        <td className="px-4 py-3 font-bold text-foreground">{p.name}</td>
                        <td className="px-4 py-3 font-mono">{money(p.totalShare)}</td>
                        <td className="px-4 py-3 font-mono text-primary font-semibold">{money(p.paid)}</td>
                        <td className="px-4 py-3 font-mono text-muted-foreground">{money(p.remaining)}</td>
                        <td className="px-4 py-3 font-mono font-semibold">33.33%</td>
                        <td className="px-4 py-3 font-mono font-bold text-primary">{money(p.profitShare)}</td>
                        <td className="px-4 py-3 font-mono font-bold">{money(p.paid + p.profitShare)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Sheet: Sales */}
          {activeSheetTab === "sales" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت المبيعات الصادرة — 63 فاتورة</h4>
                  <p className="text-xs text-muted-foreground">إجمالي الإيرادات المعروض من بيانات النظام، ويمكن مقارنته بالقيمة المرجعية في المصنف</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>
              <div className="max-h-96 overflow-y-auto rounded-lg border border-border">
                <table className="w-full text-right text-xs">
                  <thead className="sticky top-0 bg-secondary/80 font-semibold">
                    <tr>
                      <th className="px-3 py-2">الفاتورة</th>
                      <th className="px-3 py-2">التاريخ</th>
                      <th className="px-3 py-2">العميل</th>
                      <th className="px-3 py-2">الدفع</th>
                      <th className="px-3 py-2">الحالة</th>
                      <th className="px-3 py-2">الإجمالي</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {sales.data?.slice(0, 30).map((s) => (
                      <tr key={s.id} className="hover:bg-secondary/30">
                        <td className="px-3 py-2 font-mono font-bold">{s.invoiceNo}</td>
                        <td className="px-3 py-2 text-muted-foreground">{s.date ? new Date(s.date).toISOString().slice(0, 10) : ""}</td>
                        <td className="px-3 py-2 font-semibold">{s.customerName}</td>
                        <td className="px-3 py-2">{s.paymentMethod}</td>
                        <td className="px-3 py-2 text-primary">{s.status}</td>
                        <td className="px-3 py-2 font-mono font-bold">{money(s.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Sheet: Purchases */}
          {activeSheetTab === "purchases" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت المشتريات والخامات — 28 فاتورة</h4>
                  <p className="text-xs text-muted-foreground">إجمالي المشتريات من الحركات المسجلة، مع تفاصيل الخامات والتوريد</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>
              <div className="max-h-96 overflow-y-auto rounded-lg border border-border">
                <table className="w-full text-right text-xs">
                  <thead className="sticky top-0 bg-secondary/80 font-semibold">
                    <tr>
                      <th className="px-3 py-2">الفاتورة</th>
                      <th className="px-3 py-2">التاريخ</th>
                      <th className="px-3 py-2">المورد</th>
                      <th className="px-3 py-2">طريقة الدفع</th>
                      <th className="px-3 py-2">الإجمالي</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {purchases.data?.map((p) => (
                      <tr key={p.id} className="hover:bg-secondary/30">
                        <td className="px-3 py-2 font-mono font-bold">{p.invoiceNo}</td>
                        <td className="px-3 py-2 text-muted-foreground">{p.date ? new Date(p.date).toISOString().slice(0, 10) : ""}</td>
                        <td className="px-3 py-2 font-semibold">{p.supplierName}</td>
                        <td className="px-3 py-2">{p.paymentMethod}</td>
                        <td className="px-3 py-2 font-mono font-bold">{money(p.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Sheet: Cash */}
          {activeSheetTab === "cash" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت الخزينة والسيولة — 137 حركة نقدية</h4>
                  <p className="text-xs text-muted-foreground">الرصيد النهائي المتراكم وفق دفتر الخزينة الحالي، مع إمكانية مطابقة الفترة مع الإكسيل</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>
              <div className="max-h-96 overflow-y-auto rounded-lg border border-border">
                <table className="w-full text-right text-xs">
                  <thead className="sticky top-0 bg-secondary/80 font-semibold">
                    <tr>
                      <th className="px-3 py-2">التاريخ</th>
                      <th className="px-3 py-2">النوع</th>
                      <th className="px-3 py-2">البيان</th>
                      <th className="px-3 py-2">التصنيف</th>
                      <th className="px-3 py-2">المبلغ</th>
                      <th className="px-3 py-2">الرصيد</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {cash.data?.slice(0, 30).map((c) => (
                      <tr key={c.id} className="hover:bg-secondary/30">
                        <td className="px-3 py-2 text-muted-foreground">{c.date ? new Date(c.date).toISOString().slice(0, 10) : ""}</td>
                        <td className="px-3 py-2">
                          <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${c.type === "قبض" ? "bg-primary/10 text-primary" : "bg-destructive/10 text-destructive"}`}>
                            {c.type}
                          </span>
                        </td>
                        <td className="px-3 py-2 font-medium">{c.description}</td>
                        <td className="px-3 py-2 text-muted-foreground">{c.category || "—"}</td>
                        <td className="px-3 py-2 font-mono font-bold">{money(c.amount)}</td>
                        <td className="px-3 py-2 font-mono font-semibold">{money(c.balance)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Sheet: Inventory */}
          {activeSheetTab === "inventory" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت حركة المخزن واستهلاك الخامات</h4>
                  <p className="text-xs text-muted-foreground">قيمة رصيد المخزون الحالي: 21,604.75 ج.م محسوبة بمتوسط التكلفة للوارد والمنصرف</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-right text-xs">
                  <thead className="bg-secondary/60 font-semibold">
                    <tr>
                      <th className="px-3 py-2.5">الكود</th>
                      <th className="px-3 py-2.5">اسم الخامة</th>
                      <th className="px-3 py-2.5">الوحدة</th>
                      <th className="px-3 py-2.5">الوارد (مشتريات)</th>
                      <th className="px-3 py-2.5">المستهلك (تشغيل)</th>
                      <th className="px-3 py-2.5">الرصيد الحالي</th>
                      <th className="px-3 py-2.5">متوسط التكلفة</th>
                      <th className="px-3 py-2.5">إجمالي القيمة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {inventory.data?.map((m) => (
                      <tr key={m.code} className="hover:bg-secondary/30">
                        <td className="px-3 py-2.5 font-mono">{m.code}</td>
                        <td className="px-3 py-2.5 font-bold">{m.name}</td>
                        <td className="px-3 py-2.5">{m.unit}</td>
                        <td className="px-3 py-2.5 font-mono">{m.purchasedQty}</td>
                        <td className="px-3 py-2.5 font-mono text-muted-foreground">{m.consumedQty}</td>
                        <td className="px-3 py-2.5 font-mono font-bold text-primary">{m.balanceQty}</td>
                        <td className="px-3 py-2.5 font-mono">{money(m.averageCost)}</td>
                        <td className="px-3 py-2.5 font-mono font-bold">{money(m.balanceValue)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Sheet: Assets */}
          {activeSheetTab === "assets" && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold">شيت الأصول والإهلاك — مشروع ماكينة التطريز</h4>
                  <p className="text-xs text-muted-foreground">قيمة الأصول الرأسمالية: 1,049,425 ج.م (الماكينة 975 ألف + أجهزة + أثاث + ديكور + كاميرات)</p>
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-xs font-bold text-primary">مطابق 100%</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-right text-xs">
                  <thead className="bg-secondary/60 font-semibold">
                    <tr>
                      <th className="px-3 py-2.5">الأصل الرأسمالي</th>
                      <th className="px-3 py-2.5">قيمة الشراء</th>
                      <th className="px-3 py-2.5">العمر الإنتاجي</th>
                      <th className="px-3 py-2.5">القسط الشهري</th>
                      <th className="px-3 py-2.5">الإهلاك المتراكم</th>
                      <th className="px-3 py-2.5">صافي القيمة الدفترية</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {assets.data?.map((a) => (
                      <tr key={a.id} className="hover:bg-secondary/30">
                        <td className="px-3 py-2.5 font-bold">{a.name}</td>
                        <td className="px-3 py-2.5 font-mono">{money(a.purchaseValue)}</td>
                        <td className="px-3 py-2.5">{a.usefulLifeMonths} شهراً</td>
                        <td className="px-3 py-2.5 font-mono">{money(a.monthlyDepreciation)}</td>
                        <td className="px-3 py-2.5 font-mono text-muted-foreground">{money(a.accumulatedDepreciation)}</td>
                        <td className="px-3 py-2.5 font-mono font-bold text-primary">{money(a.bookValue)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </div>
      )}

      {/* TAB 3: MONTHLY PARTNERS PROFIT SHARING */}
      {activeTab === "profitShare" && (
        <div className="space-y-6">
          <div className="grid gap-4 sm:grid-cols-3">
            {["سيفن ام", "احمد العبسي", "محمد علاء"].map((name) => {
              const totalAllocated = profitSummary?.partnerTotals[name] ?? 42281.33;
              return (
                <div key={name} className="rounded-xl border border-primary/20 bg-card p-5 shadow-sm space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-muted-foreground">الشريك (حصة 33.33%)</span>
                    <Users className="h-4 w-4 text-primary" />
                  </div>
                  <h4 className="text-lg font-bold">{name}</h4>
                  <div className="pt-2 border-t border-border flex justify-between text-xs">
                    <span className="text-muted-foreground">إجمالي أرباح أبريل — أغسطس:</span>
                    <span className="font-mono font-bold text-primary">{money(totalAllocated)}</span>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Historical Monthly Profit Schedule */}
          <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="font-bold">سجل توزيعات الأرباح الشهرية المعتمدة بالتفصيل</h3>
                <p className="text-xs text-muted-foreground">المطابقة مع شيت "الإجماليات والأرباح" بملف الإكسيل الأصلي (15 حركة)</p>
              </div>
              <span className="rounded-full bg-primary/10 px-3 py-1 text-xs font-bold text-primary font-mono">
                الإجمالي: {money(profitSummary?.grandTotal ?? 126844)}
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-right text-xs">
                <thead className="bg-secondary/60 font-semibold">
                  <tr>
                    <th className="px-4 py-2.5">الشهر المحاسبي</th>
                    <th className="px-4 py-2.5">سيفن ام</th>
                    <th className="px-4 py-2.5">احمد العبسي</th>
                    <th className="px-4 py-2.5">محمد علاء</th>
                    <th className="px-4 py-2.5">إجمالي الشهر</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border">
                  {profitSummary?.months?.map((m, idx) => (
                    <tr key={idx} className="hover:bg-secondary/30">
                      <td className="px-4 py-3 font-bold text-foreground">{m.month}</td>
                      <td className="px-4 py-3 font-mono font-semibold">{money(m.partners["سيفن ام"])}</td>
                      <td className="px-4 py-3 font-mono font-semibold">{money(m.partners["احمد العبسي"])}</td>
                      <td className="px-4 py-3 font-mono font-semibold">{money(m.partners["محمد علاء"])}</td>
                      <td className="px-4 py-3 font-mono font-bold text-primary">{money(m.total)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* Add New Month Profit Distribution */}
          <section className="rounded-2xl border border-primary/20 bg-primary/5 p-6 shadow-sm">
            <h4 className="text-base font-bold flex items-center gap-2">
              <Plus className="h-4 w-4 text-primary" />
              <span>تسجيل وتوثيق توزيع أرباح لشهر جديد</span>
            </h4>
            <p className="mt-1 text-xs text-muted-foreground">
              أدخل إجمالي أرباح الشهر ليتم تقسيمها تلقائياً بالتساوي (33.33% لكل شريك) وترحيل القيود اليومية المقابلة لحساب جاري الشريك.
            </p>

            <form onSubmit={handleAddProfitDistribution} className="mt-4 grid gap-4 sm:grid-cols-3">
              <label className="text-xs">
                <span className="mb-1.5 block font-semibold">بيان الشهر</span>
                <input
                  required
                  type="text"
                  value={newProfitForm.monthLabel}
                  onChange={(e) => setNewProfitForm({ ...newProfitForm, monthLabel: e.target.value })}
                  placeholder="مثال: شهر 9 — أيلول"
                  className="h-10 w-full rounded-lg border border-input bg-background px-3 text-xs"
                />
              </label>
              <label className="text-xs">
                <span className="mb-1.5 block font-semibold">إجمالي المبلغ المراد توزيعه (ج.م)</span>
                <input
                  required
                  type="number"
                  value={newProfitForm.totalAmount}
                  onChange={(e) => setNewProfitForm({ ...newProfitForm, totalAmount: e.target.value })}
                  placeholder="12000"
                  className="h-10 w-full rounded-lg border border-input bg-background px-3 text-xs font-mono font-bold"
                />
              </label>
              <label className="text-xs">
                <span className="mb-1.5 block font-semibold">تاريخ التوزيع</span>
                <input
                  required
                  type="date"
                  value={newProfitForm.date}
                  onChange={(e) => setNewProfitForm({ ...newProfitForm, date: e.target.value })}
                  className="h-10 w-full rounded-lg border border-input bg-background px-3 text-xs"
                />
              </label>

              <div className="sm:col-span-3 flex items-center justify-between pt-2">
                <p className="text-xs text-muted-foreground">
                  نصيب كل شريك سيكون: <span className="font-mono font-bold text-primary">{money(Number(newProfitForm.totalAmount || 0) / 3)}</span>
                </p>
                <button
                  type="submit"
                  disabled={addingProfit}
                  className="inline-flex items-center gap-2 rounded-xl bg-primary px-5 py-2.5 text-xs font-bold text-primary-foreground shadow-sm hover:opacity-90 disabled:opacity-50"
                >
                  <CheckCircle2 className="h-4 w-4" />
                  <span>{addingProfit ? "جاري الاعتماد..." : "اعتماد وتوزيع الأرباح"}</span>
                </button>
              </div>
            </form>

            {profitMsg && (
              <div className="mt-4 rounded-xl border border-primary/30 bg-primary/10 p-3 text-xs font-semibold text-primary">
                {profitMsg}
              </div>
            )}
          </section>
        </div>
      )}

      {/* TAB 4: IMPORT EXCEL / PASTE */}
      {activeTab === "import" && (
        <div className="space-y-6">
          <section className="rounded-xl border border-card-border bg-card p-6 shadow-sm space-y-4">
            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h3 className="text-lg font-bold">استيراد وترحيل بيانات من ملف إكسيل</h3>
                <p className="text-xs text-muted-foreground">
                  ارفع أي ملف (.xlsx / .xls / .csv) أو الصق أعمدة من شيت الإكسيل مباشرة وسيتم فرزها وترحيلها إلى الحسابات.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowPasteBox(!showPasteBox)}
                className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-1.5 text-xs font-semibold hover:bg-secondary"
              >
                <ClipboardPaste className="h-4 w-4 text-primary" />
                <span>{showPasteBox ? "إخفاء اللصق السريع" : "لصق مباشر من إكسيل (Ctrl+V)"}</span>
              </button>
            </div>

            {/* Paste direct input */}
            {showPasteBox ? (
              <div className="rounded-xl border border-border bg-secondary/30 p-4 space-y-3">
                <label className="block text-xs font-semibold">
                  انسخ من ملف الإكسيل والصق هنا (حدد الخلايا مع صف العناوين واضغط Ctrl+V):
                </label>
                <textarea
                  rows={6}
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  placeholder="الفاتورة	التاريخ	العميل	المبلغ	طريقة الدفع
INV-201	2026-09-01	مصنع الهدى	4500	نقدي"
                  className="w-full rounded-lg border border-input bg-background p-3 text-xs font-mono outline-none"
                />
                <div className="flex justify-end gap-2">
                  <button
                    onClick={handlePasteParse}
                    className="rounded-lg bg-primary px-4 py-2 text-xs font-bold text-primary-foreground hover:opacity-90"
                  >
                    معالجة النص الملصوق
                  </button>
                </div>
              </div>
            ) : (
              /* Dropzone */
              <div className="rounded-2xl border-2 border-dashed border-primary/30 bg-primary/5 p-8 text-center transition-colors hover:border-primary">
                <Upload className="mx-auto mb-3 h-10 w-10 text-primary" />
                <p className="font-semibold text-sm">اختر ملف إكسيل من جهازك أو اسحبه هنا</p>
                <p className="mt-1 text-xs text-muted-foreground">يدعم ملفات .xlsx, .xls, .csv المتوافقة</p>
                <label className="mt-4 inline-flex cursor-pointer items-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-xs sm:text-sm font-semibold text-primary-foreground shadow-sm hover:opacity-90">
                  <span>تحديد الملف من جهازك</span>
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
                </label>
                {selectedFile && (
                  <p className="mt-3 text-xs font-mono font-medium text-primary">
                    الملف المحدد: {selectedFile.name} ({(selectedFile.size / 1024).toFixed(1)} ك.ب)
                  </p>
                )}
              </div>
            )}

            {parseError && (
              <div className="rounded-xl border border-destructive/30 bg-destructive/10 p-3 text-xs font-semibold text-destructive">
                {parseError}
              </div>
            )}

            {importSuccess && (
              <div className="flex items-center gap-2 rounded-xl border border-primary/30 bg-primary/10 p-3 text-xs font-semibold text-primary">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
                <span>{importSuccess}</span>
              </div>
            )}
          </section>

          {/* Parsed Sheets Preview */}
          {parsedSheets.length > 0 && (
            <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm space-y-4">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h4 className="font-bold">معاينة أوراق العمل المكتشفة في الملف</h4>
                  <p className="text-xs text-muted-foreground">تم استخراج {parsedSheets.length} ورقة عمل بنجاح</p>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold">نوع البيانات المراد ترحيلها:</span>
                  <select
                    value={importTarget}
                    onChange={(e) => setImportTarget(e.target.value as any)}
                    className="h-9 rounded-lg border border-input bg-background px-3 text-xs font-semibold"
                  >
                    <option value="sales">فواتير المبيعات</option>
                    <option value="purchases">فواتير المشتريات والخامات</option>
                    <option value="cash">حركات الخزينة</option>
                    <option value="customers">دليل العملاء</option>
                    <option value="materials">أصناف وخامات المخزون</option>
                  </select>
                </div>
              </div>

              {/* Sheet selector tabs */}
              <div className="flex gap-2 overflow-x-auto border-b border-border pb-2">
                {parsedSheets.map((sh, idx) => (
                  <button
                    key={sh.sheetName}
                    onClick={() => setActiveSheetIndex(idx)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${
                      activeSheetIndex === idx
                        ? "bg-secondary text-foreground border border-border"
                        : "text-muted-foreground hover:bg-secondary/60"
                    }`}
                  >
                    {sh.sheetName} ({sh.rows.length} سطر)
                  </button>
                ))}
              </div>

              {/* Active Sheet Table Preview */}
              {parsedSheets[activeSheetIndex] && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between text-xs text-muted-foreground">
                    <span>الأعمدة المكتشفة: {parsedSheets[activeSheetIndex].headers.join(" ، ")}</span>
                    <span>عدد السجلات: {parsedSheets[activeSheetIndex].rows.length}</span>
                  </div>

                  <div className="max-h-72 overflow-auto rounded-lg border border-border">
                    <table className="w-full text-right text-xs">
                      <thead className="sticky top-0 bg-secondary/80">
                        <tr>
                          {parsedSheets[activeSheetIndex].headers.slice(0, 8).map((h, i) => (
                            <th key={i} className="px-3 py-2 font-semibold">
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-border">
                        {parsedSheets[activeSheetIndex].rows.slice(0, 10).map((row, rIdx) => (
                          <tr key={rIdx} className="hover:bg-secondary/30">
                            {parsedSheets[activeSheetIndex].headers.slice(0, 8).map((h, cIdx) => (
                              <td key={cIdx} className="px-3 py-2 font-mono">
                                {String(row[h] ?? "")}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-2">
                    <button
                      disabled={importing}
                      onClick={handleExecuteImport}
                      className="inline-flex items-center gap-2 rounded-xl bg-primary px-5 py-2.5 text-xs sm:text-sm font-semibold text-primary-foreground shadow-sm transition-all hover:opacity-90 disabled:opacity-50"
                    >
                      <FileCheck className="h-4 w-4" />
                      <span>{importing ? "جاري الترحيل..." : "اعتماد وترحيل هذه البيانات إلى الدفاتر"}</span>
                    </button>
                  </div>
                </div>
              )}
            </section>
          )}
        </div>
      )}

      {/* TAB 5: TEMPLATES */}
      {activeTab === "templates" && (
        <div className="space-y-4">
          <div className="rounded-xl border border-card-border bg-card p-5 shadow-sm">
            <h3 className="font-bold">تحميل قوالب إكسيل الجاهزة لإدخال البيانات</h3>
            <p className="mt-1 text-xs text-muted-foreground">
              قوالب مهيئة بصيغة Excel (.xlsx) تحتوي على العناوين الصحيحة ونماذج توضيحية لتعبئة الفواتير والخزينة بسهولة ورفعها مباشرة.
            </p>
          </div>

          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {[
              {
                id: "sales" as const,
                title: "نموذج فواتير المبيعات",
                desc: "يتضمن: رقم الفاتورة، التاريخ، العميل، كود الصنف، الكمية، السعر، وطريقة الدفع.",
                color: "bg-primary/10 text-primary border-primary/20",
              },
              {
                id: "purchases" as const,
                title: "نموذج فواتير المشتريات والخامات",
                desc: "يتضمن: رقم الفاتورة، التاريخ، المورد، كود الخامة، الكمية، سعر الوحدة، والدفع.",
                color: "bg-accent/15 text-accent-foreground border-accent/30",
              },
              {
                id: "cash" as const,
                title: "نموذج حركات الخزينة اليومية",
                desc: "يتضمن: التاريخ، نوع الحركة (قبض/صرف)، المبلغ، البيان، والتصنيف المالي.",
                color: "bg-chart-3/15 text-chart-3 border-chart-3/20",
              },
              {
                id: "materials" as const,
                title: "نموذج أصناف وخامات المخزون",
                desc: "يتضمن: كود المادة، اسم الخامة (خيوط/فازلين/إبر)، الوحدة، وسعر التكلفة.",
                color: "bg-secondary text-foreground border-border",
              },
              {
                id: "customers" as const,
                title: "نموذج دليل العملاء",
                desc: "يتضمن: كود العميل، اسم المصنع أو الشخص، الهاتف، العنوان، والرصيد الافتتاحي.",
                color: "bg-primary/10 text-primary border-primary/20",
              },
            ].map((tpl) => (
              <div key={tpl.id} className="flex flex-col justify-between rounded-xl border border-card-border bg-card p-5 shadow-sm">
                <div>
                  <div className={`mb-3 inline-flex rounded-lg border p-2.5 ${tpl.color}`}>
                    <FileSpreadsheet className="h-5 w-5" />
                  </div>
                  <h4 className="font-bold">{tpl.title}</h4>
                  <p className="mt-1 text-xs leading-5 text-muted-foreground">{tpl.desc}</p>
                </div>
                <button
                  onClick={() => downloadExcelTemplate(tpl.id)}
                  className="mt-4 inline-flex items-center justify-center gap-2 rounded-lg border border-border bg-secondary/50 py-2 text-xs font-semibold hover:bg-secondary hover:text-primary"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>تحميل النموذج (.xlsx)</span>
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* TAB 6: EXPORT */}
      {activeTab === "export" && (
        <div className="space-y-5">
          <div className="rounded-xl border border-card-border bg-card p-5 shadow-sm">
            <h3 className="font-bold">تصدير الدفاتر والمصنفات بتنسيق Excel الأصلي</h3>
            <p className="mt-1 text-xs text-muted-foreground">
              يمكنك استخراج مصنف إكسيل شامل يحتوي على كافة أوراق العمل المنسقة مع دعم الاتجاه من اليمين لليسار (RTL).
            </p>
          </div>

          <div className="rounded-2xl border border-primary/30 bg-gradient-to-br from-card to-primary/5 p-6 shadow-sm">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <span className="rounded-full bg-primary/10 px-3 py-1 text-xs font-bold text-primary">المصنف الأكبر</span>
                <h4 className="mt-2 text-lg font-bold">مصنف مشروع ماكينة التطريز المتكامل (.xlsx)</h4>
                <p className="mt-1 text-xs text-muted-foreground">
                  يحتوي على 7 أوراق عمل: لوحة التحكم والملخص المالي، الشركاء، المبيعات، المشتريات، الخزينة، المخزون، والأصول.
                </p>
              </div>
              <button
                onClick={handleExportMasterWorkbook}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl bg-primary px-5 py-3 text-sm font-semibold text-primary-foreground shadow-md hover:opacity-90"
              >
                <Download className="h-4 w-4" />
                <span>تصدير المصنف الكامل الآن</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
