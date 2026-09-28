import { useMemo, useState, type FormEvent, type ReactNode } from 'react';
import { QueryClient, QueryClientProvider, useQueryClient, useQuery } from '@tanstack/react-query';
import { ErrorBoundary } from '@/components/error-boundary';
import { Toaster } from '@/components/ui/toaster';
import { TooltipProvider } from '@/components/ui/tooltip';
import NotFound from '@/pages/not-found';
import { ExcelHub } from '@/components/excel-hub';
import {
  ArrowDownLeft, ArrowUpRight, BarChart3, Boxes, BriefcaseBusiness, Building2, Calculator,
  CalendarDays, CheckCircle2, ChevronLeft, CircleDollarSign, ClipboardList, Download, Eye, FileBarChart, FilePlus2,
  FileSpreadsheet, Home as HomeIcon, Landmark, Menu, Package, PanelRightClose, Plus, Printer,
  RefreshCw, Search, Settings2, ShieldCheck, ShoppingBag, Sparkles, Tag, TrendingUp, Users, WalletCards, X, type LucideIcon
} from 'lucide-react';
import {
  getGetDashboardQueryKey, getListCashQueryKey, getListCustomersQueryKey, getListMaterialsQueryKey,
  getListProductsQueryKey, getListPurchasesQueryKey, getListSalesQueryKey, getListSuppliersQueryKey, getListJournalsQueryKey,
  getGetIncomeReportQueryKey, getGetReconciliationQueryKey,
  useCreateCashEntry, useCreateCustomer, useCreateMaterial, useCreateProduct, useCreatePurchase,
  useCreateSale, useCreateSupplier, useGetDashboard, useGetIncomeReport, useGetReconciliation,
  useListAssets, useListCash, useListCustomers, useListInventory, useListJournals, useListMaterials,
  useListPartners, useListProducts, useListPurchases, useListSales, useListSuppliers,
  type CashEntryInput, type CashEntryInputType, type CustomerInput, type MaterialInput, type ProductInput,
  type PurchaseInput, type SaleInput, type TransactionItem
} from '@workspace/api-client-react';
import { exportReportToExcel, printReport, type ReportCell } from '@/lib/report-export';
import {
  Route,
  Switch,
  Link,
  useLocation,
  Router as WouterRouter,
} from 'wouter';
import {
  Area, AreaChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis,
} from 'recharts';

const queryClient = new QueryClient();

type PageKey = 'dashboard'|'excel'|'sales'|'purchases'|'cash'|'inventory'|'employees'|'masters'|'balances'|'partners'|'assets'|'accounting'|'reports';
const nav: { key: PageKey; href: string; label: string; icon: LucideIcon; group: string }[] = [
  { key:'dashboard', href:'/', label:'لوحة المتابعة', icon:HomeIcon, group:'اليوم' },
  { key:'excel', href:'/excel', label:'شيتات الإكسيل والمطابقة', icon:FileSpreadsheet, group:'الإكسيل والحسابات' },
  { key:'sales', href:'/sales', label:'المبيعات', icon:FilePlus2, group:'التشغيل' },
  { key:'purchases', href:'/purchases', label:'المشتريات', icon:ShoppingBag, group:'التشغيل' },
  { key:'cash', href:'/cash', label:'الخزينة', icon:WalletCards, group:'التشغيل' },
  { key:'inventory', href:'/inventory', label:'المخزن والتكلفة', icon:Boxes, group:'التشغيل' },
  { key:'employees', href:'/employees', label:'الموظفون والمرتبات', icon:BriefcaseBusiness, group:'التشغيل' },
  { key:'masters', href:'/masters', label:'التعريفات', icon:Tag, group:'البيانات الأساسية' },
  { key:'balances', href:'/balances', label:'أرصدة العملاء والموردين', icon:ClipboardList, group:'التقارير التشغيلية' },
  { key:'partners', href:'/partners', label:'الشركاء ورأس المال', icon:Users, group:'الحسابات' },
  { key:'assets', href:'/assets', label:'الأصول والإهلاك', icon:Building2, group:'الحسابات' },
  { key:'accounting', href:'/accounting', label:'المحاسبة والتسويات', icon:Calculator, group:'الحسابات' },
  { key:'reports', href:'/reports', label:'التقارير والتصدير', icon:FileBarChart, group:'الحسابات' },
];

const money = (value?: number) => `${new Intl.NumberFormat('ar-EG', { maximumFractionDigits: 2 }).format(value ?? 0)} ج.م`;

const CASH_MAIN_CODES: Array<{ code: string; name: string; kind: 'customer'|'supplier'|'group' }> = [
  { code: '3001', name: 'مشتريات كاش', kind: 'supplier' },
  { code: '4001', name: 'الحاج علاء فكري', kind: 'customer' },
  { code: '4002', name: 'ياسر', kind: 'customer' },
  { code: '4003', name: 'ابو حمزة', kind: 'customer' },
  { code: '4004', name: 'علي رزة', kind: 'customer' },
  { code: '4005', name: 'الحاج محمود منصور', kind: 'customer' },
  { code: '4006', name: 'ابراهيم الليثي', kind: 'customer' },
  { code: '4007', name: 'ميدو عزت', kind: 'customer' },
  { code: '4008', name: 'مبيعات نقدية', kind: 'customer' },
  { code: '4009', name: 'مصطفي شاهين', kind: 'customer' },
  { code: '4010', name: 'زهران', kind: 'customer' },
  { code: '4011', name: 'محمد محسن', kind: 'customer' },
  { code: '4012', name: 'اسلام غنيم', kind: 'customer' },
  { code: '7000', name: 'مصاريف راس مالية', kind: 'group' },
  { code: '8000', name: 'مصروفات ادارية وعمومية', kind: 'group' },
  { code: '9000', name: 'مصاريف الانتاج', kind: 'group' },
];
const CASH_ANALYSIS_CODES = [
  ['6000','إيرادات'],['7001','شراء اصول'],['7002','مصروفات تاسيس'],['7004','تامين لدي الغير'],
  ['8001','ايجار المصنع'],['8002','مصروفات النظافة'],['8003','مرافق وخدمات'],['8004','التامين'],
  ['8005','اجور ومرتبات الادارة'],['8006','البوفية والضيافة'],['8007','ادوات مكتبية'],['8008','اتصالات وانترنت'],
  ['8009','البرامج والاشتراكات'],['8010','اتعاب'],['8011','مصاريف بنكية'],['8012','تجديد التراخيص'],
  ['8013','انتقالات'],['8014','شحن وتوصيل'],['8015','نثريات'],['8016','الدعاية والاعلان'],['8017','مكافات'],
  ['9001','الخامات'],['9002','قطع غيار'],['9003','كهرباء'],['9004','التعبئة والتغليف'],['9005','اجور الانتاج'],['9006','مصاريف صيانة'],
] as Array<[string,string]>;
const cashMainName = (code: string, fallback = '') => CASH_MAIN_CODES.find(x => x.code === code)?.name || fallback;
const cashAnalysisName = (code: string, fallback = '') => CASH_ANALYSIS_CODES.find(x => x[0] === code)?.[1] || fallback;

const dateLabel = (value?: string) => value ? new Intl.DateTimeFormat('ar-EG', { day:'2-digit', month:'short', year:'numeric' }).format(new Date(value)) : '—';
const todayLabel = () => new Intl.DateTimeFormat('ar-EG', { weekday:'long', day:'numeric', month:'long', year:'numeric' }).format(new Date());
const today = () => new Date().toISOString().slice(0,10);

function State({ loading, error, empty, retry, children }: { loading?: boolean; error?: boolean; empty?: boolean; retry?: () => void; children: ReactNode }) {
  if (loading) return <div className="grid gap-3 md:grid-cols-3"><div className="h-24 animate-pulse rounded-xl bg-muted" /><div className="h-24 animate-pulse rounded-xl bg-muted" /><div className="h-24 animate-pulse rounded-xl bg-muted" /></div>;
  if (error) return <div className="rounded-xl border border-destructive/30 bg-destructive/5 p-8 text-center"><ShieldCheck className="mx-auto mb-3 h-8 w-8 text-destructive" /><p className="font-semibold">تعذّر تحميل البيانات</p><p className="mt-1 text-sm text-muted-foreground">تحقق من الاتصال بالخادم ثم حاول مرة أخرى.</p>{retry && <button data-testid="button-retry" onClick={retry} className="mt-4 inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm text-primary-foreground"><RefreshCw className="h-4 w-4" />إعادة المحاولة</button>}</div>;
  if (empty) return <div className="rounded-xl border border-dashed border-border bg-card/70 p-12 text-center"><ClipboardList className="mx-auto mb-3 h-8 w-8 text-muted-foreground/50" /><p className="font-semibold">لا توجد سجلات بعد</p><p className="mt-1 text-sm text-muted-foreground">أضف أول سجل ليظهر هنا وتبدأ المتابعة.</p></div>;
  return <>{children}</>;
}

function Shell({ children }: { children: ReactNode }) {
  const [location, setLocation] = useLocation();
  const [mobileOpen, setMobileOpen] = useState(false);
  const active = nav.find((item) => item.href === location)?.key ?? 'dashboard';
  const current = nav.find((item) => item.key === active) ?? nav[0];
  const groups = Array.from(new Set(nav.map((item) => item.group)));
  return <div className="min-h-[100dvh] bg-background" dir="rtl">
    <aside className={`fixed inset-y-0 right-0 z-40 w-72 border-l border-sidebar-border bg-sidebar text-sidebar-foreground shadow-xl transition-transform lg:translate-x-0 ${mobileOpen ? 'translate-x-0' : 'translate-x-full'}`}>
      <div className="flex h-full flex-col">
        <div className="flex items-center justify-between border-b border-sidebar-border px-6 py-5">
          <Link href="/" onClick={() => setMobileOpen(false)} className="flex items-center gap-3" data-testid="link-brand">
            <div className="grid h-11 w-11 place-items-center rounded-xl bg-sidebar-primary text-sidebar-primary-foreground"><span className="font-mono text-lg font-bold">م</span></div>
            <div><div className="text-[17px] font-bold tracking-tight">ماكينة</div><div className="text-xs text-sidebar-foreground/55">الحسابات والمخزون</div></div>
          </Link>
          <button className="rounded-md p-2 text-sidebar-foreground/60 hover:bg-sidebar-accent lg:hidden" onClick={() => setMobileOpen(false)} data-testid="button-close-menu"><X className="h-5 w-5"/></button>
        </div>
        <div className="px-4 py-5">
          {groups.map((group) => <div key={group} className="mb-5">
            <p className="mb-2 px-3 text-[10px] font-bold tracking-[.16em] text-sidebar-foreground/40">{group}</p>
            <div className="space-y-1">{nav.filter((item) => item.group === group).map((item) => { const Icon = item.icon; return <Link key={item.key} href={item.href} onClick={() => setMobileOpen(false)} data-testid={`link-nav-${item.key}`} className={`flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm ${active === item.key ? 'bg-sidebar-accent font-semibold text-sidebar-accent-foreground shadow-sm' : 'text-sidebar-foreground/72 hover:bg-sidebar-accent/70 hover:text-sidebar-foreground'}`}><Icon className="h-[18px] w-[18px]" /><span>{item.label}</span>{active === item.key && <ChevronLeft className="mr-auto h-4 w-4 text-sidebar-primary" />}</Link>})}</div>
          </div>)}
        </div>
        <div className="mt-auto border-t border-sidebar-border p-4">
          <div className="rounded-xl bg-sidebar-accent/60 p-3"><div className="mb-2 flex items-center gap-2 text-xs font-semibold"><div className="h-2 w-2 rounded-full bg-sidebar-primary" />النظام يعمل</div><p className="text-[11px] leading-5 text-sidebar-foreground/55">آخر مزامنة مع الخادم<br />منذ لحظات</p></div>
        </div>
      </div>
    </aside>
    {mobileOpen && <button aria-label="إغلاق القائمة" className="fixed inset-0 z-30 bg-foreground/25 lg:hidden" onClick={() => setMobileOpen(false)} data-testid="button-overlay" />}
    <div className="lg:mr-72">
      <header className="sticky top-0 z-20 flex h-[76px] items-center justify-between border-b border-border/70 bg-background/90 px-4 backdrop-blur-md sm:px-8">
        <div className="flex items-center gap-3"><button className="rounded-lg border border-border p-2 lg:hidden" onClick={() => setMobileOpen(true)} data-testid="button-open-menu"><Menu className="h-5 w-5"/></button><div><p className="text-xs text-muted-foreground">{todayLabel()}</p><h1 className="text-lg font-bold tracking-tight">{current.label}</h1></div></div>
        <div className="flex items-center gap-2 sm:gap-4">
          <Link href="/excel" className="inline-flex items-center gap-1.5 rounded-lg border border-primary/30 bg-primary/10 px-2.5 py-1.5 text-xs font-bold text-primary hover:bg-primary/20 transition-colors">
            <FileSpreadsheet className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">مطابقة الإكسيل: 100%</span>
            <span className="sm:hidden">إكسيل</span>
          </Link>
          <div className="hidden items-center gap-2 rounded-lg border border-border bg-card px-3 py-2 sm:flex"><Search className="h-4 w-4 text-muted-foreground" /><input data-testid="input-global-search" className="w-40 border-0 bg-transparent text-xs outline-none placeholder:text-muted-foreground" placeholder="بحث سريع..." /></div><button className="rounded-lg border border-border bg-card p-2 text-muted-foreground hover:text-foreground" data-testid="button-settings"><Settings2 className="h-4 w-4"/></button><div className="flex h-9 w-9 items-center justify-center rounded-full bg-primary text-xs font-bold text-primary-foreground">م</div>
        </div>
      </header>
      <main className="paper-grid min-h-[calc(100dvh-76px)] px-4 py-6 sm:px-8 sm:py-8"><div className="mx-auto max-w-[1500px] fade-in">{children}</div></main>
    </div>
  </div>;
}

function PageHeading({ title, eyebrow, description, action, onAction, icon: Icon = BarChart3 }: { title:string; eyebrow:string; description:string; action?:string; onAction?:()=>void; icon?:LucideIcon }) {
  return <div className="mb-7 flex flex-col justify-between gap-4 sm:flex-row sm:items-end"><div><div className="mb-2 flex items-center gap-2 text-xs font-semibold text-primary"><Icon className="h-4 w-4"/><span>{eyebrow}</span></div><h2 className="text-2xl font-bold tracking-tight sm:text-3xl">{title}</h2><p className="mt-2 text-sm text-muted-foreground">{description}</p></div>{action && <button data-testid="button-page-action" onClick={onAction} className="inline-flex shrink-0 items-center justify-center gap-2 rounded-lg bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground shadow-sm hover:-translate-y-0.5"><Plus className="h-4 w-4"/>{action}</button>}</div>;
}

function Metric({ label, value, note, icon: Icon, tone='primary' }: { label:string; value:string; note?:string; icon:LucideIcon; tone?:'primary'|'accent'|'blue'|'red' }) {
  const color = { primary:'bg-primary/10 text-primary', accent:'bg-accent/20 text-foreground', blue:'bg-chart-3/15 text-chart-3', red:'bg-destructive/10 text-destructive' }[tone];
  return <div className="rounded-xl border border-card-border bg-card p-4 shadow-sm"><div className="flex items-start justify-between gap-3"><div><p className="text-xs font-medium text-muted-foreground">{label}</p><p className="mt-2 text-xl font-bold tracking-tight">{value}</p>{note && <p className="mt-1 text-[11px] text-muted-foreground">{note}</p>}</div><div className={`grid h-10 w-10 shrink-0 place-items-center rounded-lg ${color}`}><Icon className="h-5 w-5"/></div></div></div>;
}

function Dashboard() {
  const q = useGetDashboard();
  const salesQ = useListSales();
  const purchasesQ = useListPurchases();
  const auditQ = useQuery<{ overallMatchRate: number; totalChecks: number; passedChecks: number }>({
    queryKey: ['/api/excel-audit'],
    queryFn: async () => { const res = await fetch('/api/excel-audit'); if (!res.ok) throw new Error('تعذر جلب المطابقة'); return res.json(); },
  });
  const d = q.data;
  const [, navigate] = useLocation();
  const monthly = useMemo(() => {
    const map = new Map<string, { month: string; sales: number; purchases: number }>();
    for (const row of salesQ.data ?? []) {
      const key = String(row.date).slice(0, 7);
      const item = map.get(key) ?? { month: key, sales: 0, purchases: 0 };
      item.sales += row.total; map.set(key, item);
    }
    for (const row of purchasesQ.data ?? []) {
      const key = String(row.date).slice(0, 7);
      const item = map.get(key) ?? { month: key, sales: 0, purchases: 0 };
      item.purchases += row.total; map.set(key, item);
    }
    return [...map.values()].sort((a,b)=>a.month.localeCompare(b.month)).slice(-8).map(x => ({
      ...x,
      label: new Intl.DateTimeFormat('ar-EG', { month: 'short', year: '2-digit' }).format(new Date(`${x.month}-01`)),
    }));
  }, [salesQ.data, purchasesQ.data]);
  const lowStock = useListInventory();
  const lowStockRows = (lowStock.data ?? []).filter(x => Number(x.balanceQty) <= 5).slice(0, 5);
  const matchRate = auditQ.data?.overallMatchRate ?? null;
  return <>
    <PageHeading eyebrow="نظرة اليوم" title="غرفة التحكم" description="لوحة تشغيل ومالية تجمع المبيعات والمشتريات والخزينة والمخزون والمطابقة مع الإكسيل." action="تسجيل حركة" onAction={() => navigate('/cash')} icon={BarChart3}/>
    <div className="mb-5 rounded-2xl border border-primary/20 bg-card p-4 shadow-sm sm:p-5">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
        <div className="flex items-center gap-3.5">
          <div className="grid h-12 w-12 shrink-0 place-items-center rounded-xl bg-primary text-primary-foreground shadow-sm"><FileSpreadsheet className="h-5 w-5" /></div>
          <div><div className="flex flex-wrap items-center gap-2 text-xs font-bold text-primary"><span>مركز مطابقة Excel</span><span className="rounded-full bg-primary/10 px-2 py-0.5">{matchRate === null ? 'جارٍ الفحص' : `${Math.round(matchRate)}% مطابقة`}</span></div><h3 className="mt-1 text-sm font-bold">الحركات والتقارير مبنية على نفس منطق شيتات التشغيل والحسابات</h3><p className="mt-1 text-xs text-muted-foreground">راجع المطابقة التفصيلية، الاستيراد والتصدير من شاشة الإكسيل والحسابات.</p></div>
        </div>
        <Link href="/excel" className="inline-flex shrink-0 items-center justify-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-xs font-bold text-primary-foreground shadow-sm hover:opacity-90"><span>فتح مركز Excel</span><ChevronLeft className="h-4 w-4" /></Link>
      </div>
    </div>
    <State loading={q.isLoading} error={q.isError} retry={() => q.refetch()}>
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-5">
        <Metric label="إجمالي المبيعات" value={money(d?.salesTotal)} note="كل الفواتير المسجلة" icon={ArrowUpRight} tone="primary"/>
        <Metric label="إجمالي المشتريات" value={money(d?.purchasesTotal)} note="تكلفة التوريد" icon={ArrowDownLeft} tone="accent"/>
        <Metric label="رصيد الخزينة" value={money(d?.cashBalance)} note="الرصيد المتاح" icon={WalletCards} tone="blue"/>
        <Metric label="قيمة المخزون" value={money(d?.inventoryValue)} note="بسعر التكلفة" icon={Boxes} tone="accent"/>
        <Metric label="صافي النشاط" value={money(d?.netActivity)} note="مبيعات ناقص التكلفة المباشرة" icon={BarChart3} tone="primary"/>
      </div>
      <div className="mt-5 grid gap-5 xl:grid-cols-[1.45fr_.55fr]">
        <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm">
          <div className="mb-4 flex items-center justify-between"><div><h3 className="font-bold">المبيعات والمشتريات شهريًا</h3><p className="mt-1 text-xs text-muted-foreground">آخر 8 أشهر من الحركات المسجلة</p></div><TrendingUp className="h-5 w-5 text-primary"/></div>
          <div className="h-[280px]">{monthly.length ? <ResponsiveContainer width="100%" height="100%"><AreaChart data={monthly}><defs><linearGradient id="salesFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="hsl(var(--primary))" stopOpacity={0.28}/><stop offset="100%" stopColor="hsl(var(--primary))" stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))"/><XAxis dataKey="label" tick={{fontSize:11}}/><YAxis tick={{fontSize:10}} tickFormatter={(v)=>new Intl.NumberFormat('ar-EG',{notation:'compact'}).format(v)}/><Tooltip formatter={(v:number)=>money(v)} labelStyle={{direction:'rtl'}}/><Area type="monotone" dataKey="sales" name="المبيعات" stroke="hsl(var(--primary))" fill="url(#salesFill)" strokeWidth={3}/><Area type="monotone" dataKey="purchases" name="المشتريات" stroke="hsl(var(--accent))" fill="none" strokeWidth={2}/></AreaChart></ResponsiveContainer> : <div className="grid h-full place-items-center text-sm text-muted-foreground">لا توجد بيانات كافية للرسم بعد.</div>}</div>
        </section>
        <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm"><div className="mb-4 flex items-center justify-between"><div><h3 className="font-bold">تنبيهات المخزون</h3><p className="mt-1 text-xs text-muted-foreground">مواد رصيدها منخفض</p></div><Boxes className="h-5 w-5 text-accent"/></div>{lowStockRows.length ? <div className="space-y-2">{lowStockRows.map((x)=><div key={x.code} className="flex items-center justify-between rounded-lg border border-border bg-secondary/35 p-3"><div><p className="text-sm font-semibold">{x.name}</p><p className="text-[11px] text-muted-foreground">{x.code} · {x.unit}</p></div><span className="rounded-full bg-destructive/10 px-2 py-1 font-mono text-xs text-destructive">{x.balanceQty}</span></div>)}</div> : <div className="rounded-lg bg-primary/5 p-4 text-sm text-primary">لا توجد أصناف منخفضة الرصيد حاليًا.</div>}<button onClick={()=>navigate('/inventory')} className="mt-4 w-full rounded-lg border border-border py-2 text-xs font-semibold hover:bg-secondary">فتح المخزون</button></section>
      </div>
      <div className="mt-5 grid gap-5 xl:grid-cols-[1.4fr_.6fr]">
        <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm"><div className="mb-5 flex items-center justify-between"><div><h3 className="font-bold">آخر النشاطات</h3><p className="mt-1 text-xs text-muted-foreground">كل عملية مسجلة في دفتر اليوم</p></div><Link href="/accounting" className="text-xs font-semibold text-primary hover:underline">عرض الدفتر</Link></div>{!d?.recentActivity?.length ? <State empty>{null}</State> : <div className="divide-y divide-border">{d.recentActivity.map((item,index)=><div key={`${item.date}-${index}`} className="flex items-center justify-between gap-4 py-3 first:pt-0 last:pb-0"><div className="flex min-w-0 items-center gap-3"><div className={`grid h-9 w-9 shrink-0 place-items-center rounded-lg ${item.kind.includes('بيع') ? 'bg-primary/10 text-primary' : 'bg-accent/20 text-foreground'}`}>{item.kind.includes('بيع')?<ArrowUpRight className="h-4 w-4"/>:<ClipboardList className="h-4 w-4"/>}</div><div className="min-w-0"><p className="truncate text-sm font-semibold">{item.label}</p><p className="mt-0.5 text-[11px] text-muted-foreground">{dateLabel(item.date)} · {item.kind}</p></div></div><p className="whitespace-nowrap font-mono text-sm font-semibold">{money(item.amount)}</p></div>)}</div>}</section>
        <section className="rounded-xl border border-card-border bg-card p-5 shadow-sm"><div className="mb-5"><h3 className="font-bold">مؤشرات التحصيل</h3><p className="mt-1 text-xs text-muted-foreground">الأرصدة التي تحتاج متابعة</p></div><div className="space-y-5"><div><div className="mb-2 flex justify-between text-sm"><span>ذمم العملاء</span><span className="font-mono text-xs">{money(d?.customerReceivables)}</span></div><div className="h-2 overflow-hidden rounded-full bg-muted"><div className="h-full w-[68%] rounded-full bg-primary"/></div></div><div><div className="mb-2 flex justify-between text-sm"><span>مستحقات الموردين</span><span className="font-mono text-xs">{money(d?.supplierPayables)}</span></div><div className="h-2 overflow-hidden rounded-full bg-muted"><div className="h-full w-[42%] rounded-full bg-accent"/></div></div><div className="rounded-lg bg-secondary/70 p-3 text-xs leading-6 text-muted-foreground">استخدم شاشة التقارير لتحديد الفترة ومراجعة أرصدة العملاء والموردين والخزينة قبل الإقفال.</div></div></section>
      </div>
    </State>
  </>;
}

function TableShell({ title, description, search, onSearch, children, count }: { title:string; description?:string; search?:string; onSearch?:(v:string)=>void; children:ReactNode; count?:number }) {
  return <section className="rounded-xl border border-card-border bg-card shadow-sm"><div className="flex flex-col gap-3 border-b border-border p-4 sm:flex-row sm:items-center sm:justify-between"><div><h3 className="font-bold">{title}{typeof count === 'number' && <span className="mr-2 rounded-full bg-secondary px-2 py-0.5 text-[10px] text-muted-foreground">{count}</span>}</h3>{description && <p className="mt-1 text-xs text-muted-foreground">{description}</p>}</div>{onSearch && <div className="flex items-center gap-2 rounded-lg border border-border px-3 py-2 sm:w-64"><Search className="h-4 w-4 text-muted-foreground"/><input data-testid="input-table-search" value={search} onChange={(e) => onSearch(e.target.value)} className="w-full bg-transparent text-xs outline-none" placeholder="تصفية النتائج..." /></div>}</div><div className="overflow-x-auto">{children}</div></section>;
}
function Table({ headers, children }: { headers:string[]; children:ReactNode }) { return <table className="w-full min-w-[680px] text-right text-sm"><thead className="bg-secondary/55 text-[11px] text-muted-foreground"><tr>{headers.map(h => <th key={h} className="whitespace-nowrap px-4 py-3 font-semibold">{h}</th>)}</tr></thead><tbody className="divide-y divide-border">{children}</tbody></table>; }
function RowEmpty({ colSpan }: { colSpan:number }) { return <tr><td colSpan={colSpan} className="px-4 py-12 text-center text-sm text-muted-foreground">لا توجد بيانات مطابقة</td></tr>; }
function ErrorInline({ text = 'تعذر تحميل البيانات' }: { text?:string }) { return <div className="px-4 py-10 text-center text-sm text-destructive">{text}</div>; }

function InvoiceForm({ kind, onClose }: { kind:'sale'|'purchase'; onClose:()=>void }) {
  const qc = useQueryClient(); const sale = useCreateSale(); const purchase = useCreatePurchase();
  const customers = useListCustomers(); const suppliers = useListSuppliers(); const products = useListProducts(); const materials = useListMaterials();
  const isSale = kind === 'sale';
  const partyRows = isSale ? (customers.data ?? []) : (suppliers.data ?? []);
  const itemRows = isSale ? (products.data ?? []) : (materials.data ?? []);
  const [form,setForm]=useState({invoiceNo:'',date:today(),party:'',paymentMethod:'نقدي',notes:''});
  const [lines,setLines]=useState([{code:'',name:'',quantity:'1',unitPrice:'0'}]);
  const addLine=()=>setLines([...lines,{code:'',name:'',quantity:'1',unitPrice:'0'}]);
  const removeLine=(idx:number)=>setLines(lines.length===1?lines:lines.filter((_,i)=>i!==idx));
  const updateLine=(idx:number,key:string,value:string)=>setLines(lines.map((line,i)=>i===idx?{...line,[key]:value}:line));
  const selectItem=(idx:number,code:string)=>{ const row=itemRows.find(x=>x.code===code); if(!row) return; updateLine(idx,'code',row.code); updateLine(idx,'name',row.name); updateLine(idx,'unitPrice',String(isSale ? ('salePrice' in row ? row.salePrice : 0) : ('cost' in row ? row.cost : 0))); };
  const grandTotal=lines.reduce((sum,l)=>sum+Number(l.quantity||0)*Number(l.unitPrice||0),0);
  const submit=(e:FormEvent)=>{e.preventDefault(); if(!form.invoiceNo.trim()||!form.party||!lines.length||lines.some(l=>!l.code||Number(l.quantity)<=0||Number(l.unitPrice)<0))return;
    const items=lines.map(l=>({code:l.code,name:l.name,quantity:Number(l.quantity),unitPrice:Number(l.unitPrice),total:Number(l.quantity)*Number(l.unitPrice)}));
    const data={invoiceNo:form.invoiceNo,date:form.date,...(isSale?{customerName:form.party}:{supplierName:form.party}),total:grandTotal,paymentMethod:form.paymentMethod,items};
    const done=()=>{qc.invalidateQueries({queryKey:isSale?getListSalesQueryKey():getListPurchasesQueryKey()});qc.invalidateQueries({queryKey:getGetDashboardQueryKey()});qc.invalidateQueries({queryKey:getListCashQueryKey()});qc.invalidateQueries({queryKey:getListJournalsQueryKey()});qc.invalidateQueries({queryKey:getGetIncomeReportQueryKey()});qc.invalidateQueries({queryKey:getGetReconciliationQueryKey()});onClose();};
    if(isSale)sale.mutate({data:data as SaleInput},{onSuccess:done});else purchase.mutate({data:data as PurchaseInput},{onSuccess:done});
  };
  const pending=sale.isPending||purchase.isPending;
  return <Modal title={isSale?'فاتورة مبيعات جديدة':'فاتورة مشتريات جديدة'} onClose={onClose} size="wide">
    <form onSubmit={submit} className="space-y-4">
      <div className="rounded-xl bg-secondary/55 p-3 text-xs leading-6 text-muted-foreground">طريقة الإدخال مطابقة لفكرة شيت الحركة: رأس الفاتورة ثم عدة أصناف في نفس الفاتورة، مع إجمالي تلقائي. الدفع النقدي/الآجل يؤثر على الخزينة والأرصدة.</div>
      <div className="form-section"><div className="form-section-title">بيانات الفاتورة</div><div className="grid gap-3 sm:grid-cols-3"><Field label="رقم الفاتورة" value={form.invoiceNo} onChange={v=>setForm({...form,invoiceNo:v})} required testId="input-invoice-no"/><Field label="التاريخ" type="date" value={form.date} onChange={v=>setForm({...form,date:v})} required testId="input-invoice-date"/><label className="text-sm"><span className="mb-2 block font-semibold">طريقة الدفع</span><select value={form.paymentMethod} onChange={e=>setForm({...form,paymentMethod:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option>نقدي</option><option>آجل</option></select></label></div></div>
      <div className="form-section"><div className="form-section-title">الطرف</div><label className="text-sm"><span className="mb-2 block font-semibold">{isSale?'العميل':'المورد'}<span className="mr-1 text-destructive">*</span></span><select required value={form.party} onChange={e=>setForm({...form,party:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option value="">اختر {isSale?'العميل':'المورد'}</option>{partyRows.map(x=><option key={x.id} value={x.name}>{x.code} — {x.name}</option>)}</select></label></div>
      <div className="rounded-xl border border-border overflow-hidden"><div className="flex items-center justify-between bg-secondary/55 px-3 py-2"><span className="text-sm font-bold">بنود الفاتورة</span><button type="button" onClick={addLine} className="inline-flex items-center gap-1 rounded-lg bg-primary px-3 py-1.5 text-xs font-semibold text-primary-foreground"><Plus className="h-3.5 w-3.5"/>إضافة بند</button></div>
        <div className="divide-y divide-border">{lines.map((line,idx)=><div key={idx} className="grid gap-2 p-3 md:grid-cols-[1.4fr_.55fr_.7fr_.7fr_auto] md:items-end">
          <label className="text-xs"><span className="mb-1 block font-semibold">{isSale?'المنتج':'الخامة'}</span><select required value={line.code} onChange={e=>selectItem(idx,e.target.value)} className="h-9 w-full rounded-lg border border-input bg-background px-2"><option value="">اختيار</option>{itemRows.map(x=><option key={x.id} value={x.code}>{x.code} — {x.name}</option>)}</select></label>
          <Field label="الكمية" type="number" value={line.quantity} onChange={v=>updateLine(idx,'quantity',v)} required testId={`input-invoice-qty-${idx}`}/>
          <Field label={isSale?'سعر البيع':'سعر التكلفة'} type="number" value={line.unitPrice} onChange={v=>updateLine(idx,'unitPrice',v)} required testId={`input-invoice-price-${idx}`}/>
          <div className="rounded-lg border border-primary/20 bg-primary/5 p-2"><p className="text-[10px] text-muted-foreground">الإجمالي</p><p className="font-mono font-bold text-primary">{money(Number(line.quantity||0)*Number(line.unitPrice||0))}</p></div>
          <button type="button" onClick={()=>removeLine(idx)} className="h-9 rounded-lg border border-destructive/20 px-3 text-xs text-destructive">حذف</button>
        </div>)}</div>
      </div>
      <div className="form-section"><div className="form-section-title">ملاحظات الفاتورة</div><Field label="ملاحظات" value={form.notes} onChange={v=>setForm({...form,notes:v})} testId="input-invoice-notes"/> </div>
      <div className="flex items-center justify-between rounded-xl bg-primary/5 p-4"><span className="font-bold">إجمالي الفاتورة</span><span className="font-mono text-xl font-bold text-primary">{money(grandTotal)}</span></div>
      <div className="flex gap-2"><button disabled={pending} className="flex-1 rounded-lg bg-primary py-2.5 text-sm font-semibold text-primary-foreground">{pending?'جاري الحفظ...':'حفظ الفاتورة'}</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-4">إلغاء</button></div>
    </form>
  </Modal>;
}

function Field({ label, value, onChange, type='text', required=false, testId }: {label:string; value:string; onChange:(v:string)=>void; type?:string; required?:boolean; testId:string}) { return <label className="block text-sm"><span className="mb-2 block font-semibold">{label}{required && <span className="mr-1 text-destructive">*</span>}</span><input data-testid={testId} required={required} type={type} value={value} onChange={e=>onChange(e.target.value)} className="h-10 w-full rounded-lg border border-input bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-ring/20" /></label>; }
function Modal({ title, onClose, children, size='xl' }: { title:string; onClose:()=>void; children:ReactNode; size?:'xl'|'wide'|'full' }) {
  const width=size==='full'?'max-w-[1500px]':size==='wide'?'max-w-6xl':'max-w-2xl';
  return <div className="fixed inset-0 z-50 overflow-y-auto bg-foreground/45 p-2 sm:p-4">
    <div className="flex min-h-full items-start justify-center py-2 sm:py-3">
      <div className={`modal-mobile-full w-full ${width} max-h-[calc(100dvh-18px)] overflow-y-auto rounded-2xl border border-card-border bg-card p-4 shadow-2xl sm:max-h-[calc(100dvh-28px)] sm:p-6`}>
        <div className="sticky top-0 z-20 -mx-4 -mt-4 mb-5 flex items-center justify-between border-b border-border bg-card/98 px-4 py-3 backdrop-blur sm:-mx-6 sm:-mt-6 sm:px-6">
          <div><h3 className="text-lg font-bold">{title}</h3><p className="mt-0.5 text-[11px] text-muted-foreground">راجع البيانات والتفاصيل قبل الحفظ أو التصدير</p></div>
          <button onClick={onClose} className="rounded-lg p-2 text-muted-foreground hover:bg-secondary" data-testid="button-close-modal"><X className="h-4 w-4"/></button>
        </div>
        {children}
      </div>
    </div>
  </div>;
}

function Sales() {
  const q=useListSales(); const [open,setOpen]=useState(false); const [search,setSearch]=useState(''); const [selected,setSelected]=useState<any|null>(null);
  const rows=useMemo(()=>q.data?.filter(x=>`${x.invoiceNo} ${x.customerName}`.toLowerCase().includes(search.toLowerCase())) ?? [],[q.data,search]);
  const total=rows.reduce((s,r)=>s+r.total,0), cashTotal=rows.filter(r=>r.paymentMethod==='نقدي').reduce((s,r)=>s+r.total,0), creditTotal=total-cashTotal;
  return <><PageHeading eyebrow="التشغيل / المبيعات" title="فواتير المبيعات" description="سجل احترافي للفواتير مع تفاصيل الأصناف، التحصيل، وحركة العميل." action="فاتورة مبيعات" onAction={()=>setOpen(true)} icon={FilePlus2}/>
    <div className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><Metric label="عدد الفواتير" value={new Intl.NumberFormat('ar-EG').format(rows.length)} icon={ClipboardList}/><Metric label="إجمالي المبيعات" value={money(total)} icon={TrendingUp} tone="primary"/><Metric label="مبيعات نقدية" value={money(cashTotal)} icon={CircleDollarSign} tone="accent"/><Metric label="مبيعات آجلة" value={money(creditTotal)} icon={Users} tone="blue"/></div>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>{q.data && <TableShell title="سجل المبيعات" description="اضغط على أي فاتورة لفتح المعاملة كاملة ومراجعة أصنافها والقيد المحاسبي." search={search} onSearch={setSearch} count={rows.length}><Table headers={['الفاتورة','التاريخ','العميل','عدد البنود','الدفع','الحالة','الإجمالي','إجراء']} >{rows.length ? rows.map(r=><tr key={r.id} onClick={()=>setSelected(r)} tabIndex={0} onKeyDown={e=>{if(e.key==='Enter')setSelected(r)}} className="cursor-pointer hover:bg-primary/5" data-testid={`row-sale-${r.id}`}><td className="px-4 py-3 font-mono text-xs font-semibold">{r.invoiceNo}</td><td className="px-4 py-3 text-muted-foreground">{dateLabel(r.date)}</td><td className="px-4 py-3 font-semibold">{r.customerName}</td><td className="px-4 py-3">{r.items?.length??0}</td><td className="px-4 py-3">{r.paymentMethod}</td><td className="px-4 py-3"><span className="rounded-full bg-primary/10 px-2 py-1 text-[11px] text-primary">{r.status || 'مسجلة'}</span></td><td className="px-4 py-3 font-mono font-semibold">{money(r.total)}</td><td className="px-4 py-3"><button type="button" onClick={e=>{e.stopPropagation();setSelected(r)}} className="inline-flex items-center gap-1 rounded-lg bg-secondary px-2.5 py-1.5 text-xs font-bold"><Eye className="h-3.5 w-3.5"/>فتح</button></td></tr>) : <RowEmpty colSpan={8}/>}</Table></TableShell>}</State>
    {open&&<InvoiceForm kind="sale" onClose={()=>setOpen(false)}/>} {selected&&<TransactionPreview kind="sale" transaction={selected} onClose={()=>setSelected(null)}/>}</>;
}

function Purchases() {
  const q=useListPurchases(); const [open,setOpen]=useState(false); const [search,setSearch]=useState(''); const [selected,setSelected]=useState<any|null>(null);
  const rows=useMemo(()=>q.data?.filter(x=>`${x.invoiceNo} ${x.supplierName}`.toLowerCase().includes(search.toLowerCase()))??[],[q.data,search]);
  const total=rows.reduce((s,r)=>s+r.total,0), cashTotal=rows.filter(r=>r.paymentMethod==='نقدي').reduce((s,r)=>s+r.total,0), creditTotal=total-cashTotal;
  return <><PageHeading eyebrow="التشغيل / المشتريات" title="فواتير المشتريات" description="سجل التوريد مع تفاصيل الخامات، التكلفة، والسداد المرتبط بالخزينة." action="فاتورة مشتريات" onAction={()=>setOpen(true)} icon={ShoppingBag}/>
    <div className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><Metric label="عدد الفواتير" value={new Intl.NumberFormat('ar-EG').format(rows.length)} icon={ClipboardList}/><Metric label="إجمالي المشتريات" value={money(total)} icon={ShoppingBag} tone="primary"/><Metric label="مشتريات نقدية" value={money(cashTotal)} icon={CircleDollarSign} tone="accent"/><Metric label="مشتريات آجلة" value={money(creditTotal)} icon={Users} tone="blue"/></div>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>{q.data&&<TableShell title="سجل المشتريات" description="اضغط على أي فاتورة لفتح المعاملة كاملة ومراجعة الخامات والقيد المحاسبي." search={search} onSearch={setSearch} count={rows.length}><Table headers={['الفاتورة','التاريخ','المورد','عدد البنود','الدفع','الحالة','الإجمالي','إجراء']}>{rows.length?rows.map(r=><tr key={r.id} onClick={()=>setSelected(r)} tabIndex={0} onKeyDown={e=>{if(e.key==='Enter')setSelected(r)}} className="cursor-pointer hover:bg-accent/10" data-testid={`row-purchase-${r.id}`}><td className="px-4 py-3 font-mono text-xs font-semibold">{r.invoiceNo}</td><td className="px-4 py-3 text-muted-foreground">{dateLabel(r.date)}</td><td className="px-4 py-3 font-semibold">{r.supplierName}</td><td className="px-4 py-3">{r.items?.length??0}</td><td className="px-4 py-3">{r.paymentMethod}</td><td className="px-4 py-3"><span className="rounded-full bg-accent/25 px-2 py-1 text-[11px]">{r.status || 'مسجلة'}</span></td><td className="px-4 py-3 font-mono font-semibold">{money(r.total)}</td><td className="px-4 py-3"><button type="button" onClick={e=>{e.stopPropagation();setSelected(r)}} className="inline-flex items-center gap-1 rounded-lg bg-secondary px-2.5 py-1.5 text-xs font-bold"><Eye className="h-3.5 w-3.5"/>فتح</button></td></tr>):<RowEmpty colSpan={8}/>}</Table></TableShell>}</State>
    {open&&<InvoiceForm kind="purchase" onClose={()=>setOpen(false)}/>} {selected&&<TransactionPreview kind="purchase" transaction={selected} onClose={()=>setSelected(null)}/>}</>;
}

function TransactionPreview({kind,transaction,onClose}:{kind:'sale'|'purchase';transaction:any;onClose:()=>void}){
  const journals=useListJournals();
  const isSale=kind==='sale'; const title=`${isSale?'فاتورة مبيعات':'فاتورة مشتريات'} — ${transaction.invoiceNo}`; const party=isSale?transaction.customerName:transaction.supplierName;
  const journal=(journals.data??[]).find((j:any)=>j.source===(isSale?'sale':'purchase') && String(j.description).includes(String(transaction.invoiceNo)));
  const items=(transaction.items??[]).map((item:any)=>[item.code,item.name,item.quantity,money(item.unitPrice),money(item.total)] as ReportCell[]);
  const printRows: ReportCell[][] = [
    ['التاريخ', dateLabel(transaction.date), '', '', ''],
    ['الطرف', party, '', '', ''],
    ['طريقة الدفع', transaction.paymentMethod, '', '', ''],
    ...items,
    ['', '', '', 'إجمالي الفاتورة', money(transaction.total)],
  ];
  const headers=['البيان','القيمة','الكمية','سعر الوحدة','الإجمالي'];
  const print=()=>printReport(title,headers,printRows);
  const exportExcel=()=>exportReportToExcel(title,headers,printRows);
  return <Modal title={title} onClose={onClose} size="full"><div className="space-y-5">
    <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><InfoCard label="الطرف" value={party}/><InfoCard label="التاريخ" value={dateLabel(transaction.date)}/><InfoCard label="طريقة الدفع" value={transaction.paymentMethod}/><InfoCard label="إجمالي المعاملة" value={money(transaction.total)} emphasis/></div>
    <section className="rounded-2xl border border-border bg-background p-4"><div className="mb-3 flex items-center justify-between"><div><h4 className="font-bold">تفاصيل الأصناف</h4><p className="mt-1 text-xs text-muted-foreground">كل بند بالكمية والسعر والإجمالي.</p></div><span className="rounded-full bg-primary/10 px-3 py-1 text-xs font-bold text-primary">{transaction.items?.length??0} بند</span></div><div className="overflow-auto rounded-xl border border-border"><Table headers={['الكود','الصنف / الخدمة','الكمية','سعر الوحدة','الإجمالي']}>{(transaction.items??[]).map((item:any,i:number)=><tr key={i}><td className="px-4 py-3 font-mono">{item.code}</td><td className="px-4 py-3 font-semibold">{item.name}</td><td className="px-4 py-3">{item.quantity}</td><td className="px-4 py-3 font-mono">{money(item.unitPrice)}</td><td className="px-4 py-3 font-mono font-bold text-primary">{money(item.total)}</td></tr>)}<tr className="bg-primary/5"><td colSpan={4} className="px-4 py-3 text-left font-bold">إجمالي الفاتورة</td><td className="px-4 py-3 font-mono text-lg font-bold text-primary">{money(transaction.total)}</td></tr></Table></div></section>
    {transaction.notes&&<section className="rounded-xl border border-accent/30 bg-accent/10 p-4"><h4 className="font-bold">ملاحظات</h4><p className="mt-1 text-sm text-muted-foreground">{transaction.notes}</p></section>}
    <section className="rounded-2xl border border-border bg-card p-4"><div className="mb-3 flex items-center gap-2"><CheckCircle2 className="h-5 w-5 text-primary"/><div><h4 className="font-bold">القيد المحاسبي المرتبط</h4><p className="text-xs text-muted-foreground">المعاملة ليست مجرد فاتورة؛ القيد الناتج عنها ظاهر هنا للمراجعة.</p></div></div>{journal?<div className="grid gap-2 md:grid-cols-2">{journal.lines.map((line:any,i:number)=><div key={i} className="rounded-xl border border-border bg-background p-3"><div className="flex items-center justify-between"><span className="font-semibold">{line.account}</span><span className="font-mono text-sm">{line.debit>0?`مدين ${money(line.debit)}`:`دائن ${money(line.credit)}`}</span></div></div>)}</div>:<p className="text-sm text-muted-foreground">القيد غير ظاهر حاليًا في دفتر اليومية.</p>}</section>
    <div className="flex flex-col-reverse gap-2 border-t border-border pt-4 sm:flex-row sm:justify-end"><button onClick={exportExcel} className="inline-flex items-center justify-center gap-2 rounded-lg bg-primary px-5 py-3 text-sm font-bold text-primary-foreground"><FileSpreadsheet className="h-4 w-4"/>تصدير Excel</button><button onClick={print} className="inline-flex items-center justify-center gap-2 rounded-lg border border-border px-5 py-3 text-sm font-bold hover:bg-secondary"><Printer className="h-4 w-4"/>طباعة المعاملة / حفظ PDF</button><button onClick={onClose} className="rounded-lg bg-primary px-5 py-3 text-sm font-bold text-primary-foreground">إغلاق</button></div>
  </div></Modal>;
}

function InfoCard({label,value,emphasis=false}:{label:string;value:string;emphasis?:boolean}){return <div className={`rounded-xl border border-border p-3 ${emphasis?'bg-primary/5 border-primary/20':''}`}><p className="text-[11px] text-muted-foreground">{label}</p><p className={`mt-1 ${emphasis?'text-lg text-primary':'font-semibold'} font-bold`}>{value}</p></div>}

function Cash() {
  const q=useListCash(); const [open,setOpen]=useState(false); const [search,setSearch]=useState(''); const [selected,setSelected]=useState<any|null>(null);
  const rows=useMemo(()=>q.data?.filter(x=>`${x.description} ${x.category??''} ${x.partyName??''} ${x.beneficiary??''} ${x.voucherNo??''} ${x.accountCode??''} ${x.analysisCode??''}`.toLowerCase().includes(search.toLowerCase()))??[],[q.data,search]);
  const received=rows.filter(r=>r.type==='قبض').reduce((s,r)=>s+r.amount,0), paid=rows.filter(r=>r.type==='صرف').reduce((s,r)=>s+r.amount,0), net=received-paid, closing=rows.at(-1)?.balance??0;
  return <><PageHeading eyebrow="التشغيل / السيولة" title="الخزينة" description="سجل كامل مطابق لشيت حركة الخزينة: السند، مدين/دائن، الرصيد، الكود الرئيسي، جهة الإيراد/الصرف، المستفيد، البيان، وكود تحليل البند." action="حركة جديدة" onAction={()=>setOpen(true)} icon={WalletCards}/>
    <div className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><Metric label="إجمالي القبض" value={money(received)} icon={ArrowDownLeft} tone="primary"/><Metric label="إجمالي الصرف" value={money(paid)} icon={ArrowUpRight} tone="red"/><Metric label="صافي الحركة" value={money(net)} icon={TrendingUp} tone="accent"/><Metric label="الرصيد الحالي" value={money(closing)} icon={WalletCards} tone="blue"/></div>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>{q.data&&<TableShell title="حركات الخزينة" description="اضغط على أي حركة لفتح تفاصيلها كاملة، بما فيها كودي G وK والقيد المحاسبي." search={search} onSearch={setSearch} count={rows.length}><Table headers={['التاريخ','نوع الحركة','رقم السند','مدين','دائن','الرصيد','كود الحساب','جهة الإيراد/الصرف','البند/المستفيد','البيان','كود التحليل','تحليل البند','إجراء']}>{rows.length?rows.map((r:any)=><tr key={r.id} onClick={()=>setSelected(r)} tabIndex={0} onKeyDown={e=>{if(e.key==='Enter')setSelected(r)}} className="cursor-pointer hover:bg-primary/5" data-testid={`row-cash-${r.id}`}><td className="px-3 py-3 text-muted-foreground">{dateLabel(r.date)}</td><td className="px-3 py-3"><span className={`inline-flex items-center gap-1 rounded-full px-2 py-1 text-[11px] ${r.type==='قبض'?'bg-primary/10 text-primary':'bg-destructive/10 text-destructive'}`}>{r.type==='قبض'?<ArrowDownLeft className="h-3 w-3"/>:<ArrowUpRight className="h-3 w-3"/>}{r.type}</span></td><td className="px-3 py-3 font-mono text-xs">{r.voucherNo||'—'}</td><td className="px-3 py-3 font-mono">{r.type==='قبض'?money(r.amount):'—'}</td><td className="px-3 py-3 font-mono">{r.type==='صرف'?money(r.amount):'—'}</td><td className="px-3 py-3 font-mono font-bold">{money(r.balance)}</td><td className="px-3 py-3 font-mono font-semibold">{r.accountCode||'—'}</td><td className="px-3 py-3 text-muted-foreground">{r.accountName||r.category||'—'}</td><td className="px-3 py-3 font-semibold">{r.beneficiary||r.partyName||'—'}</td><td className="px-3 py-3">{r.description}</td><td className="px-3 py-3 font-mono">{r.analysisCode||'—'}</td><td className="px-3 py-3 text-muted-foreground">{r.analysisName||r.analysis||'—'}</td><td className="px-3 py-3"><button type="button" onClick={e=>{e.stopPropagation();setSelected(r)}} className="inline-flex items-center gap-1 rounded-lg bg-secondary px-2.5 py-1.5 text-xs font-bold"><Eye className="h-3.5 w-3.5"/>فتح</button></td></tr>):<RowEmpty colSpan={13}/>}</Table></TableShell>}</State>
    {open&&<CashForm onClose={()=>setOpen(false)}/>} {selected&&<CashPreview entry={selected} onClose={()=>setSelected(null)}/>}</>;
}

function CashPreview({entry,onClose}:{entry:any;onClose:()=>void}){
  const journals=useListJournals(); const journal=(journals.data??[]).find((j:any)=> entry.sourceType==='sale' ? (j.source==='sale' && String(j.description).includes(String(entry.sourceNo||''))) : entry.sourceType==='purchase' ? (j.source==='purchase' && String(j.description).includes(String(entry.sourceNo||''))) : entry.sourceType==='payroll' ? (j.source==='payroll' && String(j.description).includes(String(entry.partyName||''))) : j.source==='cash' && (j.description===entry.description || String(j.description).startsWith(String(entry.description)+' [كود')) && Math.abs(new Date(j.date).getTime()-new Date(entry.date).getTime())<24*60*60*1000 && Math.abs(Number(j.lines?.[0]?.debit||j.lines?.[0]?.credit||0)-entry.amount)<0.01);
  const title=`حركة خزينة — ${entry.description}`; const headers=['البند','القيمة']; const printRows: ReportCell[][] = [['التاريخ',reportDate(entry.date)],['نوع الحركة',entry.type],['رقم السند',entry.voucherNo||'—'],['مدين',entry.type==='قبض'?money(entry.amount):'—'],['دائن',entry.type==='صرف'?money(entry.amount):'—'],['الرصيد بعد الحركة',money(entry.balance)],['كود الحساب',entry.accountCode||'—'],['جهة الإيراد/الصرف',entry.accountName||entry.category||'—'],['البند/المستفيد',entry.beneficiary||entry.partyName||'—'],['البيان',entry.description],['كود التحليل',entry.analysisCode||'—'],['تحليل البند',entry.analysisName||entry.analysis||'—'],['ملاحظات',entry.notes||'—'],['',''],['القيد المحاسبي',''],...(journal?.lines??[]).map((line:any)=>[line.account,line.debit>0?`مدين ${money(line.debit)}`:`دائن ${money(line.credit)}`])]; const print=()=>printReport(title,headers,printRows); const exportExcel=()=>exportReportToExcel(title,headers,printRows);
  return <Modal title={title} onClose={onClose} size="wide"><div className="space-y-4"><div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4"><InfoCard label="التاريخ" value={dateLabel(entry.date)}/><InfoCard label="نوع الحركة" value={entry.type}/><InfoCard label="الطرف" value={entry.partyName||'حركة عامة'}/><InfoCard label="المبلغ" value={money(entry.amount)} emphasis/></div><section className="grid gap-3 sm:grid-cols-2"><div className="form-section"><div className="form-section-title">بيانات الحركة كما في شيت الخزينة</div><div className="space-y-3 text-sm"><div className="flex justify-between"><span className="text-muted-foreground">كود الحساب</span><b className="font-mono">{entry.accountCode||'—'}</b></div><div className="flex justify-between"><span className="text-muted-foreground">جهة الإيراد/الصرف</span><b>{entry.accountName||entry.category||'—'}</b></div><div className="flex justify-between"><span className="text-muted-foreground">البند/المستفيد</span><b>{entry.beneficiary||entry.partyName||'—'}</b></div><div className="flex justify-between"><span className="text-muted-foreground">كود تحليل البند</span><b className="font-mono">{entry.analysisCode||'—'}</b></div><div className="flex justify-between"><span className="text-muted-foreground">تحليل البند</span><b>{entry.analysisName||entry.analysis||'—'}</b></div><div className="flex justify-between"><span className="text-muted-foreground">البيان</span><b>{entry.description}</b></div><div className="flex justify-between"><span className="text-muted-foreground">الرصيد بعد الحركة</span><b className="font-mono text-primary">{money(entry.balance)}</b></div></div></div><div className="form-section"><div className="form-section-title">القيد المحاسبي</div>{journal?<div className="space-y-2">{journal.lines.map((line:any,i:number)=><div key={i} className="rounded-lg border border-border bg-background p-3 flex justify-between"><b>{line.account}</b><span className="font-mono">{line.debit>0?`مدين ${money(line.debit)}`:`دائن ${money(line.credit)}`}</span></div>)}</div>:<p className="text-sm text-destructive">لم يتم العثور على قيد مطابق لهذه الحركة.</p>}</div></section>{entry.notes&&<div className="rounded-xl border border-accent/30 bg-accent/10 p-4 text-sm"><b>ملاحظات:</b> {entry.notes}</div>}<div className="flex flex-col-reverse gap-2 border-t border-border pt-4 sm:flex-row sm:justify-end"><button onClick={exportExcel} className="inline-flex items-center justify-center gap-2 rounded-lg bg-primary px-5 py-3 text-sm font-bold text-primary-foreground"><FileSpreadsheet className="h-4 w-4"/>تصدير Excel</button><button onClick={print} className="inline-flex items-center justify-center gap-2 rounded-lg border border-border px-5 py-3 text-sm font-bold"><Printer className="h-4 w-4"/>طباعة الحركة / حفظ PDF</button><button onClick={onClose} className="rounded-lg bg-primary px-5 py-3 text-sm font-bold text-primary-foreground">إغلاق</button></div></div></Modal>;
}

function CashForm({onClose}:{onClose:()=>void}) {
  const mutation=useCreateCashEntry(); const qc=useQueryClient(); const customers=useListCustomers(); const suppliers=useListSuppliers(); const partners=useListPartners();
  const [form,setForm]=useState({date:today(),type:'قبض' as CashEntryInputType,amount:'',voucherNo:'',accountCode:'',accountName:'',beneficiary:'',analysisCode:'',analysisName:'',description:'',partyType:'عام',partyName:'',category:'',analysis:'',loading:'',notes:''});
  const dynamicCustomers=(customers.data??[]).map(x=>({code:x.code,name:x.name,kind:'customer' as const}));
  const dynamicSuppliers=(suppliers.data??[]).map(x=>({code:x.code,name:x.name,kind:'supplier' as const}));
  const mainOptions=useMemo(()=>Array.from(new Map([...CASH_MAIN_CODES,...dynamicCustomers,...dynamicSuppliers].map(x=>[x.code,x])).values()).sort((a,b)=>a.code.localeCompare(b.code)),[customers.data,suppliers.data]);
  const selectedMain=mainOptions.find(x=>x.code===form.accountCode); const selectedAnalysis=CASH_ANALYSIS_CODES.find(x=>x[0]===form.analysisCode);
  const inferredPartyType=selectedMain?.kind==='customer'?'عميل':selectedMain?.kind==='supplier'?'مورد':form.partyType;
  const setMainCode=(code:string)=>{const row=mainOptions.find(x=>x.code===code); setForm({...form,accountCode:code,accountName:row?.name||'',beneficiary:row?.name||form.beneficiary,partyType:row?.kind==='customer'?'عميل':row?.kind==='supplier'?'مورد':form.partyType,partyName:row?.kind==='customer'||row?.kind==='supplier'?row.name:form.partyName});};
  const setAnalysisCode=(code:string)=>{const row=CASH_ANALYSIS_CODES.find(x=>x[0]===code); setForm({...form,analysisCode:code,analysisName:row?.[1]||'',analysis:row?.[1]||''});};
  const partyRows=inferredPartyType==='عميل'?(customers.data??[]):inferredPartyType==='مورد'?(suppliers.data??[]):inferredPartyType==='شريك'?(partners.data??[]):[];
  const submit=(e:FormEvent)=>{e.preventDefault();if(!form.amount||!form.description.trim()||!form.accountCode||!form.analysisCode)return; mutation.mutate({data:{...form,partyType:inferredPartyType,partyName:form.partyName,category:form.accountName,analysis:form.analysisName,amount:Number(form.amount)} as CashEntryInput},{onSuccess:()=>{qc.invalidateQueries({queryKey:getListCashQueryKey()});qc.invalidateQueries({queryKey:getGetDashboardQueryKey()});qc.invalidateQueries({queryKey:getListJournalsQueryKey()});qc.invalidateQueries({queryKey:getGetIncomeReportQueryKey()});qc.invalidateQueries({queryKey:getGetReconciliationQueryKey()});onClose();}})};
  return <Modal title="تسجيل حركة الخزينة — مطابق للشيت الأصلي" onClose={onClose} size="full"><form onSubmit={submit} className="space-y-4">
    <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
      <div className="form-section"><div className="form-section-title">رقم العملية</div><Field label="رقم السند" value={form.voucherNo} onChange={v=>setForm({...form,voucherNo:v})} testId="input-cash-voucher"/></div>
      <div className="form-section"><div className="form-section-title">التاريخ والنوع</div><div className="grid gap-2 sm:grid-cols-2"><Field label="التاريخ" type="date" value={form.date} onChange={v=>setForm({...form,date:v})} required testId="input-cash-date"/><label className="text-sm"><span className="mb-2 block font-semibold">نوع الحركة</span><select value={form.type} onChange={e=>setForm({...form,type:e.target.value as CashEntryInputType})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option>قبض</option><option>صرف</option></select></label></div></div>
      <div className="form-section"><div className="form-section-title">القيمة</div><Field label={form.type==='قبض'?'المدين / المقبوض':'الدائن / المدفوع'} type="number" value={form.amount} onChange={v=>setForm({...form,amount:v})} required testId="input-cash-amount"/></div>
      <div className="form-section"><div className="form-section-title">الكود الرئيسي</div><label className="text-sm"><span className="mb-2 block font-semibold">الاكواد (G)</span><select required value={form.accountCode} onChange={e=>setMainCode(e.target.value)} className="h-10 w-full rounded-lg border border-primary/30 bg-background px-3"><option value="">اختر الكود</option>{mainOptions.map(x=><option key={x.code} value={x.code}>{x.code} — {x.name}</option>)}</select></label><p className="mt-2 text-[11px] text-muted-foreground">المسمى: {form.accountName||'سيظهر تلقائيًا'}</p></div>
      <div className="form-section"><div className="form-section-title">كود التحليل</div><label className="text-sm"><span className="mb-2 block font-semibold">كود تحليل البند (K)</span><select required value={form.analysisCode} onChange={e=>setAnalysisCode(e.target.value)} className="h-10 w-full rounded-lg border border-primary/30 bg-background px-3"><option value="">اختر التحليل</option>{CASH_ANALYSIS_CODES.map(([code,name])=><option key={code} value={code}>{code} — {name}</option>)}</select></label><p className="mt-2 text-[11px] text-muted-foreground">التحليل: {form.analysisName||'سيظهر تلقائيًا'}</p></div>
    </div>
    <div className="grid gap-4 lg:grid-cols-2">
      <div className="form-section"><div className="form-section-title">جهة الإيراد / الصرف والبند</div><div className="grid gap-3 sm:grid-cols-2"><label className="text-sm"><span className="mb-2 block font-semibold">جهة الإيراد/الصرف (H)</span><input value={form.accountName} readOnly className="h-10 w-full rounded-lg border border-input bg-muted/30 px-3"/></label><label className="text-sm"><span className="mb-2 block font-semibold">البند/المستفيد (I)</span>{partyRows.length ? <select value={form.partyName} onChange={e=>setForm({...form,partyName:e.target.value,beneficiary:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option value="">اختر المستفيد</option>{partyRows.map((x:any)=><option key={x.id} value={x.name}>{x.code} — {x.name}</option>)}</select> : <input value={form.beneficiary} onChange={e=>setForm({...form,beneficiary:e.target.value,partyName:e.target.value})} placeholder="اكتب اسم المستفيد / البند" className="h-10 w-full rounded-lg border border-input bg-background px-3" />}</label></div><div className="mt-3 rounded-lg bg-primary/5 p-3 text-xs leading-6">الطرف المحدد: <b>{inferredPartyType}</b> — {form.partyName||form.beneficiary||'حركة عامة'}</div></div>
      <div className="form-section"><div className="form-section-title">البيان والملاحظات</div><Field label="البيان (J)" value={form.description} onChange={v=>setForm({...form,description:v})} required testId="input-cash-description"/><div className="mt-3"><Field label="ملاحظات إضافية" value={form.notes} onChange={v=>setForm({...form,notes:v})} testId="input-cash-notes"/></div></div>
    </div>
    <div className="rounded-2xl border border-primary/20 bg-gradient-to-l from-primary/10 via-background to-accent/10 p-4"><div className="grid gap-3 sm:grid-cols-3"><InfoCard label="كود الحساب" value={form.accountCode||'—'}/><InfoCard label="كود التحليل" value={form.analysisCode||'—'}/><InfoCard label={form.type==='قبض'?'مدين الخزينة':'دائن الخزينة'} value={form.amount?money(Number(form.amount)):'0 ج.م'} emphasis/></div></div>
    <div className="sticky bottom-0 -mx-1 flex gap-2 border-t border-border bg-card/95 pt-3 backdrop-blur"><button disabled={mutation.isPending} className="flex-1 rounded-lg bg-primary py-3 text-sm font-semibold text-primary-foreground">{mutation.isPending?'جاري الحفظ...':'حفظ الحركة وترحيلها للحسابات'}</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-5">إلغاء</button></div>
  </form></Modal>;
}

function Inventory() {
  const q=useListInventory();
  const movements=useQuery<any[]>({queryKey:['/api/inventory-movements'],queryFn:async()=>{const r=await fetch('/api/inventory-movements');if(!r.ok)throw new Error('تعذر تحميل حركات المخزن');return r.json();}});
  const [open,setOpen]=useState(false); const [search,setSearch]=useState('');
  const rows=useMemo(()=>q.data?.filter(x=>`${x.code} ${x.name}`.toLowerCase().includes(search.toLowerCase()))??[],[q.data,search]);
  const [,navigate]=useLocation();
  return <><PageHeading eyebrow="التشغيل / المخزن" title="المخزن والتكلفة" description="رصيد كل خامة، المشتريات، المنصرف الفعلي للتشغيل، ومتوسط التكلفة مع سجل حركة كامل." action="حركة مخزن" onAction={()=>setOpen(true)} icon={Boxes}/>
    <div className="mb-5 grid gap-3 sm:grid-cols-3">
      <Metric label="قيمة المخزون" value={money(rows.reduce((s,r)=>s+r.balanceValue,0))} note="بسعر التكلفة" icon={Boxes}/>
      <Metric label="تكلفة المنصرف" value={money((movements.data??[]).filter(x=>x.type==='صرف').reduce((s,x)=>s+x.total,0))} note="مواد مباشرة على الإنتاج" icon={ArrowUpRight} tone="accent"/>
      <Metric label="عدد الحركات" value={String(movements.data?.length??0)} note="إضافة / صرف / تسوية" icon={ClipboardList} tone="blue"/>
    </div>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>{q.data&&<div className="space-y-5">
      <TableShell title="أرصدة المخزون" description="المخزون = المشتريات + الإضافات − المنصرف − التسويات الخارجة" search={search} onSearch={setSearch} count={rows.length}><Table headers={['الكود','المادة','الوحدة','المشتريات','المنصرف','الرصيد','متوسط التكلفة','القيمة','إجراء']}>{rows.length?rows.map((r,i)=><tr key={`${r.code}-${i}`}><td className="px-4 py-3 font-mono text-xs">{r.code}</td><td className="px-4 py-3 font-semibold">{r.name}</td><td className="px-4 py-3">{r.unit}</td><td className="px-4 py-3 font-mono">{r.purchasedQty}</td><td className="px-4 py-3 font-mono text-muted-foreground">{r.consumedQty}</td><td className="px-4 py-3 font-mono font-bold text-primary">{r.balanceQty}</td><td className="px-4 py-3 font-mono">{money(r.averageCost)}</td><td className="px-4 py-3 font-mono font-semibold">{money(r.balanceValue)}</td><td className="px-4 py-3"><button onClick={()=>navigate(`/reports?kind=inventory`)} className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary">معاينة</button></td></tr>):<RowEmpty colSpan={9}/>}</Table></TableShell>
      <TableShell title="سجل حركات المخزن" description="كل صرف هنا يدخل مباشرة في تكلفة البضاعة/المواد المباشرة بقائمة الدخل." count={movements.data?.length??0}><Table headers={['التاريخ','النوع','الخامة','الكمية','التكلفة','الإجمالي','السبب','تحليل المصروف']}>{(movements.data??[]).slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))).map((r:any)=><tr key={r.id}><td className="px-4 py-3">{dateLabel(r.date)}</td><td className="px-4 py-3"><span className={`rounded-full px-2 py-1 text-[11px] ${r.type==='صرف'?'bg-destructive/10 text-destructive':'bg-primary/10 text-primary'}`}>{r.type}</span></td><td className="px-4 py-3 font-semibold">{r.materialName}</td><td className="px-4 py-3 font-mono">{r.quantity}</td><td className="px-4 py-3 font-mono">{money(r.unitCost)}</td><td className="px-4 py-3 font-mono font-bold">{money(r.total)}</td><td className="px-4 py-3">{r.reason||'—'}</td><td className="px-4 py-3 text-muted-foreground">{r.analysis||'—'}</td></tr>)}</Table></TableShell>
    </div>}</State>{open&&<InventoryMovementForm onClose={()=>{setOpen(false);q.refetch();movements.refetch();}}/>}</>;
}

function InventoryMovementForm({onClose}:{onClose:()=>void}){
  const materials=useListMaterials(); const qc=useQueryClient();
  const [form,setForm]=useState({date:today(),type:'صرف',materialCode:'',quantity:'',unitCost:'',reason:'صرف خامات للتشغيل',analysis:'تكلفة إنتاج / مواد مباشرة',notes:''});
  const selected=materials.data?.find(x=>x.code===form.materialCode);
  const submit=async(e:FormEvent)=>{e.preventDefault();if(!form.materialCode||!form.quantity)return;const res=await fetch('/api/inventory-movements',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...form,quantity:Number(form.quantity),unitCost:Number(form.unitCost||selected?.cost||0)})});if(!res.ok){alert((await res.json()).error||'تعذر حفظ الحركة');return;}await qc.invalidateQueries({queryKey:getListMaterialsQueryKey()});await qc.invalidateQueries({queryKey:getGetDashboardQueryKey()});await qc.invalidateQueries({queryKey:getListJournalsQueryKey()});await qc.invalidateQueries({queryKey:getGetIncomeReportQueryKey()});await qc.invalidateQueries({queryKey:getGetReconciliationQueryKey()});onClose();};
  return <Modal title="تسجيل حركة مخزن" onClose={onClose} size="wide"><form onSubmit={submit} className="space-y-4">
    <div className="form-section"><div className="form-section-title">بيانات الحركة</div><div className="grid gap-3 md:grid-cols-3"><Field label="التاريخ" type="date" value={form.date} onChange={v=>setForm({...form,date:v})} required testId="input-inv-date"/><label className="text-sm"><span className="mb-2 block font-semibold">نوع الحركة</span><select value={form.type} onChange={e=>setForm({...form,type:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option>صرف</option><option>إضافة</option><option>تسوية</option></select></label><div className="rounded-lg border border-primary/20 bg-primary/5 px-3 py-2"><p className="text-[10px] text-muted-foreground">القيمة التقديرية</p><p className="mt-1 font-mono font-bold text-primary">{money(Number(form.quantity||0)*Number(form.unitCost||selected?.cost||0))}</p></div></div></div>
    <div className="form-section"><div className="form-section-title">الصنف والكمية والتكلفة</div><label className="text-sm"><span className="mb-2 block font-semibold">الخامة</span><select required value={form.materialCode} onChange={e=>setForm({...form,materialCode:e.target.value,unitCost:String(materials.data?.find(x=>x.code===e.target.value)?.cost??'')})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option value="">اختر الخامة</option>{(materials.data??[]).map(x=><option key={x.id} value={x.code}>{x.code} — {x.name}</option>)}</select></label><div className="mt-3 grid gap-3 sm:grid-cols-2"><Field label="الكمية" type="number" value={form.quantity} onChange={v=>setForm({...form,quantity:v})} required testId="input-inv-qty"/><Field label="تكلفة الوحدة" type="number" value={form.unitCost} onChange={v=>setForm({...form,unitCost:v})} required testId="input-inv-cost"/></div></div>
    <div className="form-section"><div className="form-section-title">السبب والتحليل</div><div className="space-y-3"><Field label="السبب" value={form.reason} onChange={v=>setForm({...form,reason:v})} required testId="input-inv-reason"/><Field label="تحليل المصروف" value={form.analysis} onChange={v=>setForm({...form,analysis:v})} testId="input-inv-analysis"/><Field label="ملاحظات" value={form.notes} onChange={v=>setForm({...form,notes:v})} testId="input-inv-notes"/></div></div>
    <div className="rounded-xl border border-primary/15 bg-primary/5 p-3 text-xs text-muted-foreground">في حالة «صرف» يتم تحميل القيمة تلقائيًا على تكلفة المواد المباشرة في قائمة الدخل، وتظهر الحركة في تقرير المخزون والحسابات.</div><div className="sticky bottom-0 -mx-1 flex gap-2 border-t border-border bg-card/95 pt-3 backdrop-blur"><button className="flex-1 rounded-lg bg-primary py-3 text-sm font-bold text-primary-foreground">حفظ حركة المخزن وترحيل التكلفة</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-5">إلغاء</button></div></form></Modal>;
}

function Employees(){
  const employees=useQuery<any[]>({queryKey:['/api/employees'],queryFn:async()=>{const r=await fetch('/api/employees');if(!r.ok)throw new Error('تعذر تحميل الموظفين');return r.json();}});
  const payroll=useQuery<any[]>({queryKey:['/api/payroll'],queryFn:async()=>{const r=await fetch('/api/payroll');if(!r.ok)throw new Error('تعذر تحميل المرتبات');return r.json();}});
  const [tab,setTab]=useState<'employees'|'payroll'>('employees'); const [open,setOpen]=useState(false); const [selectedPayroll,setSelectedPayroll]=useState<any|null>(null);
  return <><PageHeading eyebrow="التشغيل / الموارد البشرية" title="الموظفون والمرتبات" description="دليل الموظفين وكشف المرتبات الشهري، مع ربط المدفوع فعليًا بالخزينة." action={tab==='employees'?'إضافة موظف':'إضافة مسير مرتب'} onAction={()=>setOpen(true)} icon={BriefcaseBusiness}/><div className="mb-5 flex gap-1 rounded-xl border border-border bg-card p-1"><button onClick={()=>setTab('employees')} className={`flex-1 rounded-lg px-4 py-2 text-sm ${tab==='employees'?'bg-primary text-primary-foreground font-semibold':''}`}>سجل الموظفين</button><button onClick={()=>setTab('payroll')} className={`flex-1 rounded-lg px-4 py-2 text-sm ${tab==='payroll'?'bg-primary text-primary-foreground font-semibold':''}`}>المرتبات</button></div><State loading={employees.isLoading||payroll.isLoading} error={employees.isError||payroll.isError} retry={()=>{employees.refetch();payroll.refetch();}}>{tab==='employees'?<TableShell title="سجل الموظفين" description="البيانات الأساسية والراتب الأساسي" count={employees.data?.length??0}><Table headers={['الكود','الموظف','الوظيفة','الهاتف','تاريخ التعيين','الراتب','الحالة']}>{(employees.data??[]).map(x=><tr key={x.id}><td className="px-4 py-3 font-mono">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3">{x.job||'—'}</td><td className="px-4 py-3">{x.phone||'—'}</td><td className="px-4 py-3">{x.hireDate||'—'}</td><td className="px-4 py-3 font-mono font-bold">{money(x.salary)}</td><td className="px-4 py-3">{x.status}</td></tr>)}</Table></TableShell>:<TableShell title="كشف المرتبات" description="المرتب الصافي = الأساسي + الإضافات − الخصومات — اضغط على أي مسير للمعاينة." count={payroll.data?.length??0}><Table headers={['الشهر','التاريخ','الموظف','الأساسي','إضافات','خصومات','الصافي','الحالة','طريقة الدفع','إجراء']}>{(payroll.data??[]).map(x=><tr key={x.id} onClick={()=>setSelectedPayroll(x)} tabIndex={0} onKeyDown={e=>{if(e.key==='Enter')setSelectedPayroll(x)}} className="cursor-pointer hover:bg-primary/5"><td className="px-4 py-3 font-semibold">{x.month}</td><td className="px-4 py-3">{dateLabel(x.date)}</td><td className="px-4 py-3">{x.employeeName}</td><td className="px-4 py-3 font-mono">{money(x.basicSalary)}</td><td className="px-4 py-3 font-mono">{money(x.additions)}</td><td className="px-4 py-3 font-mono">{money(x.deductions)}</td><td className="px-4 py-3 font-mono font-bold">{money(x.netSalary)}</td><td className="px-4 py-3">{x.paymentStatus}</td><td className="px-4 py-3">{x.paymentMethod}</td><td className="px-4 py-3"><button type="button" onClick={e=>{e.stopPropagation();setSelectedPayroll(x)}} className="inline-flex items-center gap-1 rounded-lg bg-secondary px-2.5 py-1.5 text-xs font-bold"><Eye className="h-3.5 w-3.5"/>فتح</button></td></tr>)}</Table></TableShell>}</State>{open&&(tab==='employees'?<EmployeeForm onClose={()=>{setOpen(false);employees.refetch();}}/>:<PayrollForm employees={employees.data??[]} onClose={()=>{setOpen(false);payroll.refetch();}}/> )}{selectedPayroll&&<PayrollPreview item={selectedPayroll} onClose={()=>setSelectedPayroll(null)}/>}</>;
}

function PayrollPreview({item,onClose}:{item:any;onClose:()=>void}){
  const journals=useListJournals();
  const journal=(journals.data??[]).find((j:any)=>j.source==='payroll' && String(j.description).includes(String(item.employeeName||'')) && String(j.description).includes(String(item.month||'')));
  const rows:ReportCell[][]=[['الشهر',item.month],['تاريخ الصرف',reportDate(item.date)],['الموظف',item.employeeName],['الراتب الأساسي',money(item.basicSalary)],['الإضافات',money(item.additions)],['الخصومات',money(item.deductions)],['صافي المرتب',money(item.netSalary)],['حالة السداد',item.paymentStatus],['طريقة الدفع',item.paymentMethod],['ملاحظات',item.notes||'—'],['',''],['القيد المحاسبي',''],...(journal?.lines??[]).map((line:any)=>[line.account,line.debit>0?`مدين ${money(line.debit)}`:`دائن ${money(line.credit)}`])];
  return <Modal title={`مسير مرتب — ${item.employeeName} — ${item.month}`} onClose={onClose} size="full"><div className="space-y-4"><div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4"><InfoCard label="الموظف" value={item.employeeName}/><InfoCard label="الشهر" value={item.month}/><InfoCard label="الحالة" value={item.paymentStatus}/><InfoCard label="صافي المرتب" value={money(item.netSalary)} emphasis/></div><section className="form-section"><div className="form-section-title">تفاصيل المسير</div><div className="overflow-auto rounded-xl border border-border"><Table headers={['البند','القيمة']}><tr><td className="px-4 py-3">الراتب الأساسي</td><td className="px-4 py-3 font-mono">{money(item.basicSalary)}</td></tr><tr><td className="px-4 py-3">إضافات</td><td className="px-4 py-3 font-mono">{money(item.additions)}</td></tr><tr><td className="px-4 py-3">خصومات</td><td className="px-4 py-3 font-mono">{money(item.deductions)}</td></tr><tr className="bg-primary/5"><td className="px-4 py-3 font-bold">الصافي</td><td className="px-4 py-3 font-mono text-lg font-bold text-primary">{money(item.netSalary)}</td></tr></Table></div></section><section className="form-section"><div className="form-section-title">القيد المحاسبي</div>{journal?<div className="space-y-2">{journal.lines.map((line:any,i:number)=><div key={i} className="flex items-center justify-between rounded-lg border border-border bg-background p-3"><b>{line.account}</b><span className="font-mono">{line.debit>0?`مدين ${money(line.debit)}`:`دائن ${money(line.credit)}`}</span></div>)}</div>:<p className="text-sm text-destructive">لم يتم العثور على قيد محاسبي مرتبط.</p>}</section><div className="flex flex-col-reverse gap-2 border-t border-border pt-4 sm:flex-row sm:justify-end"><button onClick={()=>printReport(`مسير مرتب — ${item.employeeName} — ${item.month}`,['البند','القيمة'],rows)} className="inline-flex items-center justify-center gap-2 rounded-lg border border-border px-5 py-3 text-sm font-bold"><Printer className="h-4 w-4"/>طباعة / حفظ PDF</button><button onClick={onClose} className="rounded-lg bg-primary px-5 py-3 text-sm font-bold text-primary-foreground">إغلاق</button></div></div></Modal>;
}

function EmployeeForm({onClose}:{onClose:()=>void}){const [f,setF]=useState({code:'',name:'',job:'',phone:'',hireDate:today(),salary:''});const submit=async(e:FormEvent)=>{e.preventDefault();const r=await fetch('/api/employees',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...f,salary:Number(f.salary||0),status:'نشط'})});if(!r.ok){alert((await r.json()).error||'تعذر الحفظ');return;}onClose();};return <Modal title="إضافة موظف" onClose={onClose}><form onSubmit={submit} className="space-y-4"><div className="grid gap-3 sm:grid-cols-2"><Field label="كود الموظف" value={f.code} onChange={v=>setF({...f,code:v})} required testId="input-employee-code"/><Field label="اسم الموظف" value={f.name} onChange={v=>setF({...f,name:v})} required testId="input-employee-name"/></div><div className="grid gap-3 sm:grid-cols-2"><Field label="الوظيفة" value={f.job} onChange={v=>setF({...f,job:v})} testId="input-employee-job"/><Field label="الهاتف" value={f.phone} onChange={v=>setF({...f,phone:v})} testId="input-employee-phone"/></div><div className="grid gap-3 sm:grid-cols-2"><Field label="تاريخ التعيين" type="date" value={f.hireDate} onChange={v=>setF({...f,hireDate:v})} testId="input-employee-hire"/><Field label="الراتب الأساسي" type="number" value={f.salary} onChange={v=>setF({...f,salary:v})} required testId="input-employee-salary"/></div><div className="flex gap-2"><button className="flex-1 rounded-lg bg-primary py-2.5 text-sm font-semibold text-primary-foreground">حفظ الموظف</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-4">إلغاء</button></div></form></Modal>}
function PayrollForm({employees,onClose}:{employees:any[];onClose:()=>void}){
  const [f,setF]=useState({date:today(),month:today().slice(0,7),employeeId:'',basicSalary:'',additions:'0',deductions:'0',paymentStatus:'مدفوع',paymentMethod:'نقدي',notes:''});
  const selected=employees.find(x=>String(x.id)===f.employeeId);
  const net=Number(f.basicSalary||selected?.salary||0)+Number(f.additions||0)-Number(f.deductions||0);
  const submit=async(e:FormEvent)=>{e.preventDefault();const r=await fetch('/api/payroll',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...f,employeeId:Number(f.employeeId),basicSalary:Number(f.basicSalary||selected?.salary||0),additions:Number(f.additions||0),deductions:Number(f.deductions||0)})});if(!r.ok){alert((await r.json()).error||'تعذر حفظ المرتب');return;}onClose();};
  return <Modal title="إضافة مسير مرتب شهري" onClose={onClose} size="wide"><form onSubmit={submit} className="space-y-4">
    <div className="form-section"><div className="form-section-title">الفترة والموظف</div><div className="grid gap-3 md:grid-cols-3">
      <Field label="الشهر" type="month" value={f.month} onChange={v=>setF({...f,month:v})} required testId="input-payroll-month"/>
      <Field label="تاريخ الصرف" type="date" value={f.date} onChange={v=>setF({...f,date:v})} required testId="input-payroll-date"/>
      <label className="text-sm"><span className="mb-2 block font-semibold">الموظف<span className="mr-1 text-destructive">*</span></span><select required value={f.employeeId} onChange={e=>setF({...f,employeeId:e.target.value,basicSalary:String(employees.find(x=>String(x.id)===e.target.value)?.salary??'')})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option value="">اختر الموظف</option>{employees.map(x=><option key={x.id} value={x.id}>{x.code} — {x.name}</option>)}</select></label>
    </div></div>
    <div className="form-section"><div className="form-section-title">استحقاقات المرتب</div><div className="grid gap-3 md:grid-cols-4">
      <Field label="الراتب الأساسي" type="number" value={f.basicSalary} onChange={v=>setF({...f,basicSalary:v})} required testId="input-payroll-basic"/>
      <Field label="إضافات" type="number" value={f.additions} onChange={v=>setF({...f,additions:v})} testId="input-payroll-additions"/>
      <Field label="خصومات" type="number" value={f.deductions} onChange={v=>setF({...f,deductions:v})} testId="input-payroll-deductions"/>
      <div className="rounded-lg border border-primary/20 bg-primary/5 px-4 py-2.5"><p className="text-xs text-muted-foreground">صافي المرتب</p><p className="mt-1 font-mono text-xl font-bold text-primary">{money(net)}</p></div>
    </div></div>
    <div className="form-section"><div className="form-section-title">السداد والحالة</div><div className="grid gap-3 md:grid-cols-2">
      <label className="text-sm"><span className="mb-2 block font-semibold">حالة السداد</span><select value={f.paymentStatus} onChange={e=>setF({...f,paymentStatus:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option>مدفوع</option><option>غير مدفوع</option></select></label>
      <label className="text-sm"><span className="mb-2 block font-semibold">طريقة الدفع</span><select value={f.paymentMethod} onChange={e=>setF({...f,paymentMethod:e.target.value})} className="h-10 w-full rounded-lg border border-input bg-background px-3"><option>نقدي</option><option>تحويل بنكي</option></select></label>
    </div></div>
    <div className="form-section"><div className="form-section-title">ملاحظات</div>
    <Field label="ملاحظات" value={f.notes} onChange={v=>setF({...f,notes:v})} testId="input-payroll-notes"/>
    </div>
    <div className="sticky bottom-0 -mx-1 flex flex-col-reverse gap-2 border-t border-border bg-card/95 pt-3 backdrop-blur sm:flex-row"><button className="flex-1 rounded-lg bg-primary py-3 text-sm font-bold text-primary-foreground">حفظ المسير وربطه بالمصروفات</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-6 py-3">إلغاء</button></div>
  </form></Modal>
}
function Masters() { const [tab,setTab]=useState<'customers'|'suppliers'|'products'|'materials'>('customers'); const [open,setOpen]=useState(false); const tabs=[['customers','العملاء'],['suppliers','الموردون'],['products','المنتجات'],['materials','المواد'] ] as const; return <><PageHeading eyebrow="البيانات الأساسية" title="التعريفات" description="مصدر واحد موثوق للأطراف والأصناف المستخدمة في التشغيل." action={`إضافة ${tabs.find(x=>x[0]===tab)?.[1]}`} onAction={()=>setOpen(true)} icon={Tag}/><div className="mb-5 flex gap-1 overflow-x-auto rounded-xl border border-border bg-card p-1 shadow-sm">{tabs.map(([key,label])=><button key={key} onClick={()=>setTab(key)} data-testid={`button-master-tab-${key}`} className={`whitespace-nowrap rounded-lg px-4 py-2 text-sm ${tab===key?'bg-primary text-primary-foreground font-semibold':'text-muted-foreground hover:bg-secondary'}`}>{label}</button>)}</div>{tab==='customers'?<MasterCustomers/>:tab==='suppliers'?<MasterSuppliers/>:tab==='products'?<MasterProducts/>:<MasterMaterials/>}{open&&<MasterForm kind={tab} onClose={()=>setOpen(false)}/>}</>; }
function MasterCustomers(){const q=useListCustomers(); const [,navigate]=useLocation(); return <MasterTable title="سجل العملاء" rows={q.data} loading={q.isLoading} error={q.isError} cols={['الكود','الاسم','الهاتف','العنوان','الرصيد الافتتاحي','إجراء']} render={(x)=><><td className="px-4 py-3 font-mono text-xs">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3">{x.phone||'—'}</td><td className="px-4 py-3 text-muted-foreground">{x.address||'—'}</td><td className="px-4 py-3 font-mono">{money(x.openingBalance)}</td><td className="px-4 py-3"><button onClick={()=>navigate(`/customer-statement?kind=customer&name=${encodeURIComponent(x.name)}`)} className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary">كشف حساب</button></td></>}/>}; function MasterSuppliers(){const q=useListSuppliers(); const [,navigate]=useLocation(); return <MasterTable title="سجل الموردين" rows={q.data} loading={q.isLoading} error={q.isError} cols={['الكود','الاسم','الهاتف','العنوان','الرصيد الافتتاحي','إجراء']} render={(x)=><><td className="px-4 py-3 font-mono text-xs">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3">{x.phone||'—'}</td><td className="px-4 py-3 text-muted-foreground">{x.address||'—'}</td><td className="px-4 py-3 font-mono">{money(x.openingBalance)}</td><td className="px-4 py-3"><button onClick={()=>navigate(`/customer-statement?kind=supplier&name=${encodeURIComponent(x.name)}`)} className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary">كشف حساب</button></td></>}/>}; function MasterProducts(){const q=useListProducts(); return <MasterTable title="المنتجات والخدمات" rows={q.data} loading={q.isLoading} error={q.isError} cols={['الكود','الاسم','الوحدة','سعر البيع']} render={(x)=><><td className="px-4 py-3 font-mono text-xs">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3">{x.unit}</td><td className="px-4 py-3 font-mono font-semibold">{money(x.salePrice)}</td></>}/>}; function MasterMaterials(){const q=useListMaterials(); return <MasterTable title="مواد التشغيل" rows={q.data} loading={q.isLoading} error={q.isError} cols={['الكود','المادة','الوحدة','سعر التكلفة']} render={(x)=><><td className="px-4 py-3 font-mono text-xs">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3">{x.unit}</td><td className="px-4 py-3 font-mono font-semibold">{money(x.cost)}</td></>}/>}; 
function MasterTable({title,rows,loading,error,cols,render}:{title:string;rows?:any[];loading:boolean;error:boolean;cols:string[];render:(x:any)=>ReactNode}){return <State loading={loading} error={error} empty={!loading&&!error&&!rows?.length}><TableShell title={title} count={rows?.length}><Table headers={cols}>{rows?.map((x:any)=><tr key={x.id} data-testid={`row-master-${x.id}`}>{render(x)}</tr>)}</Table></TableShell></State>}
function MasterForm({kind,onClose}:{kind:'customers'|'suppliers'|'products'|'materials';onClose:()=>void}){const qc=useQueryClient(); const customer=useCreateCustomer();const supplier=useCreateSupplier();const product=useCreateProduct();const material=useCreateMaterial();const [form,setForm]=useState({code:'',name:'',phone:'',address:'',unit:'قطعة',price:'',balance:''});const party=kind==='customers'||kind==='suppliers';const submit=(e:FormEvent)=>{e.preventDefault();if(!form.code.trim()||!form.name.trim())return;const done=()=>{qc.invalidateQueries({queryKey:kind==='customers'?getListCustomersQueryKey():kind==='suppliers'?getListSuppliersQueryKey():kind==='products'?getListProductsQueryKey():getListMaterialsQueryKey()});onClose()};if(kind==='customers')customer.mutate({data:{code:form.code,name:form.name,phone:form.phone,address:form.address,openingBalance:Number(form.balance||0)} as CustomerInput},{onSuccess:done});else if(kind==='suppliers')supplier.mutate({data:{code:form.code,name:form.name,phone:form.phone,address:form.address,openingBalance:Number(form.balance||0)} as CustomerInput},{onSuccess:done});else if(kind==='products')product.mutate({data:{code:form.code,name:form.name,unit:form.unit,salePrice:Number(form.price||0)} as ProductInput},{onSuccess:done});else material.mutate({data:{code:form.code,name:form.name,unit:form.unit,cost:Number(form.price||0)} as MaterialInput},{onSuccess:done})};const pending=customer.isPending||supplier.isPending||product.isPending||material.isPending;return <Modal title={`إضافة ${kind==='customers'?'عميل':kind==='suppliers'?'مورد':kind==='products'?'منتج':'مادة'}`} onClose={onClose}><form onSubmit={submit} className="space-y-4"><div className="grid gap-3 sm:grid-cols-2"><Field label="الكود" value={form.code} onChange={v=>setForm({...form,code:v})} required testId="input-master-code"/><Field label="الاسم" value={form.name} onChange={v=>setForm({...form,name:v})} required testId="input-master-name"/></div>{party&&<div className="grid gap-3 sm:grid-cols-2"><Field label="الهاتف" value={form.phone} onChange={v=>setForm({...form,phone:v})} testId="input-master-phone"/><Field label="الرصيد الافتتاحي" type="number" value={form.balance} onChange={v=>setForm({...form,balance:v})} testId="input-master-balance"/></div>}{party?<Field label="العنوان" value={form.address} onChange={v=>setForm({...form,address:v})} testId="input-master-address"/>:<div className="grid gap-3 sm:grid-cols-2"><Field label="الوحدة" value={form.unit} onChange={v=>setForm({...form,unit:v})} required testId="input-master-unit"/><Field label={kind==='products'?'سعر البيع':'سعر التكلفة'} type="number" value={form.price} onChange={v=>setForm({...form,price:v})} required testId="input-master-price"/></div>}<div className="flex gap-2"><button disabled={pending} className="flex-1 rounded-lg bg-primary py-2.5 text-sm font-semibold text-primary-foreground" data-testid="button-submit-master">{pending?'جاري الحفظ...':'حفظ السجل'}</button><button type="button" onClick={onClose} className="rounded-lg border border-border px-4" data-testid="button-cancel-master">إلغاء</button></div></form></Modal>}

function Partners(){
  const q=useListPartners();
  const summaryQ = useQuery<{ months: Array<{ month: string; total: number; partners: Record<string, number> }>; grandTotal: number }>({
    queryKey: ['/api/profit-allocations/summary'],
    queryFn: async () => {
      const res = await fetch('/api/profit-allocations/summary');
      if (!res.ok) throw new Error('تعذر جلب ملخص الأرباح');
      return res.json();
    },
  });
  const [, navigate] = useLocation();

  return <>
    <PageHeading eyebrow="الحسابات / رأس المال" title="الشركاء ورأس المال" description="التزامات الشركاء، المسدد والمتبقي، ونسبة الشراكة (33.33%) مع كشوفات الأرباح الشهرية المعتمدة." action="مطابقة شيتات الإكسيل" onAction={() => navigate('/excel')} icon={Users}/>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>
      {q.data&&<div className="space-y-6">
        <TableShell title="كشف الشركاء ورؤوس الأموال" description="مطابق لشيت الشركاء والافتتاحي في ملف مشروع ماكينة التطريز.xlsx" count={q.data.length}>
          <Table headers={['الشريك','حصة رأس المال','المسدد نقداً','المتبقي','نسبة الشراكة','إجمالي الأرباح الموزعة','الرصيد الجاري الدائن']}>
            {q.data.map(x=><tr key={x.id} data-testid={`row-partner-${x.id}`}>
              <td className="px-4 py-3 font-bold text-foreground">{x.name}</td>
              <td className="px-4 py-3 font-mono">{money(x.totalShare)}</td>
              <td className="px-4 py-3 font-mono text-primary font-bold">{money(x.paid)}</td>
              <td className="px-4 py-3 font-mono text-muted-foreground">{money(x.remaining)}</td>
              <td className="px-4 py-3 font-mono font-semibold">33.33%</td>
              <td className="px-4 py-3 font-mono font-bold text-primary">{money(x.profitShare)}</td>
              <td className="px-4 py-3 font-mono font-bold text-foreground">{money(x.paid + x.profitShare)}</td>
            </tr>)}
          </Table>
        </TableShell>

        {summaryQ.data?.months && (
          <TableShell title="جدول توزيعات الأرباح الشهرية (أبريل — أغسطس)" description="المسجلة في قسم الإجماليات بشيت الإكسيل (15 حركة)" count={summaryQ.data.months.length}>
            <Table headers={['الشهر المحاسبي','سيفن ام (33.33%)','احمد العبسي (33.33%)','محمد علاء (33.33%)','إجمالي الشهر']}>
              {summaryQ.data.months.map((m, idx) => (
                <tr key={idx}>
                  <td className="px-4 py-3 font-bold">{m.month}</td>
                  <td className="px-4 py-3 font-mono font-semibold">{money(m.partners['سيفن ام'])}</td>
                  <td className="px-4 py-3 font-mono font-semibold">{money(m.partners['احمد العبسي'])}</td>
                  <td className="px-4 py-3 font-mono font-semibold">{money(m.partners['محمد علاء'])}</td>
                  <td className="px-4 py-3 font-mono font-bold text-primary">{money(m.total)}</td>
                </tr>
              ))}
            </Table>
          </TableShell>
        )}
      </div>}
    </State>
  </>;
}

function Assets(){
  const q=useListAssets();
  const [toggleRate, setToggleRate] = useState<'accounting'|'cashflow'>('accounting');
  const [, navigate] = useLocation();

  return <>
    <PageHeading eyebrow="الحسابات / الأصول" title="الأصول والإهلاك" description="القيمة الدفترية للأصول المستخدمة في الورشة بعد الإهلاك المتراكم مع مطابقة قسط الماكينة." action="مطابقة شيتات الإكسيل" onAction={() => navigate('/excel')} icon={Building2}/>
    <div className="mb-5 rounded-xl border border-accent/30 bg-accent/10 p-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 text-xs">
      <div>
        <span className="font-bold text-foreground">ملاحظة مطابقة قسط ماكينة التطريز:</span>
        <span className="mr-1 text-muted-foreground">القسط الدفتري المعتمد 8,125 ج/شهر (10 سنوات) مقابل 7,500 ج/شهر في جدول التدفقات النقدية (فارق 625 ج/شهر).</span>
      </div>
      <div className="flex items-center gap-1 rounded-lg border border-border bg-card p-1 shrink-0">
        <button onClick={() => setToggleRate('accounting')} className={`rounded px-2.5 py-1 text-xs font-semibold ${toggleRate==='accounting'?'bg-primary text-primary-foreground shadow-sm':'text-muted-foreground'}`}>قسط دفتري (8,125 ج)</button>
        <button onClick={() => setToggleRate('cashflow')} className={`rounded px-2.5 py-1 text-xs font-semibold ${toggleRate==='cashflow'?'bg-primary text-primary-foreground shadow-sm':'text-muted-foreground'}`}>قسط تدفقات (7,500 ج)</button>
      </div>
    </div>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!q.isLoading&&!q.isError&&!q.data?.length}>
      {q.data&&<TableShell title="سجل الأصول الرأسمالية والإهلاك" count={q.data.length}>
        <Table headers={['الأصل','قيمة الشراء','العمر الإنتاجي','إهلاك شهري','إهلاك متراكم','القيمة الدفترية']}>
          {q.data.map(x=>{
            const isMachine = x.name.includes('ماكينة');
            const monthlyDep = isMachine && toggleRate === 'cashflow' ? 7500 : x.monthlyDepreciation;
            return (
              <tr key={x.id} data-testid={`row-asset-${x.id}`}>
                <td className="px-4 py-3 font-semibold">{x.name}</td>
                <td className="px-4 py-3 font-mono">{money(x.purchaseValue)}</td>
                <td className="px-4 py-3">{x.usefulLifeMonths} شهر</td>
                <td className="px-4 py-3 font-mono font-bold text-primary">{money(monthlyDep)}</td>
                <td className="px-4 py-3 font-mono text-muted-foreground">{money(x.accumulatedDepreciation)}</td>
                <td className="px-4 py-3 font-mono font-bold text-primary">{money(x.bookValue)}</td>
              </tr>
            );
          })}
        </Table>
      </TableShell>}
    </State>
  </>;
}

function Accounting(){const journals=useListJournals();const rec=useGetReconciliation();const income=useGetIncomeReport();return <><PageHeading eyebrow="الحسابات / المراجعة" title="المحاسبة والتسويات" description="دفتر اليومية، اختبار التوازن، وصافي الدخل في شاشة واحدة." icon={Calculator}/><div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4"><Metric label="الإيرادات" value={money(income.data?.revenue)} icon={ArrowUpRight}/><Metric label="التكاليف المباشرة" value={money(income.data?.directCosts)} icon={ArrowDownLeft} tone="accent"/><Metric label="المصروفات" value={money(income.data?.expenses)} icon={FileBarChart} tone="blue"/><Metric label="صافي الدخل" value={money(income.data?.netIncome)} icon={BarChart3} tone="primary"/></div><div className="mt-5 grid gap-5 xl:grid-cols-[1.35fr_.65fr]"><State loading={journals.isLoading} error={journals.isError} retry={()=>journals.refetch()} empty={!journals.isLoading&&!journals.isError&&!journals.data?.length}>{journals.data&&<TableShell title="دفتر اليومية" description="القيد المتوازن يعرض مديناً ودائناً لكل حساب" count={journals.data.length}><Table headers={['التاريخ','البيان','المصدر','سطور القيد','الإجمالي']}>{journals.data.map(x=><tr key={x.id} data-testid={`row-journal-${x.id}`}><td className="px-4 py-3 text-muted-foreground">{dateLabel(x.date)}</td><td className="px-4 py-3 font-semibold">{x.description}</td><td className="px-4 py-3">{x.source||'يدوي'}</td><td className="px-4 py-3">{x.lines.length}</td><td className="px-4 py-3 font-mono">{money(x.lines.reduce((sum,l)=>sum+l.debit,0))}</td></tr>)}</Table></TableShell>}</State><section className="rounded-xl border border-card-border bg-card p-5 shadow-sm"><div className="mb-5 flex items-center justify-between"><div><h3 className="font-bold">حالة التسوية</h3><p className="mt-1 text-xs text-muted-foreground">فحوصات آلية على الدفاتر</p></div><span className={`rounded-full px-2 py-1 text-[11px] ${rec.data?.overallStatus==='سليم'?'bg-primary/10 text-primary':'bg-accent/30'}`}>{rec.data?.overallStatus||'قيد الفحص'}</span></div>{rec.data?.checks?.length?<div className="space-y-3">{rec.data.checks.map((x,i)=><div key={i} className="rounded-lg border border-border p-3" data-testid={`row-reconciliation-${i}`}><div className="flex items-center justify-between"><span className="text-sm font-semibold">{x.label}</span><span className={`text-[11px] ${x.status==='سليم'?'text-primary':x.status==='خطأ'?'text-destructive':'text-foreground'}`}>{x.status}</span></div><p className="mt-1 text-xs text-muted-foreground">{x.note}</p></div>)}</div>:<p className="text-sm text-muted-foreground">لا توجد نتائج تسوية.</p>}</section></div></>}

type ReportKind = 'sales'|'purchases'|'cash'|'customers'|'suppliers'|'inventory'|'inventoryMovements'|'payroll'|'partners'|'assets'|'depreciation'|'journal'|'cashflow'|'income'|'trialBalance'|'balanceSheet'|'reconciliation';
const reportLabels: Record<ReportKind, string> = {
  sales: 'كشف المبيعات',
  purchases: 'كشف المشتريات',
  cash: 'كشف الخزينة',
  customers: 'كشف أرصدة العملاء',
  suppliers: 'كشف أرصدة الموردين',
  inventory: 'كشف أرصدة المخزون',
  inventoryMovements: 'حركة المخزن',
  payroll: 'كشف المرتبات',
  partners: 'كشف الشركاء ورأس المال',
  assets: 'كشف الأصول والإهلاك',
  depreciation: 'جدول الإهلاك',
  journal: 'دفتر اليومية',
  cashflow: 'قائمة التدفقات النقدية',
  income: 'قائمة الدخل',
  trialBalance: 'ميزان المراجعة',
  balanceSheet: 'قائمة المركز المالي',
  reconciliation: 'تقرير التسويات',
};
const reportDate = (value: Date | string) => value instanceof Date ? value.toISOString().slice(0, 10) : String(value).slice(0, 10);
const inReportPeriod = (value: Date | string, from: string, to: string) => {
  const date = reportDate(value);
  return (!from || date >= from) && (!to || date <= to);
};

function PartyBalances(){
  const [,navigate]=useLocation();
  const q=useQuery<any>({queryKey:['/api/party-balances'],queryFn:async()=>{const r=await fetch('/api/party-balances');if(!r.ok)throw new Error('تعذر تحميل الأرصدة');return r.json();}});
  return <><PageHeading eyebrow="التقارير التشغيلية" title="أرصدة العملاء والموردين" description="كشف سريع لكل طرف مع الرصيد الحالي والدخول المباشر لكشف الحساب." icon={ClipboardList}/>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()}>
      <div className="grid gap-5 xl:grid-cols-2">
        <TableShell title="أرصدة العملاء" description="مدين = مستحق على العميل" count={q.data?.customers?.length}><Table headers={['الكود','العميل','إجمالي المدين','إجمالي الدائن','الرصيد','إجراء']}>{(q.data?.customers??[]).map((x:any)=><tr key={x.id}><td className="px-4 py-3 font-mono">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3 font-mono">{money(x.totalDebit)}</td><td className="px-4 py-3 font-mono">{money(x.totalCredit)}</td><td className={`px-4 py-3 font-mono font-bold ${x.closingBalance>0?'text-destructive':'text-primary'}`}>{money(x.closingBalance)}</td><td className="px-4 py-3"><button onClick={()=>navigate(`/customer-statement?kind=customer&name=${encodeURIComponent(x.name)}`)} className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary">كشف حساب</button></td></tr>)}</Table></TableShell>
        <TableShell title="أرصدة الموردين" description="الرصيد الدائن يمثل المستحق للمورد" count={q.data?.suppliers?.length}><Table headers={['الكود','المورد','إجمالي المدين','إجمالي الدائن','الرصيد','إجراء']}>{(q.data?.suppliers??[]).map((x:any)=><tr key={x.id}><td className="px-4 py-3 font-mono">{x.code}</td><td className="px-4 py-3 font-semibold">{x.name}</td><td className="px-4 py-3 font-mono">{money(x.totalDebit)}</td><td className="px-4 py-3 font-mono">{money(x.totalCredit)}</td><td className={`px-4 py-3 font-mono font-bold ${x.closingBalance<0?'text-destructive':'text-primary'}`}>{money(x.closingBalance)}</td><td className="px-4 py-3"><button onClick={()=>navigate(`/customer-statement?kind=supplier&name=${encodeURIComponent(x.name)}`)} className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary">كشف حساب</button></td></tr>)}</Table></TableShell>
      </div>
    </State></>;
}

function PartyStatement(){
  const customers=useListCustomers(); const suppliers=useListSuppliers();
  const params=new URLSearchParams(window.location.search); const initialKind=params.get('kind')==='supplier'?'supplier':'customer';
  const [kind,setKind]=useState<'customer'|'supplier'>(initialKind); const [name,setName]=useState(params.get('name')||''); const [preview,setPreview]=useState(false);
  const q=useQuery<any>({queryKey:['/api/party-statement',kind,name],enabled:!!name,queryFn:async()=>{const r=await fetch(`/api/party-statement?kind=${kind}&name=${encodeURIComponent(name)}`);if(!r.ok)throw new Error('تعذر تحميل كشف الحساب');return r.json();}});
  const options=kind==='customer'?(customers.data??[]):(suppliers.data??[]);
  const changeKind=(v:'customer'|'supplier')=>{setKind(v);setName('');};
  const title=`كشف حساب ${kind==='customer'?'عميل':'مورد'}${name?` — ${name}`:''}`;
  const headers=['التاريخ','المصدر','البيان','مدين','دائن','الرصيد'];
  const rows=(q.data?.rows??[]).map((r:any)=>[dateLabel(r.date),r.source,r.description,money(r.debit),money(r.credit),money(r.balance)]);
  const exportExcel=()=>exportReportToExcel(title,headers,rows); const print=()=>printReport(title,headers,rows);
  return <><PageHeading eyebrow="التقارير التشغيلية" title="كشف حساب طرف" description="كشف تفصيلي كامل للعميل أو المورد مع الرصيد الافتتاحي والحركة والرصيد التراكمي." icon={FileBarChart}/>
    <section className="mb-5 rounded-xl border border-card-border bg-card p-4 shadow-sm"><div className="grid gap-3 md:grid-cols-[.8fr_1.8fr_auto] md:items-end"><label className="text-sm"><span className="mb-2 block font-semibold">نوع الطرف</span><select value={kind} onChange={e=>changeKind(e.target.value as 'customer'|'supplier')} className="h-11 w-full rounded-lg border border-input bg-background px-3"><option value="customer">عميل</option><option value="supplier">مورد</option></select></label><label className="text-sm"><span className="mb-2 block font-semibold">اختر الطرف</span><select value={name} onChange={e=>setName(e.target.value)} className="h-11 w-full rounded-lg border border-input bg-background px-3"><option value="">اختر {kind==='customer'?'العميل':'المورد'}</option>{options.map(x=><option key={x.id} value={x.name}>{x.code} — {x.name}</option>)}</select></label><div className="flex gap-2"><button disabled={!name} onClick={()=>setPreview(true)} className="inline-flex h-11 items-center gap-2 rounded-lg bg-primary px-4 text-sm font-bold text-primary-foreground disabled:opacity-40"><Eye className="h-4 w-4"/>معاينة</button><button disabled={!name} onClick={exportExcel} className="inline-flex h-11 items-center gap-2 rounded-lg border border-border px-4 text-sm font-bold disabled:opacity-40"><FileSpreadsheet className="h-4 w-4"/>Excel</button><button disabled={!name} onClick={print} className="inline-flex h-11 items-center gap-2 rounded-lg border border-border px-4 text-sm font-bold disabled:opacity-40"><Printer className="h-4 w-4"/>PDF</button></div></div></section>
    <State loading={q.isLoading} error={q.isError} retry={()=>q.refetch()} empty={!name}>{q.data&&<div className="space-y-5"><div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4"><Metric label="الرصيد الافتتاحي" value={money(q.data.openingBalance)} icon={ClipboardList}/><Metric label="إجمالي المدين" value={money(q.data.totalDebit)} icon={ArrowUpRight}/><Metric label="إجمالي الدائن" value={money(q.data.totalCredit)} icon={ArrowDownLeft}/><Metric label="الرصيد الحالي" value={money(q.data.closingBalance)} icon={WalletCards}/></div><TableShell title={title} description="كل حركة بالتاريخ والمصدر والبيان والمدين والدائن والرصيد التراكمي." count={q.data.rows.length}><Table headers={headers}>{q.data.rows.map((r:any,i:number)=><tr key={i}><td className="px-4 py-3">{dateLabel(r.date)}</td><td className="px-4 py-3 font-semibold">{r.source}</td><td className="px-4 py-3">{r.description}</td><td className="px-4 py-3 font-mono">{money(r.debit)}</td><td className="px-4 py-3 font-mono">{money(r.credit)}</td><td className="px-4 py-3 font-mono font-bold">{money(r.balance)}</td></tr>)}</Table></TableShell></div>}</State>
    {preview&&q.data&&<Modal title={title} onClose={()=>setPreview(false)} size="full"><div className="mb-4 flex flex-wrap gap-2"><button onClick={exportExcel} className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-bold text-primary-foreground"><FileSpreadsheet className="h-4 w-4"/>تصدير Excel</button><button onClick={print} className="inline-flex items-center gap-2 rounded-lg border border-border px-4 py-2 text-sm font-bold"><Printer className="h-4 w-4"/>طباعة / حفظ PDF</button></div><div className="grid gap-3 sm:grid-cols-4 mb-5"><Metric label="افتتاحي" value={money(q.data.openingBalance)} icon={ClipboardList}/><Metric label="مدين" value={money(q.data.totalDebit)} icon={ArrowUpRight}/><Metric label="دائن" value={money(q.data.totalCredit)} icon={ArrowDownLeft}/><Metric label="ختامي" value={money(q.data.closingBalance)} icon={WalletCards}/></div><div className="overflow-auto rounded-xl border border-border"><Table headers={headers}>{q.data.rows.map((r:any,i:number)=><tr key={i}><td className="px-4 py-3">{dateLabel(r.date)}</td><td className="px-4 py-3">{r.source}</td><td className="px-4 py-3">{r.description}</td><td className="px-4 py-3 font-mono">{money(r.debit)}</td><td className="px-4 py-3 font-mono">{money(r.credit)}</td><td className="px-4 py-3 font-mono font-bold">{money(r.balance)}</td></tr>)}</Table></div></Modal>}
  </>;
}
function Reports() {
  const [kind, setKind] = useState<ReportKind>('sales');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [preview, setPreview] = useState(false);
  const sales = useListSales();
  const purchases = useListPurchases();
  const cash = useListCash();
  const inventory = useListInventory();
  const inventoryMovements = useQuery<any[]>({queryKey:['/api/inventory-movements'],queryFn:async()=>{const r=await fetch('/api/inventory-movements');if(!r.ok)throw new Error('تعذر تحميل حركات المخزن');return r.json();}});
  const payroll = useQuery<any[]>({queryKey:['/api/payroll'],queryFn:async()=>{const r=await fetch('/api/payroll');if(!r.ok)throw new Error('تعذر تحميل المرتبات');return r.json();}});
  const partners = useListPartners();
  const assets = useListAssets();
  const journals = useListJournals();
  const income = useGetIncomeReport();
  const reconciliation = useGetReconciliation();
  const loading = sales.isLoading || purchases.isLoading || cash.isLoading || inventory.isLoading || inventoryMovements.isLoading || payroll.isLoading || partners.isLoading || assets.isLoading || journals.isLoading || income.isLoading || reconciliation.isLoading;
  const error = sales.isError || purchases.isError || cash.isError || inventory.isError || inventoryMovements.isError || payroll.isError || partners.isError || assets.isError || journals.isError || income.isError || reconciliation.isError;

  const { headers, rows } = useMemo<{ headers: string[]; rows: ReportCell[][] }>(() => {
    if (kind === 'sales') return {
      headers: ['الفاتورة','التاريخ','العميل','كود الصنف','الصنف / الخدمة','الكمية','سعر الوحدة','إجمالي البند','طريقة الدفع'],
      rows: (sales.data ?? []).filter((row) => inReportPeriod(row.date, from, to)).flatMap((row:any) => (row.items?.length ? row.items : [{code:'',name:'',quantity:'',unitPrice:0,total:row.total}]).map((item:any) => [row.invoiceNo, reportDate(row.date), row.customerName, item.code, item.name, item.quantity, money(item.unitPrice), money(item.total), row.paymentMethod])),
    };
    if (kind === 'purchases') return {
      headers: ['الفاتورة','التاريخ','المورد','كود الخامة','الخامة','الكمية','سعر التكلفة','إجمالي البند','طريقة الدفع'],
      rows: (purchases.data ?? []).filter((row) => inReportPeriod(row.date, from, to)).flatMap((row:any) => (row.items?.length ? row.items : [{code:'',name:'',quantity:'',unitPrice:0,total:row.total}]).map((item:any) => [row.invoiceNo, reportDate(row.date), row.supplierName, item.code, item.name, item.quantity, money(item.unitPrice), money(item.total), row.paymentMethod])),
    };
    if (kind === 'cash') return {
      headers: ['رقم العملية','التاريخ','رقم السند','مدين','دائن','الرصيد','كود الحساب','جهة الإيراد/الصرف','البند/المستفيد','البيان','كود تحليل البند','تحليل البند','ملاحظات'],
      rows: (cash.data ?? []).filter((row) => inReportPeriod(row.date, from, to)).map((row:any) => [row.id, reportDate(row.date), row.voucherNo || '—', money(row.type === 'قبض' ? row.amount : 0), money(row.type === 'صرف' ? row.amount : 0), money(row.balance), row.accountCode || '—', row.accountName || row.category || '—', row.beneficiary || row.partyName || '—', row.description, row.analysisCode || '—', row.analysisName || row.analysis || '—', row.notes || '—']),
    };
    if (kind === 'customers') {
      const grouped = new Map<string, { invoices: number; sales: number; cash: number; credit: number }>();
      for (const row of (sales.data ?? []).filter((item) => inReportPeriod(item.date, from, to))) {
        const current = grouped.get(row.customerName) ?? { invoices: 0, sales: 0, cash: 0, credit: 0 };
        current.invoices += 1;
        current.sales += row.total;
        if (row.paymentMethod === 'نقدي') current.cash += row.total; else current.credit += row.total;
        grouped.set(row.customerName, current);
      }
      return {
        headers: ['العميل', 'عدد الفواتير', 'إجمالي المبيعات', 'مبيعات نقدية', 'الرصيد الآجل'],
        rows: [...grouped.entries()].sort((a, b) => b[1].credit - a[1].credit).map(([name, row]) => [name, row.invoices, money(row.sales), money(row.cash), money(row.credit)]),
      };
    }
    if (kind === 'suppliers') {
      const grouped = new Map<string, { invoices: number; purchases: number; paid: number; payable: number }>();
      for (const row of (purchases.data ?? []).filter((item) => inReportPeriod(item.date, from, to))) {
        const current = grouped.get(row.supplierName) ?? { invoices: 0, purchases: 0, paid: 0, payable: 0 };
        current.invoices += 1;
        current.purchases += row.total;
        if (row.paymentMethod === 'نقدي') current.paid += row.total; else current.payable += row.total;
        grouped.set(row.supplierName, current);
      }
      return {
        headers: ['المورد', 'عدد الفواتير', 'إجمالي المشتريات', 'مشتريات نقدية', 'الرصيد المستحق'],
        rows: [...grouped.entries()].sort((a, b) => b[1].payable - a[1].payable).map(([name, row]) => [name, row.invoices, money(row.purchases), money(row.paid), money(row.payable)]),
      };
    }
    if (kind === 'inventory') return {
      headers: ['الكود', 'المادة', 'الوحدة', 'المشتريات', 'المستهلك', 'الرصيد', 'متوسط التكلفة', 'القيمة'],
      rows: (inventory.data ?? []).map((row) => [row.code, row.name, row.unit, row.purchasedQty, row.consumedQty, row.balanceQty, money(row.averageCost), money(row.balanceValue)]),
    };
    if (kind === 'inventoryMovements') return {
      headers: ['التاريخ','النوع','الخامة','الكمية','تكلفة الوحدة','الإجمالي','السبب','تحليل المصروف'],
      rows: (inventoryMovements.data ?? []).filter((row:any)=>inReportPeriod(row.date, from, to)).map((row:any) => [reportDate(row.date), row.type, row.materialName, row.quantity, money(row.unitCost), money(row.total), row.reason, row.analysis]),
    };
    if (kind === 'payroll') return {
      headers: ['الشهر','التاريخ','الموظف','الأساسي','الإضافات','الخصومات','الصافي','الحالة'],
      rows: (payroll.data ?? []).filter((row:any)=>inReportPeriod(row.date, from, to)).map((row:any) => [row.month, reportDate(row.date), row.employeeName, money(row.basicSalary), money(row.additions), money(row.deductions), money(row.netSalary), row.paymentStatus]),
    };
    if (kind === 'partners') return {
      headers: ['الشريك', 'إجمالي الحصة', 'المدفوع', 'المتبقي', 'توزيعات الأرباح'],
      rows: (partners.data ?? []).map((row) => [row.name, money(row.totalShare), money(row.paid), money(row.remaining), money(row.profitShare)]),
    };
    if (kind === 'assets') return {
      headers: ['الأصل', 'قيمة الشراء', 'العمر بالأشهر', 'الإهلاك الشهري', 'الإهلاك المتراكم', 'القيمة الدفترية'],
      rows: (assets.data ?? []).map((row) => [row.name, money(row.purchaseValue), row.usefulLifeMonths, money(row.monthlyDepreciation), money(row.accumulatedDepreciation), money(row.bookValue)]),
    };
    if (kind === 'depreciation') return {
      headers: ['الأصل', 'قسط الشهر', 'الإهلاك المتراكم', 'القيمة الدفترية', 'العمر المتبقي التقريبي'],
      rows: (assets.data ?? []).map((row) => [row.name, money(row.monthlyDepreciation), money(row.accumulatedDepreciation), money(row.bookValue), `${Math.max(0, row.usefulLifeMonths - Math.ceil(row.accumulatedDepreciation / Math.max(row.monthlyDepreciation, 1)))} شهر`]),
    };
    if (kind === 'journal') return {
      headers: ['التاريخ', 'البيان', 'المصدر', 'سطور القيد', 'مدين', 'دائن'],
      rows: (journals.data ?? []).filter((row) => inReportPeriod(row.date, from, to)).map((row) => [reportDate(row.date), row.description, row.source || 'يدوي', row.lines.length, money(row.lines.reduce((sum, line) => sum + line.debit, 0)), money(row.lines.reduce((sum, line) => sum + line.credit, 0))]),
    };
    if (kind === 'cashflow') {
      const grouped = new Map<string, { receipts: number; payments: number }>();
      for (const row of (cash.data ?? []).filter((item) => inReportPeriod(item.date, from, to))) {
        const month = reportDate(row.date).slice(0, 7);
        const current = grouped.get(month) ?? { receipts: 0, payments: 0 };
        if (row.type === 'قبض') current.receipts += row.amount; else current.payments += row.amount;
        grouped.set(month, current);
      }
      return {
        headers: ['الشهر', 'المقبوضات', 'المدفوعات', 'صافي التدفق'],
        rows: [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([month, row]) => [month, money(row.receipts), money(row.payments), money(row.receipts - row.payments)]),
      };
    }
    if (kind === 'income') {
      const revenue = (sales.data ?? []).filter((row) => inReportPeriod(row.date, from, to)).reduce((sum, row) => sum + row.total, 0);
      const directCosts = (inventoryMovements.data ?? []).filter((row:any) => row.type === 'صرف' && inReportPeriod(row.date, from, to)).reduce((sum:number, row:any) => sum + row.total, 0);
      const expenses = (cash.data ?? []).filter((row) => row.type === 'صرف' && inReportPeriod(row.date, from, to) && !(row.category ?? '').includes('شراء اصول') && !(row.category ?? '').includes('مشتريات')).reduce((sum, row) => sum + row.amount, 0);
      const depreciation = from || to ? 0 : (income.data?.depreciation ?? 0);
      const netIncome = revenue - directCosts - expenses - depreciation;
      return { headers: ['البند', 'القيمة'], rows: [['الإيرادات', money(revenue)], ['تكلفة المواد المباشرة', money(directCosts)], ['المصروفات التشغيلية والنقدية', money(expenses)], ['الإهلاك', money(depreciation)], ['صافي الدخل', money(netIncome)]] };
    }
    if (kind === 'trialBalance') {
      const accountMap = new Map<string, { debit: number; credit: number }>();
      (journals.data ?? []).filter((j:any) => inReportPeriod(j.date, from, to)).forEach((j:any) => j.lines.forEach((line:any) => {
        const item = accountMap.get(line.account) ?? { debit: 0, credit: 0 };
        item.debit += Number(line.debit || 0); item.credit += Number(line.credit || 0); accountMap.set(line.account, item);
      }));
      const rows = Array.from(accountMap.entries()).map(([account, v]) => {
        const debit = Math.max(v.debit - v.credit, 0);
        const credit = Math.max(v.credit - v.debit, 0);
        return [account, money(debit), money(credit), debit ? 'مدين' : credit ? 'دائن' : 'متزن'];
      }).sort((a,b) => String(a[0]).localeCompare(String(b[0]), 'ar'));
      const totalDebit = accountMap.size ? Array.from(accountMap.values()).reduce((s,v)=>s+Math.max(v.debit-v.credit,0),0) : 0;
      const totalCredit = accountMap.size ? Array.from(accountMap.values()).reduce((s,v)=>s+Math.max(v.credit-v.debit,0),0) : 0;
      rows.push(['الإجمالي', money(totalDebit), money(totalCredit), Math.abs(totalDebit-totalCredit)<0.01 ? 'متوازن' : 'يحتاج مراجعة']);
      return { headers: ['الحساب', 'مدين', 'دائن', 'طبيعة الرصيد'], rows };
    }
    if (kind === 'balanceSheet') {
      const cashBalance = cash.data?.at(-1)?.balance ?? 0;
      const inventoryValue = inventory.data?.reduce((sum, row) => sum + row.balanceValue, 0) ?? 0;
      const assetsValue = assets.data?.reduce((sum, row) => sum + row.bookValue, 0) ?? 0;
      const receivables = sales.data?.filter((row) => row.paymentMethod !== 'نقدي').reduce((sum, row) => sum + row.total, 0) ?? 0;
      const capital = partners.data?.reduce((sum, row) => sum + row.totalShare, 0) ?? 0;
      const netIncome = income.data?.netIncome ?? 0;
      return {
        headers: ['القسم', 'البند', 'القيمة'],
        rows: [
          ['الأصول', 'الخزينة', money(cashBalance)],
          ['الأصول', 'العملاء', money(receivables)],
          ['الأصول', 'المخزون', money(inventoryValue)],
          ['الأصول', 'الأصول الثابتة', money(assetsValue)],
          ['حقوق الملكية', 'رأس مال الشركاء', money(capital)],
          ['حقوق الملكية', 'صافي الدخل', money(netIncome)],
          ['إجمالي الأصول', 'الإجمالي', money(cashBalance + receivables + inventoryValue + assetsValue)],
          ['إجمالي الحقوق', 'الإجمالي', money(capital + netIncome)],
        ],
      };
    }
    return {
      headers: ['الفحص', 'الحالة', 'القيمة', 'الملاحظة'],
      rows: (reconciliation.data?.checks ?? []).map((row) => [row.label, row.status, row.value, row.note]),
    };
  }, [kind, from, to, sales.data, purchases.data, cash.data, inventory.data, inventoryMovements.data, payroll.data, partners.data, assets.data, journals.data, income.data, reconciliation.data]);

  const title = reportLabels[kind];
  const reportPeriod = from || to ? `${from || 'البداية'} — ${to || 'النهاية'}` : 'كل الفترات';
  const headlineTotal = kind==='sales' ? (sales.data ?? []).filter(r=>inReportPeriod(r.date,from,to)).reduce((s,r)=>s+r.total,0) : kind==='purchases' ? (purchases.data ?? []).filter(r=>inReportPeriod(r.date,from,to)).reduce((s,r)=>s+r.total,0) : kind==='cash' ? (cash.data ?? []).filter(r=>inReportPeriod(r.date,from,to)).reduce((s,r)=>s+(r.type==='قبض'?r.amount:-r.amount),0) : null;
  const exportExcel = () => exportReportToExcel(title, headers, rows);
  const print = () => printReport(`${title}${from || to ? ` — ${from || 'البداية'} إلى ${to || 'النهاية'}` : ''}`, headers, rows);
  return <><PageHeading eyebrow="الحسابات / المخرجات" title="التقارير والتصدير" description="مركز تقارير موحد: اختر الكشف، حدد الفترة، راجع المعاينة ثم اطبعه أو صدّره." action="تصدير Excel" onAction={exportExcel} icon={FileBarChart}/>
    <section className="mb-5 rounded-2xl border border-card-border bg-card p-4 shadow-sm">
      <div className="grid gap-3 lg:grid-cols-[1.6fr_1fr_1fr_auto] lg:items-end">
        <label className="text-sm"><span className="mb-2 block font-semibold">نوع التقرير</span><select data-testid="select-report-kind" value={kind} onChange={(event) => setKind(event.target.value as ReportKind)} className="h-11 w-full rounded-lg border border-input bg-background px-3"><optgroup label="التشغيل"><option value="sales">كشف المبيعات</option><option value="purchases">كشف المشتريات</option><option value="cash">كشف الخزينة</option><option value="inventory">كشف أرصدة المخزون</option><option value="inventoryMovements">حركة المخزن</option><option value="payroll">كشف المرتبات</option></optgroup><optgroup label="العملاء والموردين"><option value="customers">كشف أرصدة العملاء</option><option value="suppliers">كشف أرصدة الموردين</option></optgroup><optgroup label="الحسابات"><option value="journal">دفتر اليومية</option><option value="income">قائمة الدخل</option><option value="trialBalance">ميزان المراجعة</option><option value="balanceSheet">قائمة المركز المالي</option><option value="cashflow">قائمة التدفقات النقدية</option><option value="partners">كشف الشركاء ورأس المال</option><option value="assets">كشف الأصول والإهلاك</option><option value="depreciation">جدول الإهلاك</option><option value="reconciliation">تقرير التسويات</option></optgroup></select></label>
        <label className="text-sm"><span className="mb-2 flex items-center gap-2 font-semibold"><CalendarDays className="h-4 w-4 text-primary"/>من تاريخ</span><input data-testid="input-report-from" type="date" value={from} onChange={(event) => setFrom(event.target.value)} className="h-11 w-full rounded-lg border border-input bg-background px-3"/></label>
        <label className="text-sm"><span className="mb-2 flex items-center gap-2 font-semibold"><CalendarDays className="h-4 w-4 text-primary"/>إلى تاريخ</span><input data-testid="input-report-to" type="date" value={to} onChange={(event) => setTo(event.target.value)} className="h-11 w-full rounded-lg border border-input bg-background px-3"/></label>
        <button onClick={() => { setFrom(''); setTo(''); }} className="h-11 rounded-lg border border-border px-4 text-sm font-semibold text-muted-foreground hover:bg-secondary">مسح الفترة</button>
      </div>
      <div className="mt-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <InfoCard label="التقرير" value={title}/><InfoCard label="الفترة" value={reportPeriod}/><InfoCard label="عدد السجلات" value={new Intl.NumberFormat('ar-EG').format(rows.length)}/><InfoCard label={kind==='cash'?'صافي حركة الخزينة':'إجمالي التقرير'} value={headlineTotal===null?'—':money(headlineTotal)} emphasis={headlineTotal!==null}/>
      </div>
      <div className="mt-4 flex flex-wrap items-center gap-2 border-t border-border pt-4">
        <button onClick={()=>setPreview(true)} className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2.5 text-sm font-bold text-primary-foreground"><Eye className="h-4 w-4"/>معاينة كاملة</button>
        <button onClick={exportExcel} className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-4 py-2.5 text-sm font-bold"><FileSpreadsheet className="h-4 w-4"/>تصدير Excel</button>
        <button onClick={print} className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-4 py-2.5 text-sm font-bold hover:bg-secondary"><Printer className="h-4 w-4"/>طباعة / حفظ PDF</button>
        <span className="mr-auto text-xs text-muted-foreground">المعاينة تعرض كامل السجل قبل التصدير.</span>
      </div>
    </section>
    <State loading={loading} error={error} retry={() => { sales.refetch(); purchases.refetch(); cash.refetch(); inventory.refetch(); inventoryMovements.refetch(); payroll.refetch(); partners.refetch(); assets.refetch(); journals.refetch(); income.refetch(); reconciliation.refetch(); }} empty={!loading && !error && !rows.length}>
      <section className="report-paper overflow-hidden">
        <div className="flex flex-col gap-3 border-b border-border bg-gradient-to-l from-primary/10 via-card to-accent/10 p-5 sm:flex-row sm:items-center sm:justify-between"><div><p className="text-[11px] font-bold tracking-wide text-primary">تقرير رسمي</p><h3 className="mt-1 text-xl font-bold">{title}</h3><p className="mt-1 text-xs text-muted-foreground">الفترة: {reportPeriod}</p></div><div className="rounded-xl border border-primary/20 bg-background px-4 py-3 text-center"><p className="text-[10px] text-muted-foreground">عدد السجلات</p><p className="mt-1 font-mono text-lg font-bold text-primary">{new Intl.NumberFormat('ar-EG').format(rows.length)}</p></div></div>
        <div className="overflow-x-auto"><table className="w-full min-w-[760px] text-right text-sm"><thead className="bg-primary text-primary-foreground"><tr>{headers.map((header) => <th key={header} className="whitespace-nowrap px-4 py-3 text-xs font-bold">{header}</th>)}</tr></thead><tbody className="divide-y divide-border">{rows.map((row, index) => <tr key={`${kind}-${index}`} className={`${index%2===0?'bg-card':'bg-secondary/25'} hover:bg-primary/5`}>{row.map((cell, cellIndex) => <td key={`${index}-${cellIndex}`} className={`px-4 py-3 ${cellIndex > 0 && (typeof cell === 'number' || String(cell).includes('ج.م')) ? 'font-mono' : ''}`}>{cellTextForReport(cell)}</td>)}</tr>)}</tbody></table></div>
      </section>
    </State>
    {preview&&<Modal title={`معاينة — ${title}`} onClose={()=>setPreview(false)} size="full"><div className="report-paper overflow-hidden">
      <div className="flex flex-col gap-4 border-b border-border bg-gradient-to-l from-primary/15 via-card to-accent/15 p-5 sm:flex-row sm:items-center sm:justify-between"><div><p className="text-[11px] font-bold text-primary">المعاينة النهائية</p><h2 className="mt-1 text-2xl font-bold">{title}</h2><p className="mt-1 text-xs text-muted-foreground">الفترة: {reportPeriod} — النظام يعرض كامل التفاصيل قبل التصدير.</p></div><div className="flex flex-wrap gap-2"><button onClick={exportExcel} className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2.5 text-sm font-bold text-primary-foreground"><FileSpreadsheet className="h-4 w-4"/>Excel</button><button onClick={print} className="inline-flex items-center gap-2 rounded-lg border border-border px-4 py-2.5 text-sm font-bold"><Printer className="h-4 w-4"/>طباعة / حفظ PDF</button></div></div>
      <div className="grid gap-3 border-b border-border p-5 sm:grid-cols-3 xl:grid-cols-4"><InfoCard label="الفترة" value={reportPeriod}/><InfoCard label="عدد السجلات" value={new Intl.NumberFormat('ar-EG').format(rows.length)}/><InfoCard label={kind==='cash'?'صافي الخزينة':'إجمالي التقرير'} value={headlineTotal===null?'—':money(headlineTotal)} emphasis={headlineTotal!==null}/><InfoCard label="حالة المعاينة" value="جاهز للتصدير" emphasis/></div>
      <div className="max-h-[62dvh] overflow-auto p-4 sm:p-5"><table className="w-full min-w-[950px] text-right text-sm"><thead className="sticky top-0 z-10 bg-primary text-primary-foreground"><tr>{headers.map((header)=><th key={header} className="whitespace-nowrap px-4 py-3 text-xs font-bold">{header}</th>)}</tr></thead><tbody className="divide-y divide-border">{rows.map((row,index)=><tr key={index} className={index%2===0?'bg-card':'bg-secondary/25'}>{row.map((cell,cellIndex)=><td key={cellIndex} className="whitespace-nowrap px-4 py-3">{cellTextForReport(cell)}</td>)}</tr>)}</tbody></table></div>
      <div className="border-t border-border bg-secondary/30 p-4 text-xs text-muted-foreground">بعد الضغط على «طباعة / حفظ PDF» ستفتح نافذة الطباعة؛ اختر منها **Save as PDF / حفظ كملف PDF**.</div>
    </div></Modal>}
  </>;
}

function cellTextForReport(value: ReportCell) {
  if (value instanceof Date) return value.toLocaleDateString('ar-EG');
  return value == null ? '—' : String(value);
}

function Router() {
  return (
    // Keep a shared shell (sidebar, navbar) outside the boundary so it
    // survives a page crash.
    <RoutedErrorBoundary>
      <Switch>
        <Route path="/excel"><Shell><ExcelHub /></Shell></Route>
        <Route path="/sales"><Shell><Sales /></Shell></Route>
        <Route path="/purchases"><Shell><Purchases /></Shell></Route>
        <Route path="/cash"><Shell><Cash /></Shell></Route>
        <Route path="/inventory"><Shell><Inventory /></Shell></Route>
        <Route path="/employees"><Shell><Employees /></Shell></Route>
        <Route path="/masters"><Shell><Masters /></Shell></Route>
        <Route path="/balances"><Shell><PartyBalances /></Shell></Route>
        <Route path="/customer-statement"><Shell><PartyStatement /></Shell></Route>
        <Route path="/partners"><Shell><Partners /></Shell></Route>
        <Route path="/assets"><Shell><Assets /></Shell></Route>
        <Route path="/accounting"><Shell><Accounting /></Shell></Route>
        <Route path="/reports"><Shell><Reports /></Shell></Route>
        <Route path="/"><Shell><Dashboard /></Shell></Route>
        <Route component={NotFound} />
      </Switch>
    </RoutedErrorBoundary>
  );
}

function RoutedErrorBoundary({ children }: { children: ReactNode }) {
  const [location] = useLocation();
  return <ErrorBoundary resetKey={location}>{children}</ErrorBoundary>;
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, '')}>
          <Router />
        </WouterRouter>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
