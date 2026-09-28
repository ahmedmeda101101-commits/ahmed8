import { Router, type IRouter } from "express";
import {
  CreateCashEntryBody,
  CreateCashEntryResponse,
  CreateCustomerBody,
  CreateCustomerResponse,
  CreateMaterialBody,
  CreateMaterialResponse,
  CreateProductBody,
  CreateProductResponse,
  CreatePurchaseBody,
  CreatePurchaseResponse,
  CreateSaleBody,
  CreateSaleResponse,
  CreateSupplierBody,
  CreateSupplierResponse,
  GetDashboardResponse,
  GetIncomeReportResponse,
  GetReconciliationResponse,
  ListAssetsResponse,
  ListCashResponse,
  ListCustomersResponse,
  ListInventoryResponse,
  ListJournalsResponse,
  ListMaterialsResponse,
  ListPartnersResponse,
  ListProductsResponse,
  ListPurchasesResponse,
  ListSalesResponse,
  ListSuppliersResponse,
  UpdateCustomerBody,
  UpdateCustomerParams,
  UpdateCustomerResponse,
} from "@workspace/api-zod";
import {
  addCash,
  addMaster,
  addProfitAllocation,
  addTransaction,
  bulkImport,
  getDashboard,
  getExcelAudit,
  getIncomeReport,
  getMachineDepreciationSchedule,
  getMonthlyProfitSummary,
  getPartyStatement,
  getBalances,
  getReconciliation,
  listAssets,
  listCash,
  listCustomers,
  listInventory,
  listInventoryMovements,
  addInventoryMovement,
  listEmployees,
  addEmployee,
  listPayroll,
  addPayroll,
  listJournals,
  listMaterials,
  listPartners,
  listProducts,
  listProfitAllocations,
  listPurchases,
  listSales,
  listSuppliers,
  resetToExcelDefaults,
  updateCustomer,
} from "../lib/accounting-store";

const router: IRouter = Router();

function badRequest(res: Parameters<NonNullable<Parameters<IRouter["get"]>[1]>>[1], error: unknown) {
  res.status(400).json({ error: error instanceof Error ? error.message : "بيانات غير صحيحة" });
}

router.get("/dashboard", async (_req, res): Promise<void> => {
  res.json(GetDashboardResponse.parse(await getDashboard()));
});

router.get("/customers", async (_req, res): Promise<void> => {
  res.json(ListCustomersResponse.parse(await listCustomers()));
});

router.post("/customers", async (req, res): Promise<void> => {
  const parsed = CreateCustomerBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreateCustomerResponse.parse(await addMaster("customers", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.patch("/customers/:id", async (req, res): Promise<void> => {
  const params = UpdateCustomerParams.safeParse(req.params);
  const body = UpdateCustomerBody.safeParse(req.body);
  if (!params.success) { badRequest(res, params.error); return; }
  if (!body.success) { badRequest(res, body.error); return; }
  const result = await updateCustomer(params.data.id, body.data);
  if (!result) { res.status(404).json({ error: "العميل غير موجود" }); return; }
  res.json(UpdateCustomerResponse.parse(result));
});

router.get("/suppliers", async (_req, res): Promise<void> => {
  res.json(ListSuppliersResponse.parse(await listSuppliers()));
});

router.post("/suppliers", async (req, res): Promise<void> => {
  const parsed = CreateSupplierBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreateSupplierResponse.parse(await addMaster("suppliers", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.get("/products", async (_req, res): Promise<void> => {
  res.json(ListProductsResponse.parse(await listProducts()));
});

router.post("/products", async (req, res): Promise<void> => {
  const parsed = CreateProductBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreateProductResponse.parse(await addMaster("products", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.get("/materials", async (_req, res): Promise<void> => {
  res.json(ListMaterialsResponse.parse(await listMaterials()));
});

router.post("/materials", async (req, res): Promise<void> => {
  const parsed = CreateMaterialBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreateMaterialResponse.parse(await addMaster("materials", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.get("/sales", async (_req, res): Promise<void> => {
  res.json(ListSalesResponse.parse(await listSales()));
});

router.post("/sales", async (req, res): Promise<void> => {
  const parsed = CreateSaleBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreateSaleResponse.parse(await addTransaction("sales", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.get("/purchases", async (_req, res): Promise<void> => {
  res.json(ListPurchasesResponse.parse(await listPurchases()));
});

router.post("/purchases", async (req, res): Promise<void> => {
  const parsed = CreatePurchaseBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  try {
    res.status(201).json(CreatePurchaseResponse.parse(await addTransaction("purchases", parsed.data)));
  } catch (error) { badRequest(res, error); }
});

router.get("/cash", async (_req, res): Promise<void> => {
  res.json(ListCashResponse.parse(await listCash()));
});

router.get("/party-balances", async (_req, res): Promise<void> => {
  res.json(await getBalances());
});

router.get("/party-statement", async (req, res): Promise<void> => {
  const kind = req.query.kind === "supplier" ? "supplier" : "customer";
  const name = String(req.query.name || "").trim();
  if (!name) { res.status(400).json({ error: "اسم العميل أو المورد مطلوب" }); return; }
  res.json(await getPartyStatement(kind, name));
});

router.post("/cash", async (req, res): Promise<void> => {
  const parsed = CreateCashEntryBody.safeParse(req.body);
  if (!parsed.success) { badRequest(res, parsed.error); return; }
  res.status(201).json(CreateCashEntryResponse.parse(await addCash(parsed.data)));
});

router.get("/inventory", async (_req, res): Promise<void> => {
  res.json(ListInventoryResponse.parse(await listInventory()));
});

router.get("/inventory-movements", async (_req, res): Promise<void> => {
  res.json(await listInventoryMovements());
});

router.post("/inventory-movements", async (req, res): Promise<void> => {
  try { res.status(201).json(await addInventoryMovement(req.body)); } catch (err) { badRequest(res, err); }
});

router.get("/employees", async (_req, res): Promise<void> => {
  res.json(await listEmployees());
});

router.post("/employees", async (req, res): Promise<void> => {
  try { res.status(201).json(await addEmployee(req.body)); } catch (err) { badRequest(res, err); }
});

router.get("/payroll", async (_req, res): Promise<void> => {
  res.json(await listPayroll());
});

router.post("/payroll", async (req, res): Promise<void> => {
  try { res.status(201).json(await addPayroll(req.body)); } catch (err) { badRequest(res, err); }
});

router.get("/partners", async (_req, res): Promise<void> => {
  res.json(ListPartnersResponse.parse(await listPartners()));
});

router.get("/assets", async (_req, res): Promise<void> => {
  res.json(ListAssetsResponse.parse(await listAssets()));
});

router.get("/journals", async (_req, res): Promise<void> => {
  res.json(ListJournalsResponse.parse(await listJournals()));
});

router.get("/reports/reconciliation", async (_req, res): Promise<void> => {
  res.json(GetReconciliationResponse.parse(await getReconciliation()));
});

router.get("/reports/income", async (_req, res): Promise<void> => {
  res.json(GetIncomeReportResponse.parse(await getIncomeReport()));
});

router.get("/profit-allocations", async (_req, res): Promise<void> => {
  res.json(await listProfitAllocations());
});

router.get("/profit-allocations/summary", async (_req, res): Promise<void> => {
  res.json(await getMonthlyProfitSummary());
});

router.post("/profit-allocations", async (req, res): Promise<void> => {
  try {
    const { date, partner, amount, description } = req.body;
    if (!partner || !amount) {
      res.status(400).json({ error: "اسم الشريك والمبلغ مطلوبان" });
      return;
    }
    const result = await addProfitAllocation({ date: date || new Date().toISOString().slice(0, 10), partner, amount: Number(amount), description });
    res.status(201).json(result);
  } catch (err) {
    badRequest(res, err);
  }
});

router.get("/machine-depreciation-schedule", async (_req, res): Promise<void> => {
  res.json(await getMachineDepreciationSchedule());
});

router.get("/excel-audit", async (_req, res): Promise<void> => {
  res.json(await getExcelAudit());
});

router.post("/reset-to-excel", async (_req, res): Promise<void> => {
  try {
    const result = await resetToExcelDefaults();
    res.json(result);
  } catch (err) {
    badRequest(res, err);
  }
});

router.post("/bulk-import", async (req, res): Promise<void> => {
  try {
    const result = await bulkImport(req.body);
    res.status(200).json(result);
  } catch (err) {
    badRequest(res, err);
  }
});

export default router;