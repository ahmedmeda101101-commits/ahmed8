import * as XLSX from "xlsx";

export type ReportCell = string | number | Date | null | undefined;

const cellText = (value: ReportCell): string => {
  if (value instanceof Date) return value.toLocaleDateString("ar-EG");
  if (typeof value === "number") return String(value);
  return value == null ? "" : String(value);
};

const escapeHtml = (value: string) =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");

export function exportReportToExcel(title: string, headers: string[], rows: ReportCell[][]) {
  try {
    const data: (string | number)[][] = [
      [title],
      headers,
      ...rows.map((row) =>
        row.map((val) => {
          if (val instanceof Date) return val.toLocaleDateString("ar-EG");
          if (typeof val === "number") return val;
          return val == null ? "" : String(val);
        })
      ),
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // Auto column widths
    const colWidths = headers.map((header, colIndex) => {
      let maxLen = header.length;
      for (const row of rows) {
        const text = cellText(row[colIndex]);
        if (text.length > maxLen) maxLen = text.length;
      }
      return { wch: Math.max(maxLen + 4, 12) };
    });
    worksheet["!cols"] = colWidths;

    // Set RTL property
    if (!worksheet["!views"]) worksheet["!views"] = [];
    worksheet["!views"].push({ rightToLeft: true });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, title.slice(0, 31));

    XLSX.writeFile(workbook, `${title.replaceAll(/\s+/g, "_")}.xlsx`);
  } catch (err) {
    console.warn("Falling back to HTML Excel export", err);
    // Fallback HTML table export
    const tableRows = [
      `<tr><th colspan="${headers.length}" style="font-size:18px;background:#163b3a;color:#fff">${escapeHtml(title)}</th></tr>`,
      `<tr>${headers.map((header) => `<th style="background:#d9e8e3;font-weight:bold">${escapeHtml(header)}</th>`).join("")}</tr>`,
      ...rows.map((row) => `<tr>${row.map((value) => `<td>${escapeHtml(cellText(value))}</td>`).join("")}</tr>`),
    ].join("");
    const html = `<!doctype html><html dir="rtl"><head><meta charset="utf-8"></head><body><table border="1">${tableRows}</table></body></html>`;
    const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${title.replaceAll(/\s+/g, "_")}.xls`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

export function exportMultiSheetExcel(
  filename: string,
  sheets: Array<{ name: string; headers: string[]; rows: ReportCell[][] }>
) {
  const workbook = XLSX.utils.book_new();

  for (const sheet of sheets) {
    const data: (string | number)[][] = [
      sheet.headers,
      ...sheet.rows.map((row) =>
        row.map((val) => {
          if (val instanceof Date) return val.toLocaleDateString("ar-EG");
          if (typeof val === "number") return val;
          return val == null ? "" : String(val);
        })
      ),
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const colWidths = sheet.headers.map((header, colIndex) => {
      let maxLen = header.length;
      for (const row of sheet.rows) {
        const text = cellText(row[colIndex]);
        if (text.length > maxLen) maxLen = text.length;
      }
      return { wch: Math.max(maxLen + 4, 12) };
    });
    ws["!cols"] = colWidths;
    if (!ws["!views"]) ws["!views"] = [];
    ws["!views"].push({ rightToLeft: true });

    const safeName = sheet.name.slice(0, 31).replaceAll(/[\\/?*[\]]/g, "_");
    XLSX.utils.book_append_sheet(workbook, ws, safeName);
  }

  XLSX.writeFile(workbook, `${filename}.xlsx`);
}

export function downloadExcelTemplate(kind: "sales" | "purchases" | "cash" | "materials" | "customers") {
  const templates: Record<string, { filename: string; headers: string[]; samples: (string | number)[][] }> = {
    sales: {
      filename: "نموذج_فواتير_المبيعات_ماكينة_التطريز",
      headers: ["رقم الفاتورة", "التاريخ (YYYY-MM-DD)", "اسم العميل", "كود الصنف", "اسم الصنف / الخدمة", "الكمية", "سعر الوحدة", "طريقة الدفع"],
      samples: [
        ["INV-101", "2026-09-01", "مصنع الهدى للملابس", "4001", "تطريز عبايات كمبيوتر", 50, 45, "آجل"],
        ["INV-102", "2026-09-02", "الحاج علاء فكري", "4002", "تطريز لوجو شتوي", 120, 25, "نقدي"],
      ],
    },
    purchases: {
      filename: "نموذج_فواتير_المشتريات_والخامات",
      headers: ["رقم الفاتورة", "التاريخ (YYYY-MM-DD)", "اسم المورد", "كود الخامة", "اسم الخامة", "الكمية", "سعر الوحدة", "طريقة الدفع"],
      samples: [
        ["PUR-501", "2026-09-01", "الشركة الهندسية للخيوط", "1001", "بكرة خيط تطريز 5000م أبيض", 20, 110, "نقدي"],
        ["PUR-502", "2026-09-03", "مكتب النصر للفازلين", "1002", "فازلين تطريز لفة 100م", 3, 380, "نقدي"],
      ],
    },
    cash: {
      filename: "نموذج_حركات_الخزينة_اليومية",
      headers: ["رقم السند", "التاريخ (YYYY-MM-DD)", "نوع الحركة (قبض/صرف)", "المبلغ", "كود الحساب (G)", "جهة الإيراد/الصرف", "البند/المستفيد", "البيان (J)", "كود تحليل البند (K)", "تحليل البند (L)", "ملاحظات"],
      samples: [
        ["RV-1001", "2026-09-01", "قبض", 5000, "4001", "الحاج علاء فكري", "الحاج علاء فكري", "دفعة نقدية من العميل", "6000", "إيرادات", "دفعة على حساب العميل"],
        ["PV-1002", "2026-09-01", "صرف", 850, "8000", "مصروفات ادارية وعمومية", "شركة الكهرباء", "فاتورة كهرباء الورشة شهر 9", "8003", "مرافق وخدمات", ""],
        ["PV-1003", "2026-09-02", "صرف", 1500, "9000", "مصاريف الانتاج", "فني الصيانة", "صيانة دورية وتزييت الماكينة", "9006", "مصاريف صيانة", ""],
      ],
    },
    materials: {
      filename: "نموذج_أصناف_الخامات_والمخزون",
      headers: ["الكود", "اسم الخامة", "الوحدة", "سعر التكلفة الافتراضي"],
      samples: [
        ["1001", "بكرة خيط تطريز بوليستر أبيض", "بكرة", 110],
        ["1002", "بكرة خيط تطريز بوليستر أسود", "بكرة", 110],
        ["1003", "فازلين مقوى 80 جم", "متر", 8.5],
        ["1004", "إبر تطريز مقاس 11/75", "علبة", 140],
        ["1005", "زيت ماكينة أبيض عالي النقاء", "لتر", 95],
      ],
    },
    customers: {
      filename: "نموذج_دليل_العملاء",
      headers: ["الكود", "اسم العميل / المصنع", "رقم الهاتف", "العنوان", "الرصيد الافتتاحي"],
      samples: [
        ["4001", "مصنع الهدى للملابس الجاهزة", "01001234567", "المحلة الكبرى", 12500],
        ["4002", "الحاج علاء فكري", "01223456789", "القاهرة - شبرا", 0],
      ],
    },
  };

  const tpl = templates[kind];
  if (!tpl) return;

  const data = [tpl.headers, ...tpl.samples];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = tpl.headers.map((h) => ({ wch: Math.max(h.length + 6, 16) }));
  if (!ws["!views"]) ws["!views"] = [];
  ws["!views"].push({ rightToLeft: true });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "النموذج");
  XLSX.writeFile(wb, `${tpl.filename}.xlsx`);
}

export interface ParsedSheetData {
  sheetName: string;
  headers: string[];
  rows: Record<string, any>[];
  rawRows: (string | number)[][];
}

export async function parseUploadedExcel(file: File): Promise<ParsedSheetData[]> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

  const result: ParsedSheetData[] = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    const rawRows = XLSX.utils.sheet_to_json<(string | number)[]>(worksheet, { header: 1, defval: "" });
    if (!rawRows || rawRows.length === 0) continue;

    const headerRowIndex = rawRows.findIndex((r) => r && r.some((c) => String(c).trim() !== ""));
    if (headerRowIndex === -1) continue;

    const headers = (rawRows[headerRowIndex] || []).map((c) => String(c ?? "").trim());
    const validRows = rawRows.slice(headerRowIndex + 1).filter((r) => r && r.some((c) => String(c).trim() !== ""));

    const objectRows: Record<string, any>[] = [];
    for (const r of validRows) {
      const obj: Record<string, any> = {};
      headers.forEach((h, idx) => {
        if (h) {
          obj[h] = r[idx] !== undefined ? r[idx] : "";
        }
      });
      objectRows.push(obj);
    }

    result.push({
      sheetName,
      headers: headers.filter(Boolean),
      rows: objectRows,
      rawRows,
    });
  }

  return result;
}

export function printReport(title: string, headers: string[], rows: ReportCell[][]) {
  const tableRows = rows.map((row) => `<tr>${row.map((value) => `<td>${escapeHtml(cellText(value))}</td>`).join('')}</tr>`).join('');
  const html = `<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title>
    <style>
      @page { size: A4 landscape; margin: 9mm; }
      * { box-sizing: border-box; }
      html, body { margin:0; padding:0; background:#fff; }
      body { font-family:"Segoe UI",Tahoma,Arial,sans-serif; color:#172b2a; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
      .report { width:100%; }
      header { display:flex; justify-content:space-between; align-items:flex-start; gap:20px; border:1px solid #d6e0dd; border-top:5px solid #28635e; border-radius:12px; padding:13px 15px; margin-bottom:12px; background:linear-gradient(180deg,#f8fbfa,#fff); }
      h1 { margin:0; font-size:21px; color:#163b3a; } p { margin:4px 0 0; color:#647572; font-size:10.5px; }
      .meta { text-align:left; white-space:nowrap; font-size:9.5px; color:#647572; }
      table { border-collapse:separate; border-spacing:0; width:100%; font-size:9.5px; table-layout:auto; page-break-inside:auto; }
      thead { display:table-header-group; }
      tr { page-break-inside:avoid; page-break-after:auto; }
      th { background:#163b3a; color:#fff; padding:7px 6px; border-bottom:1px solid #163b3a; text-align:right; font-weight:700; }
      td { padding:6px; border-bottom:1px solid #dfe7e4; text-align:right; vertical-align:top; overflow-wrap:anywhere; }
      tbody tr:nth-child(even) td { background:#f6f9f8; }
      .footer { margin-top:10px; color:#647572; font-size:8.5px; display:flex; justify-content:space-between; border-top:1px solid #e0e7e5; padding-top:6px; }
    </style></head><body><div class="report">
    <header><div><h1>${escapeHtml(title)}</h1><p>نظام ماكينة التطريز — تقرير مفصل مطابق لحركة النظام</p></div><div class="meta">تاريخ الاستخراج<br>${escapeHtml(new Date().toLocaleString('ar-EG'))}</div></header>
    <table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${tableRows}</tbody></table>
    <div class="footer"><span>تقرير رسمي من نظام ماكينة التطريز</span><span>في نافذة الطباعة اختر «Microsoft Print to PDF» أو «Save as PDF»</span></div>
    </div></body></html>`;

  const printWindow = window.open('', '_blank', 'width=1280,height=900,scrollbars=yes,resizable=yes');
  if (printWindow) {
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    window.setTimeout(() => {
      try { printWindow.print(); } catch { /* browser closed the tab */ }
    }, 450);
    return true;
  }

  // Popup blocked: fall back to a same-origin visible tab via a downloadable HTML document.
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.target = '_blank'; a.rel = 'noopener'; a.download = `${title.replaceAll(/\s+/g, '_')}.html`;
  document.body.appendChild(a); a.click(); a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 5000);
  return true;
}
