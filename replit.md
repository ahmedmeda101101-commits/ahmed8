# ماكينة التطريز — الحسابات والمخزون

نظام ويب عربي لإدارة حسابات ورشة التطريز، مع ترحيل البيانات التاريخية، الخزينة، المخزون، الشركاء، الأصول والتقارير.

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `DATABASE_URL` — Postgres connection string

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)

## Where things live

- `artifacts/maquina-accounts/src/App.tsx` — واجهة RTL ومسارات التشغيل والمحاسبة.
- `artifacts/maquina-accounts/src/index.css` — الهوية البصرية والتصميم المتجاوب.
- `artifacts/api-server/src/lib/accounting-store.ts` — ترحيل البيانات، الحسابات المشتقة، والتخزين الدائم.
- `artifacts/api-server/src/lib/data/` — نسخ مستقلة من بيانات Excel المرحّلة، دون تعديل الملفات الأصلية.
- `lib/api-spec/openapi.yaml` — عقد API ومصدر توليد hooks وZod.
- `lib/db/src/schema/accounting-state.ts` — جدول الحالة المحاسبية الدائمة.

## Architecture decisions

- تُحفظ الحالة المحاسبية في PostgreSQL كحزمة JSON واحدة حتى تظل عملية الترحيل التاريخية قابلة للتدقيق ولا تُفقد الحقول القديمة.
- تُعاد حسابات لوحة المتابعة والمخزون والدخل من الحركات المرحّلة، ولا تعتمد على نتائج Excel المحسوبة التي تحتوي أخطاء.
- اختلاف قسط إهلاك الماكينة بين مصدرين (8125 و7500 جنيه) يظهر كتدقيق تحذيري ولا يُخفى داخل رقم واحد.
- المسارات الداخلية تستخدم Wouter مع `BASE_URL` حتى تعمل داخل مسار artifact المخصص.

## Product

- لوحة متابعة للمبيعات والمشتريات والخزينة والمخزون.
- إضافة العملاء والموردين والمنتجات والخامات، وإنشاء فواتير وحركات خزينة مع تحقق من المدخلات.
- كشوف الشركاء، الأصول والإهلاك، دفتر اليومية، قائمة الدخل، وتسويات جودة البيانات.
- مركز تقارير قابل للتصفية والطباعة والتصدير بصيغة Excel، ويضم تقارير المبيعات والمشتريات والعملاء والموردين والخزينة والمخزون والشركاء والأصول والإهلاك واليومية والتدفقات والدخل وميزان المراجعة والمركز المالي.
- البيانات المرحّلة حاليًا: 63 فاتورة بيع، 28 فاتورة شراء، 137 حركة خزينة، 11 مادة، 3 شركاء و5 أصول.

## User preferences

_Populate as you build — explicit user instructions worth remembering across sessions._

## Gotchas

- ملف الإهلاك التاريخي يحتوي قسطين مختلفين للماكينة؛ لا تعتمد على إجمالي الإهلاك دون مراجعة التنبيه في شاشة المحاسبة.
- ملف المخزون الصحيح النهائي غير متاح ضمن المصادر الحالية؛ أرصدة المخزون الحالية مبنية على حركات المخزن الموروثة المتاحة وتحتاج اعتمادًا عند وصول الملف الصحيح.
- بعد تعديل `lib/api-spec/openapi.yaml` شغّل codegen قبل typecheck حتى تتطابق hooks وZod مع العقد.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
