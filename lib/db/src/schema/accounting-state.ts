import { jsonb, pgTable, serial, timestamp } from "drizzle-orm/pg-core";

export const accountingStateTable = pgTable("accounting_state", {
  id: serial("id").primaryKey(),
  payload: jsonb("payload").notNull(),
  updatedAt: timestamp("updated_at", { withTimezone: true })
    .notNull()
    .defaultNow()
    .$onUpdate(() => new Date()),
});

export type AccountingStateRow = typeof accountingStateTable.$inferSelect;