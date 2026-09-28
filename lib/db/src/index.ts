import { drizzle } from "drizzle-orm/node-postgres";
import pg from "pg";
import * as schema from "./schema/index.ts";

const { Pool } = pg;

export let pool: pg.Pool | undefined;
export let db: any;

try {
  if (process.env.DATABASE_URL) {
    pool = new Pool({ connectionString: process.env.DATABASE_URL });
    db = drizzle(pool, { schema });
  } else {
    console.warn("[AI Studio] DATABASE_URL not set — using in-memory store");
  }
} catch (err) {
  console.warn("[AI Studio] Failed to connect to database — using in-memory store", err);
}

export * from "./schema/index.ts";
