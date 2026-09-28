---
name: OpenAPI date fields
description: The workspace's generated API types and persisted JSON date values use different runtime representations.
---

Treat OpenAPI `format: date` as a runtime boundary: generated TypeScript types use `Date`, while JSONB persistence and browser requests commonly contain ISO strings. Hydrate dates after loading persisted state and serialize them at the API edge.

**Why:** The generated client and Zod schemas are not both string-based for date fields, so assuming one representation across database, server, and browser creates type errors or invalid response parsing.

**How to apply:** When adding date fields to `lib/api-spec/openapi.yaml`, rerun codegen and update the store hydration/serialization path before wiring a new endpoint.