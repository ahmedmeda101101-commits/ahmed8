---
name: Arabic report export strategy
description: Why the accounting reports use browser-compatible Excel and native print PDF output.
---

The accounting report center exports Arabic tables as a UTF-8 HTML `.xls` file and opens a print-ready Arabic document for printing or Save as PDF. This avoids adding a client PDF/font dependency while preserving RTL text and spreadsheet compatibility.

**Why:** The available source data is Arabic and the workspace did not have a safe, correctly scoped spreadsheet/PDF dependency installed; native browser output is reliable for the current report requirements.

**How to apply:** Keep report columns and values plain text/number compatible, and do not label cash sales as customer collections or cash purchases as supplier payments unless the underlying movement is linked to a party.