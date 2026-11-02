# Build your own shopbook with an AI assistant

This is the order I'd build a small-business workbook in today, with example prompts for each step. It works with any capable AI assistant. Paste the prompts, adjust the bracketed parts, and test each step in Excel before moving on.

The rule throughout: **you own the data model, and the AI writes the boilerplate.** Decide what the business needs to track. Let the assistant handle the formulas and the VBA, then check everything it gives you.

---

## 1. Data model first

Before any formulas, write down the things the business records and how they relate.

> I run a small [type of shop] that sells about [N] products. I need to track supplier orders, sales, customers, cash taken out of the till, and stock. Propose a data model as Excel tables: one table per thing I record, the columns for each, and which column links them. Keep it to the minimum a one-person business would actually fill in.

What to check:
- One row per *event* (one sale line, one order line, one withdrawal), not one row per day.
- One master **Catalogue** that every other table looks up.
- Status columns stay small, fixed lists (for example PENDING / ORDERED / RECEIVED).

## 2. The catalogue table

> Create an Excel table called tblCatalogue with columns Brand, Size, Product, Key, Cost, Price. Key should be a formula joining Brand and Size with a pipe. Give me the formulas as structured references, not cell addresses.

Make every list an **Excel table** (Ctrl+T / Cmd+T). Tables grow on their own, keep formulas consistent down a column, and let formulas say `tblSales[Quantity]` instead of `H4:H94`.

## 3. Dependent dropdowns (Brand → Size → Product)

> I have tblCatalogue with Brand, Size, Product and Key (Brand|Size). I want three dependent dropdowns on my Sales table: Brand, then the sizes for that brand, then the products for that brand and size. Adding a row to the catalogue must update the dropdowns automatically, with no named ranges per brand and no requirement that the catalogue is sorted.

Pitfall to know about: a data-validation list can't use `FILTER` directly, because Excel wants a range, not an array. The pattern that works:
1. Put spilled helper lists on the Catalogue sheet. For example, `SORTBY(tblCatalogue[Product], XMATCH(keys, keys))` groups products by key in catalogue order.
2. Name them (`dd_Key = Catalogue!$U$2#`).
3. Use `OFFSET(INDEX(dd_Product, MATCH(key, dd_Key, 0)), 0, 0, COUNTIF(dd_Key, key), 1)` as the validation source.

## 4. Macros: buttons that only touch tables

> Write VBA for an Excel table called tblSales: AddSale (adds a row with the next invoice number and today's date) and DeleteLastSale (asks for confirmation). Use ListObject.ListRows only, find columns by header name, and never select or copy worksheet rows.

Then add a workbook event:

> Write a Workbook_SheetChange handler that, when Brand, Size or Product changes in tblSales, copies Unit Price and Unit Cost from tblCatalogue into that row as values. Disable events while writing and re-enable them on error.

Stamping prices at the moment of sale means a later price change never rewrites past sales.

## 5. Stock and the dashboard

> Stock for each catalogue row = SUM of tblOrders[Quantity] where Status is RECEIVED and Brand, Size and Product match, minus SUM of tblSales[Quantity] for the same product. Write it as a calculated column in tblCatalogue.

> Build a dashboard: revenue, gross profit, net profit after write-offs, profit per partner from a settings cell, top brand and top 3 products by units (dynamic arrays), and a live list of products in stock with quantity.

Keep the dashboard *read-only*: formulas over tables and nothing typed in by hand.

## 6. Review it like someone else wrote it

This step found every bug in the 2025 version of this workbook. Give the assistant your formulas and VBA and ask:

> Review this workbook as if it's going into production for a real business. List every formula or macro that could give a wrong number, delete the wrong row, depend on a hard-coded range, or break if a column is added. For each, show a concrete input that triggers it and the fix.

Then test each fix yourself, with real clicks in Excel, before you trust it.

## 7. Sample data you can share

> Write a seeded Python script that generates realistic but fake data for these tables: [paste your column lists]. Sales may only sell stock that a RECEIVED order has delivered, and never below zero.

The same seed gives the same data every time. That's handy for screenshots, tests and a public repo.

---

### Lessons from building this one

- **Tables over ranges.** Most of the 2025 bugs came from fixed ranges and pasted rows.
- **Model the bookkeeping, not just the sales.** Write-offs and owner drawings are stock leaving the shop without revenue, so they need their own rows.
- **Freeze prices at the point of sale.**
- **Mac Excel quirk:** VBA that uses Office shape constants (`msoTrue`, `msoShapeRoundedRectangle`) needs the Office Object Library reference. Declare the values as constants instead, so the code compiles anywhere.
- **The AI is fast at boilerplate and good at reviews.** It still needs you to decide what the business actually does.
