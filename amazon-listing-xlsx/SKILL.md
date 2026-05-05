---
name: amazon-listing-xlsx
description: Convert fixed-format product information XLSX files into fixed-format Amazon listing XLSX workbooks. Use when the user provides a product source workbook like source.xlsx and wants an Amazon listing workbook like target.xlsx with title, description, search terms, and five bullet points.
---

# Amazon Listing XLSX

Convert a fixed-format product workbook into an Amazon listing workbook.

## When to Use

Use this skill when the user provides a `.xlsx` product information source file and wants a generated Amazon listing `.xlsx` file.

Expected source:

- Sheet: `Product Info`
- One product per row
- Fixed product information headers like `Item Code`, `Product Name`, `Main Color`, `Main Material`, `Description`, and `Product Features 1`

Expected output:

- Sheet: `listings`
- Headers: `itemCode`, `productTitle`, `productDescription`, `searchTerms`, `bulletPoint1`, `bulletPoint2`, `bulletPoint3`, `bulletPoint4`, `bulletPoint5`

## Workflow

1. Export source rows to JSON:

   ```bash
   python amazon-listing-xlsx/scripts/listing_workbook.py export-source <source.xlsx> --output build/source.rows.json
   ```

2. Read `references/listing_prompt_rules.md`.

3. Generate listing JSON for every exported row.

   Requirements:

   - Preserve every `rowIndex`.
   - Preserve item order.
   - Return strict JSON array only.
   - Include all target fields.
   - Use empty strings only when a field cannot be recovered.

4. Save generated JSON to `build/generated-listings.json`.

5. Write the final workbook:

   ```bash
   python amazon-listing-xlsx/scripts/listing_workbook.py write-output <source.xlsx> build/generated-listings.json --output <source.amazon-listings.xlsx>
   ```

6. Validate the final workbook:

   ```bash
   python amazon-listing-xlsx/scripts/listing_workbook.py validate-output <source.xlsx> <source.amazon-listings.xlsx>
   ```

7. Return the generated workbook path to the user.

## Rules

- Do not change the target sheet name or column order.
- Do not output CSV or JSON as the final deliverable unless the user explicitly asks.
- Do not skip source rows.
- Do not invent item codes.
- Keep generated description line breaks as `<br>`, not actual newline characters.
- If validation fails, fix the generated JSON or workbook issue and rerun validation before responding.

## References

- Listing copy rules: `references/listing_prompt_rules.md`
- Workbook utility: `scripts/listing_workbook.py`
