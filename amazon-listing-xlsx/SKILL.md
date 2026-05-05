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

1. Resolve the installed skill directory first, for example `SKILL_DIR=/path/to/amazon-listing-xlsx`.

2. Export source rows to JSON:

   ```bash
   python "$SKILL_DIR/scripts/listing_workbook.py" export-source <source.xlsx> --output build/source.rows.json
   ```

3. Read `$SKILL_DIR/references/listing_prompt_rules.md`.

4. Generate listing JSON for every exported row.

   Requirements:

   - Preserve every `rowIndex`.
   - Preserve item order.
   - Return strict JSON array only.
   - Include all target fields.
   - Use empty strings only when a field cannot be recovered.

5. Save generated JSON to `build/generated-listings.json`.

6. Write the final workbook:

   ```bash
   python "$SKILL_DIR/scripts/listing_workbook.py" write-output <source.xlsx> build/generated-listings.json --output <source.amazon-listings.xlsx>
   ```

7. Validate the final workbook:

   ```bash
   python "$SKILL_DIR/scripts/listing_workbook.py" validate-output <source.xlsx> <source.amazon-listings.xlsx>
   ```

8. Return the generated workbook path to the user.

## Rules

- Do not change the target sheet name or column order.
- Do not output CSV or JSON as the final deliverable unless the user explicitly asks.
- Do not skip source rows.
- Do not invent item codes.
- Keep generated description line breaks as `<br>`, not actual newline characters.
- If validation fails, fix the generated JSON or workbook issue and rerun validation before responding.

## References

- Listing copy rules: `$SKILL_DIR/references/listing_prompt_rules.md`
- Workbook utility: `$SKILL_DIR/scripts/listing_workbook.py`
