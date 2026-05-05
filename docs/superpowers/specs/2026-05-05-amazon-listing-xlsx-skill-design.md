# Amazon Listing XLSX Skill Design

## Goal

Create a Codex skill that converts a fixed-format product information workbook into a fixed-format Amazon listing workbook.

The user workflow must stay simple: provide one source `.xlsx` file, then receive one generated Amazon listing `.xlsx` file.

## Inputs

The skill accepts one `.xlsx` workbook that follows the same structure as `source.xlsx`.

Expected source structure:

- Sheet name: `Product Info`
- First row: fixed product information headers
- Data rows: one product per row
- Source `Item Code` values must be non-empty for every product row.
- Key source fields include:
  - `Item Code`
  - `Product Name`
  - `Main Color`
  - `Main Material`
  - Product dimensions and package dimensions
  - `Description`
  - `Product Features 1` through `Product Features 10`
  - Product image and document URL fields
  - `Notes`

## Outputs

The skill generates one `.xlsx` workbook that follows the same logical structure as `target.xlsx`.

Expected output structure:

- Sheet name: `listings`
- Header row, in this exact order:
  - `itemCode`
  - `productTitle`
  - `productDescription`
  - `searchTerms`
  - `bulletPoint1`
  - `bulletPoint2`
  - `bulletPoint3`
  - `bulletPoint4`
  - `bulletPoint5`

The output must preserve the fixed Amazon listing schema. Style can be minimal, but sheet name, column order, field names, and one-output-row-per-source-row behavior are mandatory.

## Listing Generation Rules

The skill should reuse the listing prompt rules from:

`/Users/vincent/Documents/workspace/code/personal/github/listing-generate-workflow/prompts/listing`

Core rules to preserve:

- Output listing language defaults to English.
- Title must be Amazon-ready and optimized for SEO.
- Product description must use Amazon-compatible HTML formatting such as `<b>` and `<br>`.
- Five bullet points must be generated for every row.
- Search terms must be space-separated, without commas.
- Avoid emoji and non-ASCII characters for English output.
- Avoid Amazon-sensitive wording, especially pesticide-related claims such as `anti-microbial`, `anti-bacterial`, `anti-fungal`, `harmless`, `non-toxic`, `repellent`, `virus`, and `germs`, unless the source explicitly provides required certification. For this skill, uncertified products are treated as ordinary goods.

## Architecture

The skill should be script-driven instead of freeform prompt-driven.

Components:

1. `input parser`
   - Opens the source workbook.
   - Finds the `Product Info` sheet.
   - Reads the header row.
   - Validates required columns.
   - Converts each product row into a structured record.

2. `listing generator`
   - Builds a prompt payload for each source row.
   - Uses the listing rules from the referenced prompt files.
   - Produces structured fields:
     - `itemCode`
     - `productTitle`
     - `productDescription`
     - `searchTerms`
     - `bulletPoint1` through `bulletPoint5`

3. `xlsx writer`
   - Creates a new workbook.
   - Creates one `listings` sheet.
   - Writes the fixed header row.
   - Writes one output row for each source product row.

4. `validator`
   - Confirms the generated workbook can be opened.
   - Confirms the `listings` sheet exists.
   - Confirms headers exactly match the target schema.
   - Confirms output data row count matches source data row count.
   - Confirms every output row has all required fields.

## Data Flow

1. User provides a source `.xlsx` file.
2. The skill reads `Product Info`.
3. The skill validates source structure.
4. The skill generates listing content row by row.
5. The skill writes the generated content into the fixed `listings` schema.
6. The skill validates the output workbook.
7. The skill returns the output `.xlsx` path.

## Error Handling

The skill should fail clearly before generating output when:

- The input file does not exist.
- The input file is not a readable `.xlsx` workbook.
- The workbook does not contain `Product Info`.
- Required headers are missing.
- A non-empty source data row has an empty `Item Code`; the skill must not fallback to or invent item codes.
- No product data rows are found.

For row-level weak data:

- Preserve output row count.
- Keep required columns present.
- Final output fields must be non-empty.
- If source information is insufficient to generate a required listing field, fail clearly or generate reasonable conservative listing copy from the available product context.
- Do not write empty target fields.
- Surface a concise error or warning that identifies affected item codes or source row numbers.

## Validation Requirements

A run is successful only when:

- The output file exists.
- The output workbook opens successfully.
- The output contains exactly one `listings` sheet for the generated listing data.
- The output headers match the target schema exactly.
- The output data row count matches the source product row count.
- Each row contains `itemCode`, title, description, search terms, and five bullet fields.
- Visual formatting is not a pass/fail criterion unless it affects the workbook's readability or schema.

## Scope

In scope:

- Build a Codex skill for fixed-format source-to-target `.xlsx` conversion.
- Include deterministic workbook parsing, writing, and validation instructions.
- Reuse the existing listing prompt rules as the content-generation reference.
- Preserve the `target.xlsx` logical schema exactly.

Out of scope:

- Building a web UI.
- Supporting arbitrary spreadsheet formats.
- Supporting multiple marketplace templates.
- Editing product images or downloading image files.
- Guaranteeing Amazon approval after upload.
- Recreating every visual style detail from `target.xlsx` beyond the required schema.

## Implementation Notes

The skill should prefer a bundled script for repeatability. The script can use `openpyxl` for workbook IO and validation.

The `SKILL.md` should stay concise and point to bundled references or scripts instead of embedding long prompt text. The prompt rules can be copied into a reference file or summarized with a clear pointer to the existing local prompt directory.

The default output filename should be derived from the input filename, for example:

`source.amazon-listings.xlsx`

## Open Decisions Resolved

- The target schema must strictly match `target.xlsx`.
- The user approved a script-driven approach.
- The source and target workbook formats are treated as fixed.
- The skill should output `.xlsx` directly, not JSON.
