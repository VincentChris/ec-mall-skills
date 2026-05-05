# Amazon Listing XLSX Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Codex skill that converts a fixed-format product source workbook into a fixed-format Amazon listing workbook.

**Architecture:** Create a local skill folder with concise usage instructions, bundled prompt reference, and a deterministic Python utility for workbook parsing, output writing, and validation. Codex generates listing copy row by row using the prompt rules; the script owns the `.xlsx` schema so the final file does not drift.

**Tech Stack:** Codex skill format, Markdown, Python 3, `openpyxl`, Git.

---

## File Structure

- Create: `amazon-listing-xlsx/SKILL.md`
  - Skill trigger metadata and concise workflow for converting product `.xlsx` files to Amazon listing `.xlsx`.
- Create: `amazon-listing-xlsx/references/listing_prompt_rules.md`
  - Local concise reference adapted from `/Users/vincent/Documents/workspace/code/personal/github/listing-generate-workflow/prompts/listing`.
- Create: `amazon-listing-xlsx/scripts/listing_workbook.py`
  - Deterministic workbook IO utility. Reads `Product Info`, validates headers, exports row JSON for Codex, imports generated listing JSON, writes `listings`, validates output.
- Create: `tests/test_listing_workbook.py`
  - Unit tests for parsing, writing, validation, and CLI behavior.
- Create: `requirements-dev.txt`
  - Minimal Python dependencies for local testing.
- Modify: `docs/superpowers/specs/2026-05-05-amazon-listing-xlsx-skill-design.md`
  - Only if implementation exposes a necessary clarification.

## CLI Contract

The workbook utility must support these commands:

```bash
python amazon-listing-xlsx/scripts/listing_workbook.py export-source source.xlsx --output build/source.rows.json
python amazon-listing-xlsx/scripts/listing_workbook.py write-output source.xlsx build/generated-listings.json --output source.amazon-listings.xlsx
python amazon-listing-xlsx/scripts/listing_workbook.py validate-output source.xlsx source.amazon-listings.xlsx
```

`export-source` writes a JSON array of source row records:

```json
[
  {
    "rowIndex": 2,
    "itemCode": "PP191030AAU",
    "source": {
      "Item Code": "PP191030AAU",
      "Product Name": "3 Piece Luggage Set",
      "Main Color": "Light Pink"
    }
  }
]
```

`write-output` accepts a generated JSON array with fields:

```json
[
  {
    "rowIndex": 2,
    "itemCode": "PP191030AAU",
    "productTitle": "Generated title",
    "productDescription": "<b>Generated</b><br>Description",
    "searchTerms": "keyword keyword",
    "bulletPoint1": "1. HEADER - Content",
    "bulletPoint2": "2. HEADER - Content",
    "bulletPoint3": "3. HEADER - Content",
    "bulletPoint4": "4. HEADER - Content",
    "bulletPoint5": "5. HEADER - Content"
  }
]
```

The output workbook must contain sheet `listings` and exactly these headers:

```python
["itemCode", "productTitle", "productDescription", "searchTerms", "bulletPoint1", "bulletPoint2", "bulletPoint3", "bulletPoint4", "bulletPoint5"]
```

---

### Task 1: Add Test Dependency File

**Files:**
- Create: `requirements-dev.txt`

- [ ] **Step 1: Create dependency file**

Add this exact content:

```text
openpyxl>=3.1.0
pytest>=8.0.0
```

- [ ] **Step 2: Install dependencies**

Run:

```bash
python3 -m pip install -r requirements-dev.txt
```

Expected: command exits `0`; `openpyxl` and `pytest` are importable.

- [ ] **Step 3: Commit**

```bash
git add requirements-dev.txt
git commit -m "chore: add Python test dependencies"
```

---

### Task 2: Write Failing Workbook Tests

**Files:**
- Create: `tests/test_listing_workbook.py`

- [ ] **Step 1: Create tests**

Add this complete test file:

```python
import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "amazon-listing-xlsx" / "scripts" / "listing_workbook.py"


SOURCE_HEADERS = [
    "Item Code",
    "Product Name",
    "Main Color",
    "Main Material",
    "Assembled Length(inch)",
    "Assembled Width(inch)",
    "Assembled Height(inch)",
    "Product weight(pound)",
    "Package Size-Length (inch)",
    "Package Size-Width (inch)",
    "Package Size-Height (inch)",
    "Package Size-Weight (pound)",
    "Description",
    "Product Features 1",
    "Product Features 2",
    "Product Features 3",
    "Product Features 4",
    "Product Features 5",
    "Product Features 6",
    "Product Features 7",
    "Product Features 8",
    "Product Features 9",
    "Product Features 10",
    "Product Main Image",
    "Notes",
]


TARGET_HEADERS = [
    "itemCode",
    "productTitle",
    "productDescription",
    "searchTerms",
    "bulletPoint1",
    "bulletPoint2",
    "bulletPoint3",
    "bulletPoint4",
    "bulletPoint5",
]


def create_source_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Product Info"
    ws.append(SOURCE_HEADERS)
    ws.append([
        "SKU-1",
        "Expandable Hardside Luggage",
        "Black",
        "ABS",
        20,
        12,
        28,
        22,
        30,
        18,
        12,
        24,
        "Durable suitcase with TSA lock and spinner wheels.",
        "Expandable design",
        "TSA lock",
        "Spinner wheels",
        "Nested storage",
        "Fully lined interior",
        "",
        "",
        "",
        "",
        "",
        "https://example.com/main.jpg",
        "",
    ])
    ws.append([
        "SKU-2",
        "Vintage Luggage Set",
        "Green",
        "ABS",
        "",
        "",
        "",
        "",
        27,
        17,
        11,
        12,
        "Suitcase with duffel bag and toiletry bag.",
        "Duffel integration",
        "Expandable capacity",
        "Double spinner wheels",
        "Side mounted TSA lock",
        "Silicone handles",
        "",
        "",
        "",
        "",
        "",
        "https://example.com/main2.jpg",
        "",
    ])
    wb.save(path)


def run_cli(*args: str, cwd: Path = ROOT) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        cwd=cwd,
        text=True,
        capture_output=True,
        check=False,
    )


def test_export_source_writes_row_json(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    output = tmp_path / "rows.json"
    create_source_workbook(source)

    result = run_cli("export-source", str(source), "--output", str(output))

    assert result.returncode == 0, result.stderr
    rows = json.loads(output.read_text())
    assert len(rows) == 2
    assert rows[0]["rowIndex"] == 2
    assert rows[0]["itemCode"] == "SKU-1"
    assert rows[0]["source"]["Product Name"] == "Expandable Hardside Luggage"
    assert rows[1]["rowIndex"] == 3
    assert rows[1]["itemCode"] == "SKU-2"


def test_export_source_fails_when_required_header_missing(tmp_path: Path) -> None:
    source = tmp_path / "bad-source.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Product Info"
    ws.append(["Item Code", "Product Name"])
    ws.append(["SKU-1", "Bad row"])
    wb.save(source)

    result = run_cli("export-source", str(source))

    assert result.returncode == 2
    assert "Missing required headers" in result.stderr
    assert "Description" in result.stderr


def test_write_output_creates_target_schema(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    generated = tmp_path / "generated.json"
    output = tmp_path / "amazon-listings.xlsx"
    create_source_workbook(source)
    generated.write_text(json.dumps([
        {
            "rowIndex": 2,
            "itemCode": "SKU-1",
            "productTitle": "Expandable Hardside Luggage With TSA Lock Spinner Wheels Black",
            "productDescription": "<b>Travel Ready</b><br>Durable suitcase for organized trips.",
            "searchTerms": "carryon baggage spinner suitcase travel case",
            "bulletPoint1": "1. EXPANDABLE STORAGE - Adds flexible packing room.",
            "bulletPoint2": "2. TSA SECURITY - Helps protect packed belongings.",
            "bulletPoint3": "3. SMOOTH MOBILITY - Spinner wheels move easily.",
            "bulletPoint4": "4. ORGANIZED INTERIOR - Lining helps separate items.",
            "bulletPoint5": "5. EASY STORAGE - Nested shape saves closet space.",
        },
        {
            "rowIndex": 3,
            "itemCode": "SKU-2",
            "productTitle": "Vintage Luggage Set With Duffel Toiletry Bag Spinner Wheels Green",
            "productDescription": "<b>Complete Set</b><br>Includes suitcase, duffel, and toiletry bag.",
            "searchTerms": "baggage trunk roller holiday organizer",
            "bulletPoint1": "1. COMPLETE SET - Includes coordinated travel pieces.",
            "bulletPoint2": "2. DUFFEL SLEEVE - Attaches to suitcase handle.",
            "bulletPoint3": "3. SECURE LOCK - Side lock supports travel security.",
            "bulletPoint4": "4. EASY ROLLING - Double wheels support movement.",
            "bulletPoint5": "5. ABS SHELL - Hardside body supports frequent trips.",
        },
    ]))

    result = run_cli("write-output", str(source), str(generated), "--output", str(output))

    assert result.returncode == 0, result.stderr
    wb = load_workbook(output)
    assert wb.sheetnames == ["listings"]
    ws = wb["listings"]
    assert [cell.value for cell in ws[1]] == TARGET_HEADERS
    assert ws.max_row == 3
    assert ws["A2"].value == "SKU-1"
    assert ws["B3"].value.startswith("Vintage Luggage Set")


def test_write_output_rejects_row_count_mismatch(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    generated = tmp_path / "generated.json"
    create_source_workbook(source)
    generated.write_text("[]")

    result = run_cli("write-output", str(source), str(generated))

    assert result.returncode == 2
    assert "Generated row count 0 does not match source row count 2" in result.stderr


def test_validate_output_accepts_matching_workbook(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    generated = tmp_path / "generated.json"
    output = tmp_path / "amazon-listings.xlsx"
    create_source_workbook(source)
    generated.write_text(json.dumps([
        {
            "rowIndex": 2,
            "itemCode": "SKU-1",
            "productTitle": "Title One",
            "productDescription": "<b>Description One</b>",
            "searchTerms": "alpha beta",
            "bulletPoint1": "1. ONE - Content",
            "bulletPoint2": "2. TWO - Content",
            "bulletPoint3": "3. THREE - Content",
            "bulletPoint4": "4. FOUR - Content",
            "bulletPoint5": "5. FIVE - Content",
        },
        {
            "rowIndex": 3,
            "itemCode": "SKU-2",
            "productTitle": "Title Two",
            "productDescription": "<b>Description Two</b>",
            "searchTerms": "gamma delta",
            "bulletPoint1": "1. ONE - Content",
            "bulletPoint2": "2. TWO - Content",
            "bulletPoint3": "3. THREE - Content",
            "bulletPoint4": "4. FOUR - Content",
            "bulletPoint5": "5. FIVE - Content",
        },
    ]))
    write_result = run_cli("write-output", str(source), str(generated), "--output", str(output))
    assert write_result.returncode == 0, write_result.stderr

    result = run_cli("validate-output", str(source), str(output))

    assert result.returncode == 0, result.stderr
    assert "Output workbook is valid" in result.stdout
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
pytest tests/test_listing_workbook.py -v
```

Expected: tests fail because `amazon-listing-xlsx/scripts/listing_workbook.py` does not exist.

- [ ] **Step 3: Commit failing tests**

```bash
git add tests/test_listing_workbook.py
git commit -m "test: define listing workbook utility behavior"
```

---

### Task 3: Implement Workbook CLI Utility

**Files:**
- Create: `amazon-listing-xlsx/scripts/listing_workbook.py`

- [ ] **Step 1: Create implementation**

Add this complete script:

```python
#!/usr/bin/env python3
import argparse
import json
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


SOURCE_SHEET = "Product Info"
TARGET_SHEET = "listings"

REQUIRED_SOURCE_HEADERS = [
    "Item Code",
    "Product Name",
    "Main Color",
    "Main Material",
    "Description",
]

TARGET_HEADERS = [
    "itemCode",
    "productTitle",
    "productDescription",
    "searchTerms",
    "bulletPoint1",
    "bulletPoint2",
    "bulletPoint3",
    "bulletPoint4",
    "bulletPoint5",
]


class WorkbookError(Exception):
    pass


def cell_to_json_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def load_source_sheet(path: Path) -> Worksheet:
    if not path.exists():
        raise WorkbookError(f"Input file does not exist: {path}")
    try:
        workbook = load_workbook(path, data_only=True)
    except Exception as exc:
        raise WorkbookError(f"Input file is not a readable .xlsx workbook: {path}") from exc
    if SOURCE_SHEET not in workbook.sheetnames:
        raise WorkbookError(f"Workbook does not contain required sheet: {SOURCE_SHEET}")
    return workbook[SOURCE_SHEET]


def read_headers(sheet: Worksheet) -> list[str]:
    return [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]


def validate_source_headers(headers: list[str]) -> None:
    missing = [header for header in REQUIRED_SOURCE_HEADERS if header not in headers]
    feature_headers = [f"Product Features {index}" for index in range(1, 6)]
    missing.extend(header for header in feature_headers if header not in headers)
    if missing:
        raise WorkbookError("Missing required headers: " + ", ".join(missing))


def source_rows(path: Path) -> list[dict[str, Any]]:
    sheet = load_source_sheet(path)
    headers = read_headers(sheet)
    validate_source_headers(headers)
    rows: list[dict[str, Any]] = []
    for row_index in range(2, sheet.max_row + 1):
        values = [cell_to_json_value(sheet.cell(row_index, column).value) for column in range(1, len(headers) + 1)]
        source = dict(zip(headers, values))
        if not any(value != "" for value in source.values()):
            continue
        item_code = str(source.get("Item Code", "")).strip()
        if not item_code:
            raise WorkbookError(f"Source row {row_index} is missing Item Code")
        rows.append({
            "rowIndex": row_index,
            "itemCode": item_code,
            "source": source,
        })
    if not rows:
        raise WorkbookError("No product data rows found")
    return rows


def read_generated_rows(path: Path) -> list[dict[str, Any]]:
    try:
        data = json.loads(path.read_text())
    except Exception as exc:
        raise WorkbookError(f"Generated listing JSON is not readable: {path}") from exc
    if not isinstance(data, list):
        raise WorkbookError("Generated listing JSON must be an array")
    for index, row in enumerate(data):
        if not isinstance(row, dict):
            raise WorkbookError(f"Generated row at index {index} must be an object")
    return data


def normalize_generated_row(row: dict[str, Any]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for header in TARGET_HEADERS:
        value = row.get(header, "")
        normalized[header] = "" if value is None else str(value).replace("\r", " ").replace("\n", " ").strip()
    return normalized


def validate_generated_rows(source: list[dict[str, Any]], generated: list[dict[str, Any]]) -> None:
    if len(generated) != len(source):
        raise WorkbookError(f"Generated row count {len(generated)} does not match source row count {len(source)}")
    source_by_row = {row["rowIndex"]: row for row in source}
    for row in generated:
        row_index = row.get("rowIndex")
        if row_index not in source_by_row:
            raise WorkbookError(f"Generated row has unknown rowIndex: {row_index}")
        expected_item_code = source_by_row[row_index]["itemCode"]
        actual_item_code = str(row.get("itemCode", "")).strip()
        if actual_item_code and actual_item_code != expected_item_code:
            raise WorkbookError(f"Generated itemCode {actual_item_code} does not match source itemCode {expected_item_code}")


def write_output_workbook(source_path: Path, generated_json_path: Path, output_path: Path) -> None:
    source = source_rows(source_path)
    generated = read_generated_rows(generated_json_path)
    validate_generated_rows(source, generated)

    generated_by_row = {row["rowIndex"]: row for row in generated}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = TARGET_SHEET
    sheet.append(TARGET_HEADERS)

    for source_row in source:
        generated_row = dict(generated_by_row[source_row["rowIndex"]])
        generated_row["itemCode"] = source_row["itemCode"]
        normalized = normalize_generated_row(generated_row)
        sheet.append([normalized[header] for header in TARGET_HEADERS])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    validate_output_workbook(source_path, output_path)


def validate_output_workbook(source_path: Path, output_path: Path) -> None:
    source = source_rows(source_path)
    if not output_path.exists():
        raise WorkbookError(f"Output file does not exist: {output_path}")
    try:
        workbook = load_workbook(output_path, data_only=True)
    except Exception as exc:
        raise WorkbookError(f"Output workbook is not readable: {output_path}") from exc
    if TARGET_SHEET not in workbook.sheetnames:
        raise WorkbookError(f"Output workbook does not contain sheet: {TARGET_SHEET}")
    sheet = workbook[TARGET_SHEET]
    headers = read_headers(sheet)
    if headers != TARGET_HEADERS:
        raise WorkbookError(f"Output headers do not match target schema: {headers}")
    output_count = max(sheet.max_row - 1, 0)
    if output_count != len(source):
        raise WorkbookError(f"Output row count {output_count} does not match source row count {len(source)}")
    for row_index in range(2, sheet.max_row + 1):
        values = {header: sheet.cell(row_index, column).value for column, header in enumerate(TARGET_HEADERS, start=1)}
        missing = [header for header, value in values.items() if value in (None, "")]
        if missing:
            raise WorkbookError(f"Output row {row_index} has empty required fields: {', '.join(missing)}")


def default_output_path(source_path: Path) -> Path:
    return source_path.with_name(f"{source_path.stem}.amazon-listings.xlsx")


def command_export_source(args: argparse.Namespace) -> None:
    rows = source_rows(Path(args.source))
    payload = json.dumps(rows, ensure_ascii=False, indent=2)
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload)
    else:
        print(payload)


def command_write_output(args: argparse.Namespace) -> None:
    source_path = Path(args.source)
    output_path = Path(args.output) if args.output else default_output_path(source_path)
    write_output_workbook(source_path, Path(args.generated_json), output_path)
    print(f"Wrote Amazon listing workbook: {output_path}")


def command_validate_output(args: argparse.Namespace) -> None:
    validate_output_workbook(Path(args.source), Path(args.output))
    print("Output workbook is valid")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Prepare and validate Amazon listing XLSX workbooks.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    export_parser = subparsers.add_parser("export-source", help="Export Product Info rows as JSON for listing generation.")
    export_parser.add_argument("source")
    export_parser.add_argument("--output")
    export_parser.set_defaults(func=command_export_source)

    write_parser = subparsers.add_parser("write-output", help="Write generated listing JSON into target XLSX schema.")
    write_parser.add_argument("source")
    write_parser.add_argument("generated_json")
    write_parser.add_argument("--output")
    write_parser.set_defaults(func=command_write_output)

    validate_parser = subparsers.add_parser("validate-output", help="Validate generated XLSX against source workbook.")
    validate_parser.add_argument("source")
    validate_parser.add_argument("output")
    validate_parser.set_defaults(func=command_validate_output)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.func(args)
    except WorkbookError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Run tests**

Run:

```bash
pytest tests/test_listing_workbook.py -v
```

Expected: all tests pass.

- [ ] **Step 3: Run source sample export**

Run:

```bash
SKILL_DIR=amazon-listing-xlsx
python "$SKILL_DIR/scripts/listing_workbook.py" export-source source.xlsx --output build/source.rows.json
```

Expected: command exits `0` and creates `build/source.rows.json`.

- [ ] **Step 4: Inspect exported row count**

Run:

```bash
python - <<'PY'
import json
from pathlib import Path
rows = json.loads(Path("build/source.rows.json").read_text())
print(len(rows))
print(rows[0]["itemCode"])
PY
```

Expected output starts with:

```text
20
PP191030AAU
```

- [ ] **Step 5: Commit**

```bash
git add amazon-listing-xlsx/scripts/listing_workbook.py tests/test_listing_workbook.py
git commit -m "feat: add listing workbook utility"
```

---

### Task 4: Add Prompt Reference

**Files:**
- Create: `amazon-listing-xlsx/references/listing_prompt_rules.md`

- [ ] **Step 1: Create prompt reference**

Add this content:

````markdown
# Listing Prompt Rules

Use this reference when generating Amazon listing content from exported source rows.

## Role

Act as a senior Amazon Listing copywriter and SEO specialist. Generate compliant, high-conversion listing copy from product source data.

## Language

Default language is English.

English output rules:

- Use native English tone.
- Use standard ASCII characters only.
- Do not output Chinese or other languages.
- Do not use emoji.

## Required Output Fields

For every source row, generate one object with:

- `rowIndex`
- `itemCode`
- `productTitle`
- `productDescription`
- `searchTerms`
- `bulletPoint1`
- `bulletPoint2`
- `bulletPoint3`
- `bulletPoint4`
- `bulletPoint5`

Keep `rowIndex` exactly as provided by the exported source JSON.

## Product Title

- Target length: 150 to 200 characters including spaces.
- Use readable Amazon SEO wording.
- Include core keyword, material, key feature, function or usage scenario, color, size, or pack quantity when available.
- Capitalize major words.

## Bullet Points

- Generate exactly five bullet points.
- Put the strongest benefits in the first three bullets.
- Format each bullet as `N. HEADER - Detailed explanation.`
- Focus on customer benefits, not only raw features.
- Keep each bullet on one physical line.

## Product Description

- Target length: 1500 to 1800 characters when enough source data exists.
- Use `<b>` for short section headers.
- Use `<br>` for line breaks.
- Do not use actual newline characters inside the field.
- Include product story, materials, dimensions, usage scenarios, package contents, and instructions when source data provides them.

## Search Terms

- Target length: 200 to 250 characters when enough relevant terms exist.
- Use single words separated by spaces.
- Do not use commas or other punctuation.
- Include relevant synonyms and supplemental terms not already overused in title and bullets.
- Do not use marketing words such as `best`, `sale`, or `offer`.

## Compliance

Avoid Amazon-sensitive wording unless the source explicitly provides required certification.

Never use these pesticide-related terms for ordinary products:

- `anti-microbial`
- `anti-bacterial`
- `anti-fungal`
- `harmless`
- `non-toxic`
- `repellent`
- `virus`
- `germs`

Do not use unrelated brand names, trademarked names, weapons, drug references, or unsupported medical/safety claims.

## Batch Output Format

Return strict JSON only. Do not use Markdown fences or explanatory text.

```json
[
  {
    "rowIndex": 2,
    "itemCode": "SKU-1",
    "productTitle": "Generated title",
    "productDescription": "<b>Generated</b><br>Description",
    "searchTerms": "keyword keyword",
    "bulletPoint1": "1. HEADER - Content",
    "bulletPoint2": "2. HEADER - Content",
    "bulletPoint3": "3. HEADER - Content",
    "bulletPoint4": "4. HEADER - Content",
    "bulletPoint5": "5. HEADER - Content"
  }
]
```
````

- [ ] **Step 2: Commit**

```bash
git add amazon-listing-xlsx/references/listing_prompt_rules.md
git commit -m "docs: add listing prompt reference"
```

---

### Task 5: Add Skill Instructions

**Files:**
- Create: `amazon-listing-xlsx/SKILL.md`

- [ ] **Step 1: Create skill instructions**

Add this content:

````markdown
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
- Source `Item Code` values must be non-empty for every product row.

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
   - Final target fields must be non-empty.
   - If a field lacks enough source information, stop and report the issue or write reasonable listing copy from available product context instead of using an empty string.

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
- Reject source rows with missing `Item Code`; do not create fallback item codes.
- Keep generated description line breaks as `<br>`, not actual newline characters.
- If validation fails, fix the generated JSON or workbook issue and rerun validation before responding.

## References

- Listing copy rules: `$SKILL_DIR/references/listing_prompt_rules.md`
- Workbook utility: `$SKILL_DIR/scripts/listing_workbook.py`
```
````

- [ ] **Step 2: Run skill metadata sanity check**

Run:

```bash
python - <<'PY'
from pathlib import Path
text = Path("amazon-listing-xlsx/SKILL.md").read_text()
assert text.startswith("---\nname: amazon-listing-xlsx\n")
assert "description:" in text
assert "scripts/listing_workbook.py" in text
print("SKILL.md looks valid")
PY
```

Expected:

```text
SKILL.md looks valid
```

- [ ] **Step 3: Commit**

```bash
git add amazon-listing-xlsx/SKILL.md
git commit -m "feat: add amazon listing xlsx skill"
```

---

### Task 6: End-to-End Sample Validation

**Files:**
- Create during verification only: `build/source.rows.json`
- Create during verification only: `build/generated-listings.json`
- Create during verification only: `build/source.amazon-listings.xlsx`

- [ ] **Step 1: Export sample source**

Run:

```bash
mkdir -p build
SKILL_DIR=amazon-listing-xlsx
python "$SKILL_DIR/scripts/listing_workbook.py" export-source source.xlsx --output build/source.rows.json
```

Expected: exits `0`.

- [ ] **Step 2: Create deterministic sample generated JSON from target workbook**

Run:

```bash
python - <<'PY'
import json
from pathlib import Path
from openpyxl import load_workbook

wb = load_workbook("target.xlsx", data_only=True)
ws = wb["listings"]
headers = [cell.value for cell in ws[1]]
source_rows = json.loads(Path("build/source.rows.json").read_text())
items = []
for index, source_row in enumerate(source_rows, start=2):
    row = {headers[col - 1]: ws.cell(index, col).value or "" for col in range(1, len(headers) + 1)}
    items.append({
        "rowIndex": source_row["rowIndex"],
        "itemCode": row["itemCode"],
        "productTitle": row["productTitle"],
        "productDescription": row["productDescription"],
        "searchTerms": row["searchTerms"],
        "bulletPoint1": row["bulletPoint1"],
        "bulletPoint2": row["bulletPoint2"],
        "bulletPoint3": row["bulletPoint3"],
        "bulletPoint4": row["bulletPoint4"],
        "bulletPoint5": row["bulletPoint5"],
    })
Path("build/generated-listings.json").write_text(json.dumps(items, ensure_ascii=False, indent=2))
print(len(items))
PY
```

Expected:

```text
20
```

- [ ] **Step 3: Write output workbook**

Run:

```bash
SKILL_DIR=amazon-listing-xlsx
python "$SKILL_DIR/scripts/listing_workbook.py" write-output source.xlsx build/generated-listings.json --output build/source.amazon-listings.xlsx
```

Expected output:

```text
Wrote Amazon listing workbook: build/source.amazon-listings.xlsx
```

- [ ] **Step 4: Validate output workbook**

Run:

```bash
SKILL_DIR=amazon-listing-xlsx
python "$SKILL_DIR/scripts/listing_workbook.py" validate-output source.xlsx build/source.amazon-listings.xlsx
```

Expected output:

```text
Output workbook is valid
```

- [ ] **Step 5: Run full test suite**

Run:

```bash
pytest -v
```

Expected: all tests pass.

- [ ] **Step 6: Commit verification artifacts policy**

Do not commit `build/` outputs. If `git status --short` shows `build/`, remove it before commit:

```bash
rm -rf build
git status --short
```

Expected: only intended source files are tracked or modified.

---

### Task 7: Final Review and Commit State

**Files:**
- Review: all changed files

- [ ] **Step 1: Inspect git status**

Run:

```bash
git status --short
```

Expected: clean worktree after all task commits, or only intentionally uncommitted plan updates.

- [ ] **Step 2: Inspect recent commits**

Run:

```bash
git log --oneline -5
```

Expected: recent commits include dependency, tests, workbook utility, prompt reference, and skill instructions.

- [ ] **Step 3: Confirm final verification**

Run:

```bash
pytest -v
```

Expected: all tests pass.

---

## Self-Review

- Spec coverage: The plan covers skill creation, prompt rules, workbook parsing, writing, validation, fixed target schema, error handling, and end-to-end sample validation.
- Placeholder scan: No unresolved placeholder instructions remain.
- Type consistency: CLI commands, JSON keys, workbook headers, and test expectations use the same names across tasks.
- Implementation synchronization: `Item Code` is required for every non-empty source data row; the workbook script fails instead of creating fallback item codes. Installed skill usage resolves `SKILL_DIR` before invoking bundled scripts, while repo-root script paths remain only in the historical CLI contract and file path references.
