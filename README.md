# FileMaker ExcelJS Wrapper — Developer Reference

## Overview

The ExcelJS Helper file provides a stateful, scriptable interface to ExcelJS. A WebViewer
on the **ExcelJS Worker** layout holds a live `ExcelJS.Workbook` in memory. FileMaker
scripts call JavaScript functions incrementally via **Perform JavaScript in Web Viewer**
to build up the workbook sheet-by-sheet, then trigger a final export.

FileMaker developers use the wrapper scripts without needing to know JavaScript or ExcelJS.

---

## Helper File Structure

**File name:** `ExcelJS Helper.fmp12`

### Global Fields

| Field           | Type             | Purpose                                     |
|-----------------|------------------|---------------------------------------------|
| `g_TempStorage` | Global Container | Receives the decoded `.xlsx` during export  |
| `g_Status`      | Global Text      | Last status/error string from the WebViewer |

### Layout

**ExcelJS Worker** — contains a single WebViewer pointing at `exceljs_webviewer.html`.
The WebViewer shows a status panel reflecting the last operation. No auto-run on load.

---

## Script Reference

Every script calls **Perform JavaScript in Web Viewer** and stores the JS return value
in the global variable `$$ExcelJS_LastResult` (a JSON string).

Check success with:
```
JSONGetElement ( $$ExcelJS_LastResult ; "success" )   // returns 1 (true) or 0 (false)
JSONGetElement ( $$ExcelJS_LastResult ; "error" )     // error message string (if any)
```

> **Note on `ExcelJS - Finalize`:** This script is asynchronous internally.
> The `$$ExcelJS_LastResult` will contain `{ "success": true, "pending": true }` —
> meaning the operation *started*, not that the file is ready. The actual workbook
> arrives via `ExcelJS - Receive Workbook` being called back from JavaScript.
> Check `$$ExcelJS_LastResult` from *that* script for the final outcome.

---

### ExcelJS - Init Workbook

Resets the in-memory workbook. Always call this first.

**JS function:** `excelJS_init()`

**Parameters:** _(none)_

**FileMaker step:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_init"
]
Set Variable [ $$ExcelJS_LastResult ; Result: Get(ScriptResult) ]
```

---

### ExcelJS - Add Sheet

Adds a worksheet to the workbook.

**JS function:** `excelJS_addSheet(name, optionsJson)`

| Parameter     | Type                 | Notes             |
|---------------|----------------------|-------------------|
| `name`        | Text                 | Sheet tab name    |
| `optionsJson` | JSON Text (optional) | See options below |

**Options JSON keys:**

| Key         | Type        | Description           |
|-------------|-------------|-----------------------|
| `hidden`    | bool        | Hide from tab bar     |
| `tabColor`  | `"#rrggbb"` | Tab accent colour     |
| `frozenRow` | number      | Freeze top N rows     |
| `frozenCol` | number      | Freeze left N columns |

**Example — hidden valuelist sheet:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addSheet" ;
  Parameters: "BudgetItems_VL" ; JSONSetElement ( "{}" ; "hidden" ; True ; JSONBoolean )
]
```

**Example — main sheet with frozen header row:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addSheet" ;
  Parameters: "Change Order" ; JSONSetElement ( "{}" ; "frozenRow" ; 1 ; JSONNumber )
]
```

---

### ExcelJS - Set Sheet Visibility

Show or hide an existing sheet.

**JS function:** `excelJS_setSheetVisibility(sheetName, visible)`

| Parameter   | Type    | Notes             |
|-------------|---------|-------------------|
| `sheetName` | Text    | Sheet tab name    |
| `visible`   | Boolean | `True` or `False` |

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_setSheetVisibility" ;
  Parameters: "BudgetItems_VL" ; False
]
```

---

### ExcelJS - Set Columns

Defines columns on a sheet (keys, headers, widths, hidden flag, default style).
Must be called before adding rows.

**JS function:** `excelJS_setColumns(sheetName, columnsJson)`

| Parameter     | Type            | Notes                              |
|---------------|-----------------|------------------------------------|
| `sheetName`   | Text            | Sheet tab name                     |
| `columnsJson` | JSON Array Text | Array of column definition objects |

**Column definition object keys:**

| Key      | Type             | Required  | Description                                      |
|----------|------------------|-----------|--------------------------------------------------|
| `key`    | text             | yes       | Internal key, used in row objects                |
| `header` | text             | no        | Column header label (row 1 if added before rows) |
| `width`  | number           | no        | Column width in character units                  |
| `hidden` | bool             | no        | Hide the column in Excel                         |
| `style`  | preset or object | no        | Default cell style (see Style Presets)           |

**Example:**
```
Set Variable [ $cols ;
  JSONSetElement ( "[]" ;
    [ "[0].key"    ; "budget_item" ; JSONString ] ;
    [ "[0].width"  ; 50            ; JSONNumber ] ;
    [ "[0].style"  ; "dropdown"    ; JSONString ] ;
    [ "[1].key"    ; "hours"       ; JSONString ] ;
    [ "[1].width"  ; 11            ; JSONNumber ] ;
    [ "[1].style"  ; "number"      ; JSONString ] ;
    [ "[2].key"    ; "hidden_id_0" ; JSONString ] ;
    [ "[2].width"  ; 11            ; JSONNumber ] ;
    [ "[2].hidden" ; True          ; JSONBoolean ]
  )
]
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_setColumns" ;
  Parameters: "Change Order" ; $cols
]
```

---

### ExcelJS - Add Row

Adds a single row to a sheet.

**JS function:** `excelJS_addRow(sheetName, rowJson)`

| Parameter   | Type             | Notes                       |
|-------------|------------------|-----------------------------|
| `sheetName` | Text             | Sheet tab name              |
| `rowJson`   | JSON Object Text | Keys must match column keys |

**Example — header row:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addRow" ;
  Parameters: "Change Order" ;
    JSONSetElement ( "{}" ;
      [ "budget_item" ; "Budget Item" ; JSONString ] ;
      [ "hours"       ; "Hours"       ; JSONString ] ;
      [ "hidden_id_0" ; "budget_id"   ; JSONString ]
    )
]
```

---

### ExcelJS - Add Rows

Adds multiple rows in a single JavaScript call. Use this instead of N calls to
`ExcelJS - Add Row` when inserting large datasets.

**JS function:** `excelJS_addRows(sheetName, rowsJson)`

| Parameter   | Type            | Notes                |
|-------------|-----------------|----------------------|
| `sheetName` | Text            | Sheet tab name       |
| `rowsJson`  | JSON Array Text | Array of row objects |

**Example — 50 blank data rows:**
```
Set Variable [ $rows ; While (
  [ _json = "[]" ; _i = 0 ] ;
  _i < 50 ;
  [ _json = JSONSetElement ( _json ; "[" & _i & "].budget_item" ; "" ; JSONString ) ;
    _i = _i + 1 ] ;
  _json
)]
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addRows" ;
  Parameters: "Change Order" ; $rows
]
```

---

### ExcelJS - Set Cell Value

Sets a single cell value by address.

**JS function:** `excelJS_setCellValue(sheetName, cellRef, value)`

| Parameter   | Type          | Notes               |
|-------------|---------------|---------------------|
| `sheetName` | Text          | Sheet tab name      |
| `cellRef`   | Text          | e.g. `"A1"`, `"B5"` |
| `value`     | Text / Number | Cell value          |

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_setCellValue" ;
  Parameters: "Change Order" ; "D1" ; "Total Cost"
]
```

---

### ExcelJS - Set Cell Formula

Sets a formula in a single cell.

**JS function:** `excelJS_setFormula(sheetName, cellRef, formula)`

| Parameter   | Type  | Notes                               |
|-------------|-------|-------------------------------------|
| `sheetName` | Text  | Sheet tab name                      |
| `cellRef`   | Text  | e.g. `"C52"`                        |
| `formula`   | Text  | Excel formula, leading `=` optional |

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_setFormula" ;
  Parameters: "Change Order" ; "C52" ; "=SUM(C2:C51)"
]
```

---

### ExcelJS - Set Formula For Range

Applies a per-row formula to a column across a range of rows.
Use `{ROW}` in the template as a placeholder for the current row number.

**JS function:** `excelJS_setFormulaRange(sheetName, startRow, endRow, colKey, formulaTemplate)`

| Parameter         | Type   | Notes                                |
|-------------------|--------|--------------------------------------|
| `sheetName`       | Text   | Sheet tab name                       |
| `startRow`        | Number | First row number                     |
| `endRow`          | Number | Last row number (inclusive)          |
| `colKey`          | Text   | Column key as defined in Set Columns |
| `formulaTemplate` | Text   | Formula with `{ROW}` placeholder     |

**Example — INDEX/MATCH to look up budget item ID:**
```
Set Variable [ $template ;
  "=IF(A{ROW}=\"\",\"\",INDEX(BudgetItems_VL!$A$2:$A$999,MATCH(A{ROW},BudgetItems_VL!$B$2:$B$999,0)))"
]
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_setFormulaRange" ;
  Parameters: "Change Order" ; 2 ; 51 ; "hidden_id_0" ; $template
]
```

---

### ExcelJS - Style Row

Applies a style to every cell in a row. Named presets also set row height.

**JS function:** `excelJS_styleRow(sheetName, rowNum, presetOrStyleJson)`

| Parameter           | Type   | Notes                                        |
|---------------------|--------|----------------------------------------------|
| `sheetName`         | Text   | Sheet tab name                               |
| `rowNum`            | Number | 1-based row number                           |
| `presetOrStyleJson` | Text   | Preset name string OR raw ExcelJS style JSON |

**Example — style row 1 as a header:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_styleRow" ;
  Parameters: "Change Order" ; 1 ; "header"
]
```

**Example — custom style (raw JSON):**
```
Set Variable [ $style ;
  JSONSetElement ( "{}" ;
    [ "font.bold"        ; True        ; JSONBoolean ] ;
    [ "fill.type"        ; "pattern"   ; JSONString  ] ;
    [ "fill.pattern"     ; "solid"     ; JSONString  ] ;
    [ "fill.fgColor.argb"; "FFFFF2CC"  ; JSONString  ]
  )
]
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_styleRow" ;
  Parameters: "Change Order" ; 2 ; $style
]
```

---

### ExcelJS - Style Cell

Applies a style to a single cell.

**JS function:** `excelJS_styleCell(sheetName, cellRef, presetOrStyleJson)`

| Parameter           | Type  | Notes                                        |
|---------------------|-------|----------------------------------------------|
| `sheetName`         | Text  | Sheet tab name                               |
| `cellRef`           | Text  | e.g. `"B3"`                                  |
| `presetOrStyleJson` | Text  | Preset name string OR raw ExcelJS style JSON |

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_styleCell" ;
  Parameters: "Change Order" ; "C52" ; "currency"
]
```

---

### ExcelJS - Freeze Panes

Freezes rows and/or columns. Pass `0` for an axis you don't want to freeze.

**JS function:** `excelJS_freezePane(sheetName, row, col)`

| Parameter   | Type   | Notes                                  |
|-------------|--------|----------------------------------------|
| `sheetName` | Text   | Sheet tab name                         |
| `row`       | Number | Rows to freeze from top (0 = none)     |
| `col`       | Number | Columns to freeze from left (0 = none) |

**Example — freeze top row only:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_freezePane" ;
  Parameters: "Change Order" ; 1 ; 0
]
```

> Note: `excelJS_addSheet` with `frozenRow`/`frozenCol` options is the easiest way
> to freeze at sheet-creation time. Call `excelJS_freezePane` when you need to
> adjust the freeze after the sheet already exists.

---

### ExcelJS - Add Defined Name

Registers a workbook-level named range. Use defined names in dropdown formulas to
avoid cross-sheet reference quoting issues.

**JS function:** `excelJS_addDefinedName(name, ref)`

| Parameter  | Type  | Notes                                                    |
|------------|-------|----------------------------------------------------------|
| `name`     | Text  | Defined name (no spaces; underscores OK)                 |
| `ref`      | Text  | Fully-qualified range, e.g. `"'Sheet Name'!$B$2:$B$100"` |

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addDefinedName" ;
  Parameters: "BudgetItems_range" ; "'BudgetItems_VL'!$B$2:$B$999"
]
```

---

### ExcelJS - Add Data Validation Dropdown

Adds an Excel dropdown list validation to a cell range.

**JS function:** `excelJS_addDropdown(sheetName, range, optionsJson)`

| Parameter     | Type             | Notes                         |
|---------------|------------------|-------------------------------|
| `sheetName`   | Text             | Sheet tab name                |
| `range`       | Text             | Cell range, e.g. `"A2:A51"`   |
| `optionsJson` | JSON Object Text | Either `formula1` or `values` |

**optionsJson variants:**

| Key        | Type               | Description                                      |
|------------|--------------------|--------------------------------------------------|
| `formula1` | text               | Defined name or range reference                  |
| `values`   | JSON array of text | Inline list (keep total length under ~255 chars) |

**Example — using a defined name:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addDropdown" ;
  Parameters: "Change Order" ; "A2:A51" ;
    JSONSetElement ( "{}" ; "formula1" ; "BudgetItems_range" ; JSONString )
]
```

**Example — inline list:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_addDropdown" ;
  Parameters: "Change Order" ; "B2:B51" ;
    JSONSetElement ( "{}" ; "values" ;
      JSONSetElement ( "[]" ;
        [ "[0]" ; "Pending"   ; JSONString ] ;
        [ "[1]" ; "Approved"  ; JSONString ] ;
        [ "[2]" ; "Rejected"  ; JSONString ]
      ) ; JSONArray
    )
]
```

---

### ExcelJS - Finalize

Generates the `.xlsx` buffer and delivers it to FileMaker via callback.

**JS function:** `excelJS_finalize(filename)`

| Parameter  | Type  | Notes                                        |
|------------|-------|----------------------------------------------|
| `filename` | Text  | Desired filename, e.g. `"Change Order.xlsx"` |

**Synchronous return** (stored in `$$ExcelJS_LastResult`):
```json
{ "success": true, "pending": true, "message": "..." }
```
This means the export has *started*, not completed. The workbook arrives via the
`ExcelJS - Receive Workbook` callback script (see below).

**Example:**
```
Perform JavaScript in Web Viewer [
  Object Name: "ExcelJS Worker" ;
  Function Name: "excelJS_finalize" ;
  Parameters: "Change Order " & TEXT_HELPER::formatted_date & ".xlsx"
]
Set Variable [ $$ExcelJS_LastResult ; Result: Get(ScriptResult) ]
// No pause needed — FileMaker.PerformScript triggers ExcelJS - Receive Workbook
```

---

### ExcelJS - Receive Workbook

**This script is called *by JavaScript*, not by the developer directly.**

JavaScript calls `FileMaker.PerformScript("ExcelJS - Receive Workbook", payload)` where
`payload` is a JSON string containing either:

- Success: `{ "base64": "...", "filename": "Change Order.xlsx" }`
- Error:   `{ "error": "Workbook not initialized" }`

**Script logic:**

```
Set Variable [ $payload  ; Get(ScriptParameter) ]
Set Variable [ $error    ; JSONGetElement ( $payload ; "error" ) ]

If [ $error ≠ "" ]
  Set Variable [ $$ExcelJS_LastResult ; JSONSetElement ( "{}" ;
    [ "success" ; False     ; JSONBoolean ] ;
    [ "error"   ; $error    ; JSONString  ] ) ]
  // Optionally show error dialog or log
  Exit Script [ Text: $$ExcelJS_LastResult ]
End If

Set Variable [ $b64      ; JSONGetElement ( $payload ; "base64"   ) ]
Set Variable [ $filename ; JSONGetElement ( $payload ; "filename" ) ]

// Decode base64 → container field
Set Field [ ExcelJS Helper::g_TempStorage ; Base64Decode ( $b64 ; $filename ) ]

// Prompt user to save
Export Field Contents [ ExcelJS Helper::g_TempStorage ; "$filename" ]

Set Variable [ $$ExcelJS_LastResult ; JSONSetElement ( "{}" ;
  "success" ; True ; JSONBoolean ) ]
```

---

## Style Presets Reference

Pass any preset name as a plain text string to `excelJS_styleRow`, `excelJS_styleCell`,
or the `style` key of a column definition.

| Preset           | Description                                                                | Row Height  |
|------------------|----------------------------------------------------------------------------|-------------|
| `"header"`       | Blue fill (#4472C4), white bold 11pt, center-aligned, medium bottom border | 28          |
| `"hiddenHeader"` | Gray fill (#D9D9D9), black 11pt, center-aligned, medium bottom border      | —           |
| `"listHeader"`   | Light blue fill (#BDD7EE), black 11pt, center-aligned                      | —           |
| `"currency"`     | Number format `$#,##0.00`, right-aligned                                   | —           |
| `"number"`       | Number format `#,##0.##`, right-aligned                                    | —           |
| `"date"`         | Date format `MM/DD/YYYY`, center-aligned                                   | —           |
| `"dropdown"`     | Yellow fill (#FFFF99) to indicate user-editable cells                      | —           |
| `"text"`         | Format `@` (force text), left-aligned                                      | —           |
| `"center"`       | Center-aligned, no other changes                                           | —           |
| `"bold"`         | Bold 11pt Calibri, no other changes                                        | —           |

For full control, pass a raw ExcelJS style JSON object instead of a preset name:
```json
{
  "font":      { "name": "Calibri", "size": 11, "bold": true },
  "fill":      { "type": "pattern", "pattern": "solid", "fgColor": { "argb": "FFFFE699" } },
  "alignment": { "horizontal": "center", "vertical": "middle" },
  "numFmt":    "$#,##0.00"
}
```

---

## End-to-End Example: Change Order Script

This is the refactored equivalent of the one-shot `excel_template.html` approach,
built as a sequence of wrapper script calls.

```
# ── Init ──────────────────────────────────────────────────────────────────
ExcelJS - Init Workbook

# ── Hidden valuelist sheet ────────────────────────────────────────────────
ExcelJS - Add Sheet
  name:        "BudgetItems_VL"
  optionsJson: { "hidden": true }

ExcelJS - Set Columns
  sheet:       "BudgetItems_VL"
  columns:     [ { key:"id",    width:39, hidden:true },
                 { key:"value", width:50 } ]

ExcelJS - Add Row           (header)
  sheet:       "BudgetItems_VL"
  row:         { "id": "id", "value": "value" }

ExcelJS - Style Row
  sheet:       "BudgetItems_VL"
  row:         1
  preset:      "listHeader"

ExcelJS - Add Rows          (data — use $json_budget_items from FileMaker)
  sheet:       "BudgetItems_VL"
  rows:        $json_budget_items       // [{ id:"uuid-1", value:"Contractor A | WO-1" }, ...]

# ── Defined name ──────────────────────────────────────────────────────────
ExcelJS - Add Defined Name
  name:        "BudgetItems_range"
  ref:         "'BudgetItems_VL'!$B$2:$B$999"

# ── Main worksheet ────────────────────────────────────────────────────────
ExcelJS - Add Sheet
  name:        "Change Order"
  optionsJson: { "frozenRow": 1 }

ExcelJS - Set Columns
  sheet:       "Change Order"
  columns:     [
    { key:"col_0", width:50, style:"dropdown" },  // Budget Item  (visible)
    { key:"col_1", width:11, style:"number"   },  // Hours        (visible)
    { key:"col_2", width:17, style:"currency" },  // Labour Cost  (visible)
    { key:"col_3", width:30, style:"text"     },  // Notes        (visible)
    { key:"hidden_id_0", width:11, hidden:true }  // Budget ID    (hidden)
  ]

ExcelJS - Add Row           (header row)
  sheet:       "Change Order"
  row:         { col_0:"Budget Item", col_1:"Hours", col_2:"Labour Cost",
                 col_3:"Notes", hidden_id_0:"budget_item_id" }

ExcelJS - Style Row
  sheet:       "Change Order"
  row:         1
  preset:      "header"             // also sets height to 28

ExcelJS - Add Rows          (50 blank data rows)
  sheet:       "Change Order"
  rows:        [ {}, {}, ... ]      // 50 objects with empty string values

# ── Formulas in hidden ID column ──────────────────────────────────────────
ExcelJS - Set Formula For Range
  sheet:       "Change Order"
  startRow:    2
  endRow:      51
  colKey:      "hidden_id_0"
  template:    "=IF(A{ROW}=\"\",\"\",INDEX(BudgetItems_VL!$A$2:$A$999,MATCH(A{ROW},BudgetItems_VL!$B$2:$B$999,0)))"

# ── Dropdown validation ───────────────────────────────────────────────────
ExcelJS - Add Data Validation Dropdown
  sheet:       "Change Order"
  range:       "A2:A51"
  options:     { "formula1": "BudgetItems_range" }

# ── Export ────────────────────────────────────────────────────────────────
ExcelJS - Finalize
  filename:    "Change Order " & $date & ".xlsx"
```

The `ExcelJS - Receive Workbook` callback script then decodes the base64 payload,
populates `g_TempStorage`, and presents the Save As dialog via **Export Field Contents**.

---

## Error Handling Pattern

After any wrapper script call, check `$$ExcelJS_LastResult`:

```
If [ JSONGetElement ( $$ExcelJS_LastResult ; "success" ) = False ]
  Show Custom Dialog [
    Title:   "Excel Export Error" ;
    Message: JSONGetElement ( $$ExcelJS_LastResult ; "error" )
  ]
  Exit Script [ Text: $$ExcelJS_LastResult ]
End If
```

**Common errors:**

| Error message                                 | Cause                                                    |
|-----------------------------------------------|----------------------------------------------------------|
| `Workbook not initialized`                    | Called any function before `excelJS_init`                |
| `Sheet not found: X`                          | `sheetName` argument doesn't match a sheet added earlier |
| `Column key not found: X`                     | `colKey` in `setFormulaRange` doesn't match a column key |
| `Unknown style preset: X`                     | Typo in preset name (see Style Presets table)            |
| `optionsJson must contain formula1 or values` | Dropdown options object is missing required key          |
