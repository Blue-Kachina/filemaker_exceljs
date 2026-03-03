# FileMaker ExcelJS Wrapper — Developer Reference

## Overview

We're loading up a FileMaker Webviewer with HTML and JavaScript to create a wrapper around 
the [ExcelJS](https://github.com/exceljs/exceljs) library.
Using the provided FileMaker Scripts, you can custom-build a rich Excel (XLSX) spreadsheet.  
You can do so incrementally too, because our webviewer holds a live `ExcelJS.Workbook` in memory. 

## General Workflow
1) Init Workbook
2) Incrementally build up your custom spreadsheet using provided functions (details below)
3) Finalize

---

## Helper File Structure

**File name:** `FileMaker_ExcelJS.fmp12`

### Global Fields

| Field         | Type             | Purpose                                     |
|---------------|------------------|---------------------------------------------|
| `tempStorage` | Global Container | Receives the decoded `.xlsx` during export  |

The finalize process base64-encodes the XLSX file, and sends it back to FileMaker Pro.  
FileMaker Pro then decodes it, and inserts it into the global container so that it can be downloaded

### Layout

**ExcelJS_Worker** — contains a single WebViewer that loads up the HTML from [webviewer.html](webviewer.html).
The WebViewer shows a status panel reflecting details from the last operation.

---

## Scripts
Our ultimate goal will be to provide A FileMaker Pro Script wrapper for each of the ExcelJS functions available.
This is still a work in progress, so expect this list to be udpated.

See [documentation/functions.md](documentation/functions.md) for full parameter details, examples, and style preset references.

| Script                                                                                                      | Description                                                                                                        |
|-------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------|
| [ExcelJS - Init Workbook](documentation/functions.md#exceljs---init-workbook)                               | Resets the in-memory workbook. Always call this first.                                                             |
| [ExcelJS - Add Sheet](documentation/functions.md#exceljs---add-sheet)                                       | Adds a worksheet to the workbook.                                                                                  |
| [ExcelJS - Set Sheet Visibility](documentation/functions.md#exceljs---set-sheet-visibility)                 | Show or hide an existing sheet.                                                                                    |
| [ExcelJS - Set Columns](documentation/functions.md#exceljs---set-columns)                                   | Defines columns on a sheet (keys, headers, widths, hidden flag, default style). Must be called before adding rows. |
| [ExcelJS - Add Row](documentation/functions.md#exceljs---add-row)                                           | Adds a single row to a sheet.                                                                                      |
| [ExcelJS - Add Rows](documentation/functions.md#exceljs---add-rows)                                         | Adds multiple rows in one JS call. Preferred over N calls to Add Row for large datasets.                           |
| [ExcelJS - Set Cell Value](documentation/functions.md#exceljs---set-cell-value)                             | Sets a single cell value by address.                                                                               |
| [ExcelJS - Set Cell Formula](documentation/functions.md#exceljs---set-cell-formula)                         | Sets a formula in a single cell.                                                                                   |
| [ExcelJS - Set Formula For Range](documentation/functions.md#exceljs---set-formula-for-range)               | Applies a per-row formula across a column range using a `{ROW}` placeholder.                                       |
| [ExcelJS - Style Row](documentation/functions.md#exceljs---style-row)                                       | Applies a style preset or raw JSON style to every cell in a row.                                                   |
| [ExcelJS - Style Cell](documentation/functions.md#exceljs---style-cell)                                     | Applies a style preset or raw JSON style to a single cell.                                                         |
| [ExcelJS - Freeze Panes](documentation/functions.md#exceljs---freeze-panes)                                 | Freezes rows and/or columns on a sheet.                                                                            |
| [ExcelJS - Add Defined Name](documentation/functions.md#exceljs---add-defined-name)                         | Registers a workbook-level named range.                                                                            |
| [ExcelJS - Add Data Validation Dropdown](documentation/functions.md#exceljs---add-data-validation-dropdown) | Adds a dropdown list validation to a cell range.                                                                   |
| [ExcelJS - Add Conditional Formatting](documentation/functions.md#exceljs---add-conditional-formatting)     | Applies one or more conditional formatting rules to a cell range.                                                  |
| [ExcelJS - Finalize](documentation/functions.md#exceljs---finalize)                                         | Generates the `.xlsx` buffer and delivers it to FileMaker via callback.                                            |
| [ExcelJS - Receive Workbook](documentation/functions.md#exceljs---receive-workbook)                         | Callback script invoked by JavaScript to deliver the workbook or an error.                                         |
