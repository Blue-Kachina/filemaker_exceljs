# FileMaker ExcelJS Wrapper — Developer Reference

## Overview

This is a FileMaker Solution that you can connect to your own -- just keep a copy of this file beside your own, and add it as an External Data Source.

It's a very barebones system comprised of just a single layout.
That layout loads up a FileMaker Webviewer with HTML and JavaScript that instantiates an instance of 
the [ExcelJS](https://github.com/exceljs/exceljs) library.
We've got several FileMaker Scripts, most of them mirroring an ExcelJS function. 
They can be invoked even from the context of your system, and by virtue of doing so, you custom-build a rich Excel (XLSX) spreadsheet.  
This can be done incrementally too, because our webviewer holds a live `ExcelJS.Workbook` in memory. 

## General Workflow
1) Run the `Ensure Window` Script to ensure a window appears with my webviewer on it.  I often pause for a second here too... just to give it time to load
2) Run the `Init Workbook` Script to create a workbook that you can start incrementally working on.
3) Build your custom spreadsheet by calling other FileMaker Scripts I've provided (Each of the scripts mirrors one of the functions from ExcelJS -- and you can find details below)
4) Run the `Finalize` Script to finish up with the Excel Spreadsheet and prompt the user to save it

There's a `Basic Example` and an `Advanced Example` Script that you can use to get a good understanding also.
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

### Initialization

| Script                                                                                    | Description                                            |
|-------------------------------------------------------------------------------------------|--------------------------------------------------------|
| [ExcelJS - Init Workbook](documentation/functions.md#exceljs---init-workbook)             | Resets the in-memory workbook. Always call this first. |

### Creation

| Script                                                                          | Description                                                                                                        |
|---------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------|
| [ExcelJS - Add Sheet](documentation/functions.md#exceljs---add-sheet)           | Adds a worksheet to the workbook.                                                                                  |
| [ExcelJS - Set Columns](documentation/functions.md#exceljs---set-columns)       | Defines columns on a sheet (keys, headers, widths, hidden flag, default style). Must be called before adding rows. |
| [ExcelJS - Add Row](documentation/functions.md#exceljs---add-row)               | Adds a single row to a sheet.                                                                                      |
| [ExcelJS - Add Rows](documentation/functions.md#exceljs---add-rows)             | Adds multiple rows in one JS call. Preferred over N calls to Add Row for large datasets.                           |
| [ExcelJS - Add Table](documentation/functions.md#exceljs---add-table)           | Creates an Excel Table with optional theme, banding, and structured references.                                    |

### Modification

| Script                                                                                                      | Description                                                                  |
|-------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------|
| [ExcelJS - Set Sheet Visibility](documentation/functions.md#exceljs---set-sheet-visibility)                 | Show or hide an existing sheet.                                              |
| [ExcelJS - Set Cell Value](documentation/functions.md#exceljs---set-cell-value)                             | Sets a single cell value by address.                                         |
| [ExcelJS - Set Cell Formula](documentation/functions.md#exceljs---set-cell-formula)                         | Sets a formula in a single cell.                                             |
| [ExcelJS - Set Formula For Range](documentation/functions.md#exceljs---set-formula-for-range)               | Applies a per-row formula across a column range using a `{ROW}` placeholder. |
| [ExcelJS - Style Row](documentation/functions.md#exceljs---style-row)                                       | Applies a style preset or raw JSON style to every cell in a row.             |
| [ExcelJS - Style Cell](documentation/functions.md#exceljs---style-cell)                                     | Applies a style preset or raw JSON style to a single cell.                   |
| [ExcelJS - Freeze Panes](documentation/functions.md#exceljs---freeze-panes)                                 | Freezes rows and/or columns on a sheet.                                      |
| [ExcelJS - Add Defined Name](documentation/functions.md#exceljs---add-defined-name)                         | Registers a workbook-level named range.                                      |
| [ExcelJS - Add Data Validation Dropdown](documentation/functions.md#exceljs---add-data-validation-dropdown) | Adds a dropdown list validation to a cell range.                             |
| [ExcelJS - Add Conditional Formatting](documentation/functions.md#exceljs---add-conditional-formatting)     | Applies one or more conditional formatting rules to a cell range.            |

### Finalizing

| Script                                                                                      | Description                                                                |
|---------------------------------------------------------------------------------------------|----------------------------------------------------------------------------|
| [ExcelJS - Finalize](documentation/functions.md#exceljs---finalize)                         | Generates the `.xlsx` buffer and delivers it to FileMaker via callback.    |
| [ExcelJS - Receive Workbook](documentation/functions.md#exceljs---receive-workbook)         | Callback script invoked by JavaScript to deliver the workbook or an error. |
