# CNC-DataBase

A Google Apps Script project that integrates with Google Sheets to manage a CNC file database. It provides tools for downloading DXF files, managing links to CNC element files, and organizing sets and modules of CNC parts.

## Features

- **Download DXF files by set** – retrieve all DXF files belonging to a given set (e.g. `P1608`), with optional color-based folder organization.
- **Download DXF files by module** – retrieve all DXF files belonging to a given module (e.g. `M1594`).
- **Element list generation** – generate a list of elements for a set or module directly from the spreadsheet data.
- **Link management** – set, edit, and compare links assigned to individual CNC elements.
- **Mass link check** – bulk validation and auto-fix of broken links across the entire database.
- **UCANCAM file check** – dedicated validation for UCANCAM-related file links.

## Project Structure

| File | Description |
|---|---|
| `Code.js` | Entry point with top-level prompt functions wired to the menu |
| `ui.js` | `onOpen` handler that builds the custom Sheets menu; prompt wrappers |
| `helpers.js` | Shared constants (`SHEET_ZESTAWY`, `SHEET_MODULE`) and utility functions |
| `downloads.js` | DXF download logic, recursive set/module expansion, color handling |
| `links.js` | Link sync, comparison, and mass-fix logic |
| `appsscript.json` | Google Apps Script manifest (timezone, runtime version) |
| `*.html` | HTML dialogs used by the sidebar/modal UI (color selector, conflict resolvers, download complete screen) |

## Spreadsheet Structure

The script expects a Google Spreadsheet with two sheets:

- **Zestawy CNC** – contains set definitions (set ID in column A, element/module reference with hyperlink in column B, count in column C, surface in column E, name in column H).
- **Moduły CNC** – contains module definitions with the same column layout as Zestawy CNC.

## Setup

1. Open your Google Spreadsheet.
2. Go to **Extensions → Apps Script**.
3. Copy all `.js` files into the Apps Script project (one script file per `.js` file) and upload the `.html` files.
4. Save and reload the spreadsheet — the **CNC**, **Sync**, and **Sync CNC** menus will appear.

> Alternatively, use [clasp](https://github.com/google/clasp) to push the project directly from the command line using the included `.clasp.json` configuration.

## Usage

After setup, use the custom menus in Google Sheets:

**CNC menu**
- *Pobierz pliki dla zestawu (z kolorami)…* – download DXF files for a set, sorted into color-named folders.
- *Pobierz pliki dla modułu…* – download DXF files for a module.
- *Pobierz listę elementów dla zestawu…* – generate an element list for a set.
- *Pobierz listę elementów dla modułu…* – generate an element list for a module.

**Sync menu**
- *Ustaw / edytuj link dla elementu…* – assign or update a link for a specific element.
- *Porównaj linki (SyncLinks)* – compare links between entries.
- *Masowe sprawdzenie linków* – run a bulk check and auto-fix of all links.

**Sync CNC menu**
- *Sprawdzenie plików UCANCAM* – validate UCANCAM file links.

## License

Copyright © Jakub Balcerzak. All rights reserved.
