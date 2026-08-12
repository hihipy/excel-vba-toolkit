# Excel VBA Toolkit

[![Link Check](https://github.com/hihipy/excel-vba-toolkit/actions/workflows/links.yml/badge.svg)](https://github.com/hihipy/excel-vba-toolkit/actions/workflows/links.yml)
[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)

**Built with**

[![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![VBA](https://img.shields.io/badge/VBA-5E5E5E?style=flat&logoColor=white)](https://learn.microsoft.com/office/vba/api/overview/)

A collection of Excel VBA macros, organized as .bas code modules so you can reuse them and import them into any Excel workbook.

These macros cover:

- Data cleaning and transformation
- Exports (CSV, Markdown) with detailed formatting
- Documentation of PivotTables and Excel Tables
- Utility functions for hyperlink extraction and data manipulation

All files are stored as plain-text .bas modules that import directly into Excel's VBA editor.

---

## What Are .bas Files?

.bas files are VBA code modules exported from Excel. They contain reusable procedures (macros) you can import into other Excel workbooks. Each .bas file in this repo corresponds to a specific macro or functional category (e.g., exports, utilities).

---

## How to Download Macros from GitHub

You can download and use the macros in this repository in two ways:

### Option 1: Download the Entire Toolkit

1. Click the green **Code** button at the top of this page.
2. Select "Download ZIP".
3. Once downloaded, unzip the file on your computer.
4. Inside the unzipped folder, you'll find organized folders such as `data-cleaning`, `exports`, and `utilities`, each containing .bas files.

### Option 2: Download an Individual Macro

1. Click on the .bas file you want.
2. Click the **Raw** button to view the plain text.
3. Right-click the page and select "Save As...".
4. Make sure the file is saved with a .bas extension (e.g., `ExportPivotToMarkdown.bas`).

---

## How to Enable the Developer Tab in Excel

Before you can use or import macros, the Developer tab must be visible in the Excel ribbon.

### Enabling Developer Tab (Windows & Mac)

**On Windows:**

1. Open Excel.
2. Go to **File > Options**.
3. In the left pane, click **Customize Ribbon**.
4. On the right side, check the box labeled **Developer**.
5. Click **OK**.

**On Mac:**

1. Open Excel.
2. Go to **Excel > Preferences**.
3. Select **Ribbon & Toolbar**.
4. In the Main Tabs section, check **Developer**.
5. Click **Save**.

You will now see the Developer tab in the ribbon, which gives access to the VBA editor, macro tools, and import options.

---

## How to Use These in Excel

### Step-by-Step: Importing a .bas File

1. Open your Excel workbook.
2. Click the **Developer** tab, then click **Visual Basic** (or press `Alt + F11`) to open the VBA Editor.
3. In the left pane (Project Explorer), click your workbook name.
4. From the top menu, go to **File > Import File...**.
5. Navigate to the .bas file you downloaded and select it.
6. The module will now appear under "Modules".
7. Press `Alt + Q` to return to Excel.
8. Press `Alt + F8`, select the macro, and click **Run**.

**Tip:** You can import and use multiple macros in the same workbook. For ones you use often, consider adding them to your Personal Macro Workbook (PERSONAL.XLSB) so they're available in all Excel files.

---

## Requirements

- Microsoft Excel (Windows or Mac)
- Macro-enabled workbook format (.xlsm)
- Macros must be enabled in Excel (click "Enable Content" if prompted)
- Excel 2016 or later recommended

---

## Repository Structure

Each folder groups macros by category. Here is the current structure:

```
excel-vba-toolkit/
├── data-cleaning/
│   ├── DeleteHiddenRows.bas
│   ├── FillBlanksDown.bas
│   └── WhitespaceTools.bas
├── exports/
│   ├── DocumentFormulas.bas
│   ├── DocumentTableFormulas.bas
│   ├── ExportPivotToMarkdown.bas
│   ├── ExportRangeToCSV.bas
│   ├── GenerateAdvancedPivotReport.bas
│   └── GenerateTableDoc.bas
├── utilities/
│   └── GetHyperlinkURL.bas
└── README.md
```

---

## Macro Descriptions

What each .bas file in the toolkit does:

### Data Cleaning

#### DeleteHiddenRows.bas

Deletes all hidden rows in the active worksheet using a bottom-up approach. Tracks progress in real time, reports execution time, and is tuned for large datasets (50,000+ rows). Useful for cleaning filtered data before analysis or export.

#### FillBlanksDown.bas

Fills blank cells in a selected range with the value from the cell directly above. Handles merged cells cleanly and reports on the number of cells modified. Useful for cleaning pivot table exports or grouped data where labels are omitted in repeated rows.

#### WhitespaceTools.bas

Detects and fixes whitespace issues across an entire workbook. Uses single-pass processing to find and highlight leading, trailing, and repeated internal spaces in one operation, working through all sheets with bulk operations to stay fast on large files. Includes one-click workbook-wide cleaning with a report, and no selection is required.

### Exports & Documentation

#### DocumentFormulas.bas

Documents all formulas on any Excel worksheet and exports them as a structured JSON file. Works with any sheet layout, no Excel Tables required. Includes formula categorization, business-intent detection, dependency mapping, error analysis, and optimization hints. Designed for feeding into AI tools for formula review, documentation, and troubleshooting.

#### DocumentTableFormulas.bas

Scans all Excel Tables (ListObjects) in a workbook and documents their column formulas in a Markdown file. Records formula text, formula category, and cross-table references for each column. Useful for auditing formula logic or generating technical documentation for structured data models.

#### ExportPivotToMarkdown.bas

Exports the first PivotTable on the active worksheet to GitHub-compatible Markdown. Preserves table structure with proper pipe delimiters and escapes special characters. Useful for documentation, reports, or sharing pivot analysis on markdown-friendly platforms.

#### ExportRangeToCSV.bas

CSV export tool with data type detection, configurable text quoting, and buffered writing for speed. Handles 50,000+ rows while preserving data integrity. Supports custom delimiters and reports detailed export statistics.

#### GenerateAdvancedPivotReport.bas

Documents all PivotTables (both OLAP and regular) in a workbook. Includes field configurations, data sources, OLAP connection details, MDX references, calculated fields, and slicer information. Outputs Markdown reports with full metadata.

#### GenerateTableDoc.bas

Documents all Excel Tables across every worksheet in the workbook. Includes data profiling with type detection, sample values, formula transparency with exact syntax and dependency mapping, data quality flags (CLEAN/WARNING/ERROR), and performance guidance. Generates Markdown output compatible with any AI tool. Works with datasets from 100 to 100,000+ rows.

### Utilities

#### GetHyperlinkURL.bas

Custom Excel function that extracts the actual URL from hyperlinked cells. Use it as `=GetHyperlinkURL(A1)` to pull hyperlink addresses for link inventories, validation, or export lists. Includes error handling for non-hyperlinked cells and multiple cell selections.

---

## License

This project is licensed under [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/).

You are free to:

- Use, share, and adapt this work
- Use it at your job

Under these terms:

- **Attribution:** Credit the original author
- **NonCommercial:** No selling or commercial products
- **ShareAlike:** Derivatives must use the same license
