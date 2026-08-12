# Excel Grid for SSMS 22 — DevExpress Edition

Replaces the visible **Results to Grid** surface with a DevExpress grid in SQL Server Management Studio 22. The native SSMS grid remains untouched underneath and continues to supply query data.

## Features

- Excel-style column filter dropdowns powered by DevExpress.
- Always-visible auto-filter row below the headers.
- Typed sorting for numbers, dates, text, GUIDs, and other common SQL values.
- Multiple filters, incremental search, multi-cell selection, copy, Best Fit, and Clear Filters.
- Automatic light/dark theme support with a high-contrast dark palette.
- Native SSMS right-click menu integration, including commands contributed by SQL Prompt.
- Filtering and sorting are local and never rewrite or rerun the SQL query.
- Works independently for every completed result set.

The replacement snapshots up to 250,000 rows after SSMS finishes retrieving the result set. It is intentionally read-only. Larger result sets show the first 250,000 rows and display that limit in the toolbar.

## Build

Requirements:

- Windows
- SSMS 22 installed (the build locates it with `vswhere`)
- DevExpress WinForms 25.2 installed and licensed
- .NET SDK with the .NET Framework 4.8 reference assemblies

```powershell
.\build.ps1
```

The output is `artifacts\ExcelGrid.Ssms22.vsix`.

## Install

1. Close SSMS.
2. Download the latest VSIX from [GitHub Releases](https://github.com/msse-dev/SSMS_dxGrid/releases/latest).
3. Double-click the downloaded `.vsix` file.
4. Select SQL Server Management Studio 22 in the VSIX Installer.
5. Start SSMS and run a query using Results to Grid.

To uninstall, open **Extensions > Manage Extensions > Installed**, find **Excel Grid for SSMS**, and choose **Uninstall**.

## Compatibility note

SSMS does not publish a supported results-grid extension point. This project reads the completed native result storage but does not replace or mutate that storage. An SSMS update that replaces the native grid may require an adapter update.
