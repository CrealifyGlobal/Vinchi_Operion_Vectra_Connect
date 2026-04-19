# VectraConnect — MS Project Ribbon Add-in

A VSTO add-in that adds a **"Publish Schema"** button to the Microsoft Project ribbon.
One click exports the open project's tasks, resources, and assignments to a styled
**Excel workbook** (+ optional CSV files) in a folder of your choice.

---

## What it exports

| Sheet | Contents |
|-------|----------|
| **Summary** | Project name, dates, KPIs |
| **Tasks** | ID, WBS, name, outline level, summary/milestone/critical flags, duration, start, finish, % complete, costs, slack, constraints, predecessors, notes |
| **Resources** | ID, name, type, email, rates, units, costs |
| **Assignments** | Task ↔ resource links with units, work hours, and costs |

---

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| Microsoft Project 2016 / 2019 / 2021 / 365 (desktop) | |
| Visual Studio 2022 with **"Office/SharePoint development"** workload | Installs VSTO runtime |
| .NET Framework 4.8 | Included with Windows 10/11 |

---

## Build & Install

### 1. Open the solution
```
VectraConnect.sln
```

### 2. Restore NuGet packages
Visual Studio does this automatically on build. The two packages are:
- `Microsoft.Office.Interop.MSProject` — COM interop assembly
- `ClosedXML` — Excel generation (no Office required for writing files)

### 3. Add VSTO references (one-time)

In Visual Studio, right-click **References → Add Reference → Assemblies** and add:

- `Microsoft.Office.Tools.MSProject`
- `Microsoft.Office.Tools.Common`
- `Microsoft.VisualStudio.Tools.Applications.Runtime`

These ship with the Office/SharePoint workload and are in:
```
C:\Program Files\Microsoft Visual Studio\2022\...\VSTO\
```

### 4. Build
```
Build → Build Solution  (Ctrl+Shift+B)
```

### 5. Install / debug
- Press **F5** to launch MS Project with the add-in loaded (debug mode).
- For a permanent install, use **Build → Publish** to create a ClickOnce installer,
  or manually copy the DLL and register it via the `AddInRegister` tool.

---

## Usage

1. Open any `.mpp` file in MS Project.
2. Go to the **Task** tab in the ribbon.
3. Find the **"Vectra Connect"** group on the right.
4. Click **Publish Schema**.
5. Choose an output folder the first time (saved for future sessions).
6. Files appear instantly — click **Yes** to open the folder.

### Settings
Click **Settings** in the ribbon group to:
- Change the default output folder
- Toggle CSV export on/off

Settings are saved to `HKCU\Software\VectraConnect` in the registry.

---

## Output file naming

```
{ProjectName}_{yyyyMMdd_HHmmss}.xlsx
{ProjectName}_Tasks_{yyyyMMdd_HHmmss}.csv
{ProjectName}_Resources_{yyyyMMdd_HHmmss}.csv
{ProjectName}_Assignments_{yyyyMMdd_HHmmss}.csv
```

---

## Project structure

```
VectraConnect/
├── VectraConnect.sln
└── VectraConnect/
    ├── ThisAddIn.cs            ← VSTO entry point
    ├── ProjectParser.cs        ← Reads from MSP COM object model
    ├── SchemaExporter.cs       ← Writes Excel + CSV via ClosedXML
    ├── SettingsManager.cs      ← Registry persistence
    ├── VectraConnect.csproj
    ├── Models/
    │   └── ProjectSchema.cs    ← Data model (tasks, resources, assignments)
    ├── Ribbon/
    │   ├── PublisherRibbon.xml ← Custom ribbon XML (embedded resource)
    │   └── PublisherRibbon.cs  ← IRibbonExtensibility + button handlers
    └── UI/
        └── SettingsDialog.cs   ← WinForms settings window
```

---

## Extending the schema

To add more fields, edit `Models/ProjectSchema.cs` to add properties, then:
- In `ProjectParser.cs` — read the new field from the MSP task/resource COM object
- In `SchemaExporter.cs` — add the column to the relevant `BuildXxxSheet()` method and the CSV header/row

MS Project exposes hundreds of fields on each `Task` object — check the
[MSProject object model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/project) for the full list.

---

## Building the installer (for distribution)

### Prerequisites
- [WiX Toolset v3.11](https://wixtoolset.org/releases/) installed on your build machine
- `vstor_redist.exe` (VSTO runtime) placed at `Installer/Prerequisites/vstor_redist.exe`
  — download from https://aka.ms/vstoruntime

### One-command build
```powershell
.\Build-Installer.ps1
```

This produces `bin\Release\VectraConnect_Setup.exe` — the single file to share with your planners.

### What the installer does
1. Checks for .NET 4.8 — downloads and installs silently if missing
2. Checks for VSTO runtime — installs silently if missing  
3. Runs the MSI wizard (Welcome → Licence → Install → Finish)
4. Installs the add-in to `%LOCALAPPDATA%\VectraConnect\` (no admin rights needed)
5. Registers the add-in in `HKCU` so MS Project picks it up automatically

### Distributing to planners
1. Upload `VectraConnect_Setup.exe` to SharePoint / a shared drive / email
2. Send planners the `Installer/INSTALL_GUIDE.md` instructions
3. They double-click the `.exe` — done in ~2 minutes, no IT involvement needed
