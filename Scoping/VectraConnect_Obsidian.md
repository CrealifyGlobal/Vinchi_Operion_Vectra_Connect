# Vectra Connect — Export & Schema Publisher Handoff Document
**Project:** Vectra — Part of the Vectra AI Production Chain Operating System
**Module:** V1 Push Export (MS Project → Excel/CSV Schema)
**Repo:** VectraConnect.sln
**Date:** April 2026

---

## What We Are Building

A VSTO add-in for Microsoft Project that adds a **"Publish Schema"** ribbon button
to the MS Project Task tab. One click exports the open `.mpp` file's tasks,
resources, and assignments to a structured, styled Excel workbook and optional
CSV files — ready for external review, AI analysis, or import into an Obsidian vault.

This is **Module 1 of 3** in the Vectra Connect system:

| Module | What It Does | Where Claude Lives |
|---|---|---|
| **V1 Push (this doc)** | Exports MSP schema → Excel/CSV for review | External (Obsidian / analyst) |
| V2 MCP | Embeds Claude inside MS Project via local MCP server | Claude Desktop / Cowork |
| V3 Cowork | Brings project schema into Claude via ngrok tunnel + plugin | Claude.ai / Cowork |

---

## The Product in One Sentence

> Click Publish Schema in MS Project — get a structured, VectraKey-stamped Excel
> workbook (plus CSVs) that reviewers can annotate and return, and that any AI
> tool can reason over without ever touching the .mpp file.

---

## Architecture

```
MS Project (.mpp)
    └── VSTO Add-in (Vectra Connect)
            ├── Ribbon button: "Publish Schema"
            ├── Ribbon button: "Settings"
            ├── ProjectParser.cs       — reads open .mpp via COM object model
            ├── SchemaExporter.cs      — writes styled Excel + CSVs via ClosedXML
            ├── SettingsManager.cs     — persists settings to HKCU registry
            └── Models/
                    └── ProjectSchema.cs — Tasks, Resources, Assignments data model

Output (local filesystem folder, planner-chosen):
    ConstructionPhase2_20240418_143022.xlsx
    ConstructionPhase2_Tasks_20240418_143022.csv
    ConstructionPhase2_Resources_20240418_143022.csv
    ConstructionPhase2_Assignments_20240418_143022.csv
```

---

## VectraKey — Primary Key Design

Every row in every sheet and CSV carries a **VectraKey** in Column A.

**Format:**
```
{MppFileName}_{yyyyMMdd}_{HHmmss}_{SheetCode}_{RowUID:D5}
```

**Example:**
```
ConstructionPhase2_20240418_143022_TSK_00042
ConstructionPhase2_20240418_143022_RES_00005
ConstructionPhase2_20240418_143022_ASN_00089
```

**Sheet codes:** `TSK` = Tasks | `RES` = Resources | `ASN` = Assignments

**Why this works:**
- Cross-session unique — same project exported twice gets different timestamps
- Cross-tab unique — TSK_00042 and ASN_00042 are different rows in different tables
- Human readable — project, date, time, tab, row all visible at a glance
- Joinable — every tab in one export shares the same prefix, so you can join across sheets

**Key decisions made:**
- VectraKey is **Column A** (first column) on every sheet — frozen, always visible
- MS Project native **ID and UniqueID are kept** as columns B and C alongside it
- VectraKey is styled in **dark navy / cyan Consolas** text — visually unmistakable
- Column A is **frozen** so it stays pinned while scrolling right
- For Assignments (no native UID), a **Cantor pairing** of TaskUID + ResourceUID produces a stable surrogate integer

---

## Excel Output Structure

### Sheet: Summary
- Project name, export timestamp, source file path
- Session key prefix (so you know what prefix all VectraKeys in this export share)
- KPI table: Start, Finish, % Complete, Total Tasks, Total Resources, Currency

### Sheet: Tasks
Column A: VectraKey | B: ID | C: UniqueID | D: WBS | E: Outline Level | F: Name |
G: Summary? | H: Milestone? | I: Critical? | J: Duration | K: Start | L: Finish |
M: % Complete | N: Baseline Cost | O: Actual Cost | P: Remaining Cost |
Q: Total Slack (d) | R: Constraint Type | S: Constraint Date |
T: Predecessors | U: Notes | **V: Proposed Changes** ← amber column, reviewers write here

**Row styling:**
- Summary rows: blue background, bold
- Milestone rows: yellow background
- Alternating rows: light grey / white
- VectraKey column: dark navy background, cyan text, Consolas font

### Sheet: Resources
Column A: VectraKey | B: ID | C: UniqueID | D: Name | E: Type | F: Email |
G: Max Units (%) | H: Standard Rate | I: Overtime Rate | J: Baseline Cost | K: Actual Cost |
**L: Proposed Changes** ← amber column

### Sheet: Assignments
Column A: VectraKey | B: Task ID | C: Task Name | D: Task VectraKey (FK) |
E: Resource ID | F: Resource Name | G: Resource VectraKey (FK) |
H: Units (%) | I: Work (hrs) | J: Actual Work (hrs) | K: Remaining Work (hrs) |
L: Cost | M: Actual Cost | **N: Proposed Changes** ← amber column

**Assignments carry FK VectraKeys** (cols D and G) linking back to Tasks and Resources
sheets — the data is fully relational and joinable by VectraKey alone.

---

## Proposed Changes Column

Pre-added as the **last column on every sheet** on every export. Styled in **amber**
(dark gold header, light yellow cells) so reviewers immediately see where to write.

**Format reviewers use:**
```
Duration: 10d | Start: 2024-05-01 | % Complete: 75
```

Pipe-separated. Field names match column headers (case-insensitive). One cell per row.
To propose deletion: write `DELETE` in the cell.

This column is the bridge to **V2 Pull** (see below).

---

## V2 Pull — Planned, Not Yet Built

The round-trip companion to V1 Push. Decisions are fully resolved and documented
in `V2_PULL_ARCHITECTURE.md`.

**The model:**
- Reviewers annotate the exported Excel using the Proposed Changes column
- Planner clicks **Pull Changes** ribbon button in MS Project
- Vectra Connect reads the returned Excel and writes proposed changes into a
  planner-configured MS Project custom Text field (Text1–Text30)
- MS Project is **never auto-written** — planner reviews the custom field and
  applies changes manually, then re-exports
- Single source of truth stays inside MS Project at all times

**Key V2 decisions locked:**
1. Timestamp only — no reviewer identity captured
2. Proposed Changes column pre-added on every export (amber styling)
3. Multi-round pulls: last pull wins, no duplicate warning
4. If all 30 Text fields are taken: block the pull, tell planner to free a field first
5. Partial returns (only Tasks tab returned, etc.) are handled gracefully
6. Every write is timestamped and appended — building an audit trail in the custom field

**V2 new files (8):** `SessionManifest.cs`, `ManifestWriter.cs`, `ExcelPullReader.cs`,
`ChangeSet.cs`, `ProjectWriter.cs`, `PullSummaryForm.cs`, `CustomFieldScanner.cs`,
`AnnotationFieldMapper.cs`

**V2 modified files (5):** `SchemaExporter.cs`, `PublisherRibbon.xml`,
`PublisherRibbon.cs`, `SettingsDialog.cs`, `SettingsManager.cs`

---

## Installer

Single `.exe` bootstrapper built with **WiX Toolset v3.14** + Burn bootstrapper.
Produced by GitHub Actions CI — no local build environment needed.

**What the installer does:**
1. Silently installs .NET 4.8 if missing (downloads from Microsoft CDN)
2. Silently installs VSTO runtime if missing (`vstor_redist.exe` bundled)
3. Runs MSI wizard: Welcome → Licence → Install → Finish
4. Installs add-in to `%LOCALAPPDATA%\VectraConnect\` — **no admin rights needed**
5. Registers add-in in `HKCU\Software\Microsoft\Office\MS Project\AddIns\VectraConnect`
6. Persists settings to `HKCU\Software\VectraConnect`

**Distribution:** Upload `VectraConnect_Setup.exe` to SharePoint / email link.
Planners double-click — done in ~2 minutes. No IT involvement. No admin password.

**Uninstall:** Windows Settings → Apps → Vectra Connect → Uninstall. Clean removal.

---

## Settings

Accessible via the **Settings** ribbon button. Persisted to Windows registry.

| Setting | Purpose |
|---|---|
| Output folder | Default folder for exported files |
| Include CSV | Toggle CSV export on/off alongside Excel |
| *(V2)* Proposed Changes custom field | Which Text field (Text1–Text30) to write pull results into |
| *(V2)* Prepend timestamp | Toggle timestamp prefix on pull entries |
| *(V2)* Append to existing field | Audit trail vs overwrite behaviour on pull |

---

## Existing Codebase

```
VectraConnect/
├── ThisAddIn.cs                   ← VSTO entry point
├── ProjectParser.cs               ← Reads open .mpp via MSProject COM object model
├── SchemaExporter.cs              ← Writes Excel + CSV via ClosedXML
├── SettingsManager.cs             ← Registry persistence
├── Models/
│   └── ProjectSchema.cs           ← Data model (MppFileName, Tasks, Resources, Assignments)
├── Ribbon/
│   ├── PublisherRibbon.xml        ← Custom ribbon XML (embedded resource)
│   └── PublisherRibbon.cs        ← IRibbonExtensibility + button handlers
└── UI/
    └── SettingsDialog.cs          ← WinForms settings window
```

**Key dependencies:**
- `Microsoft.Office.Interop.MSProject` — COM interop (NuGet)
- `ClosedXML` — Excel generation (NuGet, no Office required for writing)
- `.NET Framework 4.8` — required, installed silently by bootstrapper if missing
- `VSTO Runtime` — required, installed silently by bootstrapper if missing

---

## GitHub Actions CI

Workflow at `.github/workflows/build.yml`. Runs on every push to `main` and on
version tags (`v*.*.*`).

**What it does automatically:**
1. Builds VSTO add-in (MSBuild, Release)
2. Installs WiX 3.14
3. Downloads VSTO runtime redistributable
4. Generates placeholder installer assets (banner, dialog, icon)
5. Compiles MSI (candle + light)
6. Compiles Setup.exe bootstrapper (candle + light)
7. Uploads `VectraConnect_Setup.exe` as downloadable artifact (30-day retention)
8. On version tag: creates a GitHub Release with the Setup.exe attached

**To ship a release:**
```bash
git tag v1.0.0
git push origin v1.0.0
```
GitHub builds and publishes the release page automatically.

---

## Key Technical Decisions Made

1. **VSTO + C# / .NET 4.8** — native MS Project integration via COM object model
2. **ClosedXML** — Excel generation without requiring Office on the build machine
3. **VectraKey as primary key** — filename + timestamp + sheet code + row UID, Column A, frozen
4. **ID and UniqueID kept** — native MSP identifiers retained alongside VectraKey
5. **Proposed Changes column pre-added** — amber styling, last column on every sheet
6. **Per-user install** — `%LOCALAPPDATA%`, no admin rights, HKCU registry
7. **WiX Burn bootstrapper** — single `.exe` handles all prerequisites silently
8. **GitHub Actions CI** — Windows runner builds and packages automatically
9. **Name: Vectra Connect** — part of the Vectra AI Production Chain OS
10. **No auto-write on pull** — MS Project is never modified without planner approval

---

## Naming Conventions

| Old name | Current name |
|---|---|
| MppPublisher | VectraConnect |
| MS Project Schema Publisher | Vectra Connect |
| Schema Publisher (ribbon) | Vectra Connect (ribbon group) |
| MppPublisher_Setup.exe | VectraConnect_Setup.exe |
| `%LOCALAPPDATA%\MppPublisher\` | `%LOCALAPPDATA%\VectraConnect\` |
| `HKCU\Software\MppPublisher` | `HKCU\Software\VectraConnect` |

---

*Handoff document generated from design session — April 2026*
*Sister documents: VectraConnect_MSP.md (MCP module) | VectraConnect_Cowork.md (Cowork module)*
