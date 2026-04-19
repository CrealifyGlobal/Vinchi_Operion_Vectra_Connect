# Vectra Connect — Version 2 Pull Architecture
## Planning Document | Not yet built

---

## Overview

Version 1 is push-only: MS Project → Excel/CSV.

Version 2 adds a pull direction, completing the round-trip. The design principle
is that **MS Project remains the single source of truth at all times**. Excel is
a review and annotation layer only — it never overwrites MS Project data directly.

```
VERSION 1 (built)
─────────────────
MS Project (.mpp)
       │  [Publish Schema button]
       ▼
Excel / CSV output  ←── reviewers work here

VERSION 2 (planned)
───────────────────
MS Project (.mpp)  ◄─── planner manually applies changes after review
       │                         ▲
       │  [Publish Schema]       │  [Pull Changes button]
       ▼                         │
Excel / CSV  ────────────────────┘
  (reviewers add "Proposed Changes" column)
```

---

## What Reviewers Do in Excel

Reviewers work in the exported Excel file. They **do not modify** any existing
data columns. Instead, they:

### 1. Propose changes via the "Proposed Changes" column
A dedicated column (added by the reviewer, or pre-added as an empty column by
the exporter in a future minor v1 update) named exactly:

```
Proposed Changes
```

Each cell contains a pipe-separated list of field → value pairs:

```
Duration: 10d | Start: 2024-05-01 | % Complete: 75
```

Rules:
- Field names must match the column headers in the exported sheet (case-insensitive)
- Values are plain text — no formatting required
- Multiple changes for the same task go in the same cell, separated by ` | `
- Rows with no proposed changes leave the cell blank

### 2. Add new columns
Reviewers may add columns to the right of the existing data (e.g. "QA Status",
"Reviewer Notes", "Risk Flag"). These are treated as **reviewer annotations**.

### 3. Add new task rows
Reviewers may add new rows at the bottom of the Tasks sheet. New rows have no
VectraKey — this is the signal that they are reviewer-added.

### 4. Mark tasks for deletion
Reviewers add the value `DELETE` in the Proposed Changes cell for any task they
propose removing.

### 5. Partial returns
Reviewers may return only the Tasks tab, or only Resources, or any subset. Vectra
Connect handles partial files gracefully — only the tabs present are processed.

---

## What Happens on Pull

The planner clicks **Pull Changes** in the Vectra Connect ribbon group inside
MS Project. The button is only enabled when a project file is open.

### Step 1 — File selection
A file picker opens. The planner selects the returned Excel file. Vectra Connect
automatically looks for the matching `.manifest` file in the same folder (same
filename stem, `.manifest` extension). If found, it is used for VectraKey
matching. If not found, matching falls back to task name.

### Step 2 — Parse and diff
Vectra Connect reads every sheet present in the returned Excel and produces a
structured `ChangeSet`:

| Change type | Detection | Action |
|---|---|---|
| Proposed field change | "Proposed Changes" cell is non-empty | Write to custom field in MSP |
| Reviewer annotation column | New column not in original schema | Write to separate custom field in MSP (if available) |
| New task row | No VectraKey in column A | Write as proposed new task to custom field on a summary row or project-level note |
| Proposed deletion | "Proposed Changes" cell contains `DELETE` | Write `DELETE PROPOSED` to custom field |

### Step 3 — Summary screen
Before writing anything, Vectra Connect shows a read-only summary panel:

```
┌─ Vectra Connect — Pull Summary ──────────────────────────────────┐
│                                                                    │
│  File:    ConstructionPhase2_20240418_143022.xlsx                 │
│  Matched: 142 tasks via VectraKey  |  3 unmatched (name fallback)│
│                                                                    │
│  📝 Proposed changes found on:   38 tasks                        │
│  ➕ New tasks proposed:           2 rows                          │
│  🗑 Deletions proposed:           1 task                          │
│  📎 New reviewer columns found:   2  ("QA Status", "Risk Flag")  │
│                                                                    │
│  Target custom field:  Text3 — "Vectra Proposed Changes"         │
│  [ Change field... ]                                              │
│                                                                    │
│  ⚠️  Custom field limit: 2 of 30 Text fields already in use.     │
│      Reviewer columns will use Text4 and Text5.                   │
│                                                                    │
│  [ Import ]                              [ Cancel ]               │
└────────────────────────────────────────────────────────────────────┘
```

### Step 4 — Write to MS Project custom fields
On confirmation, Vectra Connect writes to the open `.mpp` via the COM object
model:

**For proposed task changes:**
Each matched task gets its configured custom field populated with the raw
proposed changes string from Excel, prefixed with the import timestamp:

```
[2024-05-03 09:14] Duration: 10d | Start: 2024-05-01 | % Complete: 75
```

If the task already has a value in that custom field (from a previous pull),
the new value is appended with a newline separator — creating an audit trail.

**For reviewer annotation columns:**
Each new column maps to the next available custom Text field. The custom field
is renamed in MS Project to match the column header (e.g. `Text4` → "QA Status").
A warning is shown if no free Text fields are available.

**For proposed new tasks:**
A structured note is appended to the project-level Notes field:

```
[VECTRA PROPOSED NEW TASK — 2024-05-03]
Name: Drainage Survey
Suggested by: (reviewer name if present in file metadata)
From row: 147
```

The planner then manually creates the task if approved.

**For proposed deletions:**
The task's custom field is set to:
```
[2024-05-03 09:14] DELETE PROPOSED
```

---

## Planner Workflow After Pull

1. Open the column containing proposed changes in MS Project (e.g. Text3)
2. Review each task's proposed changes
3. Apply agreed changes manually to the native MSP fields
4. Clear the custom field cell once actioned (or leave for audit trail)
5. Re-export with **Publish Schema** when ready for next review cycle

This keeps the planner in full control. Nothing is auto-applied to schedule data.

---

## Settings Changes for V2

The Settings dialog gains a new section: **Pull Settings**

| Setting | Type | Default |
|---|---|---|
| Proposed Changes custom field | Dropdown (Text1–Text30) | (planner picks on first pull) |
| Prepend timestamp to proposed changes | Toggle | On |
| Append to existing field value (audit trail) | Toggle | On |
| Auto-map reviewer columns to next free Text field | Toggle | On |
| Warn before using more than N custom fields | Number | 5 |

---

## New Files Required for V2

| File | Purpose |
|---|---|
| `SessionManifest.cs` | Model: records VectraKeys + original values at export time |
| `ManifestWriter.cs` | Writes `.manifest` JSON alongside the Excel on Publish |
| `ExcelPullReader.cs` | Reads returned Excel, extracts ChangeSet |
| `ChangeSet.cs` | Model: structured list of all detected changes |
| `ProjectWriter.cs` | Writes ChangeSet to MSP custom fields via COM |
| `PullSummaryForm.cs` | WinForms panel: shows summary, lets planner confirm |
| Updates to `PublisherRibbon.xml` | Add "Pull Changes" button to ribbon group |
| Updates to `PublisherRibbon.cs` | Wire up OnPullClick handler |
| Updates to `SettingsDialog.cs` | Add Pull Settings section |
| Updates to `SettingsManager.cs` | Persist new pull settings to registry |

The existing v1 files (`ProjectParser.cs`, `SchemaExporter.cs`, `ProjectSchema.cs`)
are unchanged — v2 is entirely additive.

---

## Custom Field Limit Awareness

MS Project allows up to 30 custom Text fields (Text1–Text30), 30 Number fields,
30 Cost fields, etc. Vectra Connect must:

1. At pull time, enumerate which custom fields are already in use in the open project
2. Offer only free fields as options for the Proposed Changes field
3. Warn clearly if fewer than 3 free Text fields remain (needed for annotations)
4. Never silently overwrite an existing custom field that already has data

---

## Key Design Principles (Do Not Violate in Build)

1. **MS Project is always the master.** Vectra Connect never auto-modifies
   schedule fields (dates, durations, costs, assignments). Only custom Text
   fields are written on pull.

2. **VectraKey is the join key.** Every pull operation matches rows by VectraKey
   first, task name second. Never by row number or position.

3. **Additive only.** Pull appends to custom fields; it never clears or
   overwrites existing custom field data unless the planner explicitly chooses to.

4. **Planner-only trigger.** The Pull button is only accessible from within
   MS Project by whoever has the `.mpp` file open. There is no standalone pull tool.

5. **Partial files are valid.** A returned Excel with only the Tasks tab is
   processed correctly. Missing tabs are silently skipped, not treated as errors.

6. **Audit trail by default.** Every write to a custom field is timestamped.
   The history of proposed changes accumulates in the field until the planner
   clears it.

---

## Resolved Design Decisions

All four open questions have been answered. These are now binding design rules
for the v2 build.

---

### 1. Reviewer Identity
**Decision: Timestamp only — no identity capture.**

Pull entries are stamped with date and time only. The Excel file's "Last Modified
By" metadata is ignored. Format:

```
[2024-05-03 09:14] Duration: 10d | Start: 2024-05-01
```

No reviewer name is read, stored, or displayed anywhere in the pull flow.

---

### 2. Proposed Changes Column — Pre-added on Every Export
**Decision: Yes — Vectra Connect pre-adds the column on every export.**

The `SchemaExporter` will add an empty **"Proposed Changes"** column as the
last column on the Tasks, Resources, and Assignments sheets on every export.

The column header is styled distinctly (e.g. amber background) so reviewers
immediately know where to write. The column is empty on export — reviewers
fill it in.

Impact on v1 code: a small update to `SchemaExporter.cs` — add the column
header and style it. No other changes.

---

### 3. Multi-Round Pulls
**Decision: Last pull wins — no duplicate warning.**

If the same exported Excel file is pulled a second time, Vectra Connect simply
overwrites the previous value in the custom field with the new one (still
timestamped). No warning is shown. The planner is responsible for managing
which version of the returned file they pull.

This keeps the pull flow fast and simple. The timestamp on the field value
gives the planner enough context to know when the last pull occurred.

---

### 4. Annotation Column Limit Behaviour
**Decision: Block the pull entirely if no free Text fields are available.**

If a reviewer has added annotation columns AND all 30 MS Project Text fields
are already in use, Vectra Connect will **not proceed with the pull**. Instead
it shows a clear error:

```
┌─ Vectra Connect — Cannot Import ──────────────────────────────────┐
│                                                                     │
│  ❌ No free custom Text fields available in this project.          │
│                                                                     │
│  The returned Excel contains 2 new reviewer columns               │
│  ("QA Status", "Risk Flag") that need to be imported,             │
│  but all 30 Text fields (Text1–Text30) are already in use.        │
│                                                                     │
│  Please free up at least 2 Text fields in MS Project and          │
│  try again. Go to Project → Custom Fields to manage them.         │
│                                                                     │
│  Note: If you remove the reviewer columns from the Excel file,    │
│  the pull can proceed without needing free Text fields.           │
│                                                                     │
│                                          [ OK ]                   │
└─────────────────────────────────────────────────────────────────────┘
```

The block applies only when annotation columns are present. If the returned
Excel has no new columns — only a "Proposed Changes" column — the pull
proceeds regardless of Text field availability (since Proposed Changes always
writes to the single planner-configured field, which is already reserved).

---

## Final V2 Build Scope

With all decisions resolved, the complete v2 scope is:

### New files (8)
| File | Purpose |
|---|---|
| `SessionManifest.cs` | Model: VectraKeys + field snapshot at export time |
| `ManifestWriter.cs` | Writes `.manifest` JSON alongside Excel on Publish |
| `ExcelPullReader.cs` | Reads returned Excel, extracts ChangeSet |
| `ChangeSet.cs` | Model: all detected changes, typed and structured |
| `ProjectWriter.cs` | Writes ChangeSet to MSP custom fields via COM |
| `PullSummaryForm.cs` | WinForms: pre-pull summary + confirm/cancel |
| `CustomFieldScanner.cs` | Enumerates used/free Text fields in open project |
| `AnnotationFieldMapper.cs` | Maps reviewer columns to available Text fields |

### Modified files (5)
| File | Change |
|---|---|
| `SchemaExporter.cs` | Pre-add styled empty "Proposed Changes" column on all sheets |
| `PublisherRibbon.xml` | Add "Pull Changes" button to ribbon group |
| `PublisherRibbon.cs` | Wire up `OnPullClick` handler |
| `SettingsDialog.cs` | Add Pull Settings section (custom field picker, timestamp toggle) |
| `SettingsManager.cs` | Persist pull settings to registry |

### Unchanged (all v1 files not listed above)
`ProjectParser.cs`, `ProjectSchema.cs`, `ThisAddIn.cs` — zero modifications.

---

*Document status: All decisions resolved. Ready for v2 development sprint.*
*V1 codebase: VectraConnect.sln — v2 is fully additive, v1 behaviour unchanged.*
