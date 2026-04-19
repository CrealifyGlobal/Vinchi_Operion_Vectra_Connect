# Vectra Connect — MCP Integration Handoff Document
**Project:** Vinchi Operion Vectra Connect  
**Repo:** https://github.com/CrealifyGlobal/Vinchi_Operion_Vectra_Connect  
**Date:** April 2026  

---

## What We Are Building

An extension to the existing **Vectra Connect VSTO add-in** (MS Project ribbon plugin) that embeds a local MCP (Model Context Protocol) server, allowing Claude Desktop or Cowork to read and annotate any open MS Project schedule — all from a single installable `.exe`. No extra installs, no copy-pasting, no technical knowledge required from end users.

---

## The Product in One Sentence

> Claude lives inside your MS Project schedule. Click a task row, a context panel opens showing Claude's AI recommendations as a checklist. Tick off what you've done, flag what to do later. Everything persists inside the project file.

---

## Architecture

### Overview

```
VectraConnect-Setup.exe
    ├── Installs VSTO add-in into MS Project
    ├── Auto-writes MCP config into Claude Desktop (%APPDATA%\Claude\claude_desktop_config.json)
    ├── Auto-writes MCP config into Cowork (if installed)
    └── Registers in Windows Add/Remove Programs (for clean uninstall)
```

### Runtime Architecture

```
MS Project (open .mpp)
    └── VSTO Add-in (Vectra Connect)
            ├── Ribbon UI (existing — Publish Schema, Settings)
            ├── ProjectParser.cs (existing)
            ├── SchemaExporter.cs (existing)
            ├── SettingsManager.cs (existing)
            │
            └── [NEW] Embedded MCP Server (localhost:5050)
                    ├── McpServer.cs       — HTTP listener
                    ├── McpTools.cs        — Tool implementations
                    └── ClaudeNotesManager.cs — Custom field handler

Claude Desktop / Cowork
    └── Connects to localhost:5050 via MCP config (written by installer)
            └── Uses tools to read/write the open .mpp in real time
```

### MCP Protocol

The embedded server speaks the MCP JSON-RPC protocol over HTTP on `localhost:5050`. Claude Desktop and Cowork both support this natively. No additional configuration needed by the user — the installer handles it.

---

## MCP Tools to Implement

| Tool | Method | Description |
|------|--------|-------------|
| `get_all_tasks` | GET | Returns all tasks in open project (ID, name, start, finish, duration, % complete, critical flag, WBS) |
| `get_task` | POST | Returns single task by ID |
| `get_critical_path` | GET | Returns only critical path tasks |
| `add_claude_note` | POST | Writes a bullet-point recommendation to the Claude Notes field on a specific task |
| `update_note_status` | POST | Marks a note item as done `[x]`, pending `[ ]`, or maybe later `[~]` |
| `clear_claude_notes` | POST | Clears all Claude Notes from a specific task |

---

## UI: Task Context Panel

When the user **clicks any task row** in MS Project, a WinForms panel opens (docked right or bottom) showing:

```
┌─────────────────────────────────────────────────────┐
│  🔵 Task 47 — Install HVAC Ducting                  │
│  Start: 12 May  |  Duration: 8d  |  85% Critical    │
├─────────────────────────────────────────────────────┤
│  Claude Recommendations                             │
│                                                     │
│  ☑  Confirm duct materials are on site              │  ← ticked = done [x]
│  ☐  Verify access scaffolding is still booked       │  ← unticked = pending [ ]
│  ☐  Check interference with Task 51 (Electrical)    │
│  ☐  Float buffer: only 1 day — flag to PM           │
│                                                     │
│  💬 Maybe Later                                     │
│  ⚑  Review subcontractor insurance docs            │  ← flagged = [~]
│  ⚑  Update as-built drawings after completion      │
├─────────────────────────────────────────────────────┤
│  [Ask Claude about this task...]      [Regenerate]  │
└─────────────────────────────────────────────────────┘
```

### Behaviour Rules

- **Tick** = item done → strikes through, persists as `[x]` in Claude Notes field
- **Flag as Maybe Later** = moves to Later section → persists as `[~]` in Claude Notes field
- **Ask Claude** = freetext input, sends to Claude via MCP, response appended as new bullet
- **Regenerate** = re-runs Claude analysis on this task fresh, replaces existing notes
- **State persists** in the `Claude Notes` custom text field in the `.mpp` file — survives close/reopen
- Panel is **on-demand** — only analyses a task when you click it (not pre-loading all tasks on open)

---

## Claude Notes Custom Field

- Field type: **Text field** (e.g. `Text30` — configurable, won't clash with existing fields)
- Format stored: structured markdown bullet list
  ```
  [x] Confirm duct materials are on site
  [ ] Verify access scaffolding is still booked
  [~] Review subcontractor insurance docs
  ```
- The context panel **reads and writes** this field via the MSP COM object model
- The field is **hidden from the main grid** by default (managed only via the context panel)

---

## New Files to Build

| File | Location | Purpose |
|------|----------|---------|
| `McpServer.cs` | `VectraConnect/` | Embedded `HttpListener` on `localhost:5050`. Starts/stops with ribbon button. Speaks MCP JSON-RPC protocol. |
| `McpTools.cs` | `VectraConnect/` | Implements each MCP tool. Calls `ProjectParser` to read tasks. Calls `ClaudeNotesManager` to write. |
| `ClaudeNotesManager.cs` | `VectraConnect/` | Reads/writes the Claude Notes custom text field on MSP Task objects via COM. Parses `[x]`/`[ ]`/`[~]` markdown. |
| `TaskContextPanel.cs` | `VectraConnect/UI/` | WinForms panel. Fires on task row selection event. Renders checklist UI. Handles tick/flag/ask interactions. |
| `ConfigureClaude.cs` | `Installer/` | Writes MCP config entry into Claude Desktop JSON on install. Removes it surgically on uninstall. Also handles Cowork if present. |

---

## Ribbon Changes

Add one new button to the existing `PublisherRibbon.xml` group:

| Button | Action |
|--------|--------|
| **Connect Claude** | Starts/stops the embedded MCP server. Shows green (running) / grey (stopped) indicator. |

Settings dialog (`SettingsDialog.cs`) gets one new section:

```
Claude MCP Settings
  Port: [5050        ]
  Status: ● Running
  [Copy MCP config snippet]   ← fallback for power users
```

---

## Installer Changes (Build-Installer.ps1 / WiX)

### On Install:
1. Install VSTO add-in (existing)
2. Find `%APPDATA%\Claude\claude_desktop_config.json` → inject `vectra-connect` MCP entry
3. Find Cowork config (if present) → inject same entry
4. Register in Windows Add/Remove Programs

### MCP Config Entry Written:
```json
{
  "mcpServers": {
    "vectra-connect": {
      "url": "http://localhost:5050"
    }
  }
}
```

### On Uninstall:
1. Remove VSTO add-in
2. Remove **only** the `vectra-connect` entry from Claude Desktop config (leave all other MCP servers untouched)
3. Remove from Cowork config (if present)
4. Remove from Add/Remove Programs
5. Prompt: **"Keep Claude Notes data in existing .mpp files? Yes / No"**

---

## Existing Codebase (Do Not Break)

```
VectraConnect/
├── ThisAddIn.cs            ← VSTO entry point — add MCP server start/stop here
├── ProjectParser.cs        ← Read-only MSP data — McpTools.cs will call this
├── SchemaExporter.cs       ← Excel/CSV export — leave untouched
├── SettingsManager.cs      ← Registry settings — extend for MCP port setting
├── Models/
│   └── ProjectSchema.cs    ← Data model — extend with ClaudeNotes field
├── Ribbon/
│   ├── PublisherRibbon.xml ← Add "Connect Claude" button here
│   └── PublisherRibbon.cs  ← Add button handler here
└── UI/
    └── SettingsDialog.cs   ← Add Claude MCP settings section here
```

---

## Claude Client Compatibility

| Client | MCP Support | Notes |
|--------|-------------|-------|
| Claude Desktop | ✅ Yes | Standard MCP via config file. Installer writes config automatically. |
| Cowork | ✅ Yes | Same MCP protocol. Installer writes Cowork config if present. Extra features (phone, dispatch) work automatically. |
| Claude.ai (browser) | ❌ No | Cannot connect to localhost MCP servers. Not supported. |

---

## Distribution

- Single file: `VectraConnect_Setup.exe`
- Share via link, email, SharePoint, shared drive
- User double-clicks → installs in ~2 minutes → opens MS Project → done
- No IT department, no JSON, no command line

---

## Key Technical Decisions Made

1. **Embedded MCP server** — C# `HttpListener` inside the VSTO process. No separate Node/Python process. No extra installs.
2. **Claude Notes field** — dedicated custom text field (`Text30`), not the native MSP Notes field. Keeps Claude's output separate from human notes.
3. **On-demand analysis** — context panel fires when user clicks a row. No pre-loading on project open.
4. **Structured markdown persistence** — `[x]`/`[ ]`/`[~]` format stored in the Notes field so state survives close/reopen.
5. **Surgical installer** — writes and removes only the `vectra-connect` MCP entry. Never touches other user configs.
6. **Port 5050** — configurable in settings, default 5050.

---

## Build Order (Recommended)

1. `ClaudeNotesManager.cs` — read/write Claude Notes field (foundation)
2. `McpServer.cs` — HTTP listener skeleton
3. `McpTools.cs` — tool implementations calling ProjectParser + ClaudeNotesManager
4. `TaskContextPanel.cs` — WinForms UI wired to McpTools
5. Ribbon button + ThisAddIn.cs wiring
6. `ConfigureClaude.cs` — installer config writer/remover
7. Extend `Build-Installer.ps1` to call ConfigureClaude on install/uninstall
8. End-to-end test: Install → Open MSP → Connect Claude Desktop → Click task → Panel appears → Tick item → Reopen → State persists

---

*Handoff document generated from design session — April 2026*
