# Vectra Connect V3 — Cowork Plugin Handoff

**Date:** April 19, 2026
**Status:** Architecture & Design Complete — Ready to Build
**Scope:** Cowork plugin connecting MS Project to Claude via ngrok tunnel

---

## 1. Product Overview

Vectra Connect is a Microsoft Project add-in that brings Claude AI into project planning workflows. It is part of the **Vectra AI Production Chain OS** — a suite of tools built to connect MS Project data to AI-powered analysis.

### 1.1 The Three Versions

| Version | Name | Mechanism | Direction | Where Claude Lives |
|---------|------|-----------|-----------|-------------------|
| V1 | Push / Export | VSTO → Obsidian vault | Out only | Obsidian vault |
| V2 | MCP Connector | VSTO → local MCP server | Read + Write | Claude Desktop / Cowork (local) |
| **V3** | **Cowork Plugin** | **VSTO → ngrok → Cowork** | **Read + Write** | **Cowork (cloud-accessible)** |

### 1.2 Why V3 Exists

V2 (MCP) works when Claude Desktop or Cowork connects to a localhost MCP server on the same machine. V3 solves the cloud-accessibility problem: Cowork's connectors must reach a publicly accessible endpoint. V3 uses an ngrok tunnel to bridge the local VSTO MCP server to a stable public URL that Cowork can reach from Anthropic's cloud.

---

## 2. Architecture

### 2.1 Component Diagram

```
MS Project (.mpp)
      │
VSTO Add-in (C#)          ← Ribbon UI + Settings Panel
      │
Local MCP Server          (HttpListener on localhost:5050)
      │
ngrok Agent               (Windows Service, silent, starts with Windows)
      │
https://[user].ngrok-free.app    ← stable public URL, never changes
      │
Cowork (.mcp.json connector)     ← plugin points here
      │
Claude (Cowork session)          ← planner chats here
```

### 2.2 Key Design Principles

- **One installer** — single `.exe` handles everything
- **Invisible operation** — ngrok runs as a silent Windows service, no console, no tray icon
- **Stable public URL** — free ngrok static domain, automatically assigned, never changes
- **No cloud server required** — tunnel bridges local to public; no hosted infrastructure needed
- **Planner approval required** — Claude never writes to `.mpp` without explicit user confirmation
- **Clean uninstall** — removes VSTO, ngrok service, plugin folder, and Cowork config entry

---

## 3. Installer Design

### 3.1 What the Installer Does

| Step | Action | Notes |
|------|--------|-------|
| 1 | Installs VSTO add-in | Registers with MS Project, per-user, no admin rights required |
| 2 | Installs ngrok agent | Bundled silently, installed as a Windows service |
| 3 | Shows onboarding screen | User enters their ngrok auth token (one-time, ~2 min setup) |
| 4 | Configures ngrok static domain | Writes ngrok config with the user's assigned free static domain |
| 5 | Drops Cowork plugin folder | Writes plugin files to Cowork's watched folder automatically |
| 6 | Registers in Add/Remove Programs | Standard Windows entry enabling clean uninstall |

### 3.2 ngrok Onboarding Screen

The only manual step for the user is a one-time ngrok account setup. The installer presents a simple screen:

1. User goes to ngrok.com and creates a free account (~2 minutes)
2. Copies their auth token from the ngrok dashboard
3. Their static domain is automatically assigned (e.g. `lucky-eel-42.ngrok-free.app`)
4. Pastes token into the installer onboarding screen
5. Installer writes ngrok config and starts the service
6. Installer tests the connection before completing

### 3.3 What Uninstall Does

- Removes VSTO add-in from MS Project
- Stops and removes the ngrok Windows service
- Deletes ngrok config and agent files
- Removes the Cowork plugin folder from disk
- Removes the `vectra-connect` entry from Cowork config — surgical, leaves all other Cowork config intact
- Removes from Windows Add/Remove Programs
- Optional prompt: *Keep Claude Notes data in existing .mpp files? Yes / No*

---

## 4. Cowork Plugin Structure

### 4.1 Folder Layout

```
vectra-connect/
├── .claude-plugin/
│   └── plugin.json            ← manifest: name, version, description
├── .mcp.json                  ← connector: points to user's ngrok public URL
├── skills/
│   ├── project-scheduling.md  ← MS Project domain knowledge for Claude
│   └── vectra-schema.md       ← how to read Vectra's data format
└── commands/
    ├── analyze-schedule.md    ← /vectra:analyze-schedule
    ├── flag-risks.md          ← /vectra:flag-risks
    └── add-note.md            ← /vectra:add-note
```

### 4.2 plugin.json

```json
{
  "name": "Vectra Connect",
  "description": "Brings Claude into MS Project — read schedules, analyze risks, annotate tasks",
  "version": "1.0.0",
  "author": "Vectra",
  "skills": [
    "skills/project-scheduling.md",
    "skills/vectra-schema.md"
  ],
  "commands": [
    "commands/analyze-schedule.md",
    "commands/flag-risks.md",
    "commands/add-note.md"
  ]
}
```

### 4.3 .mcp.json (Connector)

```json
{
  "mcpServers": {
    "vectra-connect": {
      "url": "https://[user-static-domain].ngrok-free.app",
      "transport": "http"
    }
  }
}
```

> The `[user-static-domain]` is written dynamically by the installer using the token entered during onboarding. Every user gets their own unique URL baked in at install time.

---

## 5. MCP Tools (What Claude Can Call)

All tools are exposed by the local MCP server running inside the VSTO add-in.

| Tool | Direction | Description |
|------|-----------|-------------|
| `get_tasks` | Read | Full task list: names, dates, % complete, resources, predecessors |
| `get_task_by_id` | Read | Single task by ID with all fields |
| `get_critical_path` | Read | Critical path tasks only |
| `get_resource_assignments` | Read | All resource-to-task assignments |
| `add_note_to_task` | Write | Writes to the Claude Notes field (Text30) on a specific task |
| `set_custom_field` | Write | Writes to any custom text field on a task |
| `flag_task` | Write | Sets a risk or attention flag on a task |
| `get_open_project_name` | Read | Name and file path of the currently open .mpp file |

> All write operations are staged — Claude proposes the change, the planner confirms in the Cowork chat, then the VSTO add-in applies it to the open `.mpp`. MS Project is never modified automatically.

---

## 6. Settings Panel (Ribbon)

A **Settings** button in the Vectra Connect ribbon group opens the management panel. This is where the user controls the tunnel and plugin without reinstalling.

### 6.1 Layout

```
┌─────────────────────────────────────────┐
│  VECTRA CONNECT — Settings              │
├─────────────────────────────────────────┤
│  Tunnel                                 │
│  Status:  ● Connected                   │
│  URL:     abc123.ngrok-free.app         │
│  Last connected: 2 min ago              │
│                                         │
│  [Reconnect Tunnel]  [Copy URL]         │
├─────────────────────────────────────────┤
│  ngrok Token                            │
│  ••••••••••••••••••  [Edit]             │
├─────────────────────────────────────────┤
│  Cowork Plugin                          │
│  Status:  ● Installed                   │
│  [Reinstall Plugin]                     │
├─────────────────────────────────────────┤
│  [Save]              [Uninstall All]    │
└─────────────────────────────────────────┘
```

### 6.2 Controls

| Control | Function |
|---------|----------|
| Tunnel Status indicator | Live green/red dot, updates in real time, shows timestamp of last connection |
| Reconnect Tunnel | Kills and restarts the ngrok process to re-establish a dropped connection — one-click fix |
| Copy URL | Copies the live public ngrok URL to clipboard for debugging or sharing |
| Edit Token | Update ngrok auth token without reinstalling |
| Reinstall Plugin | Rewrites the Cowork plugin folder to disk fresh — fixes corruption |
| Uninstall All | Triggers the full clean uninstall sequence |

---

## 7. Key Technical Decisions

| # | Decision | Rationale |
|---|----------|-----------|
| 1 | ngrok over Cloudflare Tunnel | Cloudflare proven too complex in practice. ngrok is self-contained and simpler. |
| 2 | Free ngrok static domain | One free static domain per account. URL never changes. No paid plan needed. |
| 3 | ngrok as Windows Service | Silent, starts with Windows, no console or tray icon. Invisible to user. |
| 4 | Single installer .exe | User experience: download one file, run it, done. Nothing else to install separately. |
| 5 | Plugin folder drop | Cowork watches a folder. Installer drops files there. No app store approval needed. |
| 6 | Per-user ngrok token | Each user authenticates their own ngrok account. No shared infrastructure. |
| 7 | Staged writes — approval required | Claude proposes, planner confirms. `.mpp` is never modified automatically. |
| 8 | Claude Notes in Text30 | Dedicated custom field keeps Claude output separate from human notes. |
| 9 | Settings panel in ribbon | Reconnect tunnel, edit token, reinstall plugin — all manageable without reinstalling. |
| 10 | No cloud server required | Entire system runs locally. ngrok is the only external dependency (free tier sufficient). |

---

## 8. Recommended Build Order

V3 builds on top of V2 (MCP module). Assumes the VSTO add-in and embedded MCP server from V2 already exist.

| Step | Component | What to Build |
|------|-----------|---------------|
| 1 | ngrok integration | Bundle ngrok agent, write Windows service installer, write ngrok config from token |
| 2 | Installer onboarding screen | WinForms screen: token input, connection test button, success confirmation |
| 3 | Plugin folder writer | Installer step that drops the Cowork plugin folder at the correct Cowork path |
| 4 | .mcp.json generator | Writes connector file with the user's ngrok URL baked in |
| 5 | Settings panel | WinForms panel: tunnel status, reconnect, copy URL, edit token, reinstall plugin |
| 6 | Tunnel status monitor | Background thread pinging ngrok health endpoint, updates status indicator live |
| 7 | Uninstall sequence | Extend existing uninstaller to also stop ngrok service and remove Cowork plugin folder |
| 8 | End-to-end test | Install → open MSP → Cowork connects → chat about schedule → confirm Claude sees tasks |

---

## 9. Open Questions for Next Session

- What is the exact Cowork plugin folder path on a standard Windows install? Needs verification.
- Does Cowork pick up plugin folder changes live, or does it require a restart?
- Should the ngrok tunnel start when MS Project opens, or run as a persistent always-on Windows service?
- Multi-project scenario: if two `.mpp` files are open simultaneously, how does the MCP server handle context switching?
- Should the Settings panel show Cowork connection status — i.e. is Cowork currently actively talking to the tunnel?

---

## 10. Related Documents

| Document | Covers |
|----------|--------|
| `VectraConnect_Export.md` | V1 — VSTO export to Obsidian vault. VectraKey schema, Excel output structure, pull/push design. |
| `VectraConnect_MCP_Handoff.md` | V2 — VSTO embedded MCP server. All MCP tools, TaskContextPanel UI, Claude Notes format, build order. |
| `VectraConnect_Cowork_Handoff.md` | V3 — This document. Cowork plugin, ngrok tunnel, settings panel, installer design. |

---

*Generated April 2026 | Vectra Connect V3 | Cowork Plugin Design Session*
