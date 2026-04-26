# Host Desktop Observation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add host-side desktop observation APIs for cursor, foreground, hover, window list, screen state, and VMConnect pointer mapping.

**Architecture:** Extend the existing PowerShell native interop in `src/SystemAccess.Core.ps1` with focused window inspection functions. Expose those functions through the existing stdio MCP, HTTP MCP, REST web server, docs, and smoke tests.

**Tech Stack:** Windows PowerShell 5.1, user32.dll P/Invoke, System.Drawing, System.Windows.Forms, MCP JSON-RPC, HttpListener.

---

### Task 1: Tests First

**Files:**
- Modify: `tests/Invoke-SmokeTests.ps1`

- [ ] Add smoke assertions for the new core functions and MCP tool names.
- [ ] Run `.\tests\Invoke-SmokeTests.ps1` and confirm it fails because the new functions/tools are missing.

### Task 2: Core Observation Functions

**Files:**
- Modify: `src/SystemAccess.Core.ps1`
- Modify: `src/SystemAccess.HyperV.ps1`

- [ ] Add user32 P/Invoke methods for cursor, foreground, class name, process id, visibility, parent/root lookup, and window enumeration.
- [ ] Add cursor/window/screen state PowerShell functions.
- [ ] Add Hyper-V pointer mapping helper.

### Task 3: Transports And Docs

**Files:**
- Modify: `Start-McpServer.ps1`
- Modify: `Start-WebServer.ps1`
- Modify: `README.md`
- Modify: `docs/manual-hyperv-tests.md`

- [ ] Add MCP tool schemas and call handlers.
- [ ] Add REST endpoints.
- [ ] Document tools and endpoint names.

### Task 4: Verification

**Files:**
- Run: `tests/Invoke-SmokeTests.ps1`

- [ ] Run the full smoke test suite.
- [ ] Report failures honestly if environment restrictions prevent screenshots or SendInput checks.

