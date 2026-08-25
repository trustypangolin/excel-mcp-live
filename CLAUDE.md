# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Goal

`excel-mcp-live` gives an AI assistant control of a Microsoft Excel workbook the **user** already
has open — no separate open/close lifecycle, no server-managed Excel process. It's the Excel
counterpart to [`../word-mcp-live`](../word-mcp-live)'s live COM tools, built by porting that
project's attach pattern (`word_document_server/core/word_com.py`) from Word to Excel.

This exists because the workspace's other Excel automation project,
[`../mcp-server-excel`](../mcp-server-excel), works the opposite way by design: its `SessionManager`
**launches** its own invisible `Excel.Application` process, owns it, and force-kills it on
shutdown (see `src/ExcelMcp.ComInterop/Session/SessionManager.cs`). That's the right model for
scripted, from-scratch automation, but it can't attach to a workbook the user is looking at.
`excel-mcp-live` fills that specific gap: attach to what's already open, edit it live, never touch
its process lifetime. Use `mcp-server-excel` for Power Query, DAX, PivotTables, VBA, charts, and
anything that should run start-to-finish without a human's Excel window in the loop; use
`excel-mcp-live` when the user has a workbook open and wants the assistant to edit *that* window.

## Non-negotiable design constraint

**This server must never own an Excel process.** Every tool attaches via
`excel_document_server/core/excel_com.py`'s `find_workbook()` /
`list_all_open_workbooks()`, which use `win32com.client.GetActiveObject("Excel.Application")` plus
a Running Object Table (ROT) scan — the same pattern as `word_com.py:get_word_app()` /
`_find_word_with_docs()`. No tool may call `Application.Quit()`, and no tool may close a workbook
the caller didn't explicitly ask to close via an explicit "close" tool (there isn't one yet — see
Roadmap). An `Excel.Application` process hosts *every* workbook the user has open, not just the
target one — unlike Word, killing it takes out all of the user's other open work, not just one
document. If you ever add a "start Excel" or "create workbook" tool here, it does not belong in
this project — that's `mcp-server-excel`'s job.

A corollary: because Excel commonly runs multiple separate `Excel.Application` processes at once
(more so than Word), lookups walk the full ROT rather than trusting a single `GetActiveObject()`
call — see `list_all_open_workbooks()`.

## Known limitation: no undo grouping

Word's COM API exposes `Application.UndoRecord.StartCustomRecord`/`EndCustomRecord`, letting
`word_com.py`'s `undo_record()` context manager collapse a multi-step edit into a single Ctrl+Z.
**Excel has no equivalent.** Its closest analog, `Application.OnUndo(text, procedure)`, registers a
macro to *re-run* on undo, not a way to group prior COM calls. So a tool that performs several COM
writes to accomplish one logical edit (e.g. writing headers + data + formatting for one table)
will require multiple Ctrl+Z's to fully undo, not one. Don't claim single-undo behavior in tool
descriptions or the README. If this needs solving later, the likely approach is a
snapshot-before/restore-on-demand helper rather than a true undo hook.

## Commands

```powershell
# Install (editable, for local development)
pip install -e ".[dev]"

# Run the server directly against source (no install needed)
uv run excel_mcp_server.py

# Syntax-check everything (no test suite yet — see Roadmap)
python -m py_compile excel_document_server/**/*.py
```

There is no automated test suite yet. Like `mcp-server-excel` (see its `docs/ADR-001-NO-UNIT-TESTS.md`),
COM automation isn't meaningfully mockable — real verification means opening Excel with a workbook
and exercising tools against it by hand, or wiring up integration tests that require a live Excel
install (not yet done here).

## Architecture

```
excel_document_server/
  main.py              FastMCP server — every tool registered here, description=impl.__doc__
                        pulls each tool's docstring in as its MCP description (single source
                        of truth — write real docstrings, don't duplicate them in main.py)
  defaults.py           MCP_AUTHOR env var → comment author default
  core/
    excel_com.py         COM attach layer: list_all_open_workbooks(), find_workbook(),
                          find_worksheet() — the only place GetActiveObject/ROT-scanning happens
  utils/
    conversions.py       to_jsonable() (COM tuples/dates → JSON), hex_to_bgr() (Excel's
                          Font.Color/Interior.Color use BGR int, not RGB hex)
  tools/
    workbook_tools.py    list/inspect/save open workbooks
    sheet_tools.py        add/delete/rename/activate worksheets
    range_tools.py        read/write cell values & formulas, find/replace, clear
    format_tools.py       font/fill/number-format/alignment, autofit
    comment_tools.py      cell comments (notes)
```

Every tool function follows the same shape: guard on `sys.platform == "win32"`, lazily import from
`core.excel_com` inside the function body (keeps the module importable on non-Windows), call
`find_workbook()`/`find_worksheet()` to attach, do the COM operation, return a JSON string
(`json.dumps({...})` on success, `json.dumps({"error": ...})` on failure — tools never raise).

## Roadmap / not yet implemented

- Charts, PivotTables, Power Query, DAX, VBA — intentionally out of scope; that's `mcp-server-excel`.
- No explicit "close workbook" or "detach" tool.
- No macOS support (Word's live tools have a JXA backend for macOS; Excel does not here yet).
- No screenshot/export-to-verify tool (PowerPoint's `mcp-server-powerpoint` has one; could be a
  useful port).
- No automated tests.
