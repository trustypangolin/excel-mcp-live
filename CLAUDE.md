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

## Known COM gotcha: parameterized properties/methods lie about their calling convention

Discovered three times so far while adding tools — expect to hit it again on any new COM surface:

- **`Range.Address(False, False)`** raised `'str' object is not callable`. `Address` is a COM
  property with all-optional parameters; early-bound (gencache) pywin32 dispatch resolves it
  *eagerly* to a plain string using the default arguments, rather than returning something
  callable. Fixed with `com_range_address()` in `excel_com.py`, which checks `callable()` first.
  Same issue hit `Comment.Text()`.
- **`Range.Resize(rows, cols)`** was worse — it didn't raise, it silently wrote data to the *wrong
  cell*. Same eager-property resolution meant `.Resize` was already the unresized range, and
  `(rows, cols)` got reinterpreted by win32com as `Range.Item(row, col)` — an unrelated default
  indexer call. Fixed by building the target address as a plain A1 string (`column_letter()` in
  `conversions.py`) and calling `ws.Range(that_string)` instead of `.Resize(...)`. Never call
  `.Resize(...)` in this codebase — build an address string instead.
- **`Range.Sort(..., Header=xlYes)`** as a keyword argument was silently ignored — the header row
  got sorted into the data instead of staying in place. Fixed in `structure_tools.sort_range()` by
  excluding the header row from the range ourselves before calling `Sort`, rather than trusting its
  own header-detection. Separately, passing positional `None` placeholders for `Sort`'s unused
  middle parameters (`Key2`, `Type`, `Order2`, `Key3`, `Order3`) raised `int() argument must be...
  not 'NoneType'` — pywin32's early-bound stub applies `int()` to each positional slot. Only pass
  the keyword arguments actually needed; never pad with positional `None`.
- **`Hyperlinks.Add(..., TextToDisplay=...)`** as a creation-time keyword argument was also
  silently ignored — the cell showed the raw URL instead of the requested display text. Fixed in
  `range_extras_tools.add_hyperlink()` by setting the anchor cell's `.Value` directly *after*
  calling `Add()`, rather than trusting the constructor's keyword argument. General rule this
  keeps confirming: after creating something via a COM method call with several keyword arguments,
  verify the result actually reflects what you passed — don't assume a keyword argument took
  effect just because the call didn't raise.
- **`ListObject.TableStyle`** reads back as a `TableStyle` COM object, not the plain string it
  accepts on write — `json.dumps()` raised `Object of type CDispatch is not JSON serializable`
  until `table_tools.get_table_info()` unwrapped it via `.Name`. A reminder that a COM property's
  read type and write type aren't guaranteed to match — check `hasattr(value, "Name")` (or similar)
  before assuming a property you read back is already a plain Python value.

**The pattern to follow for any new COM call added here:** don't trust a parameterized
property/method's default or keyword-argument behavior without testing it live against a real,
open Excel instance first (`git status`-clean throwaway edits work fine for this). If it's a
property, check `callable()` before invoking. If it's a method with several optional parameters,
pass only the keyword arguments you actually need — never positional placeholders, never
assume an optional flag like `Header` is actually being honored without verifying the result.

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
                          Font.Color/Interior.Color use BGR int, not RGB hex), column_letter()/
                          column_number() (A1-string column math, e.g. building "B2:D5" by hand)
  tools/
    workbook_tools.py    list/inspect/save open workbooks
    sheet_tools.py        add/delete/rename/activate worksheets
    range_tools.py        read/write cell values & formulas, find/replace, clear
    structure_tools.py    insert/delete rows & columns, sort a range
    format_tools.py       font/fill/number-format/alignment, autofit
    conditional_format_tools.py  cell-value rules, color scales
    named_range_tools.py  list/read/write/create/update/delete named ranges
    range_extras_tools.py  data validation, merge/unmerge, hyperlinks
    table_tools.py         Excel Tables (ListObjects): create/list/inspect/append/read/style/delete
    chart_tools.py         charts: create/list/retype/retitle/add series/delete
    comment_tools.py      cell comments (notes)
```

Every tool function follows the same shape: guard on `sys.platform == "win32"`, lazily import from
`core.excel_com` inside the function body (keeps the module importable on non-Windows), call
`find_workbook()`/`find_worksheet()` to attach, do the COM operation, return a JSON string
(`json.dumps({...})` on success, `json.dumps({"error": ...})` on failure — tools never raise).

## Roadmap / not yet implemented

- PivotTables, Power Query, DAX, VBA are intentionally out of scope; that's `mcp-server-excel`.
- No explicit "close workbook" or "detach" tool.
- No macOS support (Word's live tools have a JXA backend for macOS; Excel does not here yet).
- **Range/sheet screenshot was attempted and shelved** — see the unmerged `feature/screenshot`
  branch. The standard VBA technique (`Range.CopyPicture` → paste onto a temp chart → `Chart.Export`)
  hit three different failure modes across six live-test rounds: `Appearance=xlPrinter` raises
  outright; calling `Worksheet.Activate()` right before `CopyPicture(xlScreen)` also raises, even
  with the message queue pumped afterward; and `Chart.Paste()` doesn't reliably accept
  `CopyPicture`'s own clipboard format (silently produces a blank exported image — worked around by
  pasting onto the worksheet first, re-copying that shape, then pasting *that* into the chart, which
  got further but still hasn't been confirmed producing real content). If picking this back up,
  start from that branch's history rather than from scratch — each dead end is documented in its
  commit messages.
- No automated tests.
