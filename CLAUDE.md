# CLAUDE.md — mcp-access MCP Server

## Overview

MCP server for reading and editing Microsoft Access databases (`.accdb`/`.mdb`) via COM automation (pywin32). Runs as stdio MCP server. Entry point: `access_mcp_server.py`. Implementation: `mcp_access/` package (~7500 lines across 20 modules).

## Architecture

- **Singleton COM session** (`_Session`): one `Access.Application` instance shared across all tool calls. Opening a different `.accdb` closes the previous one.
- **Dedicated COM thread** (`_com_executor`): All tool calls run in a single-threaded `ThreadPoolExecutor` with `CoInitialize()`. This keeps COM in one STA thread while the asyncio event loop stays free to read/write stdio.
- **Caches**: `_parsed_controls_cache` (control parsing) and `_Session._cm_cache` (CodeModule COM objects — live COM proxies). Both invalidated on DB switch, object modification, and design operations. There is **no** Python-side cache of VBE text: `_cm_all_code()` always reads via `cm.Lines(1, total)` so external edits (manual VBE edits, Ctrl+Z, add-ins) are picked up immediately. See issue #26 for the reason this cache was removed.
- **Binary section handling**: `ac_get_code` strips PrtMip/PrtDevMode from form/report exports; `ac_set_code` restores them automatically before import.

## Tools (68 total)

| Category | Tools |
|----------|-------|
| **Database** | `access_create_database`, `access_close` |
| **Objects** | `access_list_objects`, `access_get_code`, `access_set_code`, `access_export_structure`, `access_delete_object`, `access_create_form`, `access_build_form`, `access_clone_object` |
| **SQL/Tables** | `access_execute_sql`, `access_execute_batch`, `access_table_info`, `access_search_queries`, `access_search_data`, `access_create_table`, `access_alter_table` |
| **VBE line-level** | `access_vbe_get_lines`, `access_vbe_get_proc`, `access_vbe_module_info`, `access_vbe_replace_lines`, `access_vbe_find`, `access_vbe_search_all`, `access_vbe_replace_proc`, `access_vbe_patch_proc`, `access_vbe_append` |
| **Controls** | `access_list_controls`, `access_get_control`, `access_create_control`, `access_delete_control`, `access_set_control_props`, `access_set_multiple_controls`, `access_manage_tab_order` |
| **UI lint** | `access_lint_form` |
| **DB Properties** | `access_get_db_property`, `access_set_db_property`, `access_get_form_property`, `access_set_form_property` |
| **Text Export/Import** | `access_export_text`, `access_import_text` |
| **Linked Tables** | `access_list_linked_tables`, `access_relink_table` |
| **Relationships** | `access_list_relationships`, `access_create_relationship`, `access_delete_relationship` |
| **VBA References** | `access_list_references`, `access_manage_reference` |
| **Maintenance** | `access_compact_repair`, `access_decompile_compact` |
| **Screenshot & UI** | `access_screenshot`, `access_ui_click`, `access_ui_type` |
| **Queries** | `access_manage_query` |
| **Indexes** | `access_list_indexes`, `access_manage_index` |
| **VBA Compilation** | `access_compile_vba`, `access_vbe_check_syntax` |
| **VBA Execution** | `access_run_macro`, `access_run_vba`, `access_eval_vba` |
| **Export** | `access_output_report` |
| **Data Transfer** | `access_transfer_data` |
| **Field Properties** | `access_get_field_properties`, `access_set_field_property` |
| **Startup Options** | `access_list_startup_options` |
| **Cross-reference** | `access_find_usages`, `access_find_definition` |
| **Knowledge base** | `access_tips` |

## Key Implementation Details

### Encoding in ac_set_code
- **Modules** (`.bas`): written using the system ANSI codepage (`locale.getpreferredencoding()`, typically `cp1252` on Western Windows) — no BOM.
- **Forms, reports, queries, macros**: written as `utf-16` (UTF-16LE with BOM) — Access LoadFromText expects this.

### Control parsing (_parse_controls)
The Access export format nests controls inside sections:
```
Begin Form
    Begin                    <- defaults block (NOT controls)
    End
    Begin Section            <- section (Detail, FormHeader, FormFooter)
        Begin                <- container
            Begin Label      <- REAL CONTROL
            End
            Begin Page       <- CONTAINER -- children re-scanned
                Begin        <- anonymous wrapper
                    Begin ComboBox  <- child control (parent = Page)
                    End
                End
            End
        End
    End
End Form
```
**Container types** (`_CONTAINER_TYPES = {"Page", "OptionGroup"}`): re-scanned for child controls. Children get a `"parent"` field. `container_stack` tracks nesting.

**Depth counter inside a control block must include `Property = Begin`** (e.g. `GUID = Begin`, `NameMap = Begin`, `ConditionalFormat = Begin`). These open multi-line blocks closed by their own `End`. If the parser only counts plain `Begin <Type>` it decrements depth on the closing `End` of the property block without ever incrementing — the control closes prematurely at the first such `End`, and any controls that follow inside a `Page` / `OptionGroup` are silently lost. Fixed in v0.7.34 (was: `re.match(r"^Begin\b", bl_s)` — now also matches `r"^\w+\s*=\s*Begin\s*$"`, mirroring the form-level loop).

### VBE + Design view conflict
After design operations (`ac_set_control_props`, `ac_create_control`, `ac_delete_control`), the form may remain open in Design view. All VBE write functions close the form first (DoCmd.Close with acSaveYes), invalidate `_cm_cache`, then access VBE. Without this: `"Catastrophic failure" (-2147418113)`. All design operations invalidate all three caches in their `finally` block.

### VBE procedure kinds
VBE `ProcStartLine`/`ProcBodyLine`/`ProcCountLines`/`ProcOfLine` require a `kind` argument (`_ALL_PROC_KINDS = (0, 1, 2, 3)`): 0=Sub/Function, 1=Property Let, 2=Property Set, 3=Property Get. `_proc_kind()` iterates all four. `ac_vbe_module_info` deduplicates by `(name.lower(), keyword.lower())` so paired Get/Let/Set appear separately.

### Class module format (LoadFromText vs VBE)
Two **different** export/import formats exist:
- **VBE** (`VBComponent.Export/Import`): `VERSION 1.0 CLASS` header + `Attribute VB_Name`. This is `.cls` file format.
- **Access** (`SaveAsText/LoadFromText`): just the 4 `Attribute VB_*` lines (GlobalNameSpace, Creatable, PredeclaredId, Exposed), NO version header.

Passing VBE-style headers to `LoadFromText` creates a corrupt standard module. `_ensure_class_module_header()` strips VBE headers and injects the correct LoadFromText-style attributes.

### Dialog watchdog system
Blocking COM calls (`OpenCurrentDatabase`, `CompactRepair`, `RunCommand`, `Application.Run`) are protected by polling watchdog threads that dismiss Access dialogs via `_dismiss_access_dialogs()` / `_dismiss_dialogs_by_pid()`. Button priority: Cancel > End > OK (Cancel-first avoids advancing wizards).

**Attached-instance policy (v0.7.43)**: the global watchdog also runs when we attached to the user's Access, but it only dismisses dialogs while one of OUR tool calls has been in flight longer than the grace period (5 s attached vs 3 s spawned). `_Session._tool_started` (monotonic timestamp, set/cleared by `server.call_tool` around `run_in_executor`) is the in-flight signal. A modal with no tool call in flight belongs to the interactive user — never touched. Do NOT "simplify" this back to disabling the watchdog on attach: that re-introduces the 1-hour VBE hang when a broken-reference VBA project pops "Error accessing file..." during one of our calls.

**Dismissal note (v0.7.44, issue #31)**: `_dismiss_dialogs_by_pid` (the funnel every watchdog goes through) records `(monotonic, title)` in `_Session._last_dismissed`; `server.call_tool` appends a "a modal dialog (X) was auto-dismissed during this call" note to the tool result when the timestamp falls inside the call. Cancel-first dismissal can alter outcomes (e.g. cancel a save prompt) — the note makes that traceable instead of silent.

**Eval/delete hardening (v0.7.44, issue #31)**:
- `ac_eval_vba` accepts optional `timeout` with the same `_dialog_watchdog` treatment as `ac_run_vba`, wrapping BOTH `Application.Eval` and the temp-module fallback.
- `_eval_via_temp_module` sweeps orphan `_mcp_eval_wrapper` temp modules (`_sweep_orphan_eval_modules`, marker in the first 10 lines of std modules only) before creating a new one — a failed `Remove` used to wedge every later call with "cannot find the procedure 'Module1._mcp_eval_wrapper'". Deliberately NOT in `connect()`: that would cost VBE access (Trust Center + Visible) on every tool call.
- `ac_delete_object` calls `_save_all_modules(app)` (best-effort: `RunCommand 280` = acCmdSaveAllModules, fallback per-module `DoCmd.Save`) before `DoCmd.DeleteObject` to prevent the "save changes to the design of module X?" prompt — dirty state can come from user code run via eval, so per-tool bookkeeping isn't enough. Do NOT extend this to close/quit paths: on attached instances it would silently persist the interactive user's half-finished VBE edits.

### Wedged-session detection (v0.7.43, from PR #30 by @CaptainStormfield)
A DB whose startup code closes it during the open (startup-form error path + `AllowBypassKey=False`) used to leave `_db_open` pointing at a dead database — every later call died at `CurrentDb` and reconnects re-attached to the same broken instance. Now: `_switch()` validates `CurrentDb() is not None` post-open (raises an actionable RuntimeError after `quit()`-resetting the session), `connect()` health-checks `CurrentDb()` whenever `_db_open` is set (auto-reconnect via `_force_cleanup()`), and `ac_create_database` validates its reopen. Cost: one extra `CurrentDb()` COM round-trip per tool call — accepted trade-off.

### Multi-object scans must not lie with "0 matches" (v0.7.43)
`ac_vbe_search_all` / `ac_find_usages` / `ac_find_definition` collect per-object failures into `errors` (capped at `_SEARCH_ERROR_CAP = 20`) + `objects_skipped` + a `warning`, instead of `except: continue`. A VBA project that fails to load (broken reference, Trust Center) makes EVERY object fail — a clean `total: 0` was a false "doesn't exist". Same idea in `ac_list_references`: each reference property is read defensively (broken references raise `com_error` on `FullPath`), never kill the listing.

### Application.Run via InvokeTypes
`Application.Run` has 31 params (1 required + 30 optional). pywin32's late-bound `Dispatch` can't handle this. `_invoke_app_run()` calls `_oleobj_.InvokeTypes()` directly with `pythoncom.Missing` padding. Same approach for `Application.Eval` via `_invoke_app_eval()`.

## Adding a new tool

1. Write the implementation function (e.g. `ac_new_tool()`)
2. Add a `types.Tool(...)` entry to the `TOOLS` list
3. Add an `elif name == "access_new_tool":` branch in `call_tool()`
4. Update the tool count — see the release checklist below, it lives in **five**
   places and they drift

## Release checklist (the docs drift, every time)

The tool count and the version live in more places than you remember.

**The README no longer keeps its own changelog** (removed in v0.7.53): it was a
700-line duplicate of `CHANGELOG.md` that had to be updated in parallel and
drifted every single release. It now just links to `CHANGELOG.md`. Do NOT
re-add per-version entries there.

- `CHANGELOG.md` — new `## 0.7.NN — date` entry at the top. This is the only
  changelog.
- `README.md` — **three** spots: the tagline count (line ~7), the `## Tools (N)`
  heading, and the tool tables themselves.
- `CLAUDE.md` — `## Tools (N total)` heading + the category table.
- `pyproject.toml` — `version` **and** the `— N tools.` in `description`.
  The description is baked into the published artifact and PyPI versions are
  immutable, so a miss there can only be fixed in the NEXT release.
- `server.json` — `version` appears **twice**.

Then: commit, `git tag v0.7.NN`, `git push && git push --tags`. The tag push
triggers `.github/workflows/publish.yml` (PyPI + MCP Registry). Check it landed
with `gh run list --limit 3`.

## Office version autodetect (v0.7.36+)

`_Session._office_version` / `_Session._office_msaccess` are populated by `_Session._detect_office_install()` — a one-shot probe that enumerates `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path` under HKLM, HKLM\\WOW6432Node and HKCU (per-user C2R), picks the highest matching version with a working `MSACCESS.EXE`, falls back to `App Paths\MSACCESS.EXE\(Default)`, and finally to the previous hardcoded `16.0` / `Office16` defaults. Used by:
- `_Session._suppress_recovery_dialog` — Resiliency registry key path
- `_Session._decompile` — MSACCESS.EXE path for `/decompile` subprocess
- `maintenance.ac_decompile_compact` — same

Detection is idempotent (`_office_detected` flag) and never raises. Schema of `access_decompile_compact` is unchanged.

## Macros (v0.7.36 docs upgrade)

Macros have always been fully supported via the regular code tools — no dedicated tool exists. The workflow is:
- List: `access_list_objects(object_type="macro")`
- Read: `access_get_code(object_type="macro", name=...)`
- Write: `access_set_code(object_type="macro", name=..., code=...)` — UTF-16 encoded
- Run: `access_run_macro(macro_name=...)`
- Delete: `access_delete_object(object_type="macro", object_name=..., confirm=true)`

`restore_binary_sections` does NOT apply to macros (they have no PrtMip/PrtDevMode). `access_tips('macros')` shows the workflow.

## Clone object (v0.7.36)

`access_clone_object` duplicates an object by raw `SaveAsText` → `LoadFromText`. Critical detail: it does its own `app.SaveAsText` + `read_tmp` directly (≈10 lines duplicated from `code.py:ac_get_code`) — explicitly NOT going through `strip_binary_sections`, so PrtMip / PrtDevMode / NameMap / GUID ride along inside the text. `ac_set_code` then sees the binaries are present and skips restoration (`code.py:348` — only restores when absent). For `class_module`, `_ensure_class_module_header(text, target_name)` re-runs so the implicit VB_Name stays consistent.

## Tab order (v0.7.36)

`access_manage_tab_order` uses **single-pass assignment** in target order — Access enforces TabIndex to be in `0..(N-1)` per section and auto-renumbers the rest to preserve uniqueness when you set one. Do NOT try to "park" controls at indices >= N (Access rejects with "The value you used for the TabIndex property isn't valid. The correct values are from 0 through N-1."). Skips non-tabbable types (100=Label, 101=Rectangle, 102=Line, 103=Image, 114=PageBreak, 118=Page). Optional `section` filter; defaults to all sections.

## UI design lint (v0.7.41)

`mcp_access/lint.py` is a **deterministic, pure-Python** design validator (no
LLM, no external service). `access_lint_form` returns structured JSON
violations; the same engine runs **automatically** on every design mutation.

### Why it exists
The LLM sets control coordinates/colours blind and used to accept objectively
broken layouts (white-on-white, overlap, truncation, inconsistent siblings,
out-of-bounds). The fix the user asked for: validation that lives *inside* the
MCP and **cannot be skipped or "talked past"** by the model. So the rules are
numeric and the result is attached to mutations whether the model asks or not.

### Architecture
- One SaveAsText export (via `ac_get_code`, binary sections already stripped),
  never opens Design view. `_build_model` layers a style dict (`_extract_style`,
  reading the control's `raw_block`) and section assignment onto the cached
  `_parse_controls` result, plus `_parse_geometry` (form Width/BackColor +
  per-section Height/BackColor/line-range).
- Rules: `contrast` (WCAG 2.1, `_decode_bgr`+`_contrast_ratio`), `overlap`
  (AABB, same section+parent only), `out_of_bounds`, `truncation`,
  `sibling_inconsistency`, `misalignment`, `invisible_or_zero_size`.
- `lint_compact()` (errors+warnings, heuristic measure, capped) is attached by
  `_attach_lint` to the result of `ac_set_control_props`,
  `ac_set_multiple_controls`, `ac_create_control`. Wrapped in try/except so a
  lint failure NEVER breaks the mutation. `skip_lint=true` opts out for bulk ops.

### Hard-won gotchas baked into the rules (do NOT "simplify" these away)
- **Absent dimension ≠ 0.** Access omits `Left/Top/Width/Height` (and `BackColor`)
  when they equal the form default. `_twips_opt` returns None for absent; rules
  use `_has_full_geom` and skip None. Treating absent as 0 caused false
  "zero-size" / bounds violations on inherited-default controls.
- **Opaque text control with no `BackColor` renders on white.** Access omits the
  default white BackColor — this is exactly how white-on-white slips through, so
  `_effective_background` defaults Label/TextBox/ComboBox/ListBox to white.
- **`ControlType =` is often absent** in modern exports; type comes from the
  `Begin <Type>` keyword. Rules key off `type_name`, not the int.
- **Attached labels are nested inside their control's block** → `_parse_controls`
  never enumerates them, so no overlap false positives there for free.
- **Access auto-grows form Width (and section Height) to fit controls**, so
  horizontal `out_of_bounds` rarely fires for forms (still useful for reports +
  negative coords). Not a bug — documented limitation.
- **`ConditionalFormat = Begin … End`** holds its own colours; `_extract_style`
  tracks block depth so only the control's own (depth-1) props are read.
- **System/theme colours** have the high bit `0x80000000` (e.g. `-2147483633`) —
  `_decode_bgr` flags them; contrast emits an `info` note instead of a number.
- **Conditional formatting** (`format_conditions`) overrides ForeColor/BackColor
  at runtime and is BINARY in the export — `_rule_contrast` skips those controls
  and notes them (can't verify the runtime colour statically).
- **Captions wrap.** Both Labels AND CommandButtons wrap their caption across
  lines; SaveAsText encodes the breaks as literal `\015\012`. `_caption_lines`
  splits them; truncation counts how many display lines the text needs
  (`ceil(line_width/avail)`) vs how many fit the height (`round(height/lineH)`,
  `lineH ≈ fontPt*20*1.2`). A 540-twip button shows 2 lines of 11pt — use
  `round`, not `floor`, or you under-count and false-flag.
- **Heuristic width is approximate.** Narrow UI fonts (Calibri/Tahoma) average
  ~0.46× the point size per glyph; the heuristic only flags a line as
  overflowing past **1.25×** the available width (absorbs metric error). WizHook
  uses 1.02×. Without these, bold header labels that fit get false-flagged.
- **Transparent buttons are a click layer, not an overlap.** A `Transparent=True`
  CommandButton stacked on a styled Label/Rectangle is the standard Access
  custom-button pattern (the label shows the colour, the invisible button takes
  the click) — `_rule_overlap` skips any pair where one side is a transparent
  button. (Classic command buttons ignore `BackColor` even with `UseTheme=No`
  on Win11/Office16, so this label+transparent-button trick is how you get
  coloured tiles.)
- **`sibling_inconsistency` clusters, not modes.** A form legitimately uses two
  sizes (tall main buttons + a row of short inline buttons). `_accepted_clusters`
  treats any value ≥2 controls share as a norm; only a lone outlier (and not a
  >2× different class like a memo box) is flagged. Needs ≥4 controls in the group.

### WizHook text measurement
`measure="auto"|"wizhook"|"heuristic"`. WizHook (`_measure_text_batch`) measures
exact rendered width in ONE COM round-trip via a temp std module +
`_invoke_app_run`. It REQUIRES a compiled VBA project (`Application.IsCompiled`);
during active development the ERP project is usually uncompiled, so it fails and
falls back to the conservative heuristic (a `note` is added when `measure` was
explicitly `wizhook`). The embedded lint always uses `heuristic` (fast, no Run
dependency). Default everywhere leans on the heuristic for reliability.

## Declarative form auto-layout (v0.7.45)

The LLM is poor at emitting absolute twip coordinates blind (overlaps,
out-of-bounds, ragged columns, invented colours). The fix is to take the
arithmetic away from the model — the same idea as the lint, applied at
*generation* time instead of validation time.

### Pieces
- **`mcp_access/design_defaults.py`** — the single source of truth for layout
  tokens: the 60-twip grid (`GRID`), margins/gaps, standard control sizes
  (`ROW_H=300`, `LABEL_W=1800`, `FIELD_W=2400`, `BUTTON_W/H`, `MEMO_H`…),
  fonts, and a **closed BGR palette** (`PALETTE`). `bgr(r,g,b)` builds an Access
  colour Long (BGR order, NOT RGB); `snap(v)` rounds to the grid. `lint.py`,
  `build_form.py` and `tips('layout')` all read from here — change a number once.
- **`access_build_form`** (`mcp_access/build_form.py`) — declarative form
  builder. The model passes a spec (`title`, ordered `fields`, `actions`,
  `layout` single|two-column, `theme` light|plain); `_plan_layout` (pure, unit
  tested in `tests/test_build_form_layout.py`) computes every rect from
  `design_defaults`; `ac_build_form` creates all controls in **one** Design-view
  session, sets the palette, sizes the form + header/footer sections, assigns a
  per-section tab order, then attaches the embedded lint. A form it builds passes
  the lint clean by construction.
- **`snap_to_grid`** (opt-in, default false) on `ac_create_control` /
  `ac_set_control_props` — rounds Left/Top/Width/Height to `GRID`. `-1` (auto)
  values are left untouched.

### Gotchas baked in
- `_plan_layout` references `_PREFIX` / `_looks_unbound` defined *below* it —
  fine because both are module globals resolved at call time, not import time.
- Two-column mode ignores `width_units` (keeps a strict grid); single-column
  honours it. Memo fields use `MEMO_H` and advance the running `y` cursor so
  taller rows don't overlap the next.
- A field name with spaces/operators is left **unbound** (`_looks_unbound`) — a
  plain column name gets a `ControlSource`. `theme="plain"` emits geometry only
  (no colours/fonts).
- `has_header` toggles BOTH FormHeader and FormFooter; when only one is needed
  the other section's Height is set to 0.

### New lint rules (v0.7.45) — all `info`-severity
`grid_alignment`, `spacing_consistency`, `edge_margin`, `hierarchy`. They enrich
the full `access_lint_form` report but **never** change the verdict (which only
counts errors/warnings) and **never** reach `lint_compact` (errors+warnings
only), so they can't make the embedded mutation lint noisier. Each keys off the
canonical grid/margin, so a `build_form` layout passes them clean. Deliberately
conservative: `hierarchy` only fires on an explicit FontSize inversion (action
text smaller than body), `spacing_consistency` needs ≥4 controls in a column.

## Design directions for build_form (v0.7.46)

The v0.7.45 themes (`light`/`polish`/`flat`) were invented by eye and looked
soso. v0.7.46 adds three **curated design directions** that translate real
design-system thinking into native Access. The only insertion point is
`_resolve_theme`; `_plan_layout` changed ~2 lines (`pal = T.get("palette",
D.PALETTE)`, title uses `T.get("title_font", font)`). `light`/`plain` stay
literal, so the pure tests are untouched.

### Pieces
- **`design_defaults.py`** (additive): `type_scale(base, ratio)` (modular scale
  → caption/body/subhead/title/display, whole points); `SPACE` (closed spacing
  scale — the legacy `MARGIN_X`/`GAP_LABEL`/`COL_GAP`/… are now **aliases** into
  it, same values, so old layouts/tests are byte-identical); `DENSITY`
  (compact/comfortable/spacious — margins & gaps ONLY, never control sizes);
  `DIRECTIONS` (the three bundles) + `DIRECTION_COMMON`. `PALETTE` is untouched
  (a test pins it; `light` still uses it).
- **The 3 directions** — `despacho` (Constantia serif title / Segoe UI body,
  teal band, warm paper, comfortable, no card), `panel` (Segoe UI Semibold /
  Segoe UI, slate band, white card on a cool canvas, comfortable), `archivo`
  (Cambria serif / Corbel, clay band, warm paper, spacious, no card). Palette
  keys are the canonical ones (`form_bg`/`field_bg`/`field_border`/`text`/
  `accent`; `accent` doubles as the band). Colours are built with `bgr()`
  **straight from the hex** so they can't drift — `test_directions_palette_anti_drift`
  recomputes `bgr(hex)` and `test_directions_contrast_wcag` re-derives every
  contrast with the lint's own WCAG maths.
- **Two `info` lint rules** — `generic_font` (closed list: Arial/Roboto/Inter/
  Times New Roman/MS Sans Serif) and a `type_hierarchy` extension of
  `_rule_hierarchy` (header title must be larger than body). `_parse_geometry`
  gained an additive `kind` per section (the `Begin` token) + `_assign_section_kind`
  / `section_kind` on each control, so the header title is found reliably even
  when a section exports without a `Name`.

### The two-tone band fix (v0.7.46)
The directions surfaced a latent v0.7.45 bug: `_set_section` resolved sections
via `Form.Section(index)`, which **pywin32 cannot late-bind** (every index
raises `-2147352573 "member not found"`). The call failed *silently* inside its
own try/except, so the canvas colour was never painted onto Detail and the
header/footer kept Access' oversized default heights — the themed light-blue
header then showed past the form-width accent rectangle (the "two-tone band").
Fix: `_get_section` resolves by the **named** property (`Detail`/`FormHeader`/
`FormFooter`, which DO bind) with the index as fallback; and a styled header
band is now painted on the section `BackColor` (which fills the full
document-window width) via `plan["header_backcolor"]`, with the Rectangle kept
as a fallback in case a theme overrides the section colour. Do NOT revert
`_get_section` to the indexed accessor — it re-introduces the silent failure.

## Build-a-form-from-scratch recipes (v0.7.38)

### Add VBA to a form you just created with ac_create_form

Before v0.7.38, calling `ac_set_code(form, "Option Compare Database\n...")`
on a freshly-created form failed with `errors while importing` — `LoadFromText`
was always invoked and `restore_binary_sections` had nothing to restore from.
Now `ac_set_code` detects VBA-only input (`_looks_like_vba_only`: no
`Version =` / `Begin Form`, but Option/Sub/Function/etc.) and routes through
`_inject_vba_after_import` (Design view → `HasModule=True` → VBE write). No
`LoadFromText` round-trip, layout preserved.

```
ac_create_form(db, "frmFoo")
ac_create_control(db, "form", "frmFoo", "CommandButton",
                  {"left": 100, "top": 100, "width": 1500, "height": 400},
                  control_name="btCerrar")  # NEW: top-level control_name
ac_set_code(db, "form", "frmFoo",
            "Option Compare Database\nOption Explicit\n"
            "Private Sub btCerrar_Click()\n"
            "    DoCmd.Close acForm, Me.Name\n"
            "End Sub\n")  # routes via VBE, not LoadFromText
```

If you need to write a full form export (e.g. cloning the binary sections of
another form), include `Version =NN` / `Begin Form` and the original
`LoadFromText` path runs. The two paths are mutually exclusive — the
detection in `_looks_like_vba_only` is the discriminator.

### Drop a control inside a TabControl Page

`ac_create_control` accepts `parent` (or `Parent` — case-insensitive since
v0.7.38) as a special key that maps to the 4th positional arg of
`CreateControl(form, type, section, parent, column, l, t, w, h)`. Passing
`Parent` with capital P used to fall through to `setattr(ctrl, "Parent", ...)`
which Access rejects with `"Property 'CreateControl.Parent' can not be set"`
— misleading because Parent IS available, just not via setattr.

```
ac_create_control(db, "form", "frmFoo", "CommandButton",
                  {"Parent": "tabGestion",    # case-insensitive special key
                   "Left": 100, "Top": 100, "Width": 2000, "Height": 500,
                   "Caption": "Acción", "OnClick": "[Event Procedure]"},
                  control_name="btMiAccion")
```

If `Parent` doesn't refer to an existing TabControl Page (or other container
like OptionGroup), the control lands in Detail at the requested coordinates
and `Parent` is silently ignored by CreateControl — same behaviour as VBA.

### Read VBE from a brand-new form

Before v0.7.38: `ac_vbe_module_info(form, "frmFoo")` on a form just made by
`ac_create_form` raised `Subscript out of range`. The error message blamed
the Trust Center, but the actual cause was `HasModule=False` — VBComponents
had nothing to return because the code module had not been created yet.
Now `_force_vbe_init` activates `HasModule` when opening the form in Design
view during the retry, so this works out of the box.

If you want to *be explicit*, the original workaround is still valid:

```
ac_set_form_property(db, "form", "frmFoo", {"HasModule": True})
ac_vbe_module_info(db, "form", "frmFoo")  # then this works too
```

## VBE procedure editing (v0.7.42)

Field-report fixes for `vbe.py`. Three behaviours to keep in mind:

- **`ProcStartLine` owns the blank separator above a proc** (it equals the
  previous proc's `End` + 1, so it includes the blank/comment lines VBE attributes
  to the proc). `ac_vbe_replace_proc` therefore, *when replacing*, counts the run
  of leading whitespace-only lines (`lead`) and deletes/inserts at `start + lead`
  over `count - lead` — preserving the separator. A pure delete (`new_code==""`)
  still deletes the whole `[start, count]` range (separator included) so a deleted
  proc doesn't leave an orphan blank. Do NOT "simplify" this back to
  `DeleteLines(start, count)` for the replace path — that re-introduces the
  blank-eating bug Tom reported.
- **The Option-placement health check is comment-header-aware**, not
  line-number-thresholded. `_check_module_health` flags an `Option …` line only
  when real code (non-blank, non-comment `'`/`Rem`, non-`Option`) already appeared
  above it. A banner comment header of any length is fine. Do NOT restore the old
  `i >= 5` threshold — it false-positived on long headers (e.g. `_modTest`).
- **`new_lines` is an alias for `new_code` in `access_vbe_replace_lines`.** The
  dispatcher (`_new_lines_to_code`) joins a list with `\n` (so `""` entries are
  blank lines) and tolerates a JSON-encoded string from string-serialising
  clients. A single-mode replace that deletes lines but inserts nothing appends a
  note — the silent destructive-delete footgun (wrong arg name → empty `new_code`
  → pure delete) is now surfaced, not hidden.

`start_line` vs `body_line` (get_proc / module_info): `start_line` is the VBE proc
start (includes the blank/comment lines above); `body_line` is the
`Sub`/`Function`/`Property` declaration line. Use `start_line` for whole-proc ops,
`body_line` for body line-range edits.

## VBE patching: atomic / case / (Declarations) (v0.7.52)

Field requests from @TvanStiphout-Home, all tested against a real database
before being filed. `_apply_patches` (`vbe.py`) is the pure, COM-free engine
extracted from `ac_vbe_patch_proc`'s old inline loop.

### The 4-tier match ladder — order is load-bearing
1. literal, case-sensitive → 2. ws-normalized, case-sensitive →
3. literal, case-insensitive → 4. ws-normalized, case-insensitive.
Tiers 3–4 only run when `match_case=false` (the default).

**ALL case-sensitive tiers run before ANY case-insensitive one.** That is what
makes the change byte-for-byte backwards compatible: every call that succeeds
today still lands on tier 1 or 2, exactly where it landed before. Interleaving
them (e.g. putting literal-CI between 1 and 2) would silently relocate calls
that currently succeed via the ws fallback. Do NOT reorder.

Case-insensitive replacement **cannot use `str.replace`** (it is case-sensitive).
It finds the index on a lowered copy and splices the ORIGINAL string by
position, inserting the caller's replacement text with its casing untouched.

### The Unicode length guard
`'İ'.lower()` (U+0130) returns TWO characters, so offsets computed on the
lowered copy no longer address the original — one such char in a comment would
splice the replacement into the middle of a line. `_case_insensitive_safe()`
checks `len(text.lower()) == len(text)` and the CI tiers are skipped (with a
note) when it fails. `casefold()` is worse (`ß`→`ss`); do not "improve" this.

### atomic is simulate-then-commit, NOT a pre-pass
`atomic=true` (default) decides AFTER running the whole patch loop in memory and
BEFORE any `DeleteLines`/`InsertLines`. A pre-pass validating every anchor
against the ORIGINAL text would be wrong in both directions: patch 0 can destroy
the anchor patch 3 cites (pre-pass says OK, real run half-writes) or create it
(pre-pass rejects a valid batch). Because the simulation and the commit are the
same single pass, divergence is structurally impossible.

The ABORTED message MUST keep telling the caller to re-send the **entire** batch.
A model that re-sends only the failed patches loses the ones that did match —
that failure mode is worse than the partial write atomic exists to prevent.

`require_unique` violations are collected into the same blocking list as
not-founds, so the atomic gate covers both uniformly. Occurrence counting uses
the same tier that produced the match, and reports absolute module line numbers
(`base_line + offset`). Note that `match_case=false` makes `require_unique`
*stricter* — a CI comparison can match more often than a CS one.

### (Declarations) as a target
`_is_declarations()` matches `"(declarations)"` case-insensitively after
`.strip()`. Deliberately NOT triggered by `""` — `ac_vbe_find` already reads
`""` as "the whole module".
- `ac_vbe_patch_proc` / `ac_vbe_get_proc`: resolve to `start=1`,
  `count=cm.CountOfDeclarationLines`, bypassing `_proc_bounds`.
- `count == 0` raises an actionable error pointing at
  `ac_vbe_replace_lines(start_line=1, count=0, ...)` — and it must, because
  `cm.Lines(1, 0)` and `cm.DeleteLines(1, 0)` both raise in VBE.
- `_strip_option_lines` is guarded by `not is_declarations and start > 5`. The
  old `start > 5` alone happened to be false at `start=1`, but relying on that
  coincidence would silently delete `Option Explicit` if the threshold ever moved.
- The final message reads `CountOfDeclarationLines`, never `ProcCountLines`
  (there is no proc; the bare `except` would report a bogus `0`).
- `ac_vbe_replace_proc` REFUSES `(Declarations)`: `new_code=""` would wipe
  `Option Explicit` plus every module-level `Const` in one unconfirmed call.
- `ac_vbe_module_info` gained an additive `declarations: {start_line, count}`.

### The off-by-one: cm.CountOfLines is the source of truth
`patch_proc` reported `cm.CountOfLines`; `module_info`/`get_lines` reported
`len(splitlines())`. VBE emits no trailing terminator, so a module ending in a
blank line makes `splitlines()` drop it → the two disagreed by exactly 1.

`_cm_lines_list()` splits and then **pads with `""` up to `cm.CountOfLines`**, so
`len(lines) == cm.CountOfLines` by construction. Every existing slice and bounds
check keeps working, the reported number becomes authoritative, and a trailing
blank line becomes readable via `get_lines` (it was rejected as out-of-range).

**Do NOT "fix" this the other way round** by switching `patch_proc` to
`splitlines()`. Its `count = min(count, total - start + 1)` is the clamp feeding
a destructive `DeleteLines`; changing that input to a 1-short value to make a
cosmetic message agree turns a display bug into code loss.

Related: `new_count` is now clamped like `replace_proc` does, and
`_check_module_health` receives `expected_total = total - count +
_vbe_line_count(inserted)` so its Check 3 stops being dead code.
`_vbe_line_count` counts a trailing CRLF as opening a further empty line —
`"a\r\nb\r\n"` → 3 — because that is what `InsertLines` does.

## access_vbe_check_syntax (v0.7.52)

The safe alternative to `access_compile_vba`, which is unusable as a post-edit
check: its step 0 calls `_Session._decompile` → a `MSACCESS.EXE /decompile`
subprocess, then either `Quit(1)` (= acQuitSaveNone) or `CloseCurrentDatabase()`
on the user's instance. **Unsaved VBA is discarded.** That stays as-is; the new
tool simply never goes near it.

Checks the ALREADY OPEN project: no decompile, no `RunCommand`, no Design view,
no second Access instance. Reuses the pure validators in `compile.py`
(`_check_blocks_in_module`, plus `_check_structure_in_module` extracted from
`_verify_module_structure` in this release) — `ac_compile_vba`'s behaviour is
unchanged, its wrappers just delegate now.

- Uses `_get_vb_project(app)`, **not** `VBE.ActiveVBProject`: the active project
  can be `acwzmain` (the wizard library) after a decompile/compact.
- Feeds the checkers `code.split("\n")`, not `splitlines()` — they were written
  against raw VBE text and their `" _"` continuation test sees the stray `\r`.
- **Never reports a clean 0 for something it could not read.** Per-module
  failures land in `skipped` and force `ok=false`; a project that fails to
  enumerate raises. Same rule as the multi-object scans.
- The `note` field states plainly that this is structural validation, not
  compilation: no identifier resolution, no type checking, no references. A
  caller who reads `ok=true` as "it compiles" is the failure mode to avoid.

`_check_structure_in_module` also gained an end-of-module check for an unclosed
`Type`/`Enum` block — everything below the opener is absorbed by it, so no line
inside the loop could ever have flagged it (the "Statement invalid inside Type
block" trap already documented under VBA Language Gotchas).

## access_compile_vba trigger hardening (v0.7.53, PR #35)

`ac_compile_vba` reads `Application.IsCompiled` as its success signal after
**deliberately dirtying the project** (step 0b) — so any path where the
Debug > Compile trigger silently fails leaves `IsCompiled=False` and used to be
misreported as a compile error in the user's code ("missing reference,
undeclared variable, or type mismatch") while manual Debug > Compile succeeded.

- **`_ensure_code_pane(app)`** runs before the trigger: makes a code pane of the
  CURRENT database's project active. Debug > Compile acts on the ACTIVE project
  and is only reliably enabled with a code pane focused; after a
  decompile/compact the active project is often `acwzmain`. Do NOT remove this
  step — without it `Execute()` raises DISP_E_EXCEPTION, no-ops, or compiles the
  wizard library. It short-circuits when a pane of our project is already
  active: re-`Show()`ing on every compile piles code windows into the user's VBE
  and costs a COM round-trip per component on a large project.
- Step 0b uses `_get_vb_project`, NOT `VBE.ActiveVBProject` — same wrong-project
  reasoning as `access_vbe_check_syntax`.
- **The trigger is a chain**, not an if/else: VBE menu item (unless Access
  reports `Enabled=False`) → `RunCommand(AC_CMD_COMPILE)`. The menu item is
  first because `RunCommand(126)` silently skips form/report modules. Keep the
  menu item first if you touch this.
- **`if dismissed: break` inside the chain is load-bearing.** A real compile
  error surfaces as a dialog; the watchdog dismisses it and `Execute()` can then
  raise as a side effect. Without the break we would re-fire the compile and,
  worse, report "command unavailable" for what is a genuine code error — the
  exact false alarm inverted. A trigger exception with a dismissed dialog must
  fall through to step 4.
- All triggers failed + no dialog ⇒ "could not run the compile command" (NOT a
  code error). `IsCompiled=False` with no dialog and no block mismatches ⇒ the
  message states BOTH possible causes and says to cross-check manually. Both
  carry `trigger` + `code_pane` diagnostics. Do NOT restore the old
  unconditional "missing reference…" wording — it was a repeated field false
  alarm.

`_save_all_modules` (`code.py`) runs `RunCommand(280)` under a dialog watchdog:
when Access is not foreground (VBE has focus, e.g. right after a compile
activated a code pane) "not available now" arrives as a MODAL dialog instead of
a trappable 2046, wedging `ac_delete_object`. `ran_ok and not dismissed` is the
success test — a dismissed dialog means the command never ran, so the per-module
`DoCmd.Save` loop must still execute. The watchdog waits `_SAVE_MODULES_GRACE`
(1.5 s) before dismissing anything: a working RunCommand returns in
milliseconds, so this keeps the attached-instance policy intact (a modal with
nothing blocking belongs to the interactive user). The `join()` before reading
`dismissed` is also load-bearing — dismissing the dialog is what unblocks the
COM call, so the main thread can otherwise win the race.

## SHIFT AutoExec bypass opt-out (v0.7.53, PR #34)

`security.shift_bypass_enabled()` gates the synthetic SHIFT hold behind
**`MCP_ACCESS_SHIFT_BYPASS`**. `keybd_event` is a global key-down: it shifts
whatever the human types anywhere on the machine while held (~0.3 s per open,
~3 s per decompile).

**Opposite polarity to the code-exec gate, on purpose.** That one is security
and fails CLOSED (only an explicit truthy value opens it). This one is
ergonomics and fails OPEN (only an explicit `0/false/no/off` disables it), so a
typo can't quietly let AutoExec run on someone's database. Hence no `ALLOW_`
prefix (implies default-off) and no `DISABLE_` (double negative). Do NOT flip
the default: turning the bypass off for everyone would change behaviour with no
error message pointing at the cause.

**`core._press_shift_bypass()` is the only place in the package that presses
SHIFT** (`core._release_shift()`, the pre-existing `atexit` safety net, releases
it — idempotent, so callers just guard with their own `shift_held` flag). The
three call sites (`_switch`, `_Session._decompile`,
`maintenance.ac_decompile_compact`) each carried their own copy before v0.7.53,
which is exactly how a gate gets half-applied. `test_shift_bypass_gate.py` fails
if `keybd_event` reappears outside `core.py`. (`ui.py` legitimately synthesises
keys for `access_ui_type` and is excluded.)

## Code-execution gate (v0.7.51)

`mcp_access/security.py` is the single source of truth for the opt-in gate that
closes the three code-execution sinks (`access_run_vba`, `access_eval_vba`,
`access_run_macro` — the last one because a macro can carry a `RunCode` action).
Controlled by the env var **`MCP_ACCESS_ALLOW_CODE_EXEC`** (truthy = `1/true/yes/on`,
case-insensitive, `.strip()`), read on **every call** (not at import) so tests can
`monkeypatch` it and import order is irrelevant.

Two layers:
1. **Advertise** — `server.list_tools()` omits the 3 tools when the gate is closed
   (hygiene; the model never sees them).
2. **Dispatch** — `dispatcher.call_tool_sync` rejects a gated tool *first thing in
   the `try`*, before any `_Session`/COM, returning `code_exec_denied_message`.
   This is the REAL barrier: a client can call the name directly without seeing it
   advertised.

Rationale: `confirm_*` flags stop model mistakes, not injection (injected text can
ask for `confirm=true`). Only an out-of-band env var the model can't set defends
against prompt injection. See `SECURITY.md`. Tool count stays 67 (nothing removed,
3 gated). `_TOOL_SCHEMA_INDEX` is still built from the full `TOOLS` so
`coerce_arguments` works for gated tools too — do NOT filter it.

**Enable-on-request flow** (documented, NOT a tool): when the *user explicitly asks*
to enable VBA exec, warn what it grants (arbitrary OS commands via `Shell`, treat DB
as untrusted, trusted DBs only), edit the `env` block of this server in the MCP
client config (e.g. `.mcp.json`) to add `"MCP_ACCESS_ALLOW_CODE_EXEC": "1"`, and tell
the user to **restart** the server (the var is read at startup).

**Critical DO NOTs for the gate:**
- Do NOT remove the dispatch-time enforcement in `call_tool_sync`. The advertise
  layer alone is bypassable (a client can call an unadvertised name directly).
- Do NOT ever add an MCP tool that flips the gate on at runtime. An injection would
  call it. Enabling MUST stay out of band (edit config + restart).

## Common Gotchas

- VBE line numbers are **1-based**
- `ProcCountLines` can inflate the last proc's count past end of module — always clamp with `min(count, total - start + 1)`
- Access must be `Visible = True` for VBE COM access to work
- *"Trust access to the VBA project object model"* must be enabled in Access Trust Center

### CreateForm via COM shows "Save As" MsgBox
- **Do NOT** call `CreateForm()` directly followed by `_save_and_close()`.
- Use `access_create_form` tool: `CreateForm()` -> `DoCmd.Save(acForm, autoName)` -> `DoCmd.Close(acForm, autoName, acSaveNo)` -> `DoCmd.Rename(desired, acForm, autoName)`.
- Pass `record_source` to bind the form to a table/query and `default_view` (0=Single, 1=Continuous, 2=Datasheet, ...) to set the initial view — both are applied on the live `CreateForm()` object before `DoCmd.Save`. Without `record_source`, every bound `ControlSource` on the form will render as `#Name?`.
- Alternative: export an existing form with `ac_get_code`, modify the text, reimport with `ac_set_code`.

### AutoExec / startup forms block OpenCurrentDatabase
- `_switch()` holds Shift key during `OpenCurrentDatabase` (standard Access bypass). Auto-opened forms are closed as safety net.
- `AutomationSecurity = 3` is set as defence-in-depth but does NOT suppress AutoExec macro objects (tested).
- `_Session.reopen(path)` always applies SHIFT bypass.

### Exclusive opens are a request, not a guarantee (v0.7.55)
`MCP_ACCESS_EXCLUSIVE` (off by default, fails closed) passes `Exclusive:=True`
as the 2nd positional arg of `OpenCurrentDatabase(filepath, Exclusive,
bstrPassword)`. Access reports **none** of the failure modes, so `_switch()`
verifies instead of trusting (measured on Access 2016):
- file free -> exclusive, and **no lock file is written**;
- file already open -> opened **shared**, no exception, `CurrentDb` valid, our
  entry appended to the lock file;
- file held exclusively by another -> no exception, session left with **no
  database** (this is what reaches the existing `CurrentDb() is None` check —
  in exclusive mode it must NOT blame AutoExec).

Hence `_lock_file_in_use()` before the open (refuse, session untouched) and
again after (downgrade -> `CloseCurrentDatabase`). Existence of `.laccdb` alone
proves nothing: an orphan from a crashed Access stays on disk and Access opens
exclusively over it, so the file is probed with `CreateFileW(dwShareMode=0)` —
`ERROR_SHARING_VIOLATION` means a live session. Holder names come from its
64-byte entries (32 computer + 32 security name).

### Shared opens are reported, not refused (v0.7.56)
`MCP_ACCESS_EXCLUSIVE` is off by default, so the common case is still a shared
open onto a database somebody else may have open. `_switch` reads the lock file
BEFORE closing/opening anything (our own entry would otherwise be in it) and
parks `_Session._shared_open_warning = (monotonic, msg)`; `server.call_tool`
appends it to the result of that call, same timestamp gate as
`_last_dismissed`. Skipped when `_already_open(path)` — attaching to the user's
Access must not report the user to themselves.

It must stay a warning. Turning it into a refusal is what the env var is for,
and a default-on refusal would break every existing shared workflow.

`_db_file_in_use()` (probe the `.accdb` with `dwShareMode=0`) exists because
`_lock_file_in_use` answers only for SHARED occupants: a database held
**exclusively** by another process has NO lock file at all. In shared mode that
open leaves the session with no database and no exception, and the post-open
failure used to be reported as an AutoExec/startup-form problem — the same
wrong diagnosis the exclusive path already avoids. The file probe runs first,
the AutoExec message is the fallback. Only valid after our own session has been
torn down (otherwise we are an occupant ourselves).

### Linked tables and dbAttachSavePWD
- `dbAttachSavePWD` = **131072** (0x20000), NOT 65536.
- Setting `TableDef.Attributes` from Python COM before Append does not work reliably. Use `DoCmd.TransferDatabase(acLink, ..., StoreLogin:=True)` instead.
- **`ac_list_linked_tables` filtering (v0.7.48)**: `name` (single exact/case-insensitive match), `names_only` (drop `connect_string` — a full dump of hundreds of links overflows the per-result token cap), `mask_password` (mask `PWD=` via `_mask_pwd`). All default to the pre-v0.7.48 full output so existing callers are unaffected.
- **`ac_relink_table(refresh=True)` (v0.7.48)**: `_refresh_links` calls DAO `RefreshLink()` using the table's own connect string (no delete/TransferDatabase, password never touched) — for "the server schema changed, re-read it". `new_connect` is `Optional` and only required when `refresh=False`.

### Scoped embedded lint (v0.7.48)
`_attach_lint`/`lint_compact` take `focus_controls`: the design-mutation tools (`ac_create_control`, `ac_set_control_props`, `ac_set_multiple_controls`) pass the controls they just touched so `lint.violations` isn't buried by pre-existing issues on a big inherited form. `_violation_controls` matches a violation's own `control` **plus** an overlap pair's `measured.a`/`measured.b`. The `error`/`warning`/`info` counts stay whole-form (the model still sees there are other issues). `full_lint=true` bypasses the filter.

### ac_execute_sql / ac_execute_batch
- Both use try/except retry with `dbSeeChanges` for ODBC linked tables with IDENTITY columns.
- DELETE/DROP/TRUNCATE/ALTER require `confirm_destructive=true`.

### MCP schema type coercion
- Some MCP clients serialize ALL arguments as strings. `_fixup_schema()` widens schemas to accept both native types and strings. `_coerce_arguments()` converts back before dispatch.
- Do NOT change schemas back to strict `"type": "integer"` — clients can't be trusted.

### Jet SQL DDL Gotchas
- `YESNO` is not valid in DDL — use `BIT`, or better use `access_create_table`
- `DEFAULT` is not supported in `CREATE TABLE` — use `access_set_field_property` or `access_create_table`
- Multiple JOINs need nested parens: `FROM (A INNER JOIN B ON ...) INNER JOIN C ON ...`
- `AUTOINCREMENT` works as a type in DDL
- Use `SHORT` not `SMALLINT`, `LONG` not `INT`
- Prefer `access_create_table` over `CREATE TABLE` for full type + default + description support

### VBA Language Gotchas
- **`Private Type` without `End Type`**: All code after the block remains "inside" the type. If you get "Statement invalid inside Type block" on a correct-looking line, check for missing `End Type` above.
- **`SysCmd acSysCmdInitMeter`/`acSysCmdUpdateMeter`**: Cause intermittent "Illegal function call". Use `SysCmd acSysCmdSetStatus, "..."` instead.

### ActiveX controls
- Type 119 (`acCustomControl`): pass `class_name` with ProgID (e.g. `Shell.Explorer.2`).
- Type 128 (`acWebBrowser`): native control, no ActiveX needed.

## Critical DO NOTs

- **Do NOT remove the `DispatchEx` fallback** in `_Session._launch()`. `_launch()` tries `GetActiveObject("Access.Application")` first to attach to a user's running Access (avoids spawning a second process); on failure it falls back to `DispatchEx`, which is required after `/decompile` kills to bypass stale ROT entries. Do NOT swap `DispatchEx` for `Dispatch` in the fallback — `Dispatch` latches onto the stale ROT entry. Under `MCP_ACCESS_EXCLUSIVE` the attach is skipped entirely (a running instance holds the file shared) and `DispatchEx` is the only path — do NOT "restore" attaching there.
- **Do NOT call `cls._app.Quit()` unconditionally in `_decompile()` / `ac_decompile_compact()`**. Check `_Session._attached` first — when True we attached to the user's Access and must only `CloseCurrentDatabase()`, keeping the instance alive. Only when `_attached=False` (we spawned via `DispatchEx`) is `Quit(1)` safe. Same applies to the `atexit` handler `_Session.quit()`.
- **Do NOT use `EnsureDispatch`** — it changes binding for all 61 tools and adds `gen_py` cache dependency.
- **Do NOT run `OpenCurrentDatabase` in a separate thread** — COM STA objects can only be used from the thread that created them.
- **Do NOT call `CreateForm()` directly** — use `access_create_form` tool to avoid the "Save As" MsgBox.
- **Do NOT change schemas to strict `"type": "integer"`** — MCP clients can't be trusted to send correct types.
- **Do NOT auto-decompile on DB open** — only on first compile. Auto-decompile on open caused SHIFT key stuck issues and process accumulation on MCP reconnect.

## MCP SDK Patch (local to this machine)

The MCP Python SDK (`mcp/shared/session.py`) swallows all exceptions with a generic `-32602` error. A local patch at `c:\program files\python310\lib\site-packages\mcp\shared\session.py` adds full traceback to `ErrorData.message` and `ErrorData.data`. Re-apply after `pip install --upgrade mcp`.
