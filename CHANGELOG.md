# Changelog

## 0.7.59 — 2026-09-06

**The server knew the caller had made a mistake and said nothing.** Three of the
four items in `NEXT-STEPS.md`, all from real production use. None of them change
what the server *does* — they change what it *tells you* when you are about to
have a bad day.

### The tab control was invisible

`_parse_controls` only recognises a `Begin <Type>` block whose type is in
`CTRL_TYPE`, and **123 was never there**. A form's tab control therefore had no
geometry anywhere in the package: it was absent from `access_list_controls`, and
its pages did not say which tab they belonged to.

- `CTRL_TYPE` gains `123: "Tab"` (the export token is `Tab`, not `TabCtl`), and
  `CONTAINER_TYPES` gains it too. **The second half is load-bearing**: a
  recognised *non*-container makes the parser skip to the end of its block, which
  would have swallowed every Page and every control on them — the exact
  regression fixed in v0.7.34. `tests/test_parse_controls.py` fails without it.
- `CTRL_TYPE_BY_NAME` is derived from `CTRL_TYPE`, so `"tab"` now resolves as a
  control type name as well.
- The lint exempts `Tab` from the geometry rules (`_LAYOUT_EXEMPT_TYPES`) and
  from contrast: a tab control contains its own pages and everything on them, so
  it overlaps all of it by construction.
- Visible change: `access_list_controls` now returns the tab control, and each
  Page carries `parent` with its tab control's name.

### `access_create_control` warns when a control lands on a tab by accident

`parent` (CreateControl's 4th positional argument) is optional and easy to
forget. Without it the control is created **on the section**, not on the tab
page, even when its coordinates put it right on top of one. The call succeeded,
because it *was* a success — just not the one that was meant, and nothing said
so until somebody opened the form and found the control on every tab.

The result now carries a `warning` when the new control's rectangle falls fully
inside a tab control in the same section and `parent` was empty. It names the
tab, lists the pages available on it, and spells out the actual trap: **the page
name is what `parent` wants, not the tab control's name**.

Nothing is ever re-parented automatically — a control placed over a tab on
purpose is legitimate, and silently moving things is worse than the original
problem. Shares the `skip_lint` switch, and costs nothing on top of the lint:
`_build_model` reads the same cached export.

### `access_get_control` no longer reports `control_type: -1`

Access omits `ControlType =` when it equals the default, which is the normal
case in modern exports, so the parser fell back to `-1` while `type_name` was
correct. Round-tripping a control definition therefore produced a number that
could not be fed back to `access_create_control`. The type is now resolved from
the `Begin <Type>` token via `CTRL_TYPE_BY_NAME`.

Note: for `WebBrowser` and the two `Navigation*` controls the SaveAsText number
and the AcControlType number `CreateControl` accepts differ. The resolved value
is the latter — the one that makes the round-trip work.

### Fixed: the string `"-1"` placed controls at 1 twip

`coerce_prop` maps `"-1"` to `True` (in Access -1 *is* True, and boolean
properties depend on that), and `int(True)` is 1. Since some MCP clients
serialise every argument as a string, asking for the automatic position with
`"-1"` put the control at coordinate 1 instead of letting Access decide.
`ac_create_control` now converts its four geometry arguments with a dedicated
numeric helper (`_coord`), and so does the `snap_to_grid` loop in
`ac_set_control_props`, which had the same bug: it turned an automatic `"-1"`
into a snapped 0. `coerce_prop` itself is unchanged — removing `"-1"` there would
break every boolean property that relies on it.

### Documentation, in `access_tips`

Two things the server already knew and never said:

- **`vbe`** — why writing code changes modules nobody touched. The VBA editor
  keeps one canonical spelling per identifier project-wide, so declaring a name
  that already exists with different casing rewrites those modules too. Harmless
  at runtime, loud in version control. Deliberately documented rather than
  detected: fingerprinting every module on every write would cost hundreds of COM
  round-trips to report damage that is already done and already visible in the
  diff.
- **`design_vbe`** — injecting code rewrites the form-level
  `Left/Top/Right/Bottom` and `Checksum` in the next text export. That is the
  design *window*, not the layout. If a diff touches only those and no control's
  own coordinates, nothing moved.

Also fixed in the `controls` topic: it handed out numbers belonging to other
controls (it called 106 a ComboBox; 106 is CheckBox). The `access_tips` schema
listed the topics by hand and had drifted twice. Both are now guarded by tests.


## 0.7.58 — 2026-09-03

**The server hijacked whichever Access instance answered the phone.**
Reported as issue #38 by [@Access-Abraxas](https://github.com/Access-Abraxas),
who drives this server from VBA running inside their own database:
`access_create_database` attached to that instance and closed the database out
from under it, stopping the code that made the call.

### Fixed

- **`_Session._launch()` no longer attaches to an Access instance that has a
  different database open.** It now takes the database we are about to open and
  keeps the candidate only when that instance is idle (`CurrentDb()` is `None`)
  or already holds our target (`os.path.normcase` comparison, the same rule as
  `_already_open`). Otherwise the reference is dropped and we spawn our own
  process through the existing `DispatchEx` fallback. Attach-first is kept for
  the case it was written for — the user has Access open on the database we are
  going to work on — and `_launch(None)` still attaches to anything, so callers
  with no known destination are unaffected.
- **`access_create_database` passes its target path.** The file does not exist
  yet, so a running instance can never match it: with the user's Access busy we
  always get a second window, and the `CloseCurrentDatabase()` that caused the
  report never reaches their session. No other change was needed there — the
  fresh instance has no database open, so that branch is simply skipped.

### Known limitation

`GetActiveObject` returns a single entry from the running-object table. With
several Access windows open we can inspect only one of them, so we may spawn a
process even though another window already had the target database open. Worse
than ideal, better than closing somebody's work.

### Not adopted

The report also suggested one Access instance per database. `_Session` is a
singleton by design — one COM session, one dedicated STA thread, one open
database — and per-database instances would be a different server. Marking who
opened a database so `_switch` leaves a foreign one alone was rejected too: the
database would stay open in the user's instance and we would reopen it shared
later, which brings the v0.7.56 lock warning and design-lock conflicts to the
normal workflow.

### Tests

- `tests/test_attach_policy.py` — pure Python, no COM: `win32com.client` is
  imported inside `_launch`, so a stub in `sys.modules` covers all five cases
  (foreign database, target database, idle instance, `target_path=None`, and
  `MCP_ACCESS_EXCLUSIVE`, which still never looks at the candidate).

## 0.7.57 — 2026-09-02

**The dependency floor had no ceiling, and MCP SDK v2 walked through it.**
Reported as issue #37 by [@JMarchesoniAF](https://github.com/JMarchesoniAF).
`mcp>=1.0.0` meant a clean `pip install` / `uvx` resolved to the v2 SDK
(2.0.0 landed 2026-07-28, 2.1.1 on 2026-08-25), where the server does not even
finish importing.

### Fixed

- **Pinned `mcp>=1.0.0,<2`.** Verified against a clean 2.1.1 install: the
  import dies at `tools.py` with `AttributeError: 'Tool' object has no
  attribute 'inputSchema'` (v2 renamed the field to `input_schema`), and had it
  got past that, `@server.list_tools()` and the three other decorators no
  longer exist — the low-level `Server` takes `on_list_tools` / `on_call_tool`
  / `on_list_prompts` / `on_get_prompt` constructor callbacks with
  `(context, params)` signatures instead. `mcp.shared.session` is gone too, so
  the local SDK patch documented in `CLAUDE.md` no longer applies there.

This release only restores installability. The v2 migration itself is tracked
in issue #37 and is deliberately not bundled here: it changes the handler
interface, the prompt/tool result types and the direct-callback tests, and it
deserves its own release rather than riding along in a one-line pin.

## 0.7.56 — 2026-08-30

**Lock conflicts are reported in shared mode too.** Follow-up to
[@GPGeorge](https://github.com/GPGeorge)'s issue #36: `MCP_ACCESS_EXCLUSIVE`
turns "somebody else has this open" into a refusal, but it is off by default,
and the person he is trying to protect — a newcomer pointed at a live
production front-end — is the one least likely to switch it on. So the default
path stops being silent.

### Added

- **Shared-open advisory.** When the session opens a database another Access
  session already has open, the result of that tool call now carries a warning
  naming the sessions holding it, and saying what is at risk: design changes
  (table/form/report design, Data Macros, anything routing through Design view)
  can be dropped *without an error* while another session holds the object — a
  run reports success and changed nothing. Read-only work is unaffected, and
  nothing is refused; refusing is what `MCP_ACCESS_EXCLUSIVE` is for.
  Surfaced the same way as the auto-dismissed-dialog note (timestamp-gated, so
  it appears once, on the call that actually opened the database).
- The advisory is skipped when the live Access instance already has that file
  open — attaching to the user's own Access must not report the user to
  themselves as a foreign occupant.

### Fixed

- **A database another process holds exclusively is no longer diagnosed as a
  broken AutoExec.** In shared mode that open is refused by Access with *no
  exception and no database*, and the resulting message blamed the startup
  form — sending the user through startup code for what is a lock problem. The
  lock file cannot answer this (an exclusive holder writes none), so the
  `.accdb` itself is probed with `dwShareMode = 0` after the failed open. This
  is the same silent-failure class the exclusive switch already handles, one
  level down; it was found by running the v0.7.55 verification against two live
  Access processes rather than reasoning about it.

Verified live on Access 2016, three scenarios: occupant holding the file
exclusively (lock conflict reported, no AutoExec blame), occupant holding it
shared (open succeeds, advisory set, holder named), and nobody holding it (open
succeeds, no advisory). A stale lock file from a crashed Access still counts as
free.

## 0.7.55 — 2026-08-30

**Opt-in exclusive opens.** Requested by
[@GPGeorge](https://github.com/GPGeorge) (issue #36), whose analysis of *where*
it belongs is the design used here.

### Added

- **`MCP_ACCESS_EXCLUSIVE` — open the session's database exclusively.** Shared
  opens cannot take a **design lock**, so attaching Data Macros through
  `SaveAsText`/`LoadFromText`, or anything routing through
  `DoCmd.OpenTable acViewDesign`, is refused for any table another Access
  session has open. The refusal is per table and easy to miss: the run finishes,
  reports success, and changed nothing. This turns that into one visible failure
  at open time.

  There is no explicit open tool — the open is implicit in
  `_Session.connect()` — so a per-call argument would have to be threaded
  through every tool taking `db_path` and kept consistent across all of them. A
  session either wants exclusive or it doesn't, so it is one server-level
  switch, read per call like the other two:

  ```json
  "env": { "MCP_ACCESS_EXCLUSIVE": "1" }
  ```

  **Default off, fails closed** (like `MCP_ACCESS_ALLOW_CODE_EXEC`, unlike
  `MCP_ACCESS_SHIFT_BYPASS`). Not because exclusivity endangers the file, but
  because the COM session holds the database open between tool calls: while the
  server is connected nobody else can get in. On a shared front-end that locks
  a whole workgroup out for as long as the session lives, so a typo must not be
  able to switch it on. `access_close` releases without stopping the server.

- **The switch verifies what it asked for, because Access doesn't.** Passing
  `Exclusive:=True` is a *request*, and on its own it would have delivered a new
  silent failure in place of the old one. Measured against Access 2016:

  | Situation | What Access does |
  |---|---|
  | File free | Opens exclusive, writes **no** lock file |
  | File already open by someone else | Opens it **shared**, no exception, `CurrentDb` valid, adds our own entry to the lock file |
  | File held exclusively by someone else | No exception either — leaves the session with **no database at all** |

  So the open is checked rather than assumed. Before opening, a live lock file
  means the database is busy and the call is refused without touching the
  session. After opening, a lock file that exists *and is held* means Access
  downgraded us to shared, and the database is closed again rather than left to
  run design work under a false assumption. The `CurrentDb() is None` case now
  reports the lock, not the AutoExec/startup-form diagnosis that used to be the
  only explanation offered there.

  Occupants are named in the error, read from the lock file's 64-byte entries
  (32 bytes computer name, 32 bytes security name — [documented by
  Microsoft](https://learn.microsoft.com/troubleshoot/microsoft-365-apps/access/lock-files-introduction)).
  Mere *existence* of that file is not evidence: an Access that died without
  closing leaves an orphan, and Access opens exclusively right over one and
  leaves it untouched — so the file is probed with `dwShareMode = 0`, which
  separates a live session from a stale file.

- **With the switch on, the server no longer attaches to a running Access
  instance** (`_Session._launch`). That instance holds its database shared, and
  `connect()` does not re-open a path already recorded in `_db_open` — so
  attaching would have left the session shared while reporting itself
  exclusive, the same silent-success problem one level up. `DispatchEx`
  (unchanged, and still the `/decompile` fallback) spawns a dedicated instance
  instead. Unaffected with the switch off, which is the default.

- `access_create_database` honours the switch on its post-create reopen, so a
  session doesn't silently fall back to shared after creating a database.

## 0.7.54 — 2026-08-30

### Added

- **`access_screenshot` now accepts a `view` parameter** (`normal`, `design`,
  `preview`, `datasheet`). Default is `normal` (unchanged behaviour). `design`
  opens the form/report in Design View, which never fires `Form_Open` or
  `Form_Load` events — so it cannot block on login dialogs, slow recordsets, or
  Modal/PopUp forms. The modal guard and ESC-timeout watchdog are skipped
  entirely in design mode (they are not needed). The error messages for modal
  and timeout failures now suggest `view='design'` as an alternative.

## 0.7.53 — 2026-07-30

**Two community bug reports, both reproduced on real databases before being
filed.** Thanks to [@CaptainStormfield](https://github.com/CaptainStormfield)
(PR #35) and [@CustomDataNZ](https://github.com/CustomDataNZ) (PR #34) — the
analysis, the reproductions and the test conventions here are theirs; the
implementations below are reworked versions of their patches.

### Fixed

- **`access_compile_vba` no longer misreports a failed compile *trigger* as a
  compile *error* in your code** (reported and diagnosed by
  [@CaptainStormfield](https://github.com/CaptainStormfield), PR #35). The tool
  deliberately dirties the project (step 0b) so `Application.IsCompiled` can
  serve as the success signal, then triggers compilation by `Execute()`-ing the
  VBE *Debug > Compile* menu item. That item acts on the **active** project and
  is only reliably enabled when one of its code panes has focus — so with no
  pane active the trigger could raise `DISP_E_EXCEPTION` (-2147352567), silently
  no-op, or compile the wrong project entirely (after a decompile/compact the
  active project is typically the `acwzmain` wizard library). Every one of those
  paths left the deliberately-dirtied `IsCompiled=False` standing and produced
  the false *"VBA project is NOT compiled … missing reference, undeclared
  variable, or type mismatch"* while a manual Debug > Compile of the same
  project succeeded. Now:
  - a code pane of the **current database's** project is made active before the
    trigger (`_ensure_code_pane`), which both enables the menu item and points
    it at the right project. Standard modules are preferred (no Design-view side
    effects), and a pane of our project that is *already* active short-circuits
    the whole thing — a repeat compile no longer opens code windows in your VBE;
  - the step-0b dirty-marking resolves the project via `_get_vb_project` instead
    of `VBE.ActiveVBProject`, so it can no longer dirty `acwzmain` while
    `IsCompiled` is read from your project;
  - the trigger is now a **chain**: the VBE menu item (which compiles form and
    report modules too) unless Access reports it disabled, then `RunCommand` as
    a fallback. A menu `Execute()` that fails no longer ends the attempt;
  - failing to *run* the command is reported as exactly that — *"could not run
    the compile command … NOT that the VBA code has errors"* — and the residual
    `IsCompiled=False`-with-no-dialog case now states both possible causes and
    tells you to cross-check with Debug > Compile. Both carry `trigger` and
    `code_pane` diagnostics;
  - an exception raised *after* the compile-error dialog was auto-dismissed is
    no longer mistaken for an unavailable command: the dialog text wins, so a
    real compile error is still reported as one.
- **`access_delete_object` no longer wedges behind an invisible modal**
  (reported by [@CaptainStormfield](https://github.com/CaptainStormfield), PR
  #35). `RunCommand(280)` (`acCmdSaveAllModules`) in `_save_all_modules` does
  not always report "not available now" as a trappable 2046 — when Access is not
  the foreground application (typically because the VBE has focus, e.g. straight
  after `access_compile_vba` activated a code pane) it surfaces as a **modal
  dialog** that blocks the COM call until a human clicks OK. Seen in the field
  as a *"command 'SaveAllModules' isn't available"* box during a routine module
  delete. It now runs under a dialog watchdog, and a dismissed dialog means the
  command did not run, so the per-module `DoCmd.Save` fallback executes instead
  of trusting a save that never happened. The watchdog waits out a 1.5 s grace
  period first: a working `RunCommand` returns in milliseconds, so a modal that
  appears with nothing blocking still belongs to the interactive user and is
  never clicked away.

### Added

- **`MCP_ACCESS_SHIFT_BYPASS` — opt out of the global SHIFT AutoExec bypass**
  (reported and designed by [@CustomDataNZ](https://github.com/CustomDataNZ),
  PR #34). `OpenCurrentDatabase` and `MSACCESS /decompile` hold SHIFT to skip a
  target database's AutoExec macro and startup form, but `keybd_event(VK_SHIFT,
  …)` is a **global** OS-level key-down: it is not scoped to Access, so every
  keystroke the human types anywhere on the machine during the hold arrives
  shifted — ~0.3 s on every database switch, ~3 s per decompile. On a box where
  someone is working while the server runs, that is a repeated and fairly
  baffling nuisance with nothing on screen to explain it.

  **Nothing changes by default.** With the variable unset the bypass behaves
  exactly as before; set it to `0` / `false` / `no` / `off` to turn it off. It
  is the mirror image of `MCP_ACCESS_ALLOW_CODE_EXEC`: that one is a security
  gate and fails **closed**, this one is ergonomics and fails **open**, so a
  typo or an empty value keeps the bypass rather than quietly letting AutoExec
  run on someone's database. Turning it off is the right move for databases that
  guard their own startup with `If Not Application.UserControl Then Exit
  Function`, which is the cleaner fix and belongs in the database rather than in
  a global input hack. `access_tips('vba')` and the README now say so.

### Changed

- **The SHIFT press/release sequence is no longer copy-pasted into three call
  sites.** `core._press_shift_bypass()` is now the only place in the package
  that synthesises a SHIFT key-down (`core._release_shift()`, which already
  existed as the `atexit` safety net, handles the release). `_switch`,
  `_Session._decompile` and `ac_decompile_compact` all route through it, which
  is what makes the opt-out impossible to half-apply — a structural test fails
  if a second synthesis site ever reappears.
- **The changelog now lives only in `CHANGELOG.md`.** The README carried a
  duplicate copy of the entire release history (700+ lines) that had to be
  updated in parallel and drifted every release; it now links here.

## 0.7.52 — 2026-07-21

**`access_vbe_patch_proc` stops being the one write tool without a safety net.**
Six field requests from [@TvanStiphout-Home](https://github.com/TvanStiphout-Home)
(Tom van Stiphout), every one of them tested against a real database before being
filed — **thank you Tom**. Patches are now **all-or-nothing by default**, anchors
match regardless of case, ambiguous anchors can be rejected outright, the
`(Declarations)` section is addressable, and there is finally a way to check that
what you wrote is structurally sound *without* the destructive setup
`access_compile_vba` performs. One new tool (**68 total**).

### Added

- **`access_vbe_check_syntax`** — static structural check of the VBA project that
  is **already open**: no decompile, no `RunCommand`, no Design view, no second
  Access instance, nothing discarded. `access_compile_vba` was unusable as a
  post-edit check because its step 0 shells out to `MSACCESS.EXE /decompile` and
  then either quits with `acQuitSaveNone` or closes the database — **unsaved VBA
  is lost**. The new tool catches unbalanced `If`/`For`/`Do`/`While`/`Select`/
  `With`/`Type`/`Enum` blocks, code sitting outside a procedure and misplaced
  `Option` statements. Scope it to one object with `object_type`/`object_name` or
  let it walk every standard module and form/report code-behind.
  It is **not** a compiler and says so in its own `note` field: it does not
  resolve identifiers, types or references, so `ok=true` is not proof the project
  compiles. It also never reports a clean zero for something it could not read —
  per-module failures land in `skipped` and force `ok=false`.
- **`atomic` on `access_vbe_patch_proc` (DEFAULT `true`)** — if any patch fails to
  match, **nothing is written** and the module is left byte-for-byte identical.
  Previously a batch was applied best-effort, so one stale anchor left the
  procedure in a half-edited state nobody wrote and nobody reviewed. The abort
  message lists *every* failure at once and tells the caller to re-send the
  **entire** batch, because the patches that did match were discarded too.
  `atomic=false` restores the old behaviour.
- **`require_unique` on `access_vbe_patch_proc`** — refuse to patch when the
  anchor matches more than once, reporting the count and the absolute module line
  numbers of every hit. Default `false`: replacing the first of several
  occurrences is a legitimate use and was already warned about.
- **`(Declarations)` as a target** — `proc_name='(Declarations)'` now resolves to
  the module declarations section in `access_vbe_patch_proc` and
  `access_vbe_get_proc`, and `access_vbe_module_info` gained an additive
  `declarations: {start_line, count}` key so the boundary no longer has to be
  guessed from the first procedure. `Option` lines are never stripped there.
  `access_vbe_replace_proc` **refuses** the token on purpose: `new_code=''` would
  wipe `Option Explicit` and every module-level `Const` in one unconfirmed call.

### Changed

- **Anchors are matched case-insensitively by default** (`match_case`, default
  `false`) in `access_vbe_patch_proc`. VBA is a case-insensitive language whose
  editor rewrites identifier casing on its own, so a `find` that differed only in
  case used to fail for no useful reason — the same reasoning `access_vbe_find`
  already followed. **This changes behaviour**: anchors that find nothing today
  will start finding something. `atomic=true` and the ambiguity warning bound the
  risk, and the result echoes the text as it is actually stored so a caller can
  correct their copy. Set `match_case=true` to demand an exact-case match.
  Matching runs as a fixed ladder — literal then whitespace-normalized, **all
  case-sensitive tiers before any case-insensitive one** — so every call that
  succeeds today still lands exactly where it did before.
- **`total_lines` is now `cm.CountOfLines`** in `access_vbe_module_info` and the
  bounds of `access_vbe_get_lines`. VBE emits no trailing terminator, so
  `splitlines()` silently dropped a final blank line and these tools reported one
  line fewer than `access_vbe_patch_proc` did for the same module. A trailing
  blank line is a real, addressable line in the editor, so `CountOfLines` wins.
  Side effects worth knowing about: that blank line is now readable via
  `access_vbe_get_lines` (it used to be rejected as out of range), and the `count`
  reported for the **last** procedure of a blank-terminated module can be one
  higher than before — that is the same off-by-one, not a regression.

### Fixed

- **`access_screenshot` hung for 30+ minutes on Modal/PopUp forms**
  (`mcp_access/ui.py` `ac_screenshot`): `DoCmd.OpenForm` on a Modal/PopUp form
  enters a blocking dialog loop the ESC watchdog could not break (observed: a
  ~30-minute hang). The form's Modal/PopUp state is now checked up front via a
  non-blocking design-view open, and auto-opening is refused with a message
  telling the caller to open the form manually and call `access_screenshot`
  WITHOUT `object_name`. Best-effort: if the property check itself fails, it falls
  through to the previous open + ESC-watchdog path.
- **`access_vbe_patch_proc` reported an unclamped `new_count`** (`mcp_access/vbe.py`):
  the closing message called `ProcCountLines` raw, unlike the `count` computed
  before the write and unlike `access_vbe_replace_proc`, so the last procedure of
  a module could be reported longer than the module itself.
- **`_check_module_health`'s count-sanity check was dead code on the patch path**:
  `access_vbe_patch_proc` called it without `expected_total`, disabling Check 3.
  It now passes `total - count + <lines inserted>`.
- **An unterminated `Type`/`Enum` block passed validation silently**
  (`mcp_access/compile.py` `_check_structure_in_module`): everything below the
  opener is absorbed into the block, so no line inside the scan loop could ever
  flag it. Reported at end of module now — this is the "Statement invalid inside
  Type block" trap, which used to surface only as a confusing compile error on an
  unrelated line.

### Internal

- The patch matching loop is extracted to a pure, COM-free `_apply_patches()`,
  which is what makes `atomic` a structural guarantee rather than a promise: the
  simulation and the commit are the same single pass, so they cannot diverge. A
  pre-pass validating anchors against the original text would have been wrong in
  both directions — an earlier patch can destroy *or* create the anchor a later
  one cites.
- The case-insensitive tiers are skipped, with a note, when lowercasing would
  change the text's length (`'İ'.lower()` returns two characters): offsets
  computed on a lowered copy would no longer address the original and the
  replacement could be spliced into the middle of a line.
- `_verify_module_structure` gained a pure counterpart,
  `_check_structure_in_module`, mirroring `_check_blocks_in_module`.
  `access_compile_vba`'s behaviour is unchanged — its wrappers just delegate.
- New pure test suite `tests/test_vbe_feature_requests.py` (32 tests, no COM),
  plus a live COM verification of all six acceptance scenarios against a copy of
  a real 242-module database.

## 0.7.51 — 2026-07-06

**Behaviour change: VBA/macro execution is now opt-in.** The three tools that run
arbitrary VBA/Shell — `access_run_vba`, `access_eval_vba` and `access_run_macro`
(the last one because a macro can carry a `RunCode` action) — are **disabled by
default**. A fresh PyPI install can no longer be turned into RCE by a single
prompt injection. `confirm_*` flags never defended against injection (injected
text can just ask for `confirm=true`); only an out-of-band environment variable
the model cannot set does. No tool was removed — tool count stays **67**, three
of them gated.

### Added

- **`mcp_access/security.py`** — single source of truth for the gate. Reads
  `MCP_ACCESS_ALLOW_CODE_EXEC` (truthy = `1/true/yes/on`, case-insensitive,
  stripped) on **every call** rather than at import, so tests can monkeypatch it
  and import order is irrelevant. Exposes `code_exec_denied_message` for the
  rejection text.
- **Two enforcement layers.** `server.list_tools()` omits the three gated tools
  when the gate is closed (hygiene — the model never sees them); `dispatcher.
  call_tool_sync` rejects a gated tool as the first statement inside its `try`,
  **before** any `_Session`/COM work. The dispatch layer is the real barrier: a
  client can call an unadvertised name directly. `_TOOL_SCHEMA_INDEX` is still
  built from the full `TOOLS` list so `coerce_arguments` keeps working for gated
  tools.
- **`SECURITY.md`** — the threat model written down: local stdio server, no
  network surface, no login *by design*, prompt injection as the real risk, plus
  how to report an issue. README gains a matching Security section and the
  VBA-execution table is flagged.
- **`tests/test_code_exec_gate.py`** — pure tests (no COM) covering the truthy
  parsing, the advertise filter and the dispatch rejection.

### Fixed

- **Undiscoverable v0.7.48 parameters** (`mcp_access/tools.py`): the handlers
  for `access_list_linked_tables` (`name` / `names_only` / `mask_password`),
  `access_relink_table` (`refresh`) and the control-mutation tools (`full_lint`)
  already accepted these arguments in v0.7.48, but the corresponding schema
  entries were never committed — so MCP clients could not discover them. The
  schema now matches the implementation.

### Changed

- **Enabling stays out of band, on purpose.** To re-enable, add
  `"MCP_ACCESS_ALLOW_CODE_EXEC": "1"` to this server's `env` in the MCP client
  config and **restart** the server. There is deliberately **no** MCP tool that
  flips the gate at runtime — an injection would just call it.

## 0.7.50 — 2026-07-06 — security fix (GHSA-9jp6-hph9-jm5f)

Prompt-injection fix in the `access-workflow` prompt template, reported by
[@nicoPadi1002](https://github.com/nicoPadi1002) (CobaltoSec) via responsible
disclosure — **thank you**. No new tool (still **67**).

### Fixed

- **The `access-workflow` prompt no longer reflects an untrusted `db_path`
  verbatim** (`mcp_access/server.py`, new `_sanitize_db_path`). A `db_path`
  carrying newlines could inject arbitrary text — e.g. a fake `SYSTEM OVERRIDE:`
  block — **ahead of** the prompt's `REQUIRED RULES` section, steering an agent
  towards `access_run_vba` / `access_run_macro`. A real Access file path never
  contains newlines or control characters, so the value is now collapsed to a
  single line at the first control character, capped at `MAX_PATH`, and the
  template wraps it in backticks as plain data. Legitimate paths are byte-for-byte
  unchanged. Regression tests in `tests/test_prompt_injection.py`.

## 0.7.49 — 2026-06-25

Bugfix reported by [@TvanStiphout-Home](https://github.com/TvanStiphout-Home)
(Tom van Stiphout) ([#33](https://github.com/unmateria/MCP-Access/issues/33)) —
**thank you Tom**, once again, for the laser-precise diagnosis and repro steps.
No new tool (still **67**).

### Fixed

- **VBE-write tools no longer pop a modal Access error dialog on every edit**
  (`mcp_access/vbe.py`). `access_vbe_replace_lines`, `access_vbe_replace_proc`,
  `access_vbe_patch_proc` and `access_vbe_append` all call `DoCmd.Save` to
  persist the VBE change into the `.accdb`. When the target module or form was
  already open in the VBE, Access answered with a modal *"Save isn't available
  now"* dialog and waited for a click — one dialog per write call, blocking the
  UI. The `except Exception: pass` swallowed the COM error (so the edit itself
  landed fine), but the dialog watchdog that covers the compile/eval paths was
  absent here, so nothing dismissed it. Fix: a new `_save_vbe_module()` helper
  wraps `DoCmd.Save` in a daemon watchdog thread (0.3 s grace, same pattern as
  `_call_with_dialog_watchdog` in `maintenance.py`); all four call sites now go
  through it and the save stays best-effort.

## 0.7.48 — 2026-06-24

Usability fixes that came out of a real editing session against a database with
hundreds of ODBC-linked tables, plus a scoped embedded lint so a one-control edit
on a big inherited form stops drowning in pre-existing warnings. Every default
preserves the pre-0.7.48 output, so existing callers see no change. No new tool
(still **67**).

### Added

- **`access_list_linked_tables` filtering** (`mcp_access/relations.py`): `name='X'`
  returns a single table (exact, case-insensitive), `names_only=true` gives a
  light listing without `connect_string`, and `mask_password=true` masks `PWD=`
  via `_mask_pwd`. With hundreds of links the full connect-string dump used to
  blow past the per-result token cap and force a `grep`; all three default to the
  previous full output.
- **`access_relink_table refresh=true`** (`_refresh_links`): re-reads a linked
  table's schema through DAO `RefreshLink()` using the table's **own** connect
  string — no delete/`TransferDatabase` round-trip, the password is never touched
  or dumped. This is the common *"I altered the table on the server, re-read it"*
  case. `new_connect` is now optional and only required when `refresh=false`;
  `relink_all=true` refreshes every table sharing the connect string.

### Changed

- **Scoped embedded lint** (`mcp_access/lint.py`, `mcp_access/controls.py`):
  `_attach_lint` / `lint_compact` take a `focus_controls` argument, and
  `access_create_control`, `access_set_control_props` and
  `access_set_multiple_controls` pass the controls they just touched — so
  `lint.violations` shows the change, not every inherited issue on the form.
  `_violation_controls` matches a violation's own `control` **plus** an overlap
  pair's `measured.a`/`measured.b`. The `error`/`warning`/`info` counts stay
  whole-form (the model still sees there are other problems); `full_lint=true`
  bypasses the filter.
- **`access_relink_table` description** now documents the injected
  `LoginTimeout=8`.

## 0.7.47 — 2026-06-22

Bugfix reported by [@jbchea](https://github.com/jbchea)
([#32](https://github.com/unmateria/MCP-Access/issues/32)) — thanks! No new tool
(still **67**).

### Fixed

- **Duplicate `"design"` key in `mcp_access/tips.py`.** The v0.7.45/0.7.46
  design-system work added a *second* `"design"` entry to the `_TIPS` dict,
  silently shadowing the original — Python keeps only the last assignment, so
  `access_tips('design')` returned just the new design-direction guidance and the
  earlier tip (Design view ↔ VBE close-ordering + `SaveAsText` per-object-type
  encoding) became unreachable dead code. The original tip now lives under its
  own key, `access_tips('design_vbe')`, so both are reachable again. No behaviour
  change beyond the tips topic.

## 0.7.46 — 2026-06-19

Real design taste for `access_build_form` — three curated **design directions**
replace the ad-hoc themes, plus the fix for the washed-out two-tone header band
they surfaced. No new tool (still **67**).

### Added

- **Three curated design directions** for `access_build_form` (`theme=`):
  `despacho` (serif Constantia title on warm paper, teal accent band),
  `panel` (Segoe UI Semibold, a white card on a cool canvas, slate band) and
  `archivo` (serif Cambria, warm editorial, spacious, clay band). Each is a
  *coherent bundle* — a typeface with character, an intentional modular type
  scale, a dominant+accent palette with **WCAG-verified contrast**, a spacing
  density and an accent header band — translating real design-system thinking
  into what native Access can render. Each passes the lint clean by construction.
- **Design tokens** in `mcp_access/design_defaults.py`: `type_scale(base, ratio)`
  (a modular scale rounded to whole points), `SPACE` (a closed spacing scale;
  the legacy `MARGIN_X`/`GAP_LABEL`/… are now aliases into it, same values),
  `DENSITY` (compact/comfortable/spacious — margins & gaps only, never control
  sizes) and `DIRECTIONS` (the three palettes, built with `bgr()` straight from
  the hex so they can't drift).
- **`access_tips('design')`** — the design guide the model reads: the three
  directions, the principles (typeface with character, cohesive palette,
  hierarchy, rhythm, containment, active-voice copy) and the **honest ceiling**
  (native Access has no gradients, shadows, rounded corners, blur or animation).
- **Two `info`-only lint rules:** `generic_font` (flags Arial/Roboto/Inter/Times
  New Roman/MS Sans Serif — a closed list) and a `type_hierarchy` extension of
  the `hierarchy` rule (the header title should be larger than the body text).
  Both `info`, so they never change the verdict nor reach the embedded mutation
  lint.

### Fixed

- **Two-tone header band / unpainted canvas.** `build_form`'s `_set_section`
  resolved sections via `Form.Section(index)`, which pywin32 can't late-bind
  (it raises *"member not found"* for every index) — so the call failed
  **silently**: the canvas colour was never painted and the header/footer kept
  Access' oversized default heights, leaving the themed light-blue header
  showing past the accent rectangle (the "two-tone band"). Sections are now
  resolved by their **named** properties (`Detail`/`FormHeader`/`FormFooter`),
  which bind correctly, and a header band is painted on the section BackColor
  (full document-window width) with the Rectangle kept as a fallback.

## 0.7.45 — 2026-06-19

Making the LLM design **much** better Access forms — by moving the layout
arithmetic out of the model and into the MCP (no skill, no hooks required, as
requested). One new tool — tool count goes **66 → 67**.

### Added

- **`access_build_form` — declarative form auto-layout.** Instead of many
  blind `access_create_control` calls with hand-picked twips, describe the form
  declaratively: a `title`, an ordered list of `fields` (string or
  `{field, label, control, name, control_source, row_source, width_units,
  height, props}`), a row of `actions` (footer buttons), `layout`
  (`single`|`two-column`) and `theme` (`light`|`plain`). The tool computes every
  Left/Top/Width/Height from a canonical 60-twip grid, applies a closed
  WCAG-safe palette, binds matching `record_source` columns, assigns a
  per-section tab order, sizes the form + header/footer sections, and attaches
  the embedded lint. A form it builds passes the lint clean by construction. The
  geometry is a pure function (`_plan_layout`) covered by
  `tests/test_build_form_layout.py`.
- **Design tokens (`mcp_access/design_defaults.py`).** Single source of truth
  for the grid, standard control sizes, margins/gaps, fonts and a closed BGR
  palette (`bgr(r,g,b)` builds an Access colour Long; `snap(v)` rounds to the
  grid). `lint.py`, `build_form.py` and `access_tips('layout')` all read from it.
- **`snap_to_grid` (opt-in, default false)** on `access_create_control` and
  `access_set_control_props` — rounds Left/Top/Width/Height to the 60-twip grid;
  `-1` (auto) values are left untouched.
- **`access_tips('layout')`** — the canonical numbers, the columnar/two-column
  recipe, the palette and a `build_form` example, for hand placement.

### Changed

- **Four new lint rules, all `info`-severity:** `grid_alignment` (off the
  60-twip grid), `spacing_consistency` (uneven column gaps), `edge_margin`
  (control hugging the form edge), `hierarchy` (action text smaller than body
  text). They enrich the full `access_lint_form` report but **never** change the
  PASS/REVIEW/FAIL verdict (which only counts errors/warnings) and **never**
  reach the compact lint embedded in mutation results — so the embedded path is
  not made noisier. Each is deliberately conservative.

## 0.7.44 — 2026-06-12

Follow-ups to the attached-mode dialog hangs reported by
[@CaptainStormfield](https://github.com/CaptainStormfield)
([#31](https://github.com/unmateria/MCP-Access/issues/31)). The core complaint
(global watchdog disabled on attached instances) was already fixed in v0.7.43,
released the same day as the reported incidents — the remaining gaps are
closed here. No new tools — tool count stays **66**.

### Fixed

- **`access_eval_vba` gains an optional `timeout` parameter** — same dialog
  watchdog treatment as `access_run_vba`. With it, a MsgBox/InputBox (or any
  modal) raised by the evaluated expression is auto-dismissed and an
  actionable error returned, instead of relying solely on the global
  watchdog's grace period. Covers both `Application.Eval` and the temp-module
  fallback (which runs via `Application.Run` and is just as blockable). (#31)
- **Stale `_mcp_eval_wrapper` temp modules no longer wedge the session.**
  When the eval fallback's `VBComponents.Remove` failed (e.g. a modal was
  blocking), the orphan module's dangling name broke every later call with
  *"cannot find the procedure 'Module1._mcp_eval_wrapper'"* until a full
  reconnect. The fallback now sweeps leftover marker-tagged std modules
  before creating a new one (best-effort, scans only the first lines of each
  std module). (#31)
- **`access_delete_object` no longer triggers the *"Do you want to save
  changes to the design of module X?"* prompt.** Dirty VBA modules (often
  left by user code run via eval, e.g. a `VBComponents.Add`) are persisted
  best-effort before `DoCmd.DeleteObject` — `RunCommand acCmdSaveAllModules`
  (280), falling back to per-module `DoCmd.Save`. Deliberately NOT applied to
  close/quit paths: on attached instances that would silently persist the
  interactive user's half-finished VBE edits without being asked. (#31)
- **Watchdog dismissals are now surfaced in the tool result.** Every dialog
  auto-dismissed by any watchdog records its title; if it happened while a
  tool call was in flight, the result gains a note naming the dialog —
  converting a silent dismissal (whose Cancel may have altered the outcome)
  into a traceable event. (#31)

## 0.7.43 — 2026-06-11

Wedged-session detection — thanks to
[@CaptainStormfield](https://github.com/CaptainStormfield)
([#30](https://github.com/unmateria/MCP-Access/pull/30)) — plus a usability
bughunt round across the whole server. No new tools — tool count stays **66**.

### Fixed

- **Self-closing databases no longer wedge the COM session permanently**
  (from PR #30 by @CaptainStormfield, reimplemented with refinements). Two
  holes worked together: `_switch()` recorded `_db_open` without verifying the
  database actually stayed open (a startup form erroring out — e.g. broken
  backend links with `AllowBypassKey=False` defeating the SHIFT bypass — can
  close the db during the open), and `connect()`'s health check only probed
  `app.Visible`, which passes on an Access instance whose database was closed
  under it. On attached instances, reconnects re-attached to the same broken
  instance forever. Now: `_switch()` validates `CurrentDb()` after the open and
  raises an actionable error (naming AutoExec/startup-form failure, broken
  links, `AllowBypassKey=False`) after resetting the session via `quit()`;
  `connect()` detects a dead db on an otherwise-alive instance and
  auto-reconnects with a specific log message. The same post-open validation
  was added to `access_create_database`'s reopen path (a gap the PR didn't
  cover).
- **Modal dialogs no longer hang VBE tool calls on ATTACHED Access instances.**
  The v0.7.40 global dialog watchdog was disabled entirely when the session
  attached to the user's running Access (to never dismiss an interactive
  user's dialogs) — but that left any modal provoked by OUR blocked COM call
  (e.g. *"Error accessing file. Network connection may have been lost."* from
  a VBA project with a `TYPE_E_LIBNOTREGISTERED` reference) hanging the tool
  call until a human clicked, observed for ~1 hour in the field. The watchdog
  now also runs on attached instances but only dismisses dialogs while one of
  our tool calls has been in flight longer than a conservative 5 s grace
  (vs 3 s for spawned). A dialog with no tool call in flight belongs to the
  interactive user and is never touched.
- **`access_vbe_search_all` / `access_find_usages` / `access_find_definition`
  no longer report a clean `total: 0` when modules are inaccessible.** Each
  per-object failure used to be swallowed (`except: continue`) — so a VBA
  project that fails to load returned "0 matches", a false *"it doesn't
  exist"*. Results now include `objects_skipped`, an `errors` list (capped at
  20) and a warning when anything was skipped.
- **`access_list_references` no longer dies on a broken reference.** Reading
  `FullPath` on an unregistered library raises `com_error` and killed the whole
  call — exactly when you need the tool most. Every property is now read
  defensively (`null` on failure, reference marked `is_broken`), and the
  result carries `broken_count` + a warning.
- **Unclosed `/* ...` block comment no longer slips past the destructive-SQL
  guard.** `_sql_effective_prefix` returned `""` for an unclosed comment, so
  `/*\nDELETE FROM t` was classified non-destructive. Now fails closed by
  classifying the remaining text.
- **`access_set_code` with VBA-only code for a non-existent form/report** now
  raises a clear error ("create it first with access_create_form, or pass a
  full definition") instead of falling through to `LoadFromText` and dying
  with an opaque *"errors while importing"*.
- **`access_vbe_module_info` text fallback** (used when VBE can't locate a
  proc variant) now recognises `End Sub ' comment` — a trailing comment after
  the End keyword no longer breaks the proc-length scan.

### Improved

- **`access_vbe_replace_lines`**: omitting `start_line` in single mode now
  raises *"start_line is required (1-based)…"* instead of the cryptic
  *"start_line 0 out of range (1-N)"*. Batch mode gained the same
  destructive-delete note single mode already had (operations that delete
  lines but insert nothing are called out — the misnamed-argument footgun).
- **`access_execute_batch`**: new optional `limit` parameter for SELECT rows
  per statement (1-10000, default 100 — previously hardcoded).
- **`access_vbe_get_lines`**: an empty module now reports *"module is empty
  (0 lines)"* instead of *"start_line 1 out of range (1-0)"*, and
  `end_line < start_line` is rejected with a clear message.

## 0.7.42 — 2026-06-06

VBE procedure-editing fixes from field reports (thanks to
[@TvanStiphout-Home](https://github.com/TvanStiphout-Home)). No new tools — tool
count stays **66**.

### Fixed

- **`access_vbe_replace_proc` no longer eats the blank separator line above a
  procedure.** `ProcStartLine` is the previous proc's `End` + 1, so it *includes*
  the blank line VBE attributes to the proc; the old code deleted from there and
  re-inserted code with no leading blank, consuming the separator on every
  replace. Replaces now preserve the leading blank line(s) — delete/insert happen
  below them. A pure delete (`new_code=''`) still removes the whole range incl.
  the leading blank (so deleting a proc closes its gap cleanly).
- **No more spurious "Option … expected in first 5 lines" warning on modules
  with a long comment header** (e.g. a banner block pushing `Option Compare` past
  line 5). The structural health check replaced its fixed line-number threshold
  with a rule that flags an Option statement only when real (non-comment,
  non-blank) code already appeared above it — still catching genuinely misplaced
  Option statements.

### Added

- **`access_vbe_replace_lines` accepts `new_lines`** (a list of strings, joined
  with `\n`; `''` entries become blank lines) as an alias for `new_code`. This
  closes a destructive footgun: a call that passed the code under a wrong key
  left `new_code` empty and silently degraded into a pure delete. A replace that
  deletes lines but inserts nothing now also appends an explicit note to its
  result, so a destructive no-op is never silent.

### Changed

- `access_vbe_get_proc` / `access_vbe_module_info` descriptions and docstrings
  now spell out `start_line` (VBE proc start — includes preceding blank/comment
  lines) vs `body_line` (the `Sub`/`Function`/`Property` declaration line).
  `access_tips('vbe')` documents both, the separator-preserving replace, and
  `new_code=''` deletion.

### Tests

- `tests/test_vbe_fixes.py` — 11 COM-free tests for the Option-placement check
  and the `new_lines` alias normalisation. Blank-separator preservation and the
  end-to-end `new_lines` path were verified with a live COM integration run.

## 0.7.41 — 2026-05-29

Adds **deterministic UI design validation** so the assistant stops accepting
objectively broken form/report layouts. **65 → 66 tools.**

### Added

- **`access_lint_form`** — pure-Python, rules-based lint of a form/report.
  Returns structured JSON violations with `summary.verdict` (PASS / REVIEW /
  FAIL) and per-violation `suggested_fix`. Rules: `contrast` (WCAG 2.1
  ratio — catches white-on-white & low-contrast text), `overlap`,
  `out_of_bounds`, `truncation`, `sibling_inconsistency`, `misalignment`,
  `invisible_or_zero_size`. Static — one `SaveAsText` export, never opens
  Design view. `measure="auto"|"wizhook"|"heuristic"` (WizHook gives exact
  rendered text width when the VBA project is compiled; otherwise a
  conservative heuristic).
- **Embedded enforcement.** `access_set_control_props`,
  `access_set_multiple_controls` and `access_create_control` now attach a
  compact `lint` block (errors + warnings) to their result **automatically**.
  The validation is deterministic and lives entirely inside the MCP — it
  cannot be skipped or "talked past" by the model. `skip_lint=true` opts out
  for bulk programmatic edits. A lint failure never breaks the mutation.
- `access_tips('lint')` documents the rules, thresholds and colour encoding.
- Unit tests: `tests/test_lint.py` (29 COM-free tests for colour decoding,
  WCAG contrast, geometry parsing, and every rule incl. false-positive guards).

### Notes

False-positive guards, hardened against a real, densely populated form (findings
dropped 62 → 10, remainder genuine):

- **Conditional formatting** overrides colours at runtime (binary in the export)
  → contrast skips + notes those controls.
- **Captions wrap** — Labels *and* CommandButtons; line breaks (`\015\012`) are
  split, truncation compares wrapped-line count vs lines that fit the height.
- **`sibling_inconsistency` clusters** values: two legitimate sizes (main vs
  inline buttons) are both accepted; only a lone outlier flags.
- **Transparent buttons** stacked on styled labels (the custom-button pattern)
  are not flagged as overlaps.
- **Heuristic width is calibrated** for narrow UI fonts (≈0.46×) and only flags
  overflow past 1.25× — bold header labels that fit are no longer flagged.
- Absent dimensions inherit form defaults (not zero); attached labels,
  cross-tab-page controls, container Pages and transparent layering never count
  as overlaps; icon buttons aren't measured for caption truncation.
- Access auto-grows form Width / section Height to fit controls, so horizontal
  `out_of_bounds` rarely fires for forms (still effective for reports and
  negative coordinates).

## 0.7.40 — 2026-05-29

Fixes an indefinite hang on databases whose **startup form raises a blocking
modal** during open / VBE access.

### Fixed

- **Global dialog watchdog.** Until now the dialog-dismiss watchdog only ran
  during `open` / `compile` / `run_vba`. Operations that access the VBE
  (`vbe_get_proc`, `find_definition`, `module_info`, ...) had **no** watchdog,
  so a modal raised by a DB's startup form — e.g.
  `"Error accessing file. Network connection may have been lost."` on a DB with
  `StartupForm` set — would hang the COM call forever (observed: a ~1-hour hang
  on one such database). `_Session` now starts a background watchdog in
  `_launch()` that dismisses Access-owned `#32770` / wizard dialogs which
  persist past a 3 s grace period, for the whole lifetime of the spawned Access
  process. The grace period lets operation-specific watchdogs handle (and
  screenshot) their own dialogs first; the global thread only backstops
  un-watched operations. It is **not** started when attached to an existing
  interactive user session.
- Factored the dialog enumeration into `_find_dialog_hwnds_by_pid()` (shared by
  the dismisser and the new watchdog).

## 0.7.39 — 2026-05-28

Hardening of the v0.7.38 `_looks_like_vba_only` detector. No behaviour change
for the well-formed cases that v0.7.38 already handled.

### Fixed

- **VBA comments containing "Begin Form" or "Version =" no longer
  misclassify pure VBA as a form export.** The previous detector ran
  `_FORM_EXPORT_RE.search(code)` over the whole text, so a comment like
  `' Begin Form: this sub opens it` made `_looks_like_vba_only` return
  False — sending the file through `LoadFromText` instead of VBE
  injection, which would then fail with "errors while importing". The
  new detector only inspects the first non-blank line for `Version =NN`
  and the first 20 lines for `Option Compare` / Sub/Function/etc.,
  matching how Access actually emits form text exports (the `Version`
  declaration is the very first line of any SaveAsText output).
- `_VBA_HINT_RE` now also matches `Public Static Sub` / `Public Static
  Function` (a real-world VBA pattern used by counters and singletons).

## 0.7.38 — 2026-05-28

DX fixes for `access_set_code` on freshly-created forms, `access_create_control`
with TabControl-Page parents, and the cryptic VBE error you got when the form
had no code module yet. All real-world tripping points hit while building a
tab-heavy data-entry form — see notes below for the actual reproductions.

### Fixed

- **`ac_set_code(form|report)`** failed on forms recently created via
  `ac_create_form` because `LoadFromText` was always invoked even when the
  caller passed pure VBA (`Option Compare Database` + `Private Sub …`). Pure
  VBA isn't a valid form text export, and the binary-section restoration only
  works against an already-exported baseline — so the import raised `errors
  while importing` and rolled the form back. The same code now detects
  VBA-only input and routes it through `_inject_vba_after_import` (open in
  Design view → activate `HasModule` → write via VBE), preserving the form
  layout and never touching `LoadFromText`. A full form export (containing
  `Version =` / `Begin Form`) still takes the original `LoadFromText` path.
  See `code.py:_looks_like_vba_only`.

- **`ac_vbe_module_info` / any VBE read on a brand-new form** raised
  `Subscript out of range` with a misleading error mentioning *"Trust access to
  the VBA project object model"*. The actual cause: `HasModule=False` on a
  form just made by `ac_create_form`, so `VBComponents("Form_xxx")` had nothing
  to return. `_force_vbe_init` now flips `HasModule` on when the form/report is
  opened in Design view before retrying. The fallback error message is also
  rewritten so it no longer blames Trust Center first when the obvious cause
  is a missing module. See `vbe.py:_force_vbe_init`.

- **`ac_create_control` rejected `Parent`** (capital P) with `Property
  'CreateControl.Parent' can not be set` because special keys were popped from
  `props` case-sensitively. Hand a control to a TabControl Page using
  `{"parent": "myTab", ...}` or `{"Parent": "myTab", ...}` — both work now.
  Same case-insensitive treatment for `section`, `column_name`, `left`, `top`,
  `width`, `height`. See `controls.py:_pop_ci`.

- **`ac_create_control` lost properties Access exposes only via the
  `Properties` collection.** `setattr(ctrl, "ScrollBars", 2)` raises for some
  control types even when `ctrl.Properties("ScrollBars").Value = 2` succeeds.
  The loop now retries via the Properties collection before recording an
  entry in `property_errors`. Properties that don't exist at all on a given
  control type (e.g. the `MultiLine` UserForm property on an Access TextBox)
  still fail loudly — those are legitimate user errors.

### Added

- **`ac_create_control` accepts `control_name` at the top level** (in addition
  to `{"Name": "..."}` inside `props`). Without this you had to discover that
  the control was auto-named `Command1` / `Label2` and rename it via
  `set_control_props` in a second round-trip. `props["Name"]` still wins if
  both are provided, so existing callers don't change behaviour.

### Notes

- No schema changes to existing required fields — `control_name` is optional,
  the `Parent` fix is silent, and `set_code` only takes the new path when the
  caller passes VBA-only code to an existing form/report. Old fixtures that
  pass complete form exports continue using `LoadFromText`.
- The CLAUDE.md *Recipes* section gained two entries documenting the new flows
  for callers who want to build forms from scratch in one MCP session.

## 0.7.36 — 2026-05-25

Five new capabilities. Net tool count 62 → 65 (3 new tools; macros and the Office-version autodetect refactor are non-additive). All changes are additive — no existing schema or function signature was modified.

### Added
- **`access_search_data`** — search any text string across Text/Memo fields of any local table in a single call. Skips system tables (`MSys*`, `~*`) and linked tables (querying remote SQL servers with `LIKE` per column is rarely what the caller wants). Per-table and total caps, `match_case`, optional `tables` whitelist. Returns matches grouped by table with an `_excerpt` around each hit. See `sql.py:ac_search_data`.
- **`access_clone_object`** — duplicate a form, report, module, class_module, query or macro to a new name. Internally `SaveAsText` → `LoadFromText` with binary sections (PrtMip / PrtDevMode / NameMap / GUID) preserved by using a raw read path that bypasses `strip_binary_sections`. VBA code-behind comes along via the existing `ac_set_code` injection. Refuses to overwrite unless `overwrite=true`. See `code.py:ac_clone_object`.
- **`access_manage_tab_order`** — get / set / auto_renumber the `TabIndex` of controls on a form or report. `get` returns controls grouped by section, sorted by TabIndex. `set` assigns 0..N-1 in the order of `tab_order` (two-phase write to avoid the unique-index collision Access enforces per section). `auto_renumber` re-sequences existing TabIndex values per section. Skips controls that don't support TabIndex (Label, Line, Rectangle, Image, PageBreak, Page). Optional `section` filter. See `controls.py:ac_manage_tab_order`.

### Changed
- **Macros: docs and tips upgrade only — no new tool**. Macros were already fully supported via `access_get_code` / `access_set_code` (UTF-16 encoding correctly applied for SaveAsText/LoadFromText), `access_list_objects`, `access_run_macro`, `access_delete_object`. Tool descriptions now call macros out explicitly, and `access_tips('macros')` documents the read → edit → write workflow.
- **Office version autodetect**. The hardcoded `16.0` / `Office16` references in `_Session._suppress_recovery_dialog`, `_Session._decompile` and `maintenance.ac_decompile_compact` are now driven by a one-shot registry probe (`_Session._detect_office_install`). Detection enumerates `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path` under HKLM, HKLM\\WOW6432Node and HKCU (per-user Click-to-Run) and picks the highest version with a working `MSACCESS.EXE`. Falls back to `App Paths\MSACCESS.EXE\(Default)` and finally to the previous hardcoded defaults — so machines with a normal Office install keep working unchanged, and machines with a different major version (15.0 / 14.0) or a non-default install root start working without manual edits. Schema of `access_decompile_compact` is unchanged.

### Notes
- Test passing the `access_clone_object` overwrite path on a class_module — the `_ensure_class_module_header` re-injection runs after the raw export.
- `_detect_office_install` never raises; the worst case logs `Could not detect Office install via registry — using hardcoded defaults` and behaviour is identical to v0.7.35.

## 0.7.35 — 2026-05-25

Preventive bug sweep across the codebase. No reported regressions, but several latent issues fixed.

### Fixed (critical)
- **`_Session.quit()`** captured the Access PID by calling `app.hWndAccessApp()` from the `atexit` thread, but the COM proxy was created on the COM worker — the cross-thread call returned silently and `pid` was `None`, disabling the `taskkill` fallback added recently. PID is now captured in `_launch()` on the COM worker and stored as `cls._pid`.
- **`coerce_arguments`** only widened scalar argument types — clients that serialize arrays/objects as JSON strings (common with some MCP transports) failed on every batch tool. The coerce now JSON-parses string-typed arrays/objects, and the schema fixup also widens `number`, `array`, `object`. Boolean coercion also accepts `on`/`y`/`si`/`sí`.
- **`restore_binary_sections`** matched the first `End` after `Begin Form` as the injection point. Forms with embedded subforms (their own nested `Begin Form ... End`) had `PrtMip` / `PrtDevMode` / `NameMap` injected inside the subform, corrupting the result on `ac_set_code`. Now tracks full block depth and injects at the outermost End.
- **`write_tmp`** used `errors="replace"` for ANSI codepage writes (`.bas` modules) — a non-cp1252 character (emoji, asian, `✓`) was silently replaced with `?`. Now `errors="strict"` and the resulting `UnicodeEncodeError` carries a concrete snippet of the offending text.
- **`access_compile_vba`** accepted a `timeout` parameter but never used it. Now controls the watchdog grace window (default 2s, clamped to 1–30).

### Fixed (medium)
- VBE read operations (`ac_vbe_get_lines`, `get_proc`, `module_info`, `find`, `search_all`) now close the form/report Design view before reading the CodeModule. Skipping this could surface as `Catastrophic failure` (-2147418113) when the same object was open in design mode.
- `ac_vbe_replace_lines` no longer calls `cm.DeleteLines(start, 0)` when the count clamps to zero (raised in VBE); error message now lists separate upper bounds for replace/delete vs pure insert.
- `ac_vbe_patch_proc` normalizes `find_text`/`replace_text` line endings to CRLF before the exact-match check (callers commonly send LF and were always falling through to the ws-normalized fallback), and warns when `find_text` appears more than once (only the first occurrence is replaced).
- `_proc_kind` no longer silently picks the first matching kind when a procedure name resolves to multiple kinds (a class with both `Property Get Foo` and `Property Let Foo` is normal VBA). It raises a descriptive error so the caller can disambiguate.
- `set_db_property` / `set_field_property` infer `dbDouble`, `dbDate`, `dbSingle`, `dbMemo` for float/datetime/long-string values. Previously these fell to `dbText` (stored as string).
- `_eval_via_temp_module` pre-binds `temp_name` to avoid `UnboundLocalError` in the cleanup `finally` if creation fails before `comp.Name` is read.
- Compile watchdog captures up to 3 dialog screenshots / texts and the caller picks the last one — the first dialog is often a benign "Save changes?" and the real compile error came last (and used to be discarded).
- `ac_create_relationship` validates that local/foreign fields exist on the referenced tables before `Append`, so the error names the missing field instead of a cryptic DAO message.
- `_check_module_health` and `module_info` regex now recognize `Public Static Sub`/`Function`/`Property`.
- `decompile_compact` resets `_Session._pid` and `_attached` when killing the spawned process, keeping the `quit()` fallback consistent.

### Fixed (low / hardening)
- `read_tmp` tries UTF-8 before cp1252 (cp1252 single-byte never raises and was masking real UTF-8 files as mojibake).
- `_invoke_app_run` validates `len(args) <= 30` explicitly instead of producing a confusing InvokeTypes failure via negative-multiplier padding.
- `_split_code_behind` matches `CodeBehindForm` only at the start of a line to avoid false positives from property values containing that literal.
- `_SQL_LINE_COMMENT` / `_SQL_BLOCK_COMMENT` removed (dead code) — the destructive guard already uses `_sql_effective_prefix`. `SELECT … INTO` (make-table) now flagged as destructive.
- `ac_create_database` rejects paths without `.accdb` / `.mdb` extension.
- `relink_table` UID/PWD detection uses a parameter-boundary regex (`(^|;)UID=`), not a substring check.
- Linked-table count query escapes `]` in table names by doubling.
- `compact_repair` cleans up the orphaned `_compact_tmp.accdb` if rollback succeeds.
- Safe-args logging in `server.py` / `dispatcher.py` guards against non-string `code`.
- `tools.py` docstring now reports 62 tools (was 58).

## 0.7.34 — 2026-05-05

### Fixed
- **`access_list_controls`** silently lost controls inside `Page` / `OptionGroup` containers when any earlier control in the same Page had a multi-line property block (`GUID = Begin … End`, `NameMap = Begin … End`, `ConditionalFormat = Begin … End`, etc.). The depth counter inside a control's body matched plain `Begin <Type>` but not `Property = Begin`, so the property's closing `End` was decremented without ever being incremented — the enclosing control was closed prematurely, and every control that came after it inside the Page was never enumerated. The form-level loop already handled this; the per-control loop now mirrors it.

  Visible symptom: `access_list_controls` reported a TabControl Page as a 15-line empty stub even though the Page actually contained dozens of controls. Fixed in `mcp_access/controls.py:_parse_controls`.
