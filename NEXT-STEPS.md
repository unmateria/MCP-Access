# Next steps — safety nets for form editing

Three improvements that came out of real production use, on a database large
enough that its VBA is exported to text and versioned in git.

All three share one theme: **the server already does the right thing, but it stays
silent when the caller gets it wrong, and the damage only shows up later** — in a
git diff, or when somebody opens the form and finds a control out of place.

Ordered by how much pain each one removes.

---

## 1. Warn when a new control lands inside a tab page but has no parent

### What happens

`access_create_control` takes `parent` (the 4th positional argument of
`CreateControl`) and it works. But it is optional and easy to forget. When it is
missing, the control is created **on the form's Detail section**, not inside the
tab page — even if its `left`/`top` put it visually on top of the page.

The form then looks broken in a way that is hard to attribute: the control shows
up on every tab, or floats above the tab control, and the person who opens the
form blames the last edit without knowing why.

Nothing in the current output hints at it. The call returns success, because it
*was* a success — just not the one that was meant.

### Repro

```jsonc
// A form with a tab control "tabDetails" whose page occupies, say,
// left=200 top=600 width=9000 height=5000
{ "tool": "access_create_control",
  "object_name": "frmOrders", "control_type": "CommandButton",
  "props": { "left": 1200, "top": 1500, "width": 2000, "height": 300 },
  "control_name": "btPrint" }
// -> created, but on Detail. Nobody says anything.
```

### Proposed behaviour

In `ac_create_control` (`mcp_access/controls.py`), **after** creating the control
and only when `parent` was empty:

1. Look for tab controls on the form (`ControlType == acTabCtl`, 123).
2. If the new control's rectangle is fully inside one of them, add a warning to
   the returned dict:

```jsonc
{ "created": "btPrint",
  "warning": "btPrint was created on the Detail section, but its position is
              inside the tab control 'tabDetails'. If you meant to put it on a
              tab page, pass props.parent with the PAGE name (not the tab
              control name) — e.g. \"parent\": \"pagGeneral\"." }
```

Do **not** re-parent automatically: a control deliberately placed over a tab is
legitimate, and silently moving things is worse than the original problem. A
warning is enough — the caller learns immediately instead of three commits later.

### Notes for whoever implements it

- The page name is what `CreateControl` wants, not the tab control's name. That
  distinction is the actual trap and the warning text should spell it out.
- `ac_get_control` already reports `parent` correctly (it comes out of the
  SaveAsText nesting), so the information needed to verify the fix is there.
- Same check is worth adding to `ac_build_form` if it creates controls in bulk.

### How to test

A tiny fixture form with one tab control and two pages. Create a control inside
the page area without `parent` → expect the warning. Create the same control with
`parent` → expect no warning. Create one clearly outside the tab → no warning.

---

## 2. Warn when writing code changes modules that were not touched

### What happens

The VBA editor keeps **one canonical spelling per identifier across the whole
project**. Declare `Dim Url As String` in a new module and the VBE rewrites every
existing `url` to `Url` — in every other module, silently.

VBA is case-insensitive, so nothing breaks. But in a repository where the VBA is
exported to text and versioned, a twenty-line change turns into a diff across
eight files, half of them modules the author never opened. It buries the real
change and it makes an innocent commit look like a refactor.

This is Office behaviour, **not** a bug in this server. But this server is what
the caller is holding when it happens, and it is in the best position to say so.

### Repro

1. A project where some module uses `url` in lowercase.
2. `access_set_code` a new module declaring a parameter named `Url`.
3. Export the modules to text: unrelated modules have changed.

### Proposed behaviour

In `ac_set_code` and the `ac_vbe_*` writers, take a cheap fingerprint of every
module **before** the write (`hash(cm.Lines(1, cm.CountOfLines))` per component),
take it again after, and report any module that changed without being the target:

```jsonc
{ "ok": true,
  "also_changed": ["modInvoices", "modLabels"],
  "note": "The VBA editor unifies the capitalisation of an identifier across the
           whole project, so declaring a name that already exists elsewhere with
           different casing rewrites those modules too. Harmless at runtime, but
           it shows up in version control." }
```

### Notes for whoever implements it

- Hashing the text of every module on every write costs one COM round trip per
  component. On a project with 35 modules that was not noticeable, but if it is,
  gate it behind a parameter (`report_side_effects=True`) or only run it for
  `ac_set_code` (whole-module writes), which is where new identifiers appear.
- Forms and reports have code modules too; include them if it is cheap.

### How to test

Fixture with `modA` containing `url`, then set a new `modB` declaring `Url`.
Expect `also_changed: ["modA"]`.

---

## 3. Say that injecting code rewrites the form's design-window size

### What happens

Injecting VBA into a form (through the VBE) makes Access rewrite, in the next
text export of that form, the `Left`/`Top`/`Right`/`Bottom` **of the `Begin Form`
block** — the position and size of the design window — plus `Checksum` and a few
internal hex blocks.

No control moves. But somebody reading the diff cannot tell that at a glance, and
the reasonable reaction to "my form's geometry changed" is to revert the commit.

### Proposed behaviour

Two options, cheapest first:

1. **Document it** in `access_tips` (the knowledge base tool) under form editing:
   *«After injecting code, the form-level `Left/Top/Right/Bottom` and `Checksum`
   change in the text export. That is the design window, not the layout. If the
   diff only touches those and no control's own `Left`/`Top`, nothing moved.»*
2. Optionally, have `ac_set_code`/`ac_vbe_*` read those four values before the
   write and restore them afterwards, so the export stays byte-stable.

Option 1 is worth doing regardless: it is the piece of knowledge that lets a
reviewer accept the diff without opening Access.

---

## 4. Minor: `access_get_control` returns `control_type: -1`

`ac_get_control` returns `type_name` correctly (`"CommandButton"`) but
`control_type` comes back as `-1` instead of the numeric constant (104). Since
`access_create_control` accepts either the name or the number, round-tripping a
control definition through `get_control` gives a number that cannot be fed back.

Resolve it from `type_name` via the existing `CTRL_TYPE_BY_NAME` map, or drop the
field rather than return a value that is not true.

---

## A word on scope

None of these change what the server does — they change what it *tells you* when
you are about to have a bad day. That is deliberate: the tools are already
correct, and the failures above all come from a caller not knowing something the
server already knew.
