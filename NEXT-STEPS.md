# Next steps — safety nets for form editing

Three improvements that came out of real production use, on a database large
enough that its VBA is exported to text and versioned in git.

All three shared one theme: **the server already did the right thing, but it
stayed silent when the caller got it wrong, and the damage only showed up
later** — in a git diff, or when somebody opened the form and found a control
out of place.

**Status: all four items resolved in v0.7.59.** Kept as the record of why each
one was done the way it was. See `CHANGELOG.md` for what shipped.

---

## 1. Warn when a new control lands inside a tab page but has no parent — DONE

`access_create_control` takes `parent` (the 4th positional argument of
`CreateControl`) and it works. But it is optional and easy to forget. When it is
missing, the control is created **on the form's Detail section**, not inside the
tab page — even if its `left`/`top` put it visually on top of the page. The call
returns success, because it *was* a success — just not the one that was meant.

**What it took.** More than expected: the tab control did not exist as far as
this server was concerned. `_parse_controls` only recognises a `Begin <Type>`
block whose type is in `CTRL_TYPE`, and 123 was not there — so a tab control had
no geometry anywhere in the package to compare against. It is now a first-class
control, and a container (see `CLAUDE.md` for why the container half is
load-bearing).

`ac_create_control` now returns a `warning` when the new control's rectangle
falls fully inside a tab control in the same section and `parent` was empty. It
names the tab, lists its pages, and says that the **page** name is what `parent`
wants — that distinction was the actual trap.

Nothing is re-parented automatically: a control deliberately placed over a tab
is legitimate, and silently moving things is worse than the original problem.

`ac_build_form` does not need the check — it always passes an empty parent and
its spec vocabulary has no tabs.

## 2. Warn when writing code changes modules that were not touched — DOCUMENTED, NOT DETECTED

The VBA editor keeps **one canonical spelling per identifier across the whole
project**. Declare `Dim Url As String` in a new module and the VBE rewrites every
existing `url` to `Url` — in every other module, silently. VBA is
case-insensitive, so nothing breaks; but in a repository where the VBA is
exported to text, a twenty-line change turns into a diff across files the author
never opened. This is Office behaviour, not a bug in this server.

**Why it was not implemented as a check.** The proposal was to fingerprint every
module before and after every write. That is hundreds of COM round-trips per
write on a large project, to report damage that is *already done* and *already
visible in the diff that prompted the question*. The detector would not prevent
anything. What was actually missing was the explanation, so that is what shipped:
`access_tips('vbe')` now says why those files changed and how to avoid it (match
the casing the identifier already has; `access_vbe_find` tells you which).

If this is ever revisited, the cheap version is modules only (never form/report
code-behind) and it belongs behind a parameter, not on by default.

## 3. Say that injecting code rewrites the form's design-window size — DONE

Injecting VBA into a form makes Access rewrite, in the next text export, the
`Left`/`Top`/`Right`/`Bottom` of the `Begin Form` block — the position and size
of the design window — plus `Checksum` and a few internal hex blocks. No control
moves, but a reviewer cannot tell that at a glance, and the reasonable reaction
to "my form's geometry changed" is to revert the commit.

Documented in `access_tips('design_vbe')`. The alternative (read those four
values before the write and restore them after) was rejected: it would need a
`LoadFromText` round-trip, which is exactly what `_inject_vba_after_import`
exists to avoid.

## 4. `access_get_control` returned `control_type: -1` — DONE

`type_name` was right (`"CommandButton"`) but `control_type` came back as `-1`,
because Access omits `ControlType =` when it equals the default. Round-tripping a
control definition therefore produced a number that could not be fed back to
`access_create_control`. It is now resolved from the `Begin <Type>` token via
`CTRL_TYPE_BY_NAME`.

One wrinkle worth remembering: for `WebBrowser` and the two `Navigation*`
controls, the SaveAsText number and the AcControlType number `CreateControl`
accepts are different. The resolved value is the latter — the one that makes the
round-trip work.

---

## A word on scope

None of these changed what the server does — they changed what it *tells you*
when you are about to have a bad day. That was deliberate: the tools were already
correct, and the failures above all came from a caller not knowing something the
server already knew.
