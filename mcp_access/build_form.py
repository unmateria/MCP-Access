"""
access_build_form — deterministic auto-layout for Access forms.

The recurring failure mode is the model emitting raw twip coordinates blind:
overlaps, out-of-bounds, ragged columns, inconsistent sizes. This tool removes
that whole class of error by taking the arithmetic away from the LLM. The model
describes the form *declaratively* — a title, an ordered list of fields, a row
of action buttons, single- or two-column — and this module computes every
Left/Top/Width/Height from the canonical grid in :mod:`mcp_access.design_defaults`,
applies the closed palette, assigns a sane tab order and sizes the form and its
sections. The model never touches a coordinate (it can still fine-tune
afterwards with ``access_set_control_props``).

The geometry is produced by the pure function :func:`_plan_layout` (no COM, unit
tested); :func:`ac_build_form` just walks the plan against a single Design-view
session and runs the lint at the end.
"""

import re
from typing import Any, Optional

from .core import _Session, log, invalidate_object_caches, invalidate_all_caches
from .helpers import coerce_prop
from .controls import (
    _resolve_ctrl_type, _open_in_design, _save_and_close, _get_design_obj,
    _attach_lint,
)
from . import design_defaults as D

# Form section indices (AcSection): Detail / FormHeader / FormFooter.
_AC_DETAIL, _AC_HEADER, _AC_FOOTER = 0, 1, 2
_SECTION_LABEL = {_AC_DETAIL: "Detail", _AC_HEADER: "FormHeader",
                  _AC_FOOTER: "FormFooter"}

# Spec "control" keyword → (CTRL_TYPE name, flags). flags drive styling/sizing.
_CONTROL_MAP = {
    "textbox": ("TextBox", set()),
    "text": ("TextBox", set()),
    "string": ("TextBox", set()),
    "number": ("TextBox", set()),
    "currency": ("TextBox", set()),
    "memo": ("TextBox", {"memo"}),
    "longtext": ("TextBox", {"memo"}),
    "note": ("TextBox", {"memo"}),
    "notes": ("TextBox", {"memo"}),
    "combobox": ("ComboBox", set()),
    "combo": ("ComboBox", set()),
    "dropdown": ("ComboBox", set()),
    "lookup": ("ComboBox", set()),
    "listbox": ("ListBox", set()),
    "list": ("ListBox", set()),
    "checkbox": ("CheckBox", {"checkbox"}),
    "check": ("CheckBox", {"checkbox"}),
    "bool": ("CheckBox", {"checkbox"}),
    "boolean": ("CheckBox", {"checkbox"}),
    "yesno": ("CheckBox", {"checkbox"}),
    "date": ("TextBox", {"date"}),
    "datetime": ("TextBox", {"date"}),
    "datepicker": ("TextBox", {"date"}),
}


def _ident(raw: str, prefix: str = "") -> str:
    """Turn a field caption into a safe control name (alnum, CamelCase tail)."""
    parts = re.findall(r"[A-Za-z0-9]+", str(raw or ""))
    if not parts:
        body = "Ctl"
    else:
        body = parts[0] + "".join(p[:1].upper() + p[1:] for p in parts[1:])
    name = (prefix + body[:1].upper() + body[1:]) if prefix else body
    return name


def _normalise_field(f: Any) -> dict:
    """Accept a bare string ('Nombre') or a dict spec; return a uniform dict."""
    if isinstance(f, str):
        f = {"field": f}
    elif not isinstance(f, dict):
        raise ValueError(f"Each field must be a string or object, got: {f!r}")
    return dict(f)


def _plan_layout(fields: list, actions: list, title: Optional[str],
                 layout: str, theme: str) -> dict:
    """Pure geometry planner. Returns form_width, section heights and a flat
    list of control specs with computed twip rects + styling props.

    No COM here — this is the deterministic core the LLM never has to reason
    about, and it is what the unit tests exercise.
    """
    styled = theme != "plain"
    pal = D.PALETTE
    layout = (layout or "single").lower()
    if layout not in ("single", "two-column", "two_column", "twocolumn"):
        layout = "single"
    two_col = layout != "single"
    cols = 2 if two_col else 1

    fields = [_normalise_field(f) for f in (fields or [])]
    actions = list(actions or [])
    need_header = bool(title)
    need_footer = bool(actions)

    controls: list[dict] = []
    used_names: set[str] = set()

    def _unique(name: str) -> str:
        base, n = name, 1
        while name.lower() in used_names:
            n += 1
            name = f"{base}{n}"
        used_names.add(name.lower())
        return name

    def _label_style(caption: str) -> dict:
        p: dict = {"Caption": caption}
        if styled:
            p.update({"FontName": D.BASE_FONT, "FontSize": D.LABEL_FONT_SIZE,
                      "FontWeight": D.FONT_WEIGHT_NORMAL, "ForeColor": pal["text"]})
        return p

    def _field_style(fl: dict, flags: set, bound: Optional[str]) -> dict:
        p: dict = {}
        if bound:
            p["ControlSource"] = bound
        if "date" in flags:
            p.setdefault("Format", "Short Date")
        if fl.get("row_source"):
            p["RowSource"] = str(fl["row_source"])
        if styled and "checkbox" not in flags:
            p.update({"FontName": D.BASE_FONT, "FontSize": D.FIELD_FONT_SIZE,
                      "ForeColor": pal["text"], "BackColor": pal["field_bg"],
                      "BackStyle": 1, "BorderStyle": 1,
                      "BorderColor": pal["field_border"]})
        # Checkboxes carry no colour props — CreateControl rejects ForeColor/
        # BackColor on a CheckBox ("Property 'CreateControl.ForeColor' can not
        # be set"); the box uses the theme. Leave it unstyled.
        # Per-field escape hatch: explicit props win over computed defaults.
        if isinstance(fl.get("props"), dict):
            p.update(fl["props"])
        return p

    # --- field geometry -----------------------------------------------------
    col_w = D.LABEL_W + D.GAP_LABEL + D.FIELD_W
    x_field0 = D.MARGIN_X + D.LABEL_W + D.GAP_LABEL
    y = D.MARGIN_Y
    max_right = D.MARGIN_X + col_w  # running max for single-col width

    def _emit_row(row_fields: list, top: int) -> int:
        """Place a row of 1..cols (label,field) pairs at `top`; return row height."""
        nonlocal max_right
        row_h = D.ROW_H
        for ci, fl in enumerate(row_fields):
            field = str(fl.get("field", "") or fl.get("name", "") or f"Campo{len(controls)}")
            ctl_kw = str(fl.get("control", "textbox")).lower().strip()
            type_name, flags = _CONTROL_MAP.get(ctl_kw, ("TextBox", set()))
            caption = fl.get("label")
            if caption is None:
                caption = f"{field}:"
            bound = fl.get("control_source")
            if bound is None and fl.get("bind", True) is not False:
                bound = field if not _looks_unbound(field) else None

            # widths/heights
            if "memo" in flags:
                fh = D.snap(fl.get("height", D.MEMO_H))
            else:
                fh = D.snap(fl.get("height", D.ROW_H))
            if "checkbox" in flags:
                fw = D.CHECKBOX_W
            elif two_col:
                fw = D.FIELD_W   # two-column keeps a strict grid; width_units ignored
            else:
                units = float(fl.get("width_units", 1) or 1)
                fw = D.snap(D.FIELD_W * units)

            x0 = D.MARGIN_X + ci * (col_w + D.COL_GAP)
            x_field = x0 + D.LABEL_W + D.GAP_LABEL
            lbl_name = _unique(str(fl.get("label_name") or _ident(field, "lbl")))
            fld_name = _unique(str(fl.get("name") or _ident(field, _PREFIX.get(type_name, "ctl"))))

            controls.append({
                "role": "label", "section": _AC_DETAIL, "type_name": "Label",
                "name": lbl_name, "left": x0, "top": top,
                "width": D.LABEL_W, "height": D.ROW_H,
                "props": _label_style(caption),
            })
            controls.append({
                "role": "field", "section": _AC_DETAIL, "type_name": type_name,
                "name": fld_name, "left": x_field, "top": top,
                "width": fw, "height": fh, "tab": True,
                "props": _field_style(fl, flags, bound),
            })
            row_h = max(row_h, fh)
            max_right = max(max_right, x_field + fw)
        return row_h

    if two_col:
        for r in range(0, len(fields), cols):
            row = fields[r:r + cols]
            rh = _emit_row(row, y)
            y += rh + D.ROW_GAP
    else:
        for fl in fields:
            rh = _emit_row([fl], y)
            y += rh + D.ROW_GAP

    detail_height = D.snap(max(y - D.ROW_GAP + D.MARGIN_Y, D.ROW_STRIDE))

    # --- form width ---------------------------------------------------------
    if two_col:
        form_width = D.MARGIN_X + cols * col_w + (cols - 1) * D.COL_GAP + D.MARGIN_X
    else:
        form_width = max_right + D.MARGIN_X
    form_width = D.snap(form_width)

    # --- header title -------------------------------------------------------
    header_height = 0
    if need_header:
        header_height = D.HEADER_H
        # Dark bold title on the form-header band. Access renders the header
        # with a themed (light) gradient that overrides a literal section
        # BackColor, so a white title is invisible — a dark title is readable on
        # whatever the theme paints and always clears the contrast check.
        tprops = {"Caption": title}
        if styled:
            tprops.update({"FontName": D.BASE_FONT, "FontSize": D.TITLE_FONT_SIZE,
                           "FontWeight": D.FONT_WEIGHT_BOLD,
                           "ForeColor": pal["text"]})
        controls.append({
            "role": "title", "section": _AC_HEADER, "type_name": "Label",
            "name": _unique("lblTitle"),
            "left": D.MARGIN_X, "top": D.snap((D.HEADER_H - D.ROW_H) / 2),
            "width": max(D.LABEL_W, form_width - 2 * D.MARGIN_X), "height": D.ROW_H,
            "props": tprops,
        })

    # --- footer action buttons ---------------------------------------------
    footer_height = 0
    if need_footer:
        footer_height = D.FOOTER_H
        n = len(actions)
        total_w = n * D.BUTTON_W + (n - 1) * D.BUTTON_GAP
        start_x = form_width - D.MARGIN_X - total_w
        if start_x < D.MARGIN_X:
            start_x = D.MARGIN_X          # too many buttons → left-align
        btn_top = D.snap((D.FOOTER_H - D.BUTTON_H) / 2)
        for i, act in enumerate(actions):
            if isinstance(act, str):
                act = {"caption": act}
            caption = str(act.get("caption", f"Botón{i + 1}"))
            bname = _unique(str(act.get("name") or _ident(caption, "btn")))
            bprops: dict = {"Caption": caption}
            if act.get("on_click"):
                bprops["OnClick"] = str(act["on_click"])
            if isinstance(act.get("props"), dict):
                bprops.update(act["props"])
            controls.append({
                "role": "button", "section": _AC_FOOTER, "type_name": "CommandButton",
                "name": bname, "left": D.snap(start_x + i * (D.BUTTON_W + D.BUTTON_GAP)),
                "top": btn_top, "width": D.BUTTON_W, "height": D.BUTTON_H,
                "tab": True, "props": bprops,
            })

    return {
        "form_width": form_width,
        "detail_height": detail_height,
        "header_height": header_height,
        "footer_height": footer_height,
        "need_header": need_header,
        "need_footer": need_footer,
        "controls": controls,
    }


_PREFIX = {"TextBox": "txt", "ComboBox": "cbo", "ListBox": "lst",
           "CheckBox": "chk", "Label": "lbl"}


def _looks_unbound(field: str) -> bool:
    """A field name with spaces/operators isn't a plain column → leave unbound."""
    return bool(re.search(r"[ =()\[\]!.]", field or ""))


def _apply_props(ctrl: Any, props: dict) -> dict:
    """setattr each prop; fall back to the Properties collection; collect errors.
    Mirrors ac_create_control so behaviour is identical to a hand-built control."""
    errors: dict[str, str] = {}
    for key, val in props.items():
        cv = coerce_prop(val)
        try:
            setattr(ctrl, key, cv)
            continue
        except Exception as exc_attr:
            try:
                ctrl.Properties(key).Value = cv
                continue
            except Exception:
                errors[key] = str(exc_attr)
    return errors


def ac_build_form(
    db_path: str, form_name: str, *,
    record_source: Optional[str] = None,
    title: Optional[str] = None,
    fields: Optional[list] = None,
    actions: Optional[list] = None,
    layout: str = "single",
    default_view: Optional[int] = None,
    theme: str = "light",
    overwrite: bool = False,
    skip_lint: bool = False,
) -> dict:
    """Build a complete, well-laid-out form from a declarative spec.

    fields: list of strings or objects. Object keys:
      field (column / base name), label, control
      (textbox|memo|combobox|listbox|checkbox|date), name, control_source,
      row_source, width_units (single-column only), height, props (override dict).
    actions: list of strings or objects {caption, name, on_click, props} → a row
      of buttons in the footer.
    title: a form-header band with this caption (bold dark title, readable on
      the themed header band).
    layout: 'single' or 'two-column'. theme: 'light' (palette) or 'plain' (geometry only).

    All coordinates are computed from mcp_access.design_defaults and snapped to
    the 60-twip grid. Returns the geometry, the controls created, the tab order
    and an embedded lint of the result.
    """
    from .code import ac_create_form

    plan = _plan_layout(fields, actions, title, layout, theme)
    app = _Session.connect(db_path)

    # Replace an existing form only when asked.
    if overwrite:
        try:
            app.DoCmd.Close(2, form_name, 2)  # acForm, acSaveNo
        except Exception:
            pass
        try:
            app.DoCmd.DeleteObject(2, form_name)
        except Exception:
            pass

    has_hf = plan["need_header"] or plan["need_footer"]
    ac_create_form(db_path, form_name, has_header=has_hf,
                   record_source=record_source, default_view=default_view)

    property_errors: dict[str, dict] = {}
    created: list[dict] = []

    _open_in_design(app, "form", form_name)
    try:
        obj = _get_design_obj(app, "form", form_name)
        if title:
            try:
                obj.Caption = title
            except Exception:
                pass

        # Create every control in this single Design session.
        for spec in plan["controls"]:
            ctype = _resolve_ctrl_type(spec["type_name"])
            try:
                ctrl = app.CreateControl(
                    form_name, ctype, spec["section"], "", "",
                    spec["left"], spec["top"], spec["width"], spec["height"],
                )
            except Exception as exc:
                property_errors[spec["name"]] = {"_create": str(exc)}
                continue
            try:
                ctrl.Name = spec["name"]
            except Exception as exc:
                log.warning("build_form: could not name control '%s': %s",
                            spec["name"], exc)
            errs = _apply_props(ctrl, spec["props"])
            if errs:
                property_errors[spec["name"]] = errs
            created.append({
                "name": spec["name"], "role": spec["role"],
                "type_name": spec["type_name"],
                "section": _SECTION_LABEL.get(spec["section"], str(spec["section"])),
                "left": spec["left"], "top": spec["top"],
                "width": spec["width"], "height": spec["height"],
            })

        # Form + section geometry (after controls so our sizes stick).
        try:
            obj.Width = plan["form_width"]
        except Exception as exc:
            log.warning("build_form: could not set form Width: %s", exc)
        _set_section(obj, _AC_DETAIL, plan["detail_height"],
                     D.PALETTE["form_bg"] if theme != "plain" else None)
        if plan["need_header"]:
            # Keep Access' themed header band (a literal BackColor doesn't stick
            # against the theme gradient); the dark title reads fine on it.
            _set_section(obj, _AC_HEADER, plan["header_height"], None)
        elif has_hf:
            _set_section(obj, _AC_HEADER, 0, None)
        if plan["need_footer"]:
            _set_section(obj, _AC_FOOTER, plan["footer_height"],
                         D.PALETTE["form_bg"] if theme != "plain" else None)
        elif has_hf:
            _set_section(obj, _AC_FOOTER, 0, None)

        # Tab order: data controls + buttons, per section, in spec order.
        tab_order = _assign_tab_order(obj, plan["controls"])
    finally:
        _save_and_close(app, "form", form_name)
        invalidate_object_caches("form", form_name)

    result: dict = {
        "name": form_name,
        "layout": "two-column" if _is_two_col(layout) else "single",
        "theme": theme,
        "form_width": plan["form_width"],
        "sections": _sections_summary(plan),
        "controls_created": created,
        "control_count": len(created),
        "tab_order": tab_order,
        "snapped_to_grid": D.GRID,
    }
    if record_source is not None:
        result["record_source"] = record_source
    if property_errors:
        result["property_errors"] = property_errors
    invalidate_all_caches()
    return _attach_lint(result, db_path, "form", form_name, skip_lint)


def _is_two_col(layout: str) -> bool:
    return (layout or "single").lower() not in ("single", "")


def _sections_summary(plan: dict) -> dict:
    out = {"Detail": plan["detail_height"]}
    if plan["need_header"]:
        out["FormHeader"] = plan["header_height"]
    if plan["need_footer"]:
        out["FormFooter"] = plan["footer_height"]
    return out


def _set_section(obj: Any, index: int, height: int, backcolor: Optional[int]) -> None:
    """Set a section's Height (and BackColor) defensively."""
    try:
        sec = obj.Section(index)
    except Exception:
        return
    try:
        sec.Height = int(height)
    except Exception as exc:
        log.warning("build_form: could not set section %s height: %s", index, exc)
    if backcolor is not None:
        try:
            sec.BackColor = int(backcolor)
        except Exception:
            pass


def _assign_tab_order(obj: Any, specs: list) -> list:
    """Set TabIndex per section in spec order on the tabbable controls.

    Access auto-renumbers the rest to keep indices unique (same idiom as
    ac_manage_tab_order), so a single forward pass yields the intended order.
    """
    order: list[str] = []
    per_section: dict[int, int] = {}
    for spec in specs:
        if not spec.get("tab"):
            continue
        sec = spec["section"]
        idx = per_section.get(sec, 0)
        try:
            obj.Controls(spec["name"]).TabIndex = idx
            per_section[sec] = idx + 1
            order.append(spec["name"])
        except Exception as exc:
            log.warning("build_form: TabIndex on '%s' failed: %s", spec["name"], exc)
    return order
