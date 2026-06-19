"""
access_build_form — deterministic auto-layout for Access forms.

The recurring failure mode is the model emitting raw twip coordinates blind:
overlaps, out-of-bounds, ragged columns, inconsistent sizes. This tool removes
that whole class of error by taking the arithmetic away from the LLM. The model
describes the form *declaratively* — a title, an ordered list of fields, a row
of action buttons, single- or two-column — and this module computes every
Left/Top/Width/Height from the canonical grid in :mod:`mcp_access.design_defaults`,
applies a palette, assigns a sane tab order and sizes the form and its sections.

Themes (``theme=``)
-------------------
- ``light``  — default. Calibri, white fields, dark title on the themed header.
- ``plain``  — geometry only, no colours/fonts.
- ``polish`` — Segoe UI, more air, and the database chrome (record selectors,
  navigation buttons, scrollbars, dividing lines) turned off.
- ``flat``   — modern flat look on top of ``polish``: a solid accent header band
  and flat coloured buttons (built with Rectangles + a transparent click layer,
  the only reliable way to colour these in native Access), a bordered card
  around the fields, grey canvas. Stays lint-clean by construction.

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

_WHITE = D.bgr(255, 255, 255)

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

_PREFIX = {"TextBox": "txt", "ComboBox": "cbo", "ListBox": "lst",
           "CheckBox": "chk", "Label": "lbl"}


def _resolve_theme(name: str) -> dict:
    """Resolve a theme name into a flat config the planner reads.

    ``light``/``plain``/``polish``/``flat`` are literal branches (unchanged).
    A name in :data:`design_defaults.DIRECTIONS` is a *curated design direction*
    — a coherent typeface + type scale + WCAG-verified palette + density. It is
    expanded onto the same flat keys, plus two extra keys the planner reads
    defensively: ``palette`` (the direction's BGR dict) and ``title_font`` (the
    display typeface for the form title).
    """
    name = (name or "light").lower()
    t = {
        "name": name, "styled": True, "font": D.BASE_FONT,
        "title_size": D.TITLE_FONT_SIZE, "label_size": D.LABEL_FONT_SIZE,
        "field_size": D.FIELD_FONT_SIZE,
        "margin": D.MARGIN_X, "margin_y": D.MARGIN_Y, "row_gap": D.ROW_GAP,
        "chrome_off": False, "header_band": False, "card": False,
        "flat_buttons": False, "canvas": None,
    }
    if name in D.DIRECTIONS:
        dr = D.DIRECTIONS[name]
        fonts, scale, dens = dr["fonts"], dr["scale"], D.DENSITY[dr["density"]]
        pal = dr["palette"]
        t.update(
            font=fonts["body"], title_font=fonts["display"],
            title_size=scale["title"], label_size=scale["caption"],
            field_size=scale["body"],
            margin=dens["margin"], margin_y=dens["margin_y"],
            row_gap=dens["row_gap"],
            card=dr["card"], canvas=pal["form_bg"], palette=pal,
            **D.DIRECTION_COMMON,
        )
    elif name == "plain":
        t["styled"] = False
    elif name == "polish":
        t.update(font="Segoe UI", title_size=18, margin_y=300, row_gap=180,
                 chrome_off=True)
    elif name == "flat":
        # flat_buttons (the rect+label+transparent-button hack) is OFF: native
        # command buttons centre their caption and just look cleaner. The modern
        # feel comes from the accent header band + card + grey canvas + no chrome.
        t.update(font="Segoe UI", title_size=18, margin=300, margin_y=300,
                 row_gap=180, chrome_off=True, header_band=True, card=True,
                 flat_buttons=False, canvas=D.PALETTE["form_bg"])
    return t


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


def _looks_unbound(field: str) -> bool:
    """A field name with spaces/operators isn't a plain column → leave unbound."""
    return bool(re.search(r"[ =()\[\]!.]", field or ""))


def _rect(name: str, section: int, left: int, top: int, width: int, height: int,
          fill: int, border_color: Optional[int]) -> dict:
    """A decorative Rectangle (solid fill, optional border). Rectangles honour
    BackColor reliably, unlike theme-tinted form sections."""
    props = {"BackStyle": 1, "BackColor": fill}
    if border_color is not None:
        props.update({"BorderStyle": 1, "BorderColor": border_color, "BorderWidth": 1})
    else:
        props["BorderStyle"] = 0
    return {"role": "decoration", "section": section, "type_name": "Rectangle",
            "name": name, "left": left, "top": top, "width": width,
            "height": height, "props": props}


def _plan_layout(fields: list, actions: list, title: Optional[str],
                 layout: str, theme: str) -> dict:
    """Pure geometry planner. Returns form_width, section heights, the canvas
    colour and a flat, z-ordered list of control specs with computed twip rects.

    No COM here — this is the deterministic core the LLM never reasons about, and
    it is what the unit tests exercise. Background decorations are emitted before
    their foreground siblings so Access' create-order z-stacking puts them behind.
    """
    T = _resolve_theme(theme)
    styled = T["styled"]
    pal = T.get("palette", D.PALETTE)   # a direction carries its own palette
    font = T["font"]
    title_font = T.get("title_font", font)   # display typeface for the title
    MARGIN_X, MARGIN_Y, ROW_GAP = T["margin"], T["margin_y"], T["row_gap"]

    layout = (layout or "single").lower()
    two_col = layout not in ("single", "")
    cols = 2 if two_col else 1

    fields = [_normalise_field(f) for f in (fields or [])]
    actions = list(actions or [])
    need_header = bool(title)
    need_footer = bool(actions)

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
            p.update({"FontName": font, "FontSize": T["label_size"],
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
            p.update({"FontName": font, "FontSize": T["field_size"],
                      "ForeColor": pal["text"], "BackColor": pal["field_bg"],
                      "BackStyle": 1, "BorderStyle": 1,
                      "BorderColor": pal["field_border"]})
        # CheckBoxes carry no colour props — CreateControl rejects ForeColor on
        # a CheckBox; it uses the theme.
        if isinstance(fl.get("props"), dict):
            p.update(fl["props"])
        return p

    # --- field geometry into a Detail list ---------------------------------
    col_w = D.LABEL_W + D.GAP_LABEL + D.FIELD_W
    y = MARGIN_Y
    max_right = MARGIN_X + col_w
    detail: list[dict] = []

    def _emit_row(row_fields: list, top: int) -> int:
        nonlocal max_right
        row_h = D.ROW_H
        for ci, fl in enumerate(row_fields):
            field = str(fl.get("field", "") or fl.get("name", "") or f"Campo{len(detail)}")
            ctl_kw = str(fl.get("control", "textbox")).lower().strip()
            type_name, flags = _CONTROL_MAP.get(ctl_kw, ("TextBox", set()))
            caption = fl.get("label")
            if caption is None:
                caption = f"{field}:"
            bound = fl.get("control_source")
            if bound is None and fl.get("bind", True) is not False:
                bound = field if not _looks_unbound(field) else None

            if "memo" in flags:
                fh = D.snap(fl.get("height", D.MEMO_H))
            else:
                fh = D.snap(fl.get("height", D.ROW_H))
            if "checkbox" in flags:
                fw = D.CHECKBOX_W
            elif two_col:
                fw = D.FIELD_W
            else:
                units = float(fl.get("width_units", 1) or 1)
                fw = D.snap(D.FIELD_W * units)

            x0 = MARGIN_X + ci * (col_w + D.COL_GAP)
            x_field = x0 + D.LABEL_W + D.GAP_LABEL
            lbl_name = _unique(str(fl.get("label_name") or _ident(field, "lbl")))
            fld_name = _unique(str(fl.get("name") or _ident(field, _PREFIX.get(type_name, "ctl"))))

            detail.append({
                "role": "label", "section": _AC_DETAIL, "type_name": "Label",
                "name": lbl_name, "left": x0, "top": top,
                "width": D.LABEL_W, "height": D.ROW_H,
                "props": _label_style(caption),
            })
            detail.append({
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
            y += _emit_row(fields[r:r + cols], y) + ROW_GAP
    else:
        for fl in fields:
            y += _emit_row([fl], y) + ROW_GAP

    detail_height = D.snap(max(y - ROW_GAP + MARGIN_Y, D.ROW_STRIDE))

    if two_col:
        form_width = MARGIN_X + cols * col_w + (cols - 1) * D.COL_GAP + MARGIN_X
    else:
        form_width = max_right + MARGIN_X
    form_width = D.snap(form_width)

    # --- assemble: Detail (card behind fields) -----------------------------
    controls: list[dict] = []
    if styled and T["card"]:
        controls.append(_rect("recCard", _AC_DETAIL, 120, 120,
                              form_width - 240, detail_height - 240,
                              pal["field_bg"], pal["field_border"]))
    controls.extend(detail)

    # --- Header (accent band behind a centred title) -----------------------
    header_height = 0
    header_backcolor = T["canvas"] if styled else None
    if need_header:
        header_height = D.HEADER_H
        title_fore = pal["text"]
        title_bg = None
        if styled and T["header_band"]:
            # The band is painted TWICE: the section BackColor (fills the whole
            # document-window width, so no themed colour shows beyond the form)
            # AND a Rectangle spanning form_width (a fallback in case a theme
            # overrides the section colour). Same accent → seamless either way.
            header_backcolor = pal["accent"]
            controls.append(_rect("recHeaderBand", _AC_HEADER, 0, 0,
                                  form_width, D.HEADER_H, pal["accent"], None))
            title_fore = _WHITE
            title_bg = pal["accent"]
        tprops: dict = {"Caption": title}
        if styled:
            tprops.update({"FontName": title_font, "FontSize": T["title_size"],
                           "FontWeight": D.FONT_WEIGHT_BOLD, "ForeColor": title_fore})
            if title_bg is not None:
                # Solid same-colour label so the lint sees white-on-accent (its
                # own BackColor), not white-on-section — and it reads centred.
                tprops.update({"BackStyle": 1, "BackColor": title_bg, "BorderStyle": 0})
        # Title label must be tall enough for its font or the text clips
        # vertically (16-18pt needs ~470-540 twips, not ROW_H=300).
        title_h = max(D.ROW_H, D.line_height(T["title_size"])) if styled else D.ROW_H
        controls.append({
            "role": "title", "section": _AC_HEADER, "type_name": "Label",
            "name": _unique("lblTitle"),
            "left": MARGIN_X if not T["header_band"] else 240,
            "top": D.snap((D.HEADER_H - title_h) / 2),
            "width": max(D.LABEL_W, form_width - (2 * MARGIN_X if not T["header_band"] else 480)),
            "height": title_h, "props": tprops,
        })

    # --- Footer (action buttons) -------------------------------------------
    footer_height = 0
    if need_footer:
        footer_height = D.FOOTER_H
        n = len(actions)
        total_w = n * D.BUTTON_W + (n - 1) * D.BUTTON_GAP
        start_x = form_width - MARGIN_X - total_w
        if start_x < MARGIN_X:
            start_x = MARGIN_X
        btn_top = D.snap((D.FOOTER_H - D.BUTTON_H) / 2)
        for i, act in enumerate(actions):
            if isinstance(act, str):
                act = {"caption": act}
            caption = str(act.get("caption", f"Botón{i + 1}"))
            bname = _unique(str(act.get("name") or _ident(caption, "btn")))
            bx = D.snap(start_x + i * (D.BUTTON_W + D.BUTTON_GAP))
            on_click = act.get("on_click")
            extra = act.get("props") if isinstance(act.get("props"), dict) else None

            if styled and T["flat_buttons"]:
                primary = (i == 0)
                fill = pal["accent"] if primary else pal["field_bg"]
                txt = _WHITE if primary else pal["text"]
                border = None if primary else pal["field_border"]
                # background rect (carries the fill + optional border)
                controls.append(_rect(bname + "Bg", _AC_FOOTER, bx, btn_top,
                                      D.BUTTON_W, D.BUTTON_H, fill, border))
                # caption label: inset + solid same fill so the rect border shows
                # and the lint reads text on its own (accent/white) background.
                cprops = {"Caption": caption, "FontName": font,
                          "FontSize": T["field_size"], "FontWeight": D.FONT_WEIGHT_BOLD,
                          "ForeColor": txt, "TextAlign": 2, "BackStyle": 1,
                          "BackColor": fill, "BorderStyle": 0}
                controls.append({
                    "role": "button", "section": _AC_FOOTER, "type_name": "Label",
                    "name": bname + "Cap", "left": bx + 60,
                    "top": D.snap(btn_top + (D.BUTTON_H - D.ROW_H) / 2),
                    "width": D.BUTTON_W - 120, "height": D.ROW_H, "props": cprops,
                })
                # transparent click layer on top
                bprops: dict = {"Transparent": True, "Caption": ""}
                if on_click:
                    bprops["OnClick"] = str(on_click)
                if extra:
                    bprops.update(extra)
                controls.append({
                    "role": "button", "section": _AC_FOOTER, "type_name": "CommandButton",
                    "name": bname, "left": bx, "top": btn_top,
                    "width": D.BUTTON_W, "height": D.BUTTON_H, "tab": True,
                    "props": bprops,
                })
            else:
                bprops = {"Caption": caption}
                if on_click:
                    bprops["OnClick"] = str(on_click)
                if extra:
                    bprops.update(extra)
                controls.append({
                    "role": "button", "section": _AC_FOOTER, "type_name": "CommandButton",
                    "name": bname, "left": bx, "top": btn_top,
                    "width": D.BUTTON_W, "height": D.BUTTON_H, "tab": True,
                    "props": bprops,
                })

    return {
        "form_width": form_width,
        "detail_height": detail_height,
        "header_height": header_height,
        "header_backcolor": header_backcolor,
        "footer_height": footer_height,
        "need_header": need_header,
        "need_footer": need_footer,
        "canvas": T["canvas"] if styled else None,
        "chrome_off": T["chrome_off"],
        "controls": controls,
    }


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
    title: a form-header band with this caption.
    layout: 'single' or 'two-column'. theme: light|plain|polish|flat.

    All coordinates are computed from mcp_access.design_defaults and snapped to
    the 60-twip grid. Returns the geometry, the controls created, the tab order
    and an embedded lint of the result.
    """
    from .code import ac_create_form

    plan = _plan_layout(fields, actions, title, layout, theme)
    app = _Session.connect(db_path)

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
        if plan["chrome_off"]:
            _set_form_chrome(obj)

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

        try:
            obj.Width = plan["form_width"]
        except Exception as exc:
            log.warning("build_form: could not set form Width: %s", exc)
        _set_section(obj, _AC_DETAIL, plan["detail_height"], plan["canvas"])
        if plan["need_header"]:
            # Paint the header section with its own colour (accent for a band, the
            # canvas otherwise) so the band fills the full document-window width —
            # no themed colour bleeds past the form edge.
            _set_section(obj, _AC_HEADER, plan["header_height"],
                         plan["header_backcolor"])
        elif has_hf:
            _set_section(obj, _AC_HEADER, 0, None)
        if plan["need_footer"]:
            _set_section(obj, _AC_FOOTER, plan["footer_height"], plan["canvas"])
        elif has_hf:
            _set_section(obj, _AC_FOOTER, 0, None)

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


def _set_form_chrome(obj: Any) -> None:
    """Turn off the 'database' chrome for a cleaner, app-like form."""
    for prop, val in (("NavigationButtons", False), ("RecordSelectors", False),
                      ("ScrollBars", 0), ("DividingLines", False),
                      ("AutoCenter", True), ("CloseButton", True)):
        try:
            setattr(obj, prop, val)
        except Exception:
            pass


def _get_section(obj: Any, index: int) -> Any:
    """Resolve a form section object.

    The indexed ``Form.Section(i)`` accessor is NOT reliably late-bindable via
    pywin32 — it raises ``-2147352573 'member not found'`` for every index, which
    used to make :func:`_set_section` fail silently (so the canvas colour was
    never painted and the header/footer kept Access' oversized default heights,
    which is what produced the washed-out two-tone header band). The *named*
    section properties (``Detail`` / ``FormHeader`` / ``FormFooter``) DO bind, so
    try those first; fall back to the index for any exotic section.
    """
    prop = _SECTION_LABEL.get(index)
    if prop:
        try:
            return getattr(obj, prop)
        except Exception:
            pass
    try:
        return obj.Section(index)
    except Exception:
        return None


def _set_section(obj: Any, index: int, height: int, backcolor: Optional[int]) -> None:
    """Set a section's Height (and BackColor) defensively."""
    sec = _get_section(obj, index)
    if sec is None:
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
