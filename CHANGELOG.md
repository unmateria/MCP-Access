# Changelog

## 0.7.34 — 2026-05-05

### Fixed
- **`access_list_controls`** silently lost controls inside `Page` / `OptionGroup` containers when any earlier control in the same Page had a multi-line property block (`GUID = Begin … End`, `NameMap = Begin … End`, `ConditionalFormat = Begin … End`, etc.). The depth counter inside a control's body matched plain `Begin <Type>` but not `Property = Begin`, so the property's closing `End` was decremented without ever being incremented — the enclosing control was closed prematurely, and every control that came after it inside the Page was never enumerated. The form-level loop already handled this; the per-control loop now mirrors it.

  Visible symptom: `access_list_controls` reported a TabControl Page as a 15-line empty stub even though the Page actually contained dozens of controls. Fixed in `mcp_access/controls.py:_parse_controls`.
