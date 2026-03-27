"""
Tool dispatcher: routes tool names to implementation functions.
"""

import json
import traceback

from .core import _Session, log

from .vbe import (
    ac_vbe_get_lines, ac_vbe_get_proc, ac_vbe_module_info,
    ac_vbe_replace_lines, ac_vbe_find, ac_vbe_search_all,
    ac_search_queries, ac_vbe_replace_proc, ac_vbe_patch_proc,
    ac_vbe_append, ac_find_usages,
)
from .controls import (
    ac_list_controls, ac_get_control, ac_create_control,
    ac_delete_control, ac_set_control_props, ac_set_form_property,
    ac_get_form_property, ac_set_multiple_controls,
    ac_export_text, ac_import_text,
)
from .code import (
    ac_list_objects, ac_get_code, ac_set_code, ac_delete_object,
    ac_create_form, ac_export_structure,
)
from .database import (
    ac_create_database, ac_create_table, ac_alter_table, ac_table_info,
)
from .sql import ac_execute_sql, ac_execute_batch, ac_manage_query
from .properties import (
    ac_get_db_property, ac_set_db_property,
    ac_get_field_properties, ac_set_field_property,
    ac_list_startup_options,
)
from .relations import (
    ac_list_linked_tables, ac_relink_table,
    ac_list_relationships, ac_create_relationship, ac_delete_relationship,
    ac_list_references, ac_manage_reference,
    ac_list_indexes, ac_manage_index,
)
from .maintenance import ac_compact_repair, ac_decompile_compact
from .vba_exec import ac_run_macro, ac_run_vba, ac_eval_vba
from .compile import ac_compile_vba
from .export import ac_output_report, ac_transfer_data
from .ui import ac_screenshot, ac_ui_click, ac_ui_type
from .tips import ac_tips


def call_tool_sync(name: str, arguments: dict) -> str:
    """Synchronous tool dispatcher -- runs in a thread to avoid blocking the event loop."""
    try:
        if name == "access_list_objects":
            result = ac_list_objects(
                arguments["db_path"],
                arguments.get("object_type", "all"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_get_code":
            text = ac_get_code(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )

        elif name == "access_set_code":
            text = ac_set_code(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["code"],
            )

        elif name == "access_execute_sql":
            result = ac_execute_sql(
                arguments["db_path"],
                arguments["sql"],
                int(arguments.get("limit", 500)),
                bool(arguments.get("confirm_destructive", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_table_info":
            result = ac_table_info(arguments["db_path"], arguments["table_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_export_structure":
            text = ac_export_structure(
                arguments["db_path"],
                arguments.get("output_path"),
            )

        elif name == "access_close":
            _Session.quit()
            text = "Access session closed successfully."

        # -- VBE line-level -----------------------------------------------
        elif name == "access_vbe_get_lines":
            text = ac_vbe_get_lines(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                int(arguments["start_line"]),
                int(arguments["count"]),
            )

        elif name == "access_vbe_get_proc":
            result = ac_vbe_get_proc(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["proc_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_module_info":
            result = ac_vbe_module_info(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_replace_lines":
            ops = arguments.get("operations")
            text = ac_vbe_replace_lines(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                int(arguments.get("start_line", 0)),
                int(arguments.get("count", 0)),
                arguments.get("new_code", ""),
                operations=ops,
            )

        elif name == "access_vbe_find":
            result = ac_vbe_find(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                bool(arguments.get("use_regex", False)),
                proc_name=arguments.get("proc_name", ""),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_search_all":
            result = ac_vbe_search_all(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 100)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_search_queries":
            result = ac_search_queries(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 100)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_replace_proc":
            text = ac_vbe_replace_proc(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["proc_name"],
                arguments["new_code"],
            )

        elif name == "access_vbe_patch_proc":
            text = ac_vbe_patch_proc(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["proc_name"],
                arguments["patches"],
            )

        elif name == "access_vbe_append":
            text = ac_vbe_append(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["new_code"],
            )

        # -- Control-level ------------------------------------------------
        elif name == "access_list_controls":
            result = ac_list_controls(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_get_control":
            result = ac_get_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_create_control":
            result = ac_create_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_type"],
                dict(arguments.get("props", {})),
                class_name=arguments.get("class_name"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_delete_control":
            text = ac_delete_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
            )

        elif name == "access_export_text":
            result = ac_export_text(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["output_path"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_import_text":
            result = ac_import_text(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["input_path"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_control_props":
            result = ac_set_control_props(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
                dict(arguments.get("props", {})),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_form_property":
            result = ac_set_form_property(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                dict(arguments.get("props", {})),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Database properties ------------------------------------------
        elif name == "access_get_db_property":
            result = ac_get_db_property(arguments["db_path"], arguments["name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_db_property":
            result = ac_set_db_property(
                arguments["db_path"],
                arguments["name"],
                arguments["value"],
                arguments.get("prop_type"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Linked tables ------------------------------------------------
        elif name == "access_list_linked_tables":
            result = ac_list_linked_tables(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_relink_table":
            result = ac_relink_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["new_connect"],
                bool(arguments.get("relink_all", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Relationships ------------------------------------------------
        elif name == "access_list_relationships":
            result = ac_list_relationships(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_create_relationship":
            result = ac_create_relationship(
                arguments["db_path"],
                arguments["name"],
                arguments["table"],
                arguments["foreign_table"],
                arguments["fields"],
                int(arguments.get("attributes", 0)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_delete_relationship":
            result = ac_delete_relationship(
                arguments["db_path"],
                arguments["name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- VBA References -----------------------------------------------
        elif name == "access_list_references":
            result = ac_list_references(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_manage_reference":
            result = ac_manage_reference(
                arguments["db_path"],
                arguments["action"],
                name=arguments.get("name"),
                path=arguments.get("path"),
                guid=arguments.get("guid"),
                major=int(arguments.get("major", 0)),
                minor=int(arguments.get("minor", 0)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Compact & Repair ---------------------------------------------
        elif name == "access_compact_repair":
            result = ac_compact_repair(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_decompile_compact":
            result = ac_decompile_compact(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Query management ---------------------------------------------
        elif name == "access_manage_query":
            result = ac_manage_query(
                arguments["db_path"],
                arguments["action"],
                arguments["query_name"],
                sql=arguments.get("sql"),
                new_name=arguments.get("new_name"),
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Indexes ------------------------------------------------------
        elif name == "access_list_indexes":
            result = ac_list_indexes(arguments["db_path"], arguments["table_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_manage_index":
            result = ac_manage_index(
                arguments["db_path"],
                arguments["table_name"],
                arguments["action"],
                arguments["index_name"],
                fields=arguments.get("fields"),
                primary=bool(arguments.get("primary", False)),
                unique=bool(arguments.get("unique", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Compile VBA --------------------------------------------------
        elif name == "access_compile_vba":
            result = ac_compile_vba(arguments["db_path"], timeout=arguments.get("timeout"))
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Run macro ----------------------------------------------------
        elif name == "access_run_macro":
            result = ac_run_macro(arguments["db_path"], arguments["macro_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Output report ------------------------------------------------
        elif name == "access_output_report":
            result = ac_output_report(
                arguments["db_path"],
                arguments["report_name"],
                output_path=arguments.get("output_path"),
                fmt=arguments.get("format", "pdf"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Transfer data ------------------------------------------------
        elif name == "access_transfer_data":
            result = ac_transfer_data(
                arguments["db_path"],
                arguments["action"],
                arguments["file_path"],
                arguments["table_name"],
                has_headers=bool(arguments.get("has_headers", True)),
                file_type=arguments.get("file_type", "xlsx"),
                range_=arguments.get("range"),
                spec_name=arguments.get("spec_name"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Field properties ---------------------------------------------
        elif name == "access_get_field_properties":
            result = ac_get_field_properties(
                arguments["db_path"],
                arguments["table_name"],
                arguments["field_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_field_property":
            result = ac_set_field_property(
                arguments["db_path"],
                arguments["table_name"],
                arguments["field_name"],
                arguments["property_name"],
                arguments["value"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Startup options ----------------------------------------------
        elif name == "access_list_startup_options":
            result = ac_list_startup_options(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Create database ----------------------------------------------
        elif name == "access_create_database":
            result = ac_create_database(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Create table via DAO -----------------------------------------
        elif name == "access_create_table":
            result = ac_create_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["fields"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Alter table via DAO ------------------------------------------
        elif name == "access_alter_table":
            result = ac_alter_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["action"],
                arguments["field_name"],
                new_name=arguments.get("new_name"),
                field_type=arguments.get("field_type", "text"),
                size=int(arguments.get("size", 0)),
                required=bool(arguments.get("required", False)),
                default=arguments.get("default"),
                description=arguments.get("description"),
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Create form --------------------------------------------------
        elif name == "access_create_form":
            result = ac_create_form(
                arguments["db_path"],
                arguments["form_name"],
                has_header=bool(arguments.get("has_header", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Delete object ------------------------------------------------
        elif name == "access_delete_object":
            result = ac_delete_object(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Run VBA ------------------------------------------------------
        elif name == "access_run_vba":
            result = ac_run_vba(
                arguments["db_path"],
                arguments["procedure"],
                args=arguments.get("args"),
                timeout=arguments.get("timeout"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Eval VBA -----------------------------------------------------
        elif name == "access_eval_vba":
            result = ac_eval_vba(
                arguments["db_path"],
                arguments["expression"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Find usages --------------------------------------------------
        elif name == "access_find_usages":
            result = ac_find_usages(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 200)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Batch SQL ----------------------------------------------------
        elif name == "access_execute_batch":
            result = ac_execute_batch(
                arguments["db_path"],
                arguments["statements"],
                stop_on_error=bool(arguments.get("stop_on_error", True)),
                confirm_destructive=bool(arguments.get("confirm_destructive", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Get form/report property -------------------------------------
        elif name == "access_get_form_property":
            result = ac_get_form_property(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                property_names=arguments.get("property_names"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Set multiple controls ----------------------------------------
        elif name == "access_set_multiple_controls":
            result = ac_set_multiple_controls(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["controls"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Tips ---------------------------------------------------------
        elif name == "access_tips":
            result = ac_tips(arguments.get("topic", ""))
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # -- Screenshot + UI Automation -----------------------------------
        elif name == "access_screenshot":
            result = ac_screenshot(
                arguments["db_path"],
                object_type=arguments.get("object_type", ""),
                object_name=arguments.get("object_name", ""),
                output_path=arguments.get("output_path", ""),
                wait_ms=int(arguments.get("wait_ms", 300)),
                max_width=int(arguments.get("max_width", 1920)),
                open_timeout_sec=int(arguments.get("open_timeout_sec", 30)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_ui_click":
            result = ac_ui_click(
                arguments["db_path"],
                x=int(arguments["x"]),
                y=int(arguments["y"]),
                image_width=int(arguments["image_width"]),
                click_type=arguments.get("click_type", "left"),
                wait_after_ms=int(arguments.get("wait_after_ms", 200)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_ui_type":
            result = ac_ui_type(
                arguments["db_path"],
                text=arguments.get("text", ""),
                key=arguments.get("key", ""),
                modifiers=arguments.get("modifiers", ""),
                wait_after_ms=int(arguments.get("wait_after_ms", 100)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        else:
            text = f"ERROR: unknown tool '{name}'"

    except Exception as exc:
        log.error("Error en %s: %s", name, exc, exc_info=True)

        # Build detailed error message for the LLM
        tb_lines = traceback.format_exc().splitlines()

        # Create safe representation of arguments (hide full code)
        safe_args_display = {}
        for k, v in arguments.items():
            if k == "code":
                safe_args_display[k] = f"<VBA code provided: length {len(v)} chars>"
            else:
                safe_args_display[k] = v

        error_msg = (
            f"ERROR in tool '{name}'\n"
            f"Type: {type(exc).__name__}\n"
            f"Message: {exc}\n\n"
            f"Arguments received:\n{json.dumps(safe_args_display, indent=2, ensure_ascii=False)}\n\n"
            f"Stack trace (last 5 lines):\n" + "\n".join(tb_lines[-5:])
        )
        text = error_msg

    log.info("<<< %s  OK", name)
    return text
