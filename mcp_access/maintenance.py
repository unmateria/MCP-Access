"""
Compact and repair operations.
"""

import os
from pathlib import Path

from .core import _Session, _vbe_code_cache, _parsed_controls_cache, log


def ac_compact_repair(db_path: str) -> dict:
    """Compacts and repairs the database. Closes, compacts to temp, replaces and reopens."""
    resolved = str(Path(db_path).resolve())
    app = _Session.connect(resolved)
    original_size = os.path.getsize(resolved)

    # Close current database (keep Access alive)
    try:
        app.CloseCurrentDatabase()
    except Exception as exc:
        raise RuntimeError(f"Could not close the database for compacting: {exc}")
    _Session._db_open = None
    _Session._cm_cache.clear()
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()

    # Temp/bak paths in same directory (atomic rename)
    db_dir = os.path.dirname(resolved)
    db_name, db_ext = os.path.splitext(os.path.basename(resolved))
    tmp_path = os.path.join(db_dir, f"{db_name}_compact_tmp{db_ext}")
    bak_path = os.path.join(db_dir, f"{db_name}_compact_bak{db_ext}")

    try:
        for p in (tmp_path, bak_path):
            if os.path.exists(p):
                os.unlink(p)

        try:
            app.CompactRepair(resolved, tmp_path)
        except Exception as exc:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
            raise RuntimeError(f"Error en CompactRepair: {exc}")

        if not os.path.exists(tmp_path):
            raise RuntimeError("CompactRepair did not generate the output file")
        compacted_size = os.path.getsize(tmp_path)

        # Atomic swap: original -> .bak, tmp -> original
        os.rename(resolved, bak_path)
        try:
            os.rename(tmp_path, resolved)
        except Exception:
            os.rename(bak_path, resolved)  # rollback
            raise

        try:
            os.unlink(bak_path)
        except OSError:
            pass

    except Exception:
        # Try to reopen whatever is at the original path
        try:
            if os.path.exists(resolved):
                _Session.reopen(resolved)
        except Exception:
            pass
        raise

    # Reopen compacted database (with SHIFT to bypass AutoExec/startup)
    try:
        _Session.reopen(resolved)
    except Exception as exc:
        raise RuntimeError(f"Database compacted OK but error reopening: {exc}")

    saved = original_size - compacted_size
    return {
        "original_size": original_size,
        "compacted_size": compacted_size,
        "saved_bytes": saved,
        "saved_pct": round(saved / original_size * 100, 1) if original_size > 0 else 0,
        "status": "compacted",
    }


def ac_decompile_compact(db_path: str) -> dict:
    """Removes orphaned VBA p-code (/decompile), recompiles and compacts. Typical reduction 60-70%."""
    import subprocess, time
    resolved = str(Path(db_path).resolve())
    if not os.path.exists(resolved):
        raise FileNotFoundError(f"Database not found: {resolved}")

    original_size = os.path.getsize(resolved)

    # 1. Close COM session and release the file completely
    try:
        app = _Session.connect(resolved)
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        _Session._db_open = None
        _Session._cm_cache.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()
        try:
            app.Quit(1)  # acQuitSaveNone=1
        except Exception:
            pass
        _Session._app = None
    except Exception:
        pass  # si no habia sesion abierta, continuar igualmente

    # 2. Lanzar MSACCESS /decompile
    msaccess_candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
    ]
    msaccess = next((p for p in msaccess_candidates if os.path.exists(p)), None)
    if not msaccess:
        raise RuntimeError("MSACCESS.EXE not found in known Office 16 paths")

    proc = subprocess.Popen(
        [msaccess, resolved, "/decompile"],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    # Access opens the DB in decompiled state and stays open -- wait and kill
    time.sleep(8)
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
            capture_output=True, timeout=10,
        )
    except Exception:
        pass

    decompile_size = os.path.getsize(resolved)

    # 3. Reabrir via COM y recompilar VBA
    app2 = _Session.connect(resolved)
    try:
        app2.RunCommand(137)  # acCmdCompileAllModules = 137
    except Exception:
        pass  # compiling is not critical for the compact
    try:
        app2.CloseCurrentDatabase()
    except Exception:
        pass
    _Session._db_open = None
    _Session._cm_cache.clear()
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()

    # 4. Compact & Repair
    db_dir = os.path.dirname(resolved)
    db_name, db_ext = os.path.splitext(os.path.basename(resolved))
    tmp_path = os.path.join(db_dir, f"{db_name}_compact_tmp{db_ext}")
    bak_path = os.path.join(db_dir, f"{db_name}_compact_bak{db_ext}")
    for p in (tmp_path, bak_path):
        if os.path.exists(p):
            os.unlink(p)

    try:
        app2.CompactRepair(resolved, tmp_path)
    except Exception as exc:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise RuntimeError(f"Error en CompactRepair: {exc}")

    if not os.path.exists(tmp_path):
        raise RuntimeError("CompactRepair did not generate the output file")

    compacted_size = os.path.getsize(tmp_path)
    os.rename(resolved, bak_path)
    try:
        os.rename(tmp_path, resolved)
    except Exception:
        os.rename(bak_path, resolved)
        raise
    try:
        os.unlink(bak_path)
    except OSError:
        pass

    # Reopen (with SHIFT to bypass AutoExec/startup)
    _Session.reopen(resolved)

    saved = original_size - compacted_size
    return {
        "original_size": original_size,
        "decompile_size": decompile_size,
        "compacted_size": compacted_size,
        "saved_bytes": saved,
        "saved_pct": round(saved / original_size * 100, 1) if original_size > 0 else 0,
        "status": "decompiled_and_compacted",
    }
