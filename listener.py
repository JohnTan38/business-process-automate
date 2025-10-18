# C:\apps\esker\listener.py
import sys
import time
import json
import re
import shutil
import tempfile
import subprocess
from pathlib import Path

import pythoncom
import win32com.client

# --- CONFIG ---
PYTHON_EXE = sys.executable  # uses current venv/interpreter
APP_PY     = r"C:/Users/john.tan/esker/Scripts/app.py" # <-- adjust if needed
KEYWORDS   = ["esker vendor email"]         # lower-case match

def subject_hit(subj: str) -> bool:
    s = (subj or "").lower()
    return any(k in s for k in KEYWORDS)

def write_temp_json(data: dict) -> Path:
    tmp_dir = Path(tempfile.gettempdir())
    p = tmp_dir / f"esker_{int(time.time())}.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    downloads = Path(r"C:/Users/john.tan/Downloads/")
    downloads.mkdir(parents=True, exist_ok=True)
    shutil.copy2(p, downloads / p.name)  # keep a copy available for other scripts
    return p


def get_outlook_namespace(retries: int = 1):
    """Return the Outlook MAPI namespace. Use EnsureDispatch where possible.

    On some systems win32com.gen_py caches a broken generated file without
    CLSIDToClassMap/CLSIDToPackageMap which causes AttributeError. If that happens, 
    remove the offending gen_py module and retry once.
    """
    try:
        # Prefer EnsureDispatch which generates early-binding helpers when needed
        app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        return app.GetNamespace("MAPI")
    except AttributeError as e:
        # Known failure modes: generated module missing CLSIDToClassMap/CLSIDToPackageMap
        if retries <= 0:
            # Final fallback: use basic Dispatch (late-binding, no cache)
            try:
                app = win32com.client.Dispatch("Outlook.Application")
                return app.GetNamespace("MAPI")
            except Exception:
                raise

        # Log the error and attempt cleanup
        err_msg = f"AttributeError in COM initialization: {e}"
        err_path = Path(tempfile.gettempdir()) / "esker_com_cleanup.log"
        with err_path.open("a", encoding="utf-8") as f:
            f.write(f"{time.ctime()} {err_msg}\n")

        # Try to identify and remove the bad gen_py module
        try:
            # Clear the gencache entirely - more aggressive approach
            try:
                win32com.client.gencache.is_readonly = False
                # Try to rebuild the entire cache
                win32com.client.gencache.Rebuild()
            except Exception as rebuild_err:
                # If rebuild fails, manually remove files
                try:
                    import importlib
                    gen_py = importlib.import_module("win32com.gen_py")
                    # Remove all Outlook typelib variants
                    for p in gen_py.__path__:
                        pth = Path(p)
                        removed_files = []
                        for pattern in ['00062FFF-0000-0000-C000-000000000046*', '*Outlook*']:
                            for f in pth.glob(pattern):
                                try:
                                    if f.is_file():
                                        f.unlink()
                                        removed_files.append(str(f))
                                except Exception:
                                    pass
                        # Also remove __pycache__ folders that might contain stale bytecode
                        for cache_dir in pth.glob('__pycache__'):
                            try:
                                if cache_dir.is_dir():
                                    shutil.rmtree(cache_dir)
                            except Exception:
                                pass
                        
                        with err_path.open("a", encoding="utf-8") as f:
                            f.write(f"{time.ctime()} Removed files: {removed_files}\n")
                except Exception as cleanup_err:
                    with err_path.open("a", encoding="utf-8") as f:
                        f.write(f"{time.ctime()} Cleanup error: {cleanup_err}\n")
        except Exception:
            # If we cannot clean, re-raise original error
            raise

        # Retry once
        return get_outlook_namespace(retries=retries - 1)

class InboxEvents:
    # IMPORTANT: no __init__(self, items) here! WithEvents calls with no args.
    def OnItemAdd(self, item):
        try:
            # 43 = olMailItem
            if getattr(item, "Class", None) != 43:
                return

            subj = getattr(item, "Subject", "")
            if not subject_hit(subj):
                return

            sender = getattr(item, "SenderEmailAddress", "")
            body   = getattr(item, "Body", "")
            recv   = item.ReceivedTime  # COM Date
            try:
                recv_str = recv.strftime("%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                recv_str = str(recv)

            payload = {
                "subject": subj,
                "sender_address": sender,
                "received_utc": recv_str,
                "body": body,
                "entry_id": getattr(item, "EntryID", ""),
            }

            json_path = write_temp_json(payload)
            # Launch your worker (non-blocking)
            subprocess.Popen([PYTHON_EXE, APP_PY, "--email_json", str(json_path)])

            # Optional: quick console note
            print(f"[listener] Triggered for: {subj} -> {json_path}")

        except Exception as e:
            # Minimal error logging to temp
            err_path = Path(tempfile.gettempdir()) / "esker_listener_errors.log"
            with err_path.open("a", encoding="utf-8") as f:
                f.write(f"{time.ctime()} OnItemAdd error: {e}\n")

def main():
    # Get Outlook namespace and Inbox folder
    outlook = get_outlook_namespace()
    inbox   = outlook.GetDefaultFolder(6)  # 6 = olFolderInbox

    # Keep strong references so events stay alive
    items = inbox.Items
    # NOTE: Do not Sort/Restrict here; ItemAdd fires only on the default Items collection.
    # If you need filtering, do it in OnItemAdd.

    # Hook events
    handler = win32com.client.WithEvents(items, InboxEvents)

    print("Listening for new Inbox items… (Ctrl+C to exit)")
    # Pump COM messages
    while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)

if __name__ == "__main__":
    main()
