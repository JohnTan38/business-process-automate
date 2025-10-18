# C:\apps\esker\app.py
import os, json, re, subprocess, sys, datetime
from pathlib import Path
from flask import Flask, request, jsonify, abort
from dotenv import load_dotenv

load_dotenv()

WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "")
PORT = int(os.getenv("PORT", "5055"))
BASE = Path(r"C:/Users/john.tan/esker/Scripts/") #amend if needed
LOGS = BASE / "logs"
DATA = BASE / "data"
LOGS.mkdir(exist_ok=True)
DATA.mkdir(exist_ok=True)

app = Flask(__name__)

def parse_vendor_triplet(text: str):
    """
    Extract 'SG80 10002345678 KLO PTE LTD' -> (company_code, vendor_number, name).
    Adjust regex to your exact patterns.
    """
    if not text:
        return None
    pat = re.compile(r'\b([A-Z]{2}\d{2})\s+(\d{8,15})\s+([^\r\n]+)')
    m = pat.search(text)
    return (m.group(1), m.group(2), m.group(3).strip()) if m else None

def log(name: str, data):
    ts = datetime.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    p = LOGS / f"{ts}_{name}.log"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def require_secret(req):
    # Expect a header "X-Webhook-Secret" that matches your .env
    given = req.headers.get("X-Webhook-Secret", "")
    if not WEBHOOK_SECRET or given != WEBHOOK_SECRET:
        abort(401, description="Unauthorized")

@app.post("/outlook/esker-vendor")
def esker_webhook():
    # 1) Auth
    require_secret(request)

    # 2) Payload
    payload = request.get_json(silent=True) or {}
    log("incoming", payload)

    # 3) Basic dedupe using InternetMessageId if provided
    internet_id = payload.get("internetMessageId") or payload.get("messageId") or ""
    if internet_id:
        seen_flag = DATA / (internet_id.replace("<","").replace(">","").replace("@","_") + ".seen")
        if seen_flag.exists():
            return jsonify({"ok": True, "deduped": True})
        seen_flag.touch()

    # 4) Extract useful fields
    subject = payload.get("subject", "")
    sender  = payload.get("from") or payload.get("sender") or {}
    sender_addr = sender.get("address") if isinstance(sender, dict) else sender
    body_text = payload.get("bodyText") or payload.get("bodyPreview") or ""
    body_html = payload.get("bodyHtml") or ""

    # Prefer bodyText; if empty, strip tags from HTML quickly
    if not body_text and body_html:
        # Rough HTML text fallback
        body_text = re.sub("<[^>]+>", " ", body_html)

    # 5) Try to parse your triplet
    triplet = parse_vendor_triplet(body_text)

    # 6) Persist a handoff file for downstream scripts
    handoff = {
        "subject": subject,
        "sender_address": sender_addr,
        "received_utc": payload.get("receivedDateTime"),
        "internet_message_id": internet_id,
        "body_text": body_text[:5000],
        "triplet": triplet
    }
    out_file = DATA / f"handoff_{datetime.datetime.utcnow().strftime('%Y%m%dT%H%M%S')}.json"
    out_file.write_text(json.dumps(handoff, ensure_ascii=False, indent=2), encoding="utf-8")

    # 7) (Optional) Kick off a worker for long-running tasks
    # subprocess.Popen([sys.executable, str(BASE / "app_ui.py"), "--json", str(out_file)])
    # run or execute "app_ui.py"
    return jsonify({"ok": True, "written": str(out_file), "parsed": triplet})

if __name__ == "__main__":
    # Listen on all interfaces so the tunnel can reach us
    ui_script = BASE / "app_ui.py"
    ui_proc = subprocess.Popen([sys.executable, str(ui_script)])
    try:
        app.run(host="0.0.0.0", port=PORT, debug=True)
    finally:
        ui_proc.terminate()
    # app.run(host="0.0.0.0", port=PORT, debug=True)

