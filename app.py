"""
app.py  –  Pattern Image Asset Processor  (no login, no OAuth)
"""
import json, os, queue, threading, time, uuid

from flask import Flask, Response, jsonify, render_template, request, send_file, session
from werkzeug.utils import secure_filename
import openpyxl

from core.engine import find_magick, parse_dim, run_job, validate_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(32))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

_jobs: dict = {}
_jobs_lock = threading.Lock()

def _job(jid):
    with _jobs_lock: return _jobs.get(jid)

@app.route("/")
def index():
    return render_template("index.html", magick_ok=bool(find_magick()))

@app.route("/template")
def download_template():
    return send_file(
        os.path.join(os.path.dirname(__file__), "static", "template.xlsx"),
        as_attachment=True,
        download_name="CDN_Links_SPT_Deliverables.xlsx")

@app.route("/api/validate", methods=["POST"])
def api_validate():
    if "file" not in request.files:
        return jsonify({"ok": False, "errors": ["No file uploaded."]}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"ok": False, "errors": ["Only .xlsx / .xlsm files accepted."]}), 400
    upload_choice = request.form.get("upload_choice", "1")
    path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_{secure_filename(f.filename)}")
    f.save(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        os.remove(path)
        return jsonify({"ok": False, "errors": [f"Cannot open file: {e}"]}), 400
    errs = validate_workbook(wb, upload_choice)
    if errs:
        os.remove(path)
        return jsonify({"ok": False, "errors": errs}), 422
    session["excel_path"] = path
    session["upload_choice"] = upload_choice
    return jsonify({"ok": True})

@app.route("/api/run", methods=["POST"])
def api_run():
    excel_path = session.get("excel_path")
    upload_choice = session.get("upload_choice", "1")
    if not excel_path or not os.path.isfile(excel_path):
        return jsonify({"error": "No validated file. Please upload again."}), 400
    data = request.get_json(silent=True) or {}
    min_size = parse_dim(str(data.get("min_size", "1000"))) or 1000
    max_size = parse_dim(str(data.get("max_size", "2000"))) or 2000
    if min_size > max_size:
        return jsonify({"error": "Minimum size cannot exceed maximum."}), 400
    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "queue": q,
                         "result": None, "created": time.time()}
    def worker():
        def cb(cur, tot, fn):
            q.put({"type": "progress", "current": cur, "total": tot, "file": fn})
        result = run_job(excel_path, upload_choice, min_size, max_size, progress_cb=cb)
        with _jobs_lock:
            _jobs[job_id]["result"] = result
            _jobs[job_id]["status"] = "done" if not result.get("error") else "error"
        q.put({"type": "done", "result": result})
    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})

@app.route("/api/stream/<job_id>")
def api_stream(job_id):
    def generate():
        job = _job(job_id)
        if not job:
            yield f"data: {json.dumps({'type':'error','message':'Job not found'})}\n\n"
            return
        q = job["queue"]
        while True:
            try:
                msg = q.get(timeout=30)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg.get("type") in ("done", "error"): break
            except queue.Empty:
                yield 'data: {"type":"heartbeat"}\n\n'
    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/api/download/<job_id>")
def api_download(job_id):
    job = _job(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Job not ready."}), 404
    zp = job["result"].get("zip_path", "")
    if not zp or not os.path.isfile(zp):
        return jsonify({"error": "ZIP not found."}), 404
    return send_file(zp, mimetype="application/zip",
                     as_attachment=True, download_name="processed_images.zip")

def _gc():
    while True:
        time.sleep(3600)
        cutoff = time.time() - 7200
        with _jobs_lock:
            for jid in [k for k, v in _jobs.items() if v.get("created", 0) < cutoff]:
                del _jobs[jid]
threading.Thread(target=_gc, daemon=True).start()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
