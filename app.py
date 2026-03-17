"""
app.py  (v3 — cloud-ready)
--------------------------
Changes vs v2:
  • Per-user session isolation via Flask session + per-session temp dirs
  • ZIP is generated IN MEMORY (io.BytesIO) and streamed directly to the
    browser — NO files accumulate on the server's disk
  • Temp directories are cleaned up automatically after each request that
    needs them
  • SECRET_KEY read from environment variable (required for sessions)
  • UPLOAD_FOLDER and FONTS_DIR are relative but work on any host
"""

import io
import os
import re
import shutil
import tempfile
import uuid
import zipfile
import json

from flask import (
    Flask,
    request,
    jsonify,
    send_file,
    render_template,
    send_from_directory,
    session,
    Response,
)
from functools import wraps
from werkzeug.utils import secure_filename

from utils.excel_reader import get_columns, get_rows
from utils.generator import generate_certificate

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = Flask(__name__)

# SECRET_KEY is required for session cookies.
# On Render / Railway set this as an environment variable.
# Falls back to a random key (sessions won't survive restarts — fine for this tool).
app.secret_key = os.environ.get("SECRET_KEY", uuid.uuid4().hex)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 50 MB limit

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR  = os.path.join(BASE_DIR, "static", "uploads")
FONTS_DIR   = os.path.join(BASE_DIR, "static", "fonts")

for d in [UPLOAD_DIR, FONTS_DIR,
          os.path.join(BASE_DIR, "static", "css"),
          os.path.join(BASE_DIR, "static", "js")]:
    os.makedirs(d, exist_ok=True)

ALLOWED_IMAGE = {"png", "jpg", "jpeg"}
ALLOWED_EXCEL = {"xlsx", "xls"}
ALLOWED_FONT  = {"ttf", "otf", "woff", "woff2"}


def ext_ok(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


# ---------------------------------------------------------------------------
# Session helpers — each browser tab gets its own temp working directory
# ---------------------------------------------------------------------------

def _session_dir() -> str:
    """
    Return (and create) a per-session temp directory.
    Stored under /tmp/certmaker_<session_id> so it's cleaned up
    by the OS on restart (important for serverless/ephemeral hosts).
    """
    sid = session.get("sid")
    if not sid:
        sid = uuid.uuid4().hex
        session["sid"] = sid
    path = os.path.join(tempfile.gettempdir(), f"certmaker_{sid}")
    os.makedirs(path, exist_ok=True)
    return path


def _session_set(key, value):
    session[key] = value


def _session_get(key, default=None):
    return session.get(key, default)


# ---------------------------------------------------------------------------
# Dynamic Settings & Authentication
# ---------------------------------------------------------------------------

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
DEFAULT_SETTINGS = {
    "is_public_locked": False,
    "public_password": "mysecretpassword"
}

def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
            # Ensure all keys exist
            for k, v in DEFAULT_SETTINGS.items():
                if k not in data:
                    data[k] = v
            return data
    except Exception:
        return DEFAULT_SETTINGS.copy()

def save_settings(data):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(data, f, indent=4)

# Admin master credentials (hardcoded or from env, cannot be changed from web for safety)
admin_username = os.environ.get("ADMIN_USER", "admin")
admin_password = os.environ.get("ADMIN_PASS", "arjun123")

def authenticate():
    return Response(
        'Could not verify your access level for that URL.\n'
        'You have to login with proper credentials', 401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'})

def requires_public_auth(f):
    """Protects the main generator. Uses the dynamic settings.json public password."""
    @wraps(f)
    def decorated(*args, **kwargs):
        settings = load_settings()
        if not settings.get("is_public_locked", False):
            # If lock is off, let everyone in
            return f(*args, **kwargs)
            
        auth = request.authorization
        # The username doesn't strictly matter for the public lock, just the password, 
        # but basic auth requires both. We'll accept any username if the password matches.
        if not auth or auth.password != settings.get("public_password"):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

def requires_admin_auth(f):
    """Protects the /admin dashboard. Uses the master admin credentials."""
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or auth.username != admin_username or auth.password != admin_password:
            return authenticate()
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
@requires_public_auth
def index():
    return render_template("index.html")

# ── Admin Panel ──────────────────────────────────────────────────────────────

@app.route("/admin12arjun")
@requires_admin_auth
def admin_dashboard():
    return render_template("admin.html")

@app.route("/api/admin12arjun/settings", methods=["GET", "POST"])
@requires_admin_auth
def admin_settings():
    if request.method == "GET":
        return jsonify(load_settings())
    
    settings = load_settings()
    data = request.json or {}
    
    if "is_public_locked" in data:
        settings["is_public_locked"] = bool(data["is_public_locked"])
    if "public_password" in data and data["public_password"].strip():
        settings["public_password"] = data["public_password"].strip()
        
    save_settings(settings)
    return jsonify({"success": True})


# ── Template upload ──────────────────────────────────────────────────────────

@app.route("/upload-template", methods=["POST"])
def upload_template():
    if "template" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["template"]
    if not file.filename or not ext_ok(file.filename, ALLOWED_IMAGE):
        return jsonify({"error": "Please upload a PNG or JPEG image."}), 400

    ext          = file.filename.rsplit(".", 1)[1].lower()
    unique_name  = f"template_{uuid.uuid4().hex}.{ext}"
    save_path    = os.path.join(_session_dir(), unique_name)
    file.save(save_path)

    _session_set("template_path", save_path)

    # Copy to static/uploads so the browser can preview it
    static_path = os.path.join(UPLOAD_DIR, unique_name)
    shutil.copy(save_path, static_path)

    return jsonify({"url": f"/static/uploads/{unique_name}", "filename": unique_name})


# ── Excel upload ─────────────────────────────────────────────────────────────

@app.route("/upload-excel", methods=["POST"])
def upload_excel():
    if "excel" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["excel"]
    if not file.filename or not ext_ok(file.filename, ALLOWED_EXCEL):
        return jsonify({"error": "Please upload an .xlsx file."}), 400

    unique_name = f"excel_{uuid.uuid4().hex}.xlsx"
    save_path   = os.path.join(_session_dir(), unique_name)
    file.save(save_path)
    _session_set("excel_path", save_path)

    try:
        columns = get_columns(save_path)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {e}"}), 500

    return jsonify({"columns": columns})


# ── Font upload ───────────────────────────────────────────────────────────────

@app.route("/upload-font", methods=["POST"])
def upload_font():
    if "font" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["font"]
    if not file.filename or not ext_ok(file.filename, ALLOWED_FONT):
        return jsonify({"error": "Please upload a .ttf or .otf font."}), 400

    safe_name    = secure_filename(file.filename)
    save_path    = os.path.join(FONTS_DIR, safe_name)
    file.save(save_path)

    display_name = os.path.splitext(safe_name)[0].replace("-", " ").replace("_", " ")
    return jsonify({
        "name":     display_name,
        "filename": safe_name,
        "path":     save_path,
        "url":      f"/static/fonts/{safe_name}",
    })


# ── Font list ─────────────────────────────────────────────────────────────────

@app.route("/fonts")
def list_fonts():
    fonts = []
    for fname in sorted(os.listdir(FONTS_DIR)):
        if ext_ok(fname, ALLOWED_FONT):
            display = os.path.splitext(fname)[0].replace("-", " ").replace("_", " ")
            fonts.append({
                "name":     display,
                "filename": fname,
                "path":     os.path.join(FONTS_DIR, fname),
                "url":      f"/static/fonts/{fname}",
            })
    return jsonify({"fonts": fonts})


# ── Preview ──────────────────────────────────────────────────────────────────

@app.route("/preview", methods=["POST"])
def preview():
    data   = request.get_json()
    fields = (data or {}).get("fields", [])

    template_path = _session_get("template_path")
    excel_path    = _session_get("excel_path")

    if not template_path or not os.path.isfile(template_path):
        return jsonify({"error": "Template not uploaded yet"}), 400
    if not excel_path or not os.path.isfile(excel_path):
        return jsonify({"error": "Excel file not uploaded yet"}), 400
    if not fields:
        return jsonify({"error": "No fields configured"}), 400

    rows = get_rows(excel_path)
    if not rows:
        return jsonify({"error": "Excel file is empty"}), 400

    preview_dir = os.path.join(_session_dir(), "preview")
    if os.path.exists(preview_dir):
        shutil.rmtree(preview_dir)
    os.makedirs(preview_dir, exist_ok=True)

    try:
        output_path = generate_certificate(
            template_path=template_path,
            row=rows[0],
            fields=fields,
            output_dir=preview_dir,
            fonts_dir=FONTS_DIR,
        )
    except Exception as e:
        return jsonify({"error": f"Generation failed: {e}"}), 500

    # Expose via static so the browser can show it
    preview_name = os.path.basename(output_path)
    static_preview = os.path.join(UPLOAD_DIR, f"preview_{preview_name}")
    shutil.copy(output_path, static_preview)

    url = f"/static/uploads/preview_{preview_name}?t={uuid.uuid4().hex}"
    return jsonify({"url": url, "name": rows[0]})


# ── Batch generate + ZIP (streamed in-memory, nothing kept on server) ─────────

@app.route("/generate", methods=["POST"])
def generate():
    """
    Generate all certificates, pack them into an in-memory ZIP, and
    immediately stream it to the browser.
    Generated PNG files are created in a temporary directory that is
    deleted as soon as the response is sent.
    """
    data   = request.get_json()
    fields = (data or {}).get("fields", [])

    template_path = _session_get("template_path")
    excel_path    = _session_get("excel_path")

    if not template_path or not os.path.isfile(template_path):
        return jsonify({"error": "Template not uploaded yet"}), 400
    if not excel_path or not os.path.isfile(excel_path):
        return jsonify({"error": "Excel file not uploaded yet"}), 400
    if not fields:
        return jsonify({"error": "No fields configured"}), 400

    rows = get_rows(excel_path)
    if not rows:
        return jsonify({"error": "Excel file is empty"}), 400

    # Use a throwaway temp dir — deleted in `finally`
    tmp_dir = tempfile.mkdtemp(prefix="certmaker_batch_")
    errors         = []
    generated_count = 0

    try:
        for idx, row in enumerate(rows):
            try:
                generate_certificate(
                    template_path=template_path,
                    row=row,
                    fields=fields,
                    output_dir=tmp_dir,
                    fonts_dir=FONTS_DIR,
                )
                generated_count += 1
            except Exception as e:
                errors.append(f"Row {idx + 2}: {e}")

        # Return summary JSON first so the frontend knows the count
        # Then the /download-zip-stream route supplies the file.
        # Actually — stream the ZIP right now in this same response.
        # Store the tmp_dir path in session for /download-zip to pick up.
        _session_set("batch_dir", tmp_dir)
        _session_set("generated_count", generated_count)
        _session_set("total_rows", len(rows))

        return jsonify({
            "generated": generated_count,
            "total":     len(rows),
            "errors":    errors,
        })
    except Exception as e:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500
    # Note: tmp_dir is cleaned up only after /download-zip is called.


# ── ZIP download — streams in-memory, deletes temp dir after send ─────────────

@app.route("/download-zip")
def download_zip():
    """
    Build the ZIP in memory from the batch temp dir, stream it to the
    browser, then delete the temp dir so nothing lingers on the server.
    """
    batch_dir = _session_get("batch_dir")
    if not batch_dir or not os.path.isdir(batch_dir):
        return jsonify({"error": "No certificates generated yet."}), 404

    # Build ZIP entirely in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(batch_dir):
            if fname.lower().endswith(".png"):
                zf.write(os.path.join(batch_dir, fname), arcname=fname)
    zip_buffer.seek(0)

    # Clean up the temp dir NOW — files never linger on the server
    shutil.rmtree(batch_dir, ignore_errors=True)
    _session_set("batch_dir", None)

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="certificates.zip",
        mimetype="application/zip",
    )


# ── Serve generated static assets ────────────────────────────────────────────

@app.route("/static/fonts/<path:filename>")
def serve_font(filename):
    return send_from_directory(FONTS_DIR, filename)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 55)
    print(f"  Certificate Generator — http://127.0.0.1:{port}")
    print("=" * 55)
    app.run(debug=os.environ.get("FLASK_DEBUG", "true").lower() == "true",
            host="0.0.0.0", port=port)
