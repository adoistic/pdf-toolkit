#!/usr/bin/env python3
"""
app.py - Flask web server for PDF Toolkit.

Run: python app.py
Then open http://127.0.0.1:5000 in your browser.

Thread safety: Flask runs with threaded=True. All fitz (PyMuPDF)
calls are serialised via editor.lock. Batch processing uses
subprocesses via _run_one.py (each with its own fitz instance).
"""

import fitz
import subprocess
import sys
import os
import json
import threading
import queue
import time
import gc
import webbrowser
from pathlib import Path
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import (
    Flask, render_template, request, jsonify,
    Response, stream_with_context,
)

import license as lic

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.config["LICENSE_VALID"] = False

SCRIPT_DIR = Path(__file__).parent
HELPER = str(SCRIPT_DIR / "_run_one.py")
TIMEOUT = 300  # seconds per file


# ── Folder State ────────────────────────────────────────────────────

class FolderState:
    def __init__(self):
        self.lock = threading.Lock()
        self.folder_path = None
        self.files = []            # [{name, path, size, size_str}, ...]
        self.processed = set()     # filenames that have been saved/skipped

    def load_folder(self, path):
        with self.lock:
            self.folder_path = path
            self.processed = set()
            self.files = []
            p = Path(path)
            for pdf in sorted(p.glob("*.pdf")):
                sz = pdf.stat().st_size
                self.files.append({
                    "name": pdf.name,
                    "path": str(pdf),
                    "size": sz,
                    "size_str": human_size(sz),
                })

    def remaining(self):
        with self.lock:
            return [f for f in self.files if f["name"] not in self.processed]

    def mark_processed(self, name):
        with self.lock:
            self.processed.add(name)

    def find_file(self, name):
        with self.lock:
            for f in self.files:
                if f["name"] == name:
                    return f
        return None


folder = FolderState()


# ── Batch Processing State ──────────────────────────────────────────

class BatchState:
    def __init__(self):
        self.running = False
        self.cancelled = False
        self.proc = None
        self.log_queue = queue.Queue()
        self.total = 0
        self.current = 0

    def reset(self):
        self.running = False
        self.cancelled = False
        self.proc = None
        self.total = 0
        self.current = 0
        # drain queue
        while not self.log_queue.empty():
            try:
                self.log_queue.get_nowait()
            except queue.Empty:
                break


batch = BatchState()


# ── Page Editor State ────────────────────────────────────────────────

class EditorState:
    def __init__(self):
        self.lock = threading.Lock()
        self.doc = None
        self.path = None
        self.modified = False
        self.page_sizes = []
        self.thumb_cache = OrderedDict()
        self.cache_max = 200
        self.generation = 0  # incremented on open/close/modify

    def close(self):
        if self.doc:
            self.doc.close()
        self.doc = None
        self.path = None
        self.modified = False
        self.page_sizes = []
        self.thumb_cache.clear()
        self.generation += 1
        fitz.TOOLS.store_shrink(100)
        gc.collect()

    def invalidate_cache(self):
        self.thumb_cache.clear()
        self.generation += 1

    def get_thumb(self, page_idx, width=180):
        key = (page_idx, width)
        if key in self.thumb_cache:
            self.thumb_cache.move_to_end(key)
            return self.thumb_cache[key]

        page = self.doc[page_idx]
        zoom = width / page.rect.width
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("jpeg", jpg_quality=65)
        pix = None  # release native pixmap memory immediately

        self.thumb_cache[key] = img_bytes
        if len(self.thumb_cache) > self.cache_max:
            self.thumb_cache.popitem(last=False)

        return img_bytes


editor = EditorState()


# ── Helpers ──────────────────────────────────────────────────────────

def human_size(nbytes):
    for unit in ("B", "KB", "MB", "GB"):
        if abs(nbytes) < 1024:
            return f"{nbytes:.1f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} TB"


def _run_toc_bg(file_path, toc_out_dir, h1_only, count_toc):
    """Background wrapper for _run_toc — logs result to console."""
    try:
        result, log = _run_toc(file_path, toc_out_dir, h1_only, count_toc)
        name = os.path.basename(file_path)
        if result == "success":
            print(f"  TOC added: {name}")
        else:
            print(f"  TOC {result}: {name}")
            if log:
                print(f"    {log[:200]}")
    except Exception as e:
        print(f"  TOC error for {os.path.basename(file_path)}: {e}")


def _run_toc(file_path, toc_out_dir, h1_only, count_toc):
    """Run the TOC subprocess. Returns (toc_result, toc_log)."""
    os.makedirs(toc_out_dir, exist_ok=True)
    base = os.path.basename(file_path)
    toc_dst = os.path.join(toc_out_dir, base)

    cmd = [sys.executable, "-u", HELPER, "toc", file_path, toc_dst]
    if h1_only:
        cmd.append("--h1-only")
    if not count_toc:
        cmd.append("--no-count-toc")

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True,
            timeout=TIMEOUT, env=env,
        )
        toc_result = "failed"
        toc_log = []
        for line in proc.stdout.splitlines():
            if line.strip().startswith("__RESULT__:"):
                toc_result = line.strip().split(":", 1)[1]
            else:
                toc_log.append(line)
        if proc.stderr:
            toc_log.append(proc.stderr)
        return toc_result, "\n".join(toc_log).strip()
    except Exception as e:
        return "failed", str(e)


# ── License Gate ─────────────────────────────────────────────────────

LICENSE_EXEMPT = ("/", "/static/", "/api/license/")


@app.before_request
def license_gate():
    """Block all API routes if the app is not licensed."""
    path = request.path
    if any(path.startswith(p) for p in LICENSE_EXEMPT):
        return None
    if not app.config.get("LICENSE_VALID"):
        return jsonify(error="unlicensed"), 403


@app.route("/api/license/status")
def license_status():
    """Return current license state for the frontend."""
    valid, message, needs_key = lic.check_license()
    app.config["LICENSE_VALID"] = valid
    return jsonify(valid=valid, message=message, needs_key=needs_key)


@app.route("/api/license/activate", methods=["POST"])
def license_activate():
    """Attempt to activate a license key."""
    data = request.json or {}
    key = data.get("key", "").strip()
    if not key:
        return jsonify(success=False, message="Please enter a license key.")

    success, message = lic.activate_key(key)
    if success:
        app.config["LICENSE_VALID"] = True
    return jsonify(success=success, message=message)


# ── Routes: Pages ───────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


# ── Routes: Browse ──────────────────────────────────────────────────

@app.route("/api/browse", methods=["POST"])
def browse():
    data = request.json or {}
    mode = data.get("mode", "folder")
    initial = data.get("initial", str(SCRIPT_DIR))

    # Run tkinter file dialog in a subprocess to avoid
    # "main thread is not in main loop" crashes.
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); "
    )
    if mode == "folder":
        script += (
            f"p = filedialog.askdirectory(title='Select PDF Folder', "
            f"initialdir=r'{initial}'); "
        )
    else:
        script += (
            f"p = filedialog.askopenfilename(title='Open PDF', "
            f"filetypes=[('PDF files', '*.pdf')], initialdir=r'{initial}'); "
        )
    script += "print(p or ''); root.destroy()"

    try:
        result = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=120,
        )
        path = result.stdout.strip()
    except Exception:
        path = ""

    return jsonify(path=path or "")


# ── Routes: Folder Management ──────────────────────────────────────

@app.route("/api/folder/load", methods=["POST"])
def folder_load():
    data = request.json or {}
    path = data.get("path", "")
    if not path or not os.path.isdir(path):
        return jsonify(error="Not a valid directory"), 400

    # Close any open editor document
    with editor.lock:
        editor.close()

    folder.load_folder(path)
    remaining = folder.remaining()
    return jsonify(files=remaining, folder=path, count=len(remaining))


@app.route("/api/folder/select", methods=["POST"])
def folder_select():
    data = request.json or {}
    name = data.get("name", "")
    if not name:
        return jsonify(error="No filename specified"), 400

    file_info = folder.find_file(name)
    if not file_info:
        return jsonify(error="File not found in folder"), 404

    with editor.lock:
        # Check for unsaved changes
        if editor.doc and editor.modified:
            return jsonify(
                error="unsaved_changes",
                current_file=os.path.basename(editor.path)
            ), 409

        editor.close()

        try:
            editor.doc = fitz.open(file_info["path"])
        except Exception as e:
            return jsonify(error=f"Cannot open PDF: {e}"), 400

        editor.path = file_info["path"]
        editor.modified = False
        editor.invalidate_cache()

        page_count = len(editor.doc)
        editor.page_sizes = []
        for i in range(page_count):
            r = editor.doc[i].rect
            editor.page_sizes.append({"w": r.width, "h": r.height})

        return jsonify(
            page_count=page_count,
            path=file_info["path"],
            name=name,
            page_sizes=editor.page_sizes,
        )


@app.route("/api/folder/skip", methods=["POST"])
def folder_skip():
    data = request.json or {}
    name = data.get("name", "")
    h1_only = data.get("h1_only", False)
    count_toc = data.get("count_toc", True)

    if not name:
        return jsonify(error="No filename specified"), 400

    file_info = folder.find_file(name)
    if not file_info:
        return jsonify(error="File not found in folder"), 404

    # Close editor if this file is currently open
    with editor.lock:
        if editor.path and os.path.basename(editor.path) == name:
            editor.close()

    folder.mark_processed(name)
    remaining = folder.remaining()

    # Run TOC in background thread (don't block the HTTP response)
    toc_out_dir = os.path.join(folder.folder_path, "with_toc")
    threading.Thread(
        target=_run_toc_bg,
        args=(file_info["path"], toc_out_dir, h1_only, count_toc),
        daemon=True,
    ).start()

    return jsonify(
        ok=True,
        toc_result="pending",
        remaining=len(remaining),
    )


@app.route("/api/folder/save-and-finish", methods=["POST"])
def folder_save_and_finish():
    data = request.json or {}
    h1_only = data.get("h1_only", False)
    count_toc = data.get("count_toc", True)

    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400

        save_path = editor.path
        name = os.path.basename(save_path)

        try:
            # Atomic save: write to .tmp then replace
            tmp_path = save_path + ".tmp"
            editor.doc.save(tmp_path, deflate=True, garbage=4)
            editor.doc.close()
            os.replace(tmp_path, save_path)
        except Exception as e:
            return jsonify(error=f"Save failed: {e}"), 500

        # Fully close editor (don't reopen)
        editor.doc = None
        editor.path = None
        editor.modified = False
        editor.page_sizes = []
        editor.thumb_cache.clear()

    folder.mark_processed(name)
    remaining = folder.remaining()

    # Suggest next file
    next_file = remaining[0]["name"] if remaining else None

    # Run TOC in background thread (don't block the HTTP response)
    base_dir = folder.folder_path or os.path.dirname(save_path)
    toc_out_dir = os.path.join(base_dir, "with_toc")
    threading.Thread(
        target=_run_toc_bg,
        args=(save_path, toc_out_dir, h1_only, count_toc),
        daemon=True,
    ).start()

    return jsonify(
        ok=True,
        path=save_path,
        toc_result="pending",
        remaining=len(remaining),
        next_file=next_file,
    )


@app.route("/api/folder/skip-all", methods=["POST"])
def folder_skip_all():
    """Start batch-skipping all remaining files (add TOC to each)."""
    if batch.running:
        return jsonify(error="Batch already running"), 409

    data = request.json or {}
    h1_only = data.get("h1_only", False)
    count_toc = data.get("count_toc", True)

    remaining = folder.remaining()
    if not remaining:
        return jsonify(error="No files remaining"), 400

    # Close any open editor
    with editor.lock:
        editor.close()

    batch.reset()
    batch.running = True
    batch.total = len(remaining)

    t = threading.Thread(
        target=_skip_all_worker,
        args=(remaining, h1_only, count_toc),
        daemon=True,
    )
    t.start()

    return jsonify(total=len(remaining))


def _skip_all_worker(files, h1_only, count_toc):
    total = len(files)
    toc_out_dir = os.path.join(folder.folder_path, "with_toc")
    os.makedirs(toc_out_dir, exist_ok=True)

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    success, skipped, crashed, failed = 0, 0, 0, 0

    for idx, f in enumerate(files, 1):
        if batch.cancelled:
            batch.log_queue.put(("log", "\n--- Stopped by user ---\n"))
            break

        batch.current = idx
        batch.log_queue.put((
            "progress",
            {"current": idx, "total": total,
             "file": f["name"], "size": f["size_str"]},
        ))

        dst = os.path.join(toc_out_dir, f["name"])
        cmd = [sys.executable, "-u", HELPER, "toc", f["path"], dst]
        if h1_only:
            cmd.append("--h1-only")
        if not count_toc:
            cmd.append("--no-count-toc")

        try:
            batch.proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, bufsize=0,
            )

            result = None
            for raw in batch.proc.stdout:
                line = raw.decode("utf-8", errors="replace")
                if line.strip().startswith("__RESULT__:"):
                    result = line.strip().split(":", 1)[1]
                else:
                    batch.log_queue.put(("log", line))

            try:
                batch.proc.wait(timeout=TIMEOUT)
            except subprocess.TimeoutExpired:
                batch.proc.kill()
                batch.proc.wait()
                batch.log_queue.put(("log", f"  TIMEOUT after {TIMEOUT}s\n"))
                crashed += 1
                folder.mark_processed(f["name"])
                continue

            rc = batch.proc.returncode
            if rc != 0:
                batch.log_queue.put(("log", f"  CRASHED (exit code {rc})\n"))
                crashed += 1
            elif result == "success":
                success += 1
            elif result in ("scanned", "skipped"):
                skipped += 1
            else:
                failed += 1

        except Exception as e:
            batch.log_queue.put(("log", f"  ERROR: {e}\n"))
            failed += 1
        finally:
            batch.proc = None

        # Mark processed regardless of outcome
        folder.mark_processed(f["name"])
        batch.log_queue.put(("processed", f["name"]))

    batch.log_queue.put(("done", {
        "success": success, "skipped": skipped,
        "crashed": crashed, "failed": failed,
        "output_dir": toc_out_dir,
    }))
    fitz.TOOLS.store_shrink(100)
    gc.collect()
    batch.running = False


# ── Routes: Batch Processing (DOCX only) ───────────────────────────

@app.route("/api/batch/start", methods=["POST"])
def batch_start():
    if batch.running:
        return jsonify(error="Batch already running"), 409

    data = request.json
    input_dir = Path(data.get("dir", folder.folder_path or ""))
    mode = data.get("mode", "docx")
    h1_only = data.get("h1_only", False)
    count_toc = data.get("count_toc", True)

    if not input_dir.is_dir():
        return jsonify(error="Not a valid directory"), 400

    pdfs = sorted(input_dir.glob("*.pdf"))
    if not pdfs:
        return jsonify(error="No PDF files found"), 400

    batch.reset()
    batch.running = True
    batch.total = len(pdfs)

    out_dir = input_dir / ("with_toc" if mode == "toc" else "docx_output")
    out_dir.mkdir(exist_ok=True)

    t = threading.Thread(
        target=_batch_worker,
        args=(mode, input_dir, out_dir, pdfs, h1_only, count_toc),
        daemon=True,
    )
    t.start()

    return jsonify(total=len(pdfs))


def _batch_worker(mode, input_dir, out_dir, pdfs, h1_only, count_toc):
    total = len(pdfs)
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    success, skipped, crashed, failed = 0, 0, 0, 0

    for idx, pdf in enumerate(pdfs, 1):
        if batch.cancelled:
            batch.log_queue.put(("log", "\n--- Stopped by user ---\n"))
            break

        batch.current = idx
        size = pdf.stat().st_size
        batch.log_queue.put((
            "progress",
            {"current": idx, "total": total,
             "file": pdf.name, "size": human_size(size)},
        ))

        if mode == "toc":
            dst = str(out_dir / pdf.name)
        else:
            dst = str(out_dir / (pdf.stem + ".docx"))

        cmd = [sys.executable, "-u", HELPER, mode, str(pdf), dst]
        if mode == "toc" and h1_only:
            cmd.append("--h1-only")
        if mode == "toc" and not count_toc:
            cmd.append("--no-count-toc")

        try:
            batch.proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, bufsize=0,
            )

            result = None
            for raw in batch.proc.stdout:
                line = raw.decode("utf-8", errors="replace")
                if line.strip().startswith("__RESULT__:"):
                    result = line.strip().split(":", 1)[1]
                else:
                    batch.log_queue.put(("log", line))

            try:
                batch.proc.wait(timeout=TIMEOUT)
            except subprocess.TimeoutExpired:
                batch.proc.kill()
                batch.proc.wait()
                batch.log_queue.put(("log", f"  TIMEOUT after {TIMEOUT}s\n"))
                crashed += 1
                continue

            rc = batch.proc.returncode
            if rc != 0:
                batch.log_queue.put(("log", f"  CRASHED (exit code {rc})\n"))
                crashed += 1
            elif result == "success":
                success += 1
            elif result in ("scanned", "skipped"):
                skipped += 1
            else:
                failed += 1

        except Exception as e:
            batch.log_queue.put(("log", f"  ERROR: {e}\n"))
            failed += 1
        finally:
            batch.proc = None

    batch.log_queue.put(("done", {
        "success": success, "skipped": skipped,
        "crashed": crashed, "failed": failed,
        "output_dir": str(out_dir),
    }))
    fitz.TOOLS.store_shrink(100)
    gc.collect()
    batch.running = False


@app.route("/api/batch/events")
def batch_events():
    def generate():
        while True:
            try:
                msg_type, data = batch.log_queue.get(timeout=1)
            except queue.Empty:
                yield ": keepalive\n\n"
                if not batch.running:
                    yield "event: done\ndata: {}\n\n"
                    return
                continue

            payload = json.dumps(data) if isinstance(data, dict) else json.dumps(data)
            yield f"event: {msg_type}\ndata: {payload}\n\n"

            if msg_type == "done":
                return

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/batch/stop", methods=["POST"])
def batch_stop():
    batch.cancelled = True
    if batch.proc and batch.proc.poll() is None:
        batch.proc.terminate()
    return jsonify(ok=True)


# ── Routes: Page Editor ─────────────────────────────────────────────

@app.route("/api/editor/thumb/<int:page>")
def editor_thumb(page):
    with editor.lock:
        if not editor.doc:
            return "", 204  # doc closed — no content
        if page < 0 or page >= len(editor.doc):
            return "", 204  # page gone (deleted)

        width = request.args.get("w", 180, type=int)
        width = max(60, min(width, 600))

        try:
            img_bytes = editor.get_thumb(page, width)
        except Exception:
            return "", 204  # render failed — silently skip
    return Response(img_bytes, mimetype="image/jpeg", headers={
        "Cache-Control": "private, max-age=3600",
    })


@app.route("/api/editor/preview/<int:page>")
def editor_preview(page):
    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400
        if page < 0 or page >= len(editor.doc):
            return jsonify(error="Page out of range"), 400

        dpi = request.args.get("dpi", 150, type=int)
        dpi = max(72, min(dpi, 300))

        try:
            p = editor.doc[page]
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = p.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("jpeg", jpg_quality=75)
            pix = None  # release native pixmap memory immediately
        except Exception:
            return jsonify(error="Render failed"), 500

    return Response(img_bytes, mimetype="image/jpeg")


@app.route("/api/editor/delete", methods=["POST"])
def editor_delete():
    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400

        pages = sorted(request.json.get("pages", []), reverse=True)
        if not pages:
            return jsonify(error="No pages specified"), 400

        deleted = 0
        for p in pages:
            if 0 <= p < len(editor.doc):
                editor.doc.delete_page(p)
                deleted += 1

        editor.modified = True
        editor.invalidate_cache()

        # Rebuild page_sizes
        editor.page_sizes = []
        for i in range(len(editor.doc)):
            r = editor.doc[i].rect
            editor.page_sizes.append({"w": r.width, "h": r.height})

        return jsonify(page_count=len(editor.doc), deleted=deleted)


@app.route("/api/editor/save", methods=["POST"])
def editor_save():
    """Intermediate save — saves in-place, keeps doc open. No TOC."""
    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400

        save_path = editor.path

        try:
            tmp_path = save_path + ".tmp"
            editor.doc.save(tmp_path, deflate=True, garbage=4)
            editor.doc.close()
            os.replace(tmp_path, save_path)
            editor.doc = fitz.open(save_path)

            editor.modified = False
            editor.invalidate_cache()
        except Exception as e:
            return jsonify(error=f"Save failed: {e}"), 500

    return jsonify(ok=True, path=save_path)


@app.route("/api/editor/redact", methods=["POST"])
def editor_redact():
    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400

        data = request.json or {}
        rect = data.get("rect")
        page_idx = data.get("page", 0)
        scope = data.get("scope", "page")

        if not rect:
            return jsonify(error="No rect specified"), 400

        r = fitz.Rect(rect["x0"], rect["y0"], rect["x1"], rect["y1"])

        if scope == "all":
            pages = range(len(editor.doc))
        else:
            pages = [page_idx]

        affected = 0
        for p in pages:
            if 0 <= p < len(editor.doc):
                page = editor.doc[p]
                clipped = r & page.rect
                if clipped.is_empty:
                    continue
                page.add_redact_annot(clipped, fill=(1, 1, 1))
                page.apply_redactions(images=2, graphics=1, text=0)
                affected += 1

        editor.modified = True
        editor.invalidate_cache()

        return jsonify(ok=True, affected_pages=affected,
                       page_count=len(editor.doc))


@app.route("/api/editor/page-numbers", methods=["POST"])
def editor_page_numbers():
    with editor.lock:
        if not editor.doc:
            return jsonify(error="No document open"), 400

        data = request.json or {}
        point = data.get("point")
        if not point:
            return jsonify(error="No point specified"), 400

        fontsize = data.get("fontsize", 10)
        fontsize = max(4, min(fontsize, 72))
        font = data.get("font", "helv")
        if font not in ("helv", "tiro", "cour"):
            font = "helv"
        start = data.get("start", 1)
        color = tuple(data.get("color", [0, 0, 0]))

        px, py = float(point["x"]), float(point["y"])

        for i in range(len(editor.doc)):
            page = editor.doc[i]
            num_text = str(start + i)
            page.insert_text(
                (px, py),
                num_text,
                fontsize=fontsize,
                fontname=font,
                color=color,
                overlay=True,
            )

        editor.modified = True
        editor.invalidate_cache()

        return jsonify(ok=True, page_count=len(editor.doc))


@app.route("/api/folder/page-numbers-all", methods=["POST"])
def folder_page_numbers_all():
    """Add page numbers to ALL PDFs in the folder. Each book starts from 1."""
    if batch.running:
        return jsonify(error="Batch already running"), 409

    data = request.json or {}
    point = data.get("point")
    if not point:
        return jsonify(error="No point specified"), 400

    fontsize = data.get("fontsize", 10)
    fontsize = max(4, min(fontsize, 72))
    font = data.get("font", "helv")
    if font not in ("helv", "tiro", "cour"):
        font = "helv"
    color = tuple(data.get("color", [0, 0, 0]))

    if not folder.folder_path:
        return jsonify(error="No folder loaded"), 400

    pdfs = sorted(Path(folder.folder_path).glob("*.pdf"))
    if not pdfs:
        return jsonify(error="No PDF files found"), 400

    # Close any open editor to avoid file lock conflicts
    with editor.lock:
        editor.close()

    batch.reset()
    batch.running = True
    batch.total = len(pdfs)

    t = threading.Thread(
        target=_page_numbers_all_worker,
        args=(pdfs, float(point["x"]), float(point["y"]),
              fontsize, font, color),
        daemon=True,
    )
    t.start()

    return jsonify(total=len(pdfs))


def _page_numbers_all_worker(pdfs, px, py, fontsize, font, color):
    """Background worker: add page numbers to every PDF in the list.

    Uses a thread pool for concurrency.  Each thread opens its own
    fitz.Document on a separate file, so there is no shared mutable
    state — fully safe with PyMuPDF.
    """
    total = len(pdfs)
    success = 0
    failed = 0
    done_count = 0
    counter_lock = threading.Lock()

    def _number_one(pdf_path):
        """Process a single PDF — called from a pool thread."""
        doc = fitz.open(str(pdf_path))
        try:
            for i in range(len(doc)):
                page = doc[i]
                page.insert_text(
                    (px, py), str(1 + i),      # every book starts from 1
                    fontsize=fontsize, fontname=font,
                    color=color, overlay=True,
                )
            tmp = str(pdf_path) + ".tmp"
            doc.save(tmp, deflate=True, garbage=4)
        finally:
            doc.close()
        os.replace(tmp, str(pdf_path))

    workers = min(4, max(1, os.cpu_count() or 2))

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_number_one, p): p for p in pdfs}

        for future in as_completed(futures):
            pdf_path = futures[future]

            with counter_lock:
                done_count += 1
                batch.current = done_count

            size = pdf_path.stat().st_size
            batch.log_queue.put((
                "progress",
                {"current": done_count, "total": total,
                 "file": pdf_path.name, "size": human_size(size)},
            ))

            try:
                future.result()            # raises if _number_one failed
                batch.log_queue.put(("log", f"  Numbered: {pdf_path.name}\n"))
                with counter_lock:
                    success += 1
            except Exception as e:
                batch.log_queue.put(
                    ("log", f"  FAILED: {pdf_path.name} — {e}\n"))
                with counter_lock:
                    failed += 1

            # Periodic cleanup every 10 files to prevent memory buildup
            if done_count % 10 == 0:
                fitz.TOOLS.store_shrink(100)

            if batch.cancelled:
                for f in futures:
                    f.cancel()
                batch.log_queue.put(("log", "\n--- Stopped by user ---\n"))
                break

    batch.log_queue.put(("done", {
        "success": success, "skipped": 0,
        "crashed": 0, "failed": failed,
        "output_dir": str(pdfs[0].parent) if pdfs else "",
    }))
    fitz.TOOLS.store_shrink(100)
    gc.collect()
    batch.running = False


@app.route("/api/folder/redact-all", methods=["POST"])
def folder_redact_all():
    """Remove a region from all pages of ALL PDFs in the folder."""
    if batch.running:
        return jsonify(error="Batch already running"), 409

    data = request.json or {}
    rect = data.get("rect")
    if not rect:
        return jsonify(error="No rect specified"), 400

    if not folder.folder_path:
        return jsonify(error="No folder loaded"), 400

    pdfs = sorted(Path(folder.folder_path).glob("*.pdf"))
    if not pdfs:
        return jsonify(error="No PDF files found"), 400

    with editor.lock:
        editor.close()

    batch.reset()
    batch.running = True
    batch.total = len(pdfs)

    t = threading.Thread(
        target=_redact_all_worker,
        args=(pdfs, rect),
        daemon=True,
    )
    t.start()

    return jsonify(total=len(pdfs))


def _redact_all_worker(pdfs, rect):
    """Background worker: redact a region from every page of every PDF."""
    total = len(pdfs)
    success = 0
    failed = 0
    done_count = 0
    counter_lock = threading.Lock()

    def _redact_one(pdf_path):
        doc = fitz.open(str(pdf_path))
        try:
            r = fitz.Rect(rect["x0"], rect["y0"], rect["x1"], rect["y1"])
            affected = 0
            for i in range(len(doc)):
                page = doc[i]
                clipped = r & page.rect
                if clipped.is_empty:
                    continue
                page.add_redact_annot(clipped, fill=(1, 1, 1))
                page.apply_redactions(images=2, graphics=1, text=0)
                affected += 1
            if affected > 0:
                tmp = str(pdf_path) + ".tmp"
                doc.save(tmp, deflate=True, garbage=4)
        finally:
            doc.close()
        if affected > 0:
            os.replace(tmp, str(pdf_path))
        return affected

    # Fewer workers for redaction — apply_redactions is memory-intensive
    workers = min(2, max(1, os.cpu_count() or 2))

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_redact_one, p): p for p in pdfs}

        for future in as_completed(futures):
            pdf_path = futures[future]

            with counter_lock:
                done_count += 1
                batch.current = done_count

            size = pdf_path.stat().st_size
            batch.log_queue.put((
                "progress",
                {"current": done_count, "total": total,
                 "file": pdf_path.name, "size": human_size(size)},
            ))

            try:
                affected = future.result()
                batch.log_queue.put(
                    ("log", f"  Redacted {affected} pages: {pdf_path.name}\n"))
                with counter_lock:
                    success += 1
            except Exception as e:
                batch.log_queue.put(
                    ("log", f"  FAILED: {pdf_path.name} — {e}\n"))
                with counter_lock:
                    failed += 1

            # Periodic cleanup every 10 files to prevent memory buildup
            if done_count % 10 == 0:
                fitz.TOOLS.store_shrink(100)

            if batch.cancelled:
                for f in futures:
                    f.cancel()
                batch.log_queue.put(("log", "\n--- Stopped by user ---\n"))
                break

    batch.log_queue.put(("done", {
        "success": success, "skipped": 0,
        "crashed": 0, "failed": failed,
        "output_dir": str(pdfs[0].parent) if pdfs else "",
    }))
    fitz.TOOLS.store_shrink(100)
    gc.collect()
    batch.running = False


@app.route("/api/editor/close", methods=["POST"])
def editor_close():
    with editor.lock:
        editor.close()
    return jsonify(ok=True)


@app.route("/api/editor/status")
def editor_status():
    with editor.lock:
        return jsonify(
            is_open=editor.doc is not None,
            path=editor.path,
            name=os.path.basename(editor.path) if editor.path else None,
            page_count=len(editor.doc) if editor.doc else 0,
            modified=editor.modified,
        )


@app.route("/api/cleanup", methods=["POST"])
def cleanup():
    """Release PyMuPDF internal caches and run Python GC."""
    fitz.TOOLS.store_shrink(100)
    gc.collect()
    return jsonify(ok=True)


# ── Main ─────────────────────────────────────────────────────────────

def main():
    port = 5000
    url = f"http://127.0.0.1:{port}"
    print(f"Starting PDF Toolkit at {url}")

    # Warm up PyMuPDF's internal freetype/font caches so first PDF opens fast
    _warmup = fitz.open()
    _warmup.new_page()
    _warmup.close()
    del _warmup

    # Initial license check (sets LICENSE_VALID for the before_request gate)
    valid, msg, _ = lic.check_license()
    app.config["LICENSE_VALID"] = valid
    if valid:
        print(f"  License: {msg}")
    else:
        print(f"  License: {msg} — app will prompt for activation.")

    webbrowser.open(url)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
