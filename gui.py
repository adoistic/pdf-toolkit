#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
# All rights reserved.  Proprietary and confidential.
# Unauthorized copying or distribution is strictly prohibited.
# ─────────────────────────────────────────────────────────────────────
"""
gui.py — Tkinter GUI for pdf-toolkit.
Each PDF is processed in its own subprocess so a crash on one file
(corrupt PDF, OOM, segfault) cannot take down the whole application.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import subprocess
import sys
import os
import shutil
from pathlib import Path

import license as lic

IS_FROZEN  = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(__file__).parent

if IS_FROZEN:
    BUNDLE_DIR = Path(sys._MEIPASS)
    HELPER = str(SCRIPT_DIR / "_run_one.exe")
else:
    HELPER = str(SCRIPT_DIR / "_run_one.py")

TIMEOUT = 300  # seconds per file (5 min)


# ─── License activation dialog ───────────────────────────────────────────────

class LicenseDialog:
    """Modal dialog for license key entry. Blocks the main window until
    a valid license is activated."""

    def __init__(self, root, message="Please enter your license key."):
        self.activated = False

        self.win = tk.Toplevel(root)
        self.win.title("License Activation")
        self.win.geometry("460x260")
        self.win.resizable(False, False)
        self.win.transient(root)
        self.win.grab_set()

        # Prevent closing without activation
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        # Center on screen
        self.win.update_idletasks()
        x = (self.win.winfo_screenwidth() - 460) // 2
        y = (self.win.winfo_screenheight() - 260) // 2
        self.win.geometry(f"+{x}+{y}")

        self._build_ui(message)

    def _build_ui(self, message):
        frame = ttk.Frame(self.win, padding=24)
        frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title = ttk.Label(frame, text="PDF Toolkit — License",
                          font=("Segoe UI", 14, "bold"))
        title.pack(pady=(0, 12))

        # Status message
        self.msg_var = tk.StringVar(value=message)
        self.msg_label = ttk.Label(frame, textvariable=self.msg_var,
                                   wraplength=400, foreground="gray")
        self.msg_label.pack(pady=(0, 12))

        # Key entry
        key_frame = ttk.Frame(frame)
        key_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(key_frame, text="License Key:").pack(side=tk.LEFT)
        self.key_var = tk.StringVar()
        self.key_entry = ttk.Entry(key_frame, textvariable=self.key_var,
                                   width=30, font=("Consolas", 11))
        self.key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        self.key_entry.focus_set()
        self.key_entry.bind("<Return>", lambda e: self._activate())

        # Auto-format: uppercase + insert dashes (XXXX-XXXX-XXXX-XXXX)
        self.key_var.trace_add("write", self._format_key)

        # Activate button
        self.activate_btn = ttk.Button(frame, text="Activate",
                                       command=self._activate)
        self.activate_btn.pack(pady=(4, 0))

    def _format_key(self, *_):
        raw = self.key_var.get().upper().replace("-", "").replace(" ", "")
        # Only keep alphanumeric
        raw = "".join(c for c in raw if c.isalnum())[:16]
        # Insert dashes every 4 chars
        parts = [raw[i:i+4] for i in range(0, len(raw), 4)]
        formatted = "-".join(parts)
        # Avoid infinite recursion from trace
        if formatted != self.key_var.get():
            self.key_var.set(formatted)
            self.key_entry.icursor(len(formatted))

    def _activate(self):
        key = self.key_var.get().strip()
        if not key:
            self.msg_var.set("Please enter a license key.")
            self.msg_label.configure(foreground="#CC3333")
            return

        self.activate_btn.configure(state="disabled")
        self.msg_var.set("Validating...")
        self.msg_label.configure(foreground="gray")
        self.win.update()

        # Run activation (may take a moment for network call)
        success, message = lic.activate_key(key)

        if success:
            self.activated = True
            self.msg_var.set(message)
            self.msg_label.configure(foreground="#228B22")
            self.win.after(600, self.win.destroy)
        else:
            self.msg_var.set(message)
            self.msg_label.configure(foreground="#CC3333")
            self.activate_btn.configure(state="normal")

    def _on_close(self):
        if not self.activated:
            if messagebox.askokcancel(
                "Exit",
                "A valid license is required to use PDF Toolkit.\n\n"
                "Exit the application?",
                parent=self.win,
            ):
                self.win.destroy()
                self.win.master.destroy()
        else:
            self.win.destroy()

    def wait(self):
        """Block until the dialog is closed."""
        self.win.wait_window()
        return self.activated


# ─── Main GUI ────────────────────────────────────────────────────────────────

class PDFToolkitGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("820x560")
        self.root.minsize(620, 420)

        self.input_dir = tk.StringVar(value=str(SCRIPT_DIR))
        self.h1_only = tk.BooleanVar(value=False)
        self.running = False
        self.cancelled = False
        self.proc = None

        # ── License check on startup ──────────────────────────────────
        valid, msg, needs_key = lic.check_license()

        if not valid:
            # Hide main window while license dialog is shown
            self.root.withdraw()

            if needs_key:
                dialog = LicenseDialog(self.root, msg)
                activated = dialog.wait()
                if not activated:
                    return  # User closed without activating → app exits
            else:
                # Offline lock (no key entry possible)
                messagebox.showerror(
                    "License Error", msg,
                )
                self.root.destroy()
                return

            self.root.deiconify()

        self._build_ui()
        self._update_count()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Directory selector
        top = ttk.Frame(self.root, padding=(14, 14, 14, 6))
        top.pack(fill=tk.X)
        ttk.Label(top, text="PDF Folder:").pack(side=tk.LEFT)
        self.dir_entry = ttk.Entry(top, textvariable=self.input_dir)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        ttk.Button(top, text="Browse", command=self._browse).pack(side=tk.LEFT)

        # Action buttons + count
        btn = ttk.Frame(self.root, padding=(14, 4, 14, 4))
        btn.pack(fill=tk.X)
        self.toc_btn = ttk.Button(btn, text="Add TOC to All PDFs",
                                  command=lambda: self._run("toc"))
        self.toc_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.docx_btn = ttk.Button(btn, text="Convert All PDFs to DOCX",
                                   command=lambda: self._run("docx"))
        self.docx_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.stop_btn = ttk.Button(btn, text="Stop", command=self._stop,
                                   state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        self.count_var = tk.StringVar()
        ttk.Label(btn, textvariable=self.count_var).pack(side=tk.RIGHT)

        # Options
        opts = ttk.Frame(self.root, padding=(14, 2, 14, 2))
        opts.pack(fill=tk.X)
        self.h1_check = ttk.Checkbutton(
            opts, text="H1 headings only (skip sub-sections)",
            variable=self.h1_only,
        )
        self.h1_check.pack(side=tk.LEFT)

        # Progress bar
        prog = ttk.Frame(self.root, padding=(14, 2, 14, 4))
        prog.pack(fill=tk.X)
        self.progress = ttk.Progressbar(prog, mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 8))
        self.prog_label = tk.StringVar(value="")
        ttk.Label(prog, textvariable=self.prog_label, width=22,
                  anchor=tk.E).pack(side=tk.RIGHT)

        # Log area
        log_frame = ttk.LabelFrame(self.root, text="Output", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))
        self.log = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 9),
                           state=tk.DISABLED)
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.input_dir.trace_add("write", lambda *_: self._update_count())

    # ── Helpers ───────────────────────────────────────────────────────────

    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.input_dir.get())
        if d:
            self.input_dir.set(d)

    def _update_count(self):
        try:
            n = len(list(Path(self.input_dir.get()).glob("*.pdf")))
            self.count_var.set(f"{n} PDF{'s' if n != 1 else ''} found")
        except Exception:
            self.count_var.set("")

    def _append(self, text):
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _set_running(self, running):
        self.running = running
        st = "disabled" if running else "normal"
        self.toc_btn.configure(state=st)
        self.docx_btn.configure(state=st)
        self.dir_entry.configure(state=st)
        self.h1_check.configure(state=st)
        self.stop_btn.configure(state="normal" if running else "disabled")
        if not running:
            self.progress["value"] = 0
            self.prog_label.set("")

    def _stop(self):
        self.cancelled = True
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()

    @staticmethod
    def _human_size(nbytes):
        for unit in ("B", "KB", "MB", "GB"):
            if abs(nbytes) < 1024:
                return f"{nbytes:.1f} {unit}"
            nbytes /= 1024
        return f"{nbytes:.1f} TB"

    # ── Run ───────────────────────────────────────────────────────────────

    def _run(self, mode):
        if self.running:
            return
        input_dir = Path(self.input_dir.get())
        if not input_dir.is_dir():
            self._append(f"ERROR: Not a valid directory: {input_dir}\n")
            return
        pdfs = sorted(input_dir.glob("*.pdf"))
        if not pdfs:
            self._append(f"No PDF files found in {input_dir}\n")
            return

        # Clear
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)
        self.cancelled = False

        # Prepare output dir
        if mode == "toc":
            out_dir = input_dir / "with_toc"
        else:
            out_dir = input_dir / "docx_output"
        out_dir.mkdir(exist_ok=True)

        self._set_running(True)
        threading.Thread(target=self._worker,
                         args=(mode, input_dir, out_dir, pdfs),
                         daemon=True).start()

    def _worker(self, mode, input_dir, out_dir, pdfs):
        total = len(pdfs)
        self.root.after(0, self._append,
                        f"Processing {total} PDFs  ({mode.upper()})\n")
        self.root.after(0, lambda: self.progress.configure(maximum=total))

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        success, skipped, crashed, failed = 0, 0, 0, 0

        for idx, pdf in enumerate(pdfs, 1):
            if self.cancelled:
                self.root.after(0, self._append, "\n--- Stopped by user ---\n")
                break

            # Progress
            self.root.after(0, lambda i=idx: self.progress.configure(value=i))
            self.root.after(0, lambda i=idx, t=total, n=pdf.name:
                            self.prog_label.set(f"{i}/{t}"))

            size = pdf.stat().st_size
            self.root.after(0, self._append,
                            f"\n[{idx}/{total}]  {pdf.name}  "
                            f"({self._human_size(size)})\n")

            # Output path
            if mode == "toc":
                dst = str(out_dir / pdf.name)
            else:
                dst = str(out_dir / (pdf.stem + ".docx"))

            # ── Spawn isolated subprocess for this single PDF ────────
            if IS_FROZEN:
                cmd = [HELPER, mode, str(pdf), dst]
            else:
                cmd = [sys.executable, "-u", HELPER, mode, str(pdf), dst]
            if mode == "toc" and self.h1_only.get():
                cmd.append("--h1-only")
            try:
                self.proc = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    env=env, bufsize=0,
                )

                result = None
                for raw in self.proc.stdout:
                    try:
                        line = raw.decode("utf-8", errors="replace")
                    except Exception:
                        line = str(raw)

                    # Parse machine-readable result tag
                    if line.strip().startswith("__RESULT__:"):
                        result = line.strip().split(":", 1)[1]
                    else:
                        self.root.after(0, self._append, line)

                try:
                    self.proc.wait(timeout=TIMEOUT)
                except subprocess.TimeoutExpired:
                    self.proc.kill()
                    self.proc.wait()
                    self.root.after(0, self._append,
                                    f"  TIMEOUT after {TIMEOUT}s — killed\n")
                    crashed += 1
                    continue

                rc = self.proc.returncode
                if rc != 0:
                    self.root.after(0, self._append,
                                    f"  CRASHED  (exit code {rc})\n")
                    crashed += 1
                elif result == "success":
                    success += 1
                elif result in ("scanned", "skipped"):
                    skipped += 1
                else:
                    failed += 1

            except Exception as e:
                self.root.after(0, self._append, f"  ERROR: {e}\n")
                failed += 1
            finally:
                self.proc = None

        # ── Summary ──────────────────────────────────────────────────
        parts = [f"{success} succeeded"]
        if skipped:
            parts.append(f"{skipped} skipped")
        if crashed:
            parts.append(f"{crashed} crashed")
        if failed:
            parts.append(f"{failed} failed")
        sep = "=" * 50
        self.root.after(0, self._append,
                        f"\n{sep}\nDone.  {' | '.join(parts)}\n"
                        f"Output: {out_dir}\n")
        self.root.after(0, self._set_running, False)
        self.root.after(0, self._update_count)


def main():
    root = tk.Tk()
    PDFToolkitGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
