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
from tkinter import ttk, filedialog
import threading
import subprocess
import sys
import os
import shutil
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
HELPER     = str(SCRIPT_DIR / "_run_one.py")
TIMEOUT    = 300  # seconds per file (5 min)


class PDFToolkitGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("820x560")
        self.root.minsize(620, 420)

        self.input_dir = tk.StringVar(value=str(SCRIPT_DIR))
        self.h1_only = tk.BooleanVar(value=False)
        self.count_toc = tk.BooleanVar(value=True)
        self.running = False
        self.cancelled = False
        self.proc = None

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
        self.toc_count_check = ttk.Checkbutton(
            opts, text="Include TOC in page count",
            variable=self.count_toc,
        )
        self.toc_count_check.pack(side=tk.LEFT, padx=(16, 0))

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
        self.toc_count_check.configure(state=st)
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
            cmd = [sys.executable, "-u", HELPER, mode, str(pdf), dst]
            if mode == "toc" and self.h1_only.get():
                cmd.append("--h1-only")
            if mode == "toc" and not self.count_toc.get():
                cmd.append("--no-count-toc")
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
