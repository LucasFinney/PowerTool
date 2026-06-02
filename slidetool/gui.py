"""Minimal Tkinter GUI: file pickers + output-mode selector + Run."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .app import BuildRequest, run_build


def launch() -> None:
    root = tk.Tk()
    root.title("SlideTool")
    root.geometry("640x420")
    SlideToolUI(root)
    root.mainloop()


class SlideToolUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.pptx_var = tk.StringVar()
        self.markers_var = tk.StringVar()
        self.ref_var = tk.StringVar()
        self.duration_var = tk.StringVar()
        self.fps_var = tk.StringVar(value="24")
        self.mode_var = tk.StringVar(value="powerpoint")
        self.status_var = tk.StringVar(value="Ready.")

        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(root)
        frm.pack(fill="both", expand=True, **pad)

        self._row(frm, 0, "PowerPoint (.pptx):", self.pptx_var,
                  lambda: self._pick(self.pptx_var,
                                     [("PowerPoint", "*.pptx")]))
        self._row(frm, 1, "Markers CSV:", self.markers_var,
                  lambda: self._pick(self.markers_var,
                                     [("CSV", "*.csv"), ("All", "*.*")]))
        self._row(frm, 2, "Reference audio/video (optional):", self.ref_var,
                  lambda: self._pick(self.ref_var,
                                     [("Media", "*.wav *.mp3 *.mp4 *.mov *.aac *.m4a"),
                                      ("All", "*.*")]))

        ttk.Label(frm, text="Total duration (s, if no reference):").grid(
            row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.duration_var, width=14).grid(
            row=3, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Timeline FPS:").grid(
            row=4, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.fps_var, width=14).grid(
            row=4, column=1, sticky="w", **pad)

        mode_frame = ttk.LabelFrame(frm, text="Output mode")
        mode_frame.grid(row=5, column=0, columnspan=3, sticky="we", **pad)
        ttk.Radiobutton(
            mode_frame,
            text="PowerPoint native (preserves animations & transitions; slower)",
            variable=self.mode_var, value="powerpoint",
        ).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(
            mode_frame,
            text="Flat slides (fast preview; final state only, no animations)",
            variable=self.mode_var, value="flat",
        ).pack(anchor="w", padx=8, pady=2)

        self.run_btn = ttk.Button(frm, text="Run", command=self._on_run)
        self.run_btn.grid(row=6, column=1, sticky="w", **pad)

        ttk.Label(frm, textvariable=self.status_var, foreground="#444",
                  wraplength=600).grid(
            row=7, column=0, columnspan=3, sticky="we", **pad)

        frm.columnconfigure(1, weight=1)

    def _row(self, parent, r, label, var, browse):
        pad = {"padx": 8, "pady": 4}
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(parent, textvariable=var).grid(
            row=r, column=1, sticky="we", **pad)
        ttk.Button(parent, text="Browse...", command=browse).grid(
            row=r, column=2, **pad)

    def _pick(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def _on_run(self):
        try:
            req = self._build_request()
        except ValueError as e:
            messagebox.showerror("SlideTool", str(e))
            return

        self.run_btn.config(state="disabled")
        self.status_var.set("Running...")
        threading.Thread(target=self._do_run, args=(req,), daemon=True).start()

    def _build_request(self) -> BuildRequest:
        pptx = self.pptx_var.get().strip()
        markers = self.markers_var.get().strip()
        if not pptx or not markers:
            raise ValueError("Pick a .pptx and a markers CSV.")

        ref = self.ref_var.get().strip() or None
        dur_text = self.duration_var.get().strip()
        total = float(dur_text) if dur_text else None
        if not ref and total is None:
            raise ValueError("Provide a reference media file or total duration.")

        try:
            fps = float(self.fps_var.get().strip() or "24")
        except ValueError as e:
            raise ValueError(f"FPS must be a number: {e}") from e

        return BuildRequest(
            pptx_path=Path(pptx),
            markers_csv=Path(markers),
            fps=fps,
            total_duration_s=total,
            reference_media=Path(ref) if ref else None,
            output_mode=self.mode_var.get(),  # type: ignore[arg-type]
        )

    def _do_run(self, req: BuildRequest):
        try:
            result = run_build(req, on_progress=self._set_status_threadsafe)
        except Exception as e:
            self.root.after(0, lambda: self._on_done_error(e))
        else:
            self.root.after(0, lambda: self._on_done_ok(result))

    def _set_status_threadsafe(self, msg: str):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _on_done_ok(self, result):
        self.run_btn.config(state="normal")
        self.status_var.set(f"Done: {result.out_path}")
        messagebox.showinfo(
            "SlideTool",
            f"Wrote {result.out_path}\n"
            f"Mode: {result.output_mode}, cues: {result.n_cues}, "
            f"total: {result.total_duration_s:.2f}s",
        )

    def _on_done_error(self, exc: BaseException):
        self.run_btn.config(state="normal")
        self.status_var.set(f"Failed: {exc}")
        messagebox.showerror("SlideTool", str(exc))
