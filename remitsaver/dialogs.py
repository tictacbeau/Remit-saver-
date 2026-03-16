# dialogs.py
# All tkinter UI dialogs for RemitSaver.
# Each dialog is non-blocking relative to Outlook – callers run these on
# background threads (see remitsaver_addin.py _run_async).

from __future__ import annotations

import logging
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger("RemitSaver.dialogs")

# ── Utility: pick a local folder via Windows FolderBrowserDialog ──────────────

def _pick_folder(title: str = "Select Folder", initial: str = "",
                 parent: tk.Misc | None = None) -> str:
    """Open a native folder-picker; return selected path or ''.

    BUG FIX #11: The old implementation created a new tk.Tk() root every
    time it was called.  Creating a second Tk root while other tkinter
    windows are already open corrupts the widget hierarchy and can cause
    blank dialogs or crashes.  Pass the caller's window as *parent* instead
    so the dialog is correctly anchored to the existing Tk instance.
    """
    path = filedialog.askdirectory(
        title=title,
        initialdir=initial or os.path.expanduser("~"),
        parent=parent,
    )
    return path or ""


# ── Utility: pick an Outlook MAPI folder ─────────────────────────────────────

def _pick_outlook_folder(application) -> tuple[str, str]:
    """
    Invoke the native Outlook folder picker.
    Returns (display_path, entry_id).  Both empty strings on cancel.
    """
    try:
        ns     = application.GetNamespace("MAPI")
        folder = ns.PickFolder()
        if folder is None:
            return "", ""
        # Build a display path for the UI
        parts  = []
        f      = folder
        while f is not None:
            parts.insert(0, f.Name)
            try:
                f = f.Parent
                if not hasattr(f, "Name"):
                    break
            except Exception:
                break
        return "\\".join(parts), folder.EntryID
    except Exception as exc:
        logger.exception("PickFolder error: %s", exc)
        return "", ""


# ── Centered window helper ────────────────────────────────────────────────────

def _center(win: tk.Toplevel | tk.Tk, w: int, h: int) -> None:
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x  = (sw - w) // 2
    y  = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


# ═══════════════════════════════════════════════════════════════════════════════
# ExtractNowDialog
# ═══════════════════════════════════════════════════════════════════════════════

class ExtractNowDialog:
    """
    Dialog shown when user clicks 'Extract Now'.
    Lists payers as checkboxes; user picks which to run.
    Shows progress bar + live status log.
    """

    def __init__(self, master: tk.Tk | tk.Toplevel, application):
        self._app      = application
        self._win      = tk.Toplevel(master)
        self._win.title("RemitSaver – Extract Now")
        self._win.resizable(True, True)
        _center(self._win, 640, 520)
        self._vars: Dict[str, tk.BooleanVar] = {}
        self._build()

    def _build(self) -> None:
        from config import load_payers, load_settings
        self._payers   = load_payers()
        self._settings = load_settings()

        win = self._win
        win.columnconfigure(0, weight=1)
        win.rowconfigure(2, weight=1)

        # ── Payer checkboxes ──────────────────────────────────────────────
        lf = ttk.LabelFrame(win, text="Select Payers")
        lf.grid(row=0, column=0, padx=10, pady=(10, 4), sticky="ew")
        lf.columnconfigure(0, weight=1)

        if not self._payers:
            ttk.Label(lf, text="No payers configured. Use 'Manage Payers' to add one.").pack(pady=8)
        else:
            for payer in self._payers:
                v = tk.BooleanVar(value=True)
                self._vars[payer["name"]] = v
                ttk.Checkbutton(lf, text=payer["name"], variable=v).pack(anchor="w", padx=8)

        # ── Options ───────────────────────────────────────────────────────
        opt_frame = ttk.Frame(win)
        opt_frame.grid(row=1, column=0, padx=10, pady=2, sticky="ew")
        self._unread_only = tk.BooleanVar(value=self._settings.get("scan_unread_only", False))
        ttk.Checkbutton(opt_frame, text="Unread email only", variable=self._unread_only).pack(side="left")

        # ── Status log ────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(win, text="Status")
        log_frame.grid(row=2, column=0, padx=10, pady=4, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self._log_text = tk.Text(log_frame, state="disabled", wrap="word", height=12, font=("Consolas", 9))
        sb = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=sb.set)
        self._log_text.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # ── Progress bar ──────────────────────────────────────────────────
        self._progress = ttk.Progressbar(win, mode="indeterminate")
        self._progress.grid(row=3, column=0, padx=10, pady=4, sticky="ew")

        # ── Buttons ───────────────────────────────────────────────────────
        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=4, column=0, pady=(4, 10))
        self._run_btn = ttk.Button(btn_frame, text="Run Extraction", command=self._on_run)
        self._run_btn.pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Close", command=self._win.destroy).pack(side="left", padx=6)

        # ── Summary label ─────────────────────────────────────────────────
        self._summary = ttk.Label(win, text="")
        self._summary.grid(row=5, column=0, pady=(0, 6))

    def _append_log(self, msg: str) -> None:
        """Thread-safe append to the status log."""
        def _do():
            self._log_text.configure(state="normal")
            self._log_text.insert("end", msg + "\n")
            self._log_text.see("end")
            self._log_text.configure(state="disabled")
        try:
            self._win.after(0, _do)
        except Exception:
            pass

    def _on_run(self) -> None:
        selected = [name for name, v in self._vars.items() if v.get()]
        if not selected:
            messagebox.showinfo("RemitSaver", "No payers selected.", parent=self._win)
            return
        self._run_btn.configure(state="disabled")
        self._progress.start(10)
        t = threading.Thread(target=self._run_extraction, args=(selected,), daemon=True)
        t.start()

    def _run_extraction(self, selected: List[str]) -> None:
        # BUG FIX #4 (dialog side): initialise COM for this background thread
        # before touching any Outlook COM objects via self._app.
        import pythoncom
        pythoncom.CoInitialize()
        try:
            from config import load_payers, load_settings
            from extractor import run_extraction_for_payer

            payers   = {p["name"]: p for p in load_payers()}
            settings = load_settings()
            total_saved = total_skipped = total_errors = 0

            for name in selected:
                payer = payers.get(name)
                if not payer:
                    continue
                res = run_extraction_for_payer(
                    payer,
                    self._app,
                    settings,
                    unread_only=self._unread_only.get(),
                    progress_cb=self._append_log,
                )
                total_saved   += res.saved
                total_skipped += res.skipped
                total_errors  += res.errors

            summary = (
                f"Complete — Saved: {total_saved}, "
                f"Skipped: {total_skipped}, Errors: {total_errors}"
            )
            self._append_log(summary)

            def _done():
                self._progress.stop()
                self._run_btn.configure(state="normal")
                self._summary.configure(text=summary)
            try:
                self._win.after(0, _done)
            except Exception:
                pass
        finally:
            pythoncom.CoUninitialize()

    def run(self) -> None:
        self._win.grab_set()
        self._win.wait_window()


# ═══════════════════════════════════════════════════════════════════════════════
# PayerEditDialog  (Add / Edit a single payer)
# ═══════════════════════════════════════════════════════════════════════════════

class PayerEditDialog:
    """
    Modal dialog for adding or editing a single payer configuration.
    Returns the edited payer dict via .result, or None on cancel.
    """

    def __init__(
        self,
        master: tk.Toplevel | tk.Tk,
        application,
        payer: Optional[Dict[str, Any]] = None,
    ):
        self._app    = application
        self._result: Optional[Dict[str, Any]] = None
        self._win    = tk.Toplevel(master)
        self._win.title("Add Payer" if payer is None else "Edit Payer")
        self._win.resizable(False, False)
        _center(self._win, 520, 480)
        self._payer  = payer or {}
        self._build()

    @property
    def result(self) -> Optional[Dict[str, Any]]:
        return self._result

    def _build(self) -> None:
        p   = self._payer
        win = self._win
        fr  = ttk.Frame(win, padding=12)
        fr.pack(fill="both", expand=True)
        fr.columnconfigure(1, weight=1)

        def lbl(text, row):
            ttk.Label(fr, text=text).grid(row=row, column=0, sticky="e", padx=(0, 6), pady=4)

        # Name
        lbl("Name:", 0)
        self._name = tk.StringVar(value=p.get("name", ""))
        ttk.Entry(fr, textvariable=self._name, width=36).grid(row=0, column=1, sticky="ew")

        # Sender filters
        lbl("Sender filter(s):", 1)
        self._filters = tk.StringVar(value=", ".join(p.get("sender_filters", [])))
        ttk.Entry(fr, textvariable=self._filters, width=36).grid(row=1, column=1, sticky="ew")
        ttk.Label(fr, text="(comma-separated email substrings, e.g. @nationwide.com)",
                  font=("TkDefaultFont", 8), foreground="gray"
                  ).grid(row=2, column=1, sticky="w")

        # Outlook watch folder
        lbl("Watch folder:", 3)
        folder_fr = ttk.Frame(fr)
        folder_fr.grid(row=3, column=1, sticky="ew")
        folder_fr.columnconfigure(0, weight=1)
        self._folder_path  = tk.StringVar(value=p.get("watch_folder_path", ""))
        self._folder_id    = p.get("watch_folder_id", "")
        ttk.Entry(folder_fr, textvariable=self._folder_path, state="readonly", width=28).grid(
            row=0, column=0, sticky="ew")
        ttk.Button(folder_fr, text="Browse…", command=self._pick_outlook_folder).grid(
            row=0, column=1, padx=(4, 0))

        # Local save path
        lbl("Save path:", 4)
        save_fr = ttk.Frame(fr)
        save_fr.grid(row=4, column=1, sticky="ew")
        save_fr.columnconfigure(0, weight=1)
        self._save_path = tk.StringVar(value=p.get("save_path", ""))
        ttk.Entry(save_fr, textvariable=self._save_path, width=28).grid(row=0, column=0, sticky="ew")
        ttk.Button(save_fr, text="Browse…", command=self._pick_local_folder).grid(
            row=0, column=1, padx=(4, 0))

        # Extension filter
        lbl("File types:", 5)
        self._ext_var = tk.StringVar(value=", ".join(p.get("extensions", ["all"])))
        ttk.Entry(fr, textvariable=self._ext_var, width=36).grid(row=5, column=1, sticky="ew")
        ttk.Label(fr, text="all  or  pdf, xlsx, csv …", font=("TkDefaultFont", 8),
                  foreground="gray").grid(row=6, column=1, sticky="w")

        # Amount location
        lbl("Amount location:", 7)
        loc_fr = ttk.Frame(fr)
        loc_fr.grid(row=7, column=1, sticky="w")
        self._amt_loc = tk.StringVar(value=p.get("amount_location", "auto"))
        for val, lbl_txt in [("auto", "Auto"), ("subject", "Subject"),
                              ("body", "Body"), ("attachment", "Attachment")]:
            ttk.Radiobutton(loc_fr, text=lbl_txt, variable=self._amt_loc, value=val).pack(
                side="left", padx=4)

        # Amount strategy
        lbl("Amount strategy:", 8)
        strat_fr = ttk.Frame(fr)
        strat_fr.grid(row=8, column=1, sticky="w")
        self._amt_strat = tk.StringVar(value=p.get("amount_strategy", "largest"))
        for val, lbl_txt in [("largest", "Largest"), ("last", "Last")]:
            ttk.Radiobutton(strat_fr, text=lbl_txt, variable=self._amt_strat, value=val).pack(
                side="left", padx=4)

        # Buttons
        btn_fr = ttk.Frame(fr)
        btn_fr.grid(row=9, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn_fr, text="Save", command=self._on_save).pack(side="left", padx=6)
        ttk.Button(btn_fr, text="Cancel", command=self._win.destroy).pack(side="left", padx=6)

    def _pick_outlook_folder(self) -> None:
        if self._app is None:
            messagebox.showwarning("RemitSaver",
                "Outlook application is not available.", parent=self._win)
            return
        path, entry_id = _pick_outlook_folder(self._app)
        if path:
            self._folder_path.set(path)
            self._folder_id = entry_id

    def _pick_local_folder(self) -> None:
        path = _pick_folder("Select Save Folder", self._save_path.get(), parent=self._win)
        if path:
            self._save_path.set(path)

    def _on_save(self) -> None:
        name = self._name.get().strip()
        if not name:
            messagebox.showerror("RemitSaver", "Payer name is required.", parent=self._win)
            return
        raw_ext = [e.strip().lstrip(".").lower() for e in self._ext_var.get().split(",") if e.strip()]
        self._result = {
            "name":              name,
            "sender_filters":    [f.strip() for f in self._filters.get().split(",") if f.strip()],
            "watch_folder_path": self._folder_path.get(),
            "watch_folder_id":   self._folder_id,
            "save_path":         self._save_path.get(),
            "extensions":        raw_ext if raw_ext else ["all"],
            "amount_location":   self._amt_loc.get(),
            "amount_strategy":   self._amt_strat.get(),
        }
        self._win.destroy()

    def run(self) -> Optional[Dict[str, Any]]:
        self._win.grab_set()
        self._win.wait_window()
        return self._result


# ═══════════════════════════════════════════════════════════════════════════════
# ManagePayersDialog
# ═══════════════════════════════════════════════════════════════════════════════

class ManagePayersDialog:
    """
    Dialog for Add / Edit / Delete / Test payers.
    """

    def __init__(self, master: tk.Tk | tk.Toplevel, application):
        self._app = application
        self._win = tk.Toplevel(master)
        self._win.title("RemitSaver – Manage Payers")
        self._win.resizable(True, True)
        _center(self._win, 560, 420)
        self._build()
        self._reload_list()

    def _build(self) -> None:
        win = self._win
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)

        # Listbox
        list_fr = ttk.Frame(win)
        list_fr.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        list_fr.columnconfigure(0, weight=1)
        list_fr.rowconfigure(0, weight=1)

        self._listbox = tk.Listbox(list_fr, selectmode="single", width=50, height=16,
                                   font=("TkDefaultFont", 10))
        sb = ttk.Scrollbar(list_fr, command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=sb.set)
        self._listbox.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # Buttons
        btn_fr = ttk.Frame(win)
        btn_fr.grid(row=1, column=0, pady=(0, 10))
        ttk.Button(btn_fr, text="Add",    command=self._on_add).pack(side="left", padx=4)
        ttk.Button(btn_fr, text="Edit",   command=self._on_edit).pack(side="left", padx=4)
        ttk.Button(btn_fr, text="Delete", command=self._on_delete).pack(side="left", padx=4)
        ttk.Button(btn_fr, text="Test (dry-run)", command=self._on_test).pack(side="left", padx=4)
        ttk.Button(btn_fr, text="Close",  command=self._win.destroy).pack(side="left", padx=4)

        # Dry-run output
        self._test_text = tk.Text(win, height=6, state="disabled",
                                  wrap="word", font=("Consolas", 9))
        self._test_text.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

    def _reload_list(self) -> None:
        from config import load_payers
        self._payers = load_payers()
        self._listbox.delete(0, "end")
        for p in self._payers:
            self._listbox.insert("end", p["name"])

    def _selected_payer(self) -> Optional[Dict[str, Any]]:
        sel = self._listbox.curselection()
        if not sel:
            return None
        return self._payers[sel[0]]

    def _on_add(self) -> None:
        dlg = PayerEditDialog(self._win, self._app)
        result = dlg.run()
        if result:
            from config import upsert_payer
            upsert_payer(result)
            self._reload_list()

    def _on_edit(self) -> None:
        payer = self._selected_payer()
        if not payer:
            messagebox.showinfo("RemitSaver", "Select a payer first.", parent=self._win)
            return
        dlg = PayerEditDialog(self._win, self._app, payer)
        result = dlg.run()
        if result:
            from config import upsert_payer
            upsert_payer(result)
            self._reload_list()

    def _on_delete(self) -> None:
        payer = self._selected_payer()
        if not payer:
            messagebox.showinfo("RemitSaver", "Select a payer first.", parent=self._win)
            return
        if messagebox.askyesno("RemitSaver",
                f"Delete payer '{payer['name']}'?", parent=self._win):
            from config import delete_payer
            delete_payer(payer["name"])
            self._reload_list()

    def _on_test(self) -> None:
        payer = self._selected_payer()
        if not payer:
            messagebox.showinfo("RemitSaver", "Select a payer first.", parent=self._win)
            return
        self._test_text.configure(state="normal")
        self._test_text.delete("1.0", "end")
        self._test_text.insert("end", f"Dry-run for '{payer['name']}' (last 5 matching emails)…\n")
        self._test_text.configure(state="disabled")

        def _run():
            # BUG FIX #10: tkinter widgets must only be touched from the
            # thread that owns the event loop.  Using after(0, ...) safely
            # schedules each update on the main tkinter thread.
            import pythoncom
            pythoncom.CoInitialize()
            try:
                from config import load_settings
                from extractor import run_extraction_for_payer

                def _log(msg: str):
                    def _do():
                        self._test_text.configure(state="normal")
                        self._test_text.insert("end", msg + "\n")
                        self._test_text.see("end")
                        self._test_text.configure(state="disabled")
                    try:
                        self._win.after(0, _do)
                    except Exception:
                        pass

                settings = load_settings()
                run_extraction_for_payer(
                    payer,
                    self._app,
                    settings,
                    dry_run=True,
                    progress_cb=_log,
                    limit=5,
                )
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=_run, daemon=True).start()

    def run(self) -> None:
        self._win.grab_set()
        self._win.wait_window()


# ═══════════════════════════════════════════════════════════════════════════════
# SettingsDialog
# ═══════════════════════════════════════════════════════════════════════════════

class SettingsDialog:
    """
    Global settings dialog.
    """

    def __init__(self, master: tk.Tk | tk.Toplevel):
        self._win = tk.Toplevel(master)
        self._win.title("RemitSaver – Settings")
        self._win.resizable(False, False)
        _center(self._win, 520, 320)
        self._build()

    def _build(self) -> None:
        from config import load_settings
        settings = load_settings()

        win = self._win
        fr  = ttk.Frame(win, padding=14)
        fr.pack(fill="both", expand=True)
        fr.columnconfigure(1, weight=1)

        def lbl(text: str, row: int) -> None:
            ttk.Label(fr, text=text).grid(row=row, column=0, sticky="e", padx=(0, 6), pady=5)

        # Default save root
        lbl("Default save root:", 0)
        save_fr = ttk.Frame(fr)
        save_fr.grid(row=0, column=1, sticky="ew")
        save_fr.columnconfigure(0, weight=1)
        self._save_root = tk.StringVar(value=settings.get("default_save_root", ""))
        ttk.Entry(save_fr, textvariable=self._save_root, width=30).grid(row=0, column=0, sticky="ew")
        ttk.Button(save_fr, text="Browse…",
                   command=lambda: self._save_root.set(
                       _pick_folder("Default Save Root", self._save_root.get(),
                                    parent=self._win) or self._save_root.get()
                   )).grid(row=0, column=1, padx=(4, 0))

        # Log file path
        lbl("Log file path:", 1)
        log_fr = ttk.Frame(fr)
        log_fr.grid(row=1, column=1, sticky="ew")
        log_fr.columnconfigure(0, weight=1)
        self._log_path = tk.StringVar(value=settings.get("log_file_path", ""))
        ttk.Entry(log_fr, textvariable=self._log_path, width=30).grid(row=0, column=0, sticky="ew")
        ttk.Button(log_fr, text="Browse…",
                   command=lambda: self._log_path.set(
                       filedialog.asksaveasfilename(
                           title="Log File", defaultextension=".log",
                           filetypes=[("Log files", "*.log"), ("All", "*.*")],
                           initialfile=self._log_path.get(),
                       ) or self._log_path.get()
                   )).grid(row=0, column=1, padx=(4, 0))

        # Checkboxes
        self._skip_inline = tk.BooleanVar(value=settings.get("skip_inline_attachments", True))
        ttk.Checkbutton(fr, text="Skip inline attachments",
                        variable=self._skip_inline).grid(row=2, column=0, columnspan=2, sticky="w", pady=4)

        self._auto_run = tk.BooleanVar(value=settings.get("auto_run_on_startup", False))
        ttk.Checkbutton(fr, text="Auto-run extraction on Outlook startup",
                        variable=self._auto_run).grid(row=3, column=0, columnspan=2, sticky="w", pady=4)

        self._unread_only = tk.BooleanVar(value=settings.get("scan_unread_only", False))
        ttk.Checkbutton(fr, text="Scan unread emails only (default)",
                        variable=self._unread_only).grid(row=4, column=0, columnspan=2, sticky="w", pady=4)

        # Buttons
        btn_fr = ttk.Frame(fr)
        btn_fr.grid(row=5, column=0, columnspan=2, pady=(14, 0))
        ttk.Button(btn_fr, text="Save", command=self._on_save).pack(side="left", padx=6)
        ttk.Button(btn_fr, text="Cancel", command=self._win.destroy).pack(side="left", padx=6)

    def _on_save(self) -> None:
        from config import save_settings, load_settings
        settings = load_settings()
        settings["default_save_root"]       = self._save_root.get().strip()
        settings["log_file_path"]           = self._log_path.get().strip()
        settings["skip_inline_attachments"] = self._skip_inline.get()
        settings["auto_run_on_startup"]     = self._auto_run.get()
        settings["scan_unread_only"]        = self._unread_only.get()
        save_settings(settings)
        messagebox.showinfo("RemitSaver", "Settings saved.", parent=self._win)
        self._win.destroy()

    def run(self) -> None:
        self._win.grab_set()
        self._win.wait_window()
