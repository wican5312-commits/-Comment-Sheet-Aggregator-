import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk
import aggregator
import os
import sys
import threading
import datetime
import platform

# ---------------------------------------------------------------------------
# Config file path — next to the exe when frozen, or at project root otherwise
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):          # PyInstaller exe
    _APP_DIR = os.path.dirname(sys.executable)
else:                                       # Running as script
    _APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CONFIG_PATH = os.path.join(_APP_DIR, 'user_config.json')

# Load persisted user settings immediately (before the GUI opens)
aggregator.load_config(_CONFIG_PATH)

# ---------------------------------------------------------------------------
# Translation table
# ---------------------------------------------------------------------------
TRANSLATIONS = {
    "JP": {
        "title":                  "コメントシート集計ツール",
        "step1":                  "1. 入力データ",
        "select_files":           "コメントシートを選択...",
        "no_files":               "ファイルが選択されていません",
        "select_att":             "出席表を選択 (任意)...",
        "no_att":                 "出席表が選択されていません",
        "att_selected":           "出席表: ",
        "step2":                  "2. 設定",
        "target_year":            "対象年度 (任意):",
        "year_desc":              "(この年度で始まる科目を抽出)",
        "step3":                  "3. 実行",
        "run":                    "集計開始",
        "ready":                  "準備完了。",
        "processing":             "処理中... 保存先: ",
        "success":                "成功: ",
        "success_title":          "完了",
        "success_msg":            "集計が完了しました！\nファイルが保存されました。",
        "error":                  "エラー",
        "warning":                "警告",
        "select_files_warn":      "ファイルを選択してください。",
        "year_empty_title":       "警告: 対象年度が空です",
        "year_empty_msg":         (
            "対象年度が入力されていません。\n\n"
            "ヘッダー行や関係ない科目を含む「すべての行」が処理されます。\n\n"
            "本当に続行しますか？"
        ),
        "files_selected":         " 個のファイルを選択",
        "file_selection_cancelled":"ファイルの選択がキャンセルされました。",
        "att_selection_cancelled": "出席表の選択がキャンセルされました。",
        "critical_error":         "致命的なエラー: ",
        "save_as":                "保存先を指定",
        "footer":                 "制作：2025年度院生（有志）",
        "adv_settings_btn":       "⚙ 詳細設定",
    },
    "EN": {
        "title":                  "Comment Sheet Aggregator",
        "step1":                  "1. Input Data",
        "select_files":           "Select Comment Sheets...",
        "no_files":               "No files selected",
        "select_att":             "Select Attendance Sheet (Optional)...",
        "no_att":                 "No attendance sheet selected",
        "att_selected":           "Attendance: ",
        "step2":                  "2. Settings",
        "target_year":            "Target Year (Optional):",
        "year_desc":              "(Filters courses starting with this year)",
        "step3":                  "3. Execution",
        "run":                    "Run Aggregation",
        "ready":                  "Ready.",
        "processing":             "Processing... saving to ",
        "success":                "SUCCESS: ",
        "success_title":          "Success",
        "success_msg":            "Aggregation complete!\nFile saved.",
        "error":                  "Error",
        "warning":                "Warning",
        "select_files_warn":      "Please select files first.",
        "year_empty_title":       "Warning: Target Year Empty",
        "year_empty_msg":         (
            "Target Year is empty.\n\n"
            "This will process ALL rows, which might include header rows "
            "or unrelated courses.\n\nAre you sure you want to proceed?"
        ),
        "files_selected":         " files selected",
        "file_selection_cancelled":"File selection cancelled.",
        "att_selection_cancelled": "Attendance selection cancelled.",
        "critical_error":         "CRITICAL ERROR: ",
        "save_as":                "Save Output As",
        "footer":                 "Developed by 2025 Graduate Students",
        "adv_settings_btn":       "⚙ Advanced Settings",
    },
}


# ===========================================================================
# Advanced Settings Dialog
# ===========================================================================

class AdvancedSettingsDialog(tk.Toplevel):
    """
    Modal dialog that lets users edit every CONFIG parameter without
    touching the source code.  Changes are persisted to user_config.json.
    """

    # Excel column letters A-Z (covers all realistic use cases)
    _LETTERS = [chr(ord('A') + i) for i in range(26)]

    # ── Parameter definitions ────────────────────────────────────────────────
    # Each entry: (config_key, JP_label, EN_label, JP_desc, EN_desc, widget_type)
    # widget_type: 'col'  → Combobox A-Z  (stored as 0-based index)
    #              'row'  → Spinbox       (displayed as ATT_SKIP_ROWS+1)
    _CS_PARAMS = [
        ("COL_SUB_ID",
         "提出ID列",        "Submission ID col",
         "提出IDが記録されている列",
         "Column containing the Submission ID",
         "col"),
        ("COL_COURSE",
         "科目名列",        "Course name col",
         "科目名（年度フィルタリング対象）が記録されている列",
         "Column with the course name (used for year filter)",
         "col"),
        ("COL_NAME",
         "学生氏名列",      "Student name col",
         "学生氏名が記録されている列",
         "Column containing the student's name",
         "col"),
        ("COL_ID",
         "学生番号列",      "Student ID col",
         "学生番号が記録されている列",
         "Column containing the student ID",
         "col"),
        ("COL_COMMENT",
         "コメント列",      "Comment col",
         "コメント本文が記録されている列",
         "Column containing the comment text",
         "col"),
    ]

    _ATT_PARAMS = [
        ("ATT_SKIP_ROWS",
         "データ開始行",    "Data start row",
         "学生データが始まる行番号（ヘッダー6行なら「7」）",
         "Row where student data begins (e.g. 7 if there are 6 header rows)",
         "row"),
        ("ATT_COL_ID",
         "学生番号列",      "Student ID col",
         "出席簿で学生番号が記録されている列",
         "Column with student ID in the attendance sheet",
         "col"),
        ("ATT_COL_NAME",
         "学生氏名列",      "Student name col",
         "出席簿で学生氏名が記録されている列",
         "Column with student name in the attendance sheet",
         "col"),
    ]

    def __init__(self, parent, lang, default_font, header_font):
        super().__init__(parent)
        self.lang         = lang
        self.default_font = default_font
        self.header_font  = header_font
        self._vars        = {}   # key → StringVar

        jp = (lang == "JP")
        self.title(("⚙  高度な設定" if jp else "⚙  Advanced Settings"))
        self.resizable(False, False)
        self.grab_set()   # modal
        self.focus_set()

        self._build(jp)

        # Centre over parent
        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h   = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w)//2}+{py + (ph - h)//2}")

    # ── UI construction ──────────────────────────────────────────────────────

    def _build(self, jp):
        pad = ttk.Frame(self, padding=14)
        pad.pack(fill=tk.BOTH, expand=True)

        # ── Comment sheet section ─────────────────────────────────────────
        cs_title = "コメントシートの列設定" if jp else "Comment Sheet — Column Settings"
        cs_frame = ttk.LabelFrame(pad, text=cs_title, padding="8 6 8 8")
        cs_frame.pack(fill=tk.X, pady=(0, 10))
        self._build_grid(cs_frame, self._CS_PARAMS, jp)

        # ── Attendance sheet section ──────────────────────────────────────
        att_title = "出席簿の設定" if jp else "Attendance Sheet Settings"
        att_frame = ttk.LabelFrame(pad, text=att_title, padding="8 6 8 8")
        att_frame.pack(fill=tk.X, pady=(0, 10))
        self._build_grid(att_frame, self._ATT_PARAMS, jp)

        # ── Divider + note ────────────────────────────────────────────────
        ttk.Separator(pad, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 6))
        note = ("⚠  変更は次回の集計から反映されます"
                if jp else
                "⚠  Changes take effect on the next aggregation run")
        ttk.Label(pad, text=note, foreground="#999999",
                  font=(self.default_font[0], self.default_font[1] - 1)
                  ).pack(anchor=tk.W, pady=(0, 8))

        # ── Button row ────────────────────────────────────────────────────
        btn_row = ttk.Frame(pad)
        btn_row.pack(fill=tk.X)

        ttk.Button(
            btn_row,
            text="デフォルトに戻す" if jp else "Reset to Defaults",
            command=self._reset,
        ).pack(side=tk.LEFT)

        ttk.Button(
            btn_row,
            text="  保存  " if jp else "  Save  ",
            command=self._save,
        ).pack(side=tk.RIGHT)

        ttk.Button(
            btn_row,
            text="キャンセル" if jp else "Cancel",
            command=self.destroy,
        ).pack(side=tk.RIGHT, padx=(0, 6))

    def _build_grid(self, parent, params, jp):
        """
        Render one section's parameters as a tidy label/widget/desc grid.
        """
        small_font = (self.default_font[0], max(self.default_font[1] - 1, 8))

        for row_idx, entry in enumerate(params):
            key, jp_lbl, en_lbl, jp_desc, en_desc, wtype = entry
            label = jp_lbl if jp else en_lbl
            desc  = jp_desc if jp else en_desc

            # Column 0 — parameter name
            ttk.Label(parent, text=label, width=16,
                      anchor=tk.W).grid(row=row_idx, column=0,
                                        sticky=tk.W, pady=3)

            # Column 1 — widget
            var = tk.StringVar()
            if wtype == "col":
                raw_idx = aggregator.CONFIG.get(key, 0)
                var.set(self._LETTERS[raw_idx] if raw_idx < 26 else str(raw_idx))
                w = ttk.Combobox(parent, textvariable=var,
                                 values=self._LETTERS,
                                 state="readonly", width=5)
            else:  # 'row' — display as human row number (skip+1)
                raw_skip = aggregator.CONFIG.get(key, 6)
                var.set(str(raw_skip + 1))
                w = ttk.Spinbox(parent, textvariable=var,
                                from_=2, to=50, width=5)

            w.grid(row=row_idx, column=1, sticky=tk.W, padx=(6, 10), pady=3)
            self._vars[key] = (var, wtype)

            # Column 2 — description
            ttk.Label(parent, text=desc, foreground="#777777",
                      font=small_font).grid(row=row_idx, column=2,
                                            sticky=tk.W, pady=3)

        parent.columnconfigure(2, weight=1)

    # ── Actions ──────────────────────────────────────────────────────────────

    def _reset(self):
        """Reset every widget to the factory default value."""
        defaults = aggregator._DEFAULT_CONFIG
        for key, (var, wtype) in self._vars.items():
            raw = defaults.get(key, 0)
            if wtype == "col":
                var.set(self._LETTERS[raw] if raw < 26 else str(raw))
            else:
                var.set(str(raw + 1))  # display as row number

    def _save(self):
        """Validate inputs, apply to CONFIG, and persist to disk."""
        jp     = (self.lang == "JP")
        errors = []
        new_cfg = dict(aggregator.CONFIG)

        for key, (var, wtype) in self._vars.items():
            raw_str = var.get().strip().upper()
            if wtype == "col":
                if len(raw_str) == 1 and raw_str.isalpha():
                    new_cfg[key] = ord(raw_str) - ord('A')
                else:
                    label = key
                    errors.append(
                        f"{label}: A〜Z の列を選択してください"
                        if jp else
                        f"{label}: please select a column A–Z"
                    )
            else:  # row spinbox
                try:
                    row_num = int(raw_str)
                    if row_num < 2:
                        raise ValueError
                    new_cfg[key] = row_num - 1   # store as skip count
                except ValueError:
                    lbl = "データ開始行" if jp else "Data start row"
                    errors.append(
                        f"{lbl}: 2 以上の整数を入力してください"
                        if jp else
                        f"{lbl}: must be an integer ≥ 2"
                    )

        if errors:
            messagebox.showerror(
                "入力エラー" if jp else "Input Error",
                "\n".join(errors),
                parent=self,
            )
            return

        # Auto-derive MIN_COLS so users never have to touch it
        new_cfg["MIN_COLS"] = max(
            new_cfg["COL_SUB_ID"], new_cfg["COL_COURSE"],
            new_cfg["COL_NAME"],   new_cfg["COL_ID"],
            new_cfg["COL_COMMENT"],
        ) + 1

        # Apply in-memory and persist
        aggregator.CONFIG.update(new_cfg)
        aggregator.save_config(_CONFIG_PATH, new_cfg)

        self.destroy()


# ===========================================================================
# Main Application
# ===========================================================================

class CommentAggregatorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x480")

        # OS Detection
        self.is_windows = platform.system() == "Windows"

        # Style Configuration
        self.style = ttk.Style()
        if self.is_windows:
            try:
                self.style.theme_use('vista')
            except Exception:
                self.style.theme_use('clam')
            self.default_font = ("Segoe UI", 9)
            self.header_font  = ("Segoe UI", 10, "bold")
        else:
            self.style.theme_use('clam')
            self.default_font = ("Helvetica", 11)
            self.header_font  = ("Helvetica", 12, "bold")

        self.style.configure('.',                font=self.default_font)
        self.style.configure('TButton',          font=self.default_font, padding=3)
        self.style.configure('TLabel',           font=self.default_font)
        self.style.configure('TLabelframe.Label',font=self.header_font,
                             foreground="#003399")
        self.style.configure('Small.TLabel',
                             font=(self.default_font[0], self.default_font[1] - 1))
        self.style.configure('Gear.TButton',     font=self.default_font, padding=2)

        # Language
        self.lang = "JP"

        # State
        self.selected_files  = []
        self.attendance_file = None

        # ── Main container ───────────────────────────────────────────────
        main_frame = ttk.Frame(root, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── Header row: title area + controls ────────────────────────────
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        # Right side: gear button + language switcher
        right_frame = ttk.Frame(header_frame)
        right_frame.pack(side=tk.RIGHT)

        self.btn_settings = ttk.Button(
            right_frame,
            style='Gear.TButton',
            command=self.open_settings,
        )
        self.btn_settings.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(right_frame, text="Language: ").pack(side=tk.LEFT)
        self.combo_lang = ttk.Combobox(
            right_frame,
            values=["日本語 (JP)", "English (EN)"],
            state="readonly", width=10,
        )
        self.combo_lang.current(0)
        self.combo_lang.pack(side=tk.LEFT)
        self.combo_lang.bind("<<ComboboxSelected>>", self.change_language)

        # ── Section 1: Input ─────────────────────────────────────────────
        self.frame_input = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_input.pack(fill=tk.X, pady=(0, 10))

        self.btn_select = ttk.Button(self.frame_input, command=self.select_files)
        self.btn_select.pack(anchor=tk.W)

        self.label_file_count = ttk.Label(self.frame_input, foreground="gray")
        self.label_file_count.pack(anchor=tk.W, pady=(2, 0))

        self.btn_attendance = ttk.Button(self.frame_input, command=self.select_attendance)
        self.btn_attendance.pack(anchor=tk.W, pady=(5, 0))

        self.label_attendance = ttk.Label(self.frame_input, foreground="gray")
        self.label_attendance.pack(anchor=tk.W, pady=(2, 0))

        # ── Section 2: Settings ──────────────────────────────────────────
        self.frame_settings = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_settings.pack(fill=tk.X, pady=(0, 10))

        frame_year = ttk.Frame(self.frame_settings)
        frame_year.pack(fill=tk.X)

        self.label_year = ttk.Label(frame_year)
        self.label_year.pack(side=tk.LEFT)

        self.entry_year = ttk.Entry(frame_year, width=8)
        self.entry_year.pack(side=tk.LEFT, padx=5)
        self.entry_year.insert(0, str(datetime.datetime.now().year))

        self.label_year_desc = ttk.Label(frame_year, foreground="gray",
                                         font=(self.default_font[0],
                                               self.default_font[1] - 1))
        self.label_year_desc.pack(side=tk.LEFT)

        # ── Section 3: Run ───────────────────────────────────────────────
        self.frame_action = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_action.pack(fill=tk.BOTH, expand=True)

        self.btn_run = ttk.Button(
            self.frame_action, command=self.run_aggregation, state=tk.DISABLED,
        )
        self.btn_run.pack(fill=tk.X, pady=(0, 5))

        self.log_area = scrolledtext.ScrolledText(
            self.frame_action, height=6, state='disabled',
            font=("Consolas", 9), bg="#f8f9fa", relief=tk.FLAT,
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

        # ── Footer ───────────────────────────────────────────────────────
        self.label_footer = ttk.Label(main_frame, style='Small.TLabel',
                                      foreground="gray")
        self.label_footer.pack(side=tk.BOTTOM, anchor=tk.E, pady=(5, 0))

        self.update_ui_text()
        self.log(TRANSLATIONS[self.lang]["ready"])

    # ── Language ─────────────────────────────────────────────────────────────

    def change_language(self, event=None):
        self.lang = "JP" if "JP" in self.combo_lang.get() else "EN"
        self.update_ui_text()

    def update_ui_text(self):
        t = TRANSLATIONS[self.lang]
        self.root.title(t["title"])
        self.frame_input.config(text=t["step1"])
        self.btn_select.config(text=t["select_files"])

        if self.selected_files:
            self.label_file_count.config(
                text=f"{len(self.selected_files)}{t['files_selected']}")
        else:
            self.label_file_count.config(text=t["no_files"])

        self.btn_attendance.config(text=t["select_att"])
        if self.attendance_file:
            self.label_attendance.config(
                text=f"{t['att_selected']}{os.path.basename(self.attendance_file)}")
        else:
            self.label_attendance.config(text=t["no_att"])

        self.frame_settings.config(text=t["step2"])
        self.label_year.config(text=t["target_year"])
        self.label_year_desc.config(text=t["year_desc"])

        self.frame_action.config(text=t["step3"])
        self.btn_run.config(text=t["run"])
        self.btn_settings.config(text=t["adv_settings_btn"])
        self.label_footer.config(text=t["footer"])

    # ── Settings dialog ───────────────────────────────────────────────────────

    def open_settings(self):
        AdvancedSettingsDialog(
            self.root, self.lang,
            self.default_font, self.header_font,
        )

    # ── File selection ────────────────────────────────────────────────────────

    def select_files(self):
        t = TRANSLATIONS[self.lang]
        files = filedialog.askopenfilenames(
            title=t["select_files"],
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if files:
            self.selected_files = list(files)
            self.label_file_count.config(
                text=f"{len(self.selected_files)}{t['files_selected']}",
                foreground="#008800",
            )
            self.btn_run.config(state=tk.NORMAL)
            self.log(f"Selected {len(files)} files.")
        else:
            self.log(t["file_selection_cancelled"])

    def select_attendance(self):
        t = TRANSLATIONS[self.lang]
        file = filedialog.askopenfilename(
            title=t["select_att"],
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if file:
            self.attendance_file = file
            self.label_attendance.config(
                text=f"{t['att_selected']}{os.path.basename(file)}",
                foreground="#000088",
            )
            self.log(f"Selected attendance sheet: {os.path.basename(file)}")
        else:
            self.log(t["att_selection_cancelled"])

    # ── Run ───────────────────────────────────────────────────────────────────

    def run_aggregation(self):
        t = TRANSLATIONS[self.lang]
        if not self.selected_files:
            messagebox.showwarning(t["warning"], t["select_files_warn"])
            return

        target_year = self.entry_year.get().strip()
        if not target_year:
            if not messagebox.askyesno(t["year_empty_title"], t["year_empty_msg"]):
                return

        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="summary_output.xlsx",
            title=t["save_as"],
        )
        if not output_file:
            return

        self.btn_run.config(state=tk.DISABLED)
        self.log(f"{t['processing']}{os.path.basename(output_file)}")

        threading.Thread(
            target=self.process_thread,
            args=(output_file, target_year, self.attendance_file),
            daemon=True,
        ).start()

    def process_thread(self, output_file, target_year, attendance_file):
        t = TRANSLATIONS[self.lang]
        try:
            success, message = aggregator.process_files(
                self.selected_files, output_file, target_year, attendance_file,
            )
            if success:
                self.log(t["success"] + message)
                messagebox.showinfo(t["success_title"], t["success_msg"])
            else:
                self.log("FAILED: " + message)
                messagebox.showerror(t["error"], message)
        except Exception as e:
            self.log(t["critical_error"] + str(e))
            messagebox.showerror(t["error"], str(e))
        finally:
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))

    # ── Log ───────────────────────────────────────────────────────────────────

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app  = CommentAggregatorApp(root)
    root.mainloop()
