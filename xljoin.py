import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import io
import os
import threading


# ═════════════════════════════════════════════════════════════════════════════
# THEME / COLOURS
# ═════════════════════════════════════════════════════════════════════════════
BG          = "#f0f4f8"
CARD_BG     = "#ffffff"
PRIMARY     = "#0066cc"
PRIMARY_LT  = "#e8f0fe"
SUCCESS     = "#198754"
WARNING     = "#ffc107"
DANGER      = "#dc3545"
TEXT        = "#212529"
MUTED       = "#6c757d"
BORDER      = "#dee2e6"

FONT        = ("Segoe UI", 10)
FONT_SM     = ("Segoe UI", 9)
FONT_LG     = ("Segoe UI", 12, "bold")
FONT_HEADER = ("Segoe UI", 11, "bold")
FONT_TITLE  = ("Segoe UI", 16, "bold")


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def make_card(parent, **kwargs):
    kw = dict(bg=CARD_BG, relief="flat", bd=0, padx=12, pady=10)
    kw.update(kwargs)
    f = tk.LabelFrame(parent, **kw)
    f.configure(
        font=FONT_HEADER,
        fg=PRIMARY,
        highlightbackground=BORDER,
        highlightthickness=1,
    )
    return f


def styled_btn(parent, text, command, primary=False, danger=False, **kwargs):
    bg  = PRIMARY  if primary else (DANGER if danger else CARD_BG)
    fg  = "white"  if (primary or danger) else PRIMARY
    abg = "#0052a3" if primary else ("#b02a37" if danger else PRIMARY_LT)
    btn = tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=abg, activeforeground="white",
        font=FONT, relief="flat", bd=0, padx=14, pady=6, cursor="hand2",
        **kwargs,
    )
    btn.bind("<Enter>", lambda e: btn.config(bg=abg, fg="white"))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg,  fg=fg))
    return btn


def separator(parent, pady=6):
    ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=pady)


def step_header(parent, number, title):
    f = tk.Frame(parent, bg=PRIMARY, pady=6)
    f.pack(fill="x", pady=(10, 6))
    tk.Label(
        f,
        text=f"  Step {number} — {title}  ",
        bg=PRIMARY, fg="white",
        font=FONT_HEADER,
    ).pack(side="left")


# ═════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═════════════════════════════════════════════════════════════════════════════
class ExcelJoinApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("🔗 Excel Join Tool")
        self.state("zoomed")          # ← maximized on launch
        self.minsize(900, 700)
        self.configure(bg=BG)

        # ── State ─────────────────────────────────────────────────────────────
        self.df_left            = None
        self.df_right           = None
        self.df_result          = None
        self.left_bytes         = None
        self.right_bytes        = None
        self.left_sheets        = []
        self.right_sheets       = []
        self.join_type          = tk.StringVar(value="inner")
        self.left_file_label    = tk.StringVar(value="No file selected")
        self.right_file_label   = tk.StringVar(value="No file selected")
        self.left_sheet_var     = tk.StringVar()
        self.right_sheet_var    = tk.StringVar()
        self.left_skip_var      = tk.IntVar(value=0)
        self.right_skip_var     = tk.IntVar(value=0)
        self.output_name_var    = tk.StringVar(value="joined_result")
        self.status_var         = tk.StringVar(value="Ready.")
        self.mapping_rows       = []
        self.left_col_vars      = {}
        self.right_col_vars     = {}
        self.result_label_var   = tk.StringVar(value="")

        self._build_ui()

    # ─────────────────────────────────────────────────────────────────────────
    # UI BUILD
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Title bar
        title_bar = tk.Frame(self, bg=PRIMARY, pady=10)
        title_bar.pack(fill="x")
        tk.Label(
            title_bar,
            text="  🔗  Excel Join Tool",
            bg=PRIMARY, fg="white",
            font=FONT_TITLE,
        ).pack(side="left")
        tk.Label(
            title_bar,
            text="Upload · Configure · Export  ",
            bg=PRIMARY, fg="#cce0ff",
            font=FONT_SM,
        ).pack(side="right")

        # Scrollable canvas
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        sb     = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.scroll_frame = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(e):
            canvas.itemconfig(win_id, width=e.width)

        self.scroll_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"),
        )

        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()

        # Status bar
        status_bar = tk.Frame(self, bg=BORDER, pady=4)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(
            status_bar,
            textvariable=self.status_var,
            bg=BORDER, fg=MUTED,
            font=FONT_SM, anchor="w",
        ).pack(fill="x", padx=10)

    # ── STEP 1 ────────────────────────────────────────────────────────────────
    def _build_step1(self):
        step_header(self.scroll_frame, 1, "Upload Excel Files")

        row = tk.Frame(self.scroll_frame, bg=BG)
        row.pack(fill="x", expand=True, padx=12, pady=4)
        # uniform="half" forces both columns to be exactly equal width
        row.columnconfigure(0, weight=1, uniform="half")
        row.columnconfigure(1, weight=1, uniform="half")
        row.rowconfigure(0, weight=1)

        self._file_card(row, "left",  0)
        self._file_card(row, "right", 1)

    def _file_card(self, parent, side, col):
        color = "🔵" if side == "left" else "🟠"
        label = side.capitalize()

        card = make_card(parent, text=f"  {color} {label} File  ")
        card.grid(row=0, column=col, padx=8, pady=4, sticky="nsew")
        # Ensure the card stretches inside its grid cell
        parent.rowconfigure(0, weight=1)

        # Make internal columns of card expand to fill card width
        card.columnconfigure(0, weight=1)

        # ── File picker row ───────────────────────────────────────────────────
        pick_row = tk.Frame(card, bg=CARD_BG)
        pick_row.pack(fill="x", pady=(0, 6))
        pick_row.columnconfigure(0, weight=1)

        lbl_var = self.left_file_label if side == "left" else self.right_file_label
        tk.Label(
            pick_row,
            textvariable=lbl_var,
            bg=CARD_BG, fg=MUTED,
            font=FONT_SM,
            wraplength=500,
            anchor="w",
            justify="left",
        ).pack(side="left", fill="x", expand=True)

        styled_btn(
            pick_row, "📂 Browse…",
            command=lambda s=side: self._browse_file(s),
            primary=True,
        ).pack(side="right")

        # ── Worksheet selector ────────────────────────────────────────────────
        tk.Label(
            card, text="Worksheet:",
            bg=CARD_BG, fg=TEXT, font=FONT,
        ).pack(anchor="w", pady=(4, 0))

        sheet_var = self.left_sheet_var if side == "left" else self.right_sheet_var
        cb = ttk.Combobox(
            card, textvariable=sheet_var,
            state="readonly", font=FONT,
        )
        cb.pack(fill="x", pady=(0, 6))          # fill="x" → full card width
        cb.bind(
            "<<ComboboxSelected>>",
            lambda e, s=side: self._reload_file(s),
        )
        if side == "left":
            self.left_sheet_cb = cb
        else:
            self.right_sheet_cb = cb

        # ── Skip rows ─────────────────────────────────────────────────────────
        skip_row = tk.Frame(card, bg=CARD_BG)
        skip_row.pack(fill="x", pady=(0, 6))

        tk.Label(
            skip_row, text="⏩ Skip first N rows:",
            bg=CARD_BG, fg=TEXT, font=FONT,
        ).pack(side="left")

        skip_var = self.left_skip_var if side == "left" else self.right_skip_var
        tk.Spinbox(
            skip_row, from_=0, to=999, width=5,
            textvariable=skip_var, font=FONT,
            command=lambda s=side: self._reload_file(s),
        ).pack(side="left", padx=8)

        tk.Label(
            skip_row, text="(row N+1 becomes header)",
            bg=CARD_BG, fg=MUTED, font=FONT_SM,
        ).pack(side="left")

        # ── Info label ────────────────────────────────────────────────────────
        info_lbl = tk.Label(
            card, text="",
            bg=CARD_BG, fg=SUCCESS, font=FONT_SM,
            wraplength=500, justify="left", anchor="w",
        )
        info_lbl.pack(fill="x", pady=(2, 4))
        if side == "left":
            self.left_info_lbl = info_lbl
        else:
            self.right_info_lbl = info_lbl

        # ── Mini preview treeview ─────────────────────────────────────────────
        tv_frame = tk.Frame(card, bg=CARD_BG)
        tv_frame.pack(fill="both", expand=True)

        tv      = ttk.Treeview(tv_frame, height=5, show="headings")
        tv_sb_x = ttk.Scrollbar(tv_frame, orient="horizontal", command=tv.xview)
        tv_sb_y = ttk.Scrollbar(tv_frame, orient="vertical",   command=tv.yview)
        tv.configure(xscrollcommand=tv_sb_x.set, yscrollcommand=tv_sb_y.set)
        tv_sb_y.pack(side="right",  fill="y")
        tv_sb_x.pack(side="bottom", fill="x")
        tv.pack(fill="both", expand=True)

        if side == "left":
            self.left_preview_tv = tv
        else:
            self.right_preview_tv = tv

    # ── STEP 2 ────────────────────────────────────────────────────────────────
    def _build_step2(self):
        step_header(self.scroll_frame, 2, "Select Join Type")

        card = make_card(self.scroll_frame, text="  Join Type  ")
        card.pack(fill="x", padx=12, pady=4)

        join_row = tk.Frame(card, bg=CARD_BG)
        join_row.pack(fill="x")

        JOIN_OPTIONS = [
            ("inner", "Inner Join",      "Only matching rows\nfrom both files"  ),
            ("left",  "Left Join",       "All left rows +\nmatched right rows"  ),
            ("right", "Right Join",      "All right rows +\nmatched left rows"  ),
            ("outer", "Full Outer Join", "All rows from\nboth files"            ),
        ]

        self._join_btns = {}
        for jt, jlabel, jdesc in JOIN_OPTIONS:
            f = tk.Frame(
                join_row, bg=CARD_BG,
                relief="solid", bd=1,
                padx=10, pady=10,
                cursor="hand2",
            )
            f.pack(side="left", expand=True, fill="both", padx=6, pady=4)

            title_lbl = tk.Label(f, text=jlabel, bg=CARD_BG, fg=PRIMARY,
                                 font=("Segoe UI", 10, "bold"))
            title_lbl.pack()
            desc_lbl  = tk.Label(f, text=jdesc,  bg=CARD_BG, fg=MUTED,
                                 font=FONT_SM, justify="center")
            desc_lbl.pack(pady=(2, 6))
            sel_lbl   = tk.Label(f, text="",      bg=CARD_BG, fg=PRIMARY,
                                 font=FONT_SM)
            sel_lbl.pack()

            self._join_btns[jt] = (f, title_lbl, desc_lbl, sel_lbl)

            for widget in (f, title_lbl, desc_lbl, sel_lbl):
                widget.bind("<Button-1>", lambda e, j=jt: self._select_join(j))

        self._select_join("inner", init=True)

    def _select_join(self, jt, init=False):
        self.join_type.set(jt)
        for key, (f, tl, dl, sl) in self._join_btns.items():
            if key == jt:
                f.config(bg=PRIMARY_LT, relief="solid", bd=2)
                tl.config(bg=PRIMARY_LT, fg=PRIMARY)
                dl.config(bg=PRIMARY_LT)
                sl.config(bg=PRIMARY_LT, text="✔ Selected")
            else:
                f.config(bg=CARD_BG, relief="solid", bd=1)
                tl.config(bg=CARD_BG, fg=PRIMARY)
                dl.config(bg=CARD_BG)
                sl.config(bg=CARD_BG, text="")
        if not init:
            self._set_status(f"Join type set to: {jt.upper()}")

    # ── STEP 3 ────────────────────────────────────────────────────────────────
    def _build_step3(self):
        step_header(self.scroll_frame, 3, "Map Join Fields")

        self.mapping_card = make_card(self.scroll_frame, text="  Join Key Mappings  ")
        self.mapping_card.pack(fill="x", padx=12, pady=4)

        hdr = tk.Frame(self.mapping_card, bg=CARD_BG)
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="🔵 Left Column",  bg=CARD_BG, fg=PRIMARY,
                 font=FONT_HEADER, width=28, anchor="w").pack(side="left", padx=(0, 4))
        tk.Label(hdr, text="🔗",              bg=CARD_BG, fg=PRIMARY,
                 font=FONT_LG,     width=3              ).pack(side="left")
        tk.Label(hdr, text="🟠 Right Column", bg=CARD_BG, fg=PRIMARY,
                 font=FONT_HEADER, width=28, anchor="w").pack(side="left", padx=(4, 0))

        self.mapping_rows_frame = tk.Frame(self.mapping_card, bg=CARD_BG)
        self.mapping_rows_frame.pack(fill="x")

        btn_row = tk.Frame(self.mapping_card, bg=CARD_BG)
        btn_row.pack(fill="x", pady=(6, 0))
        styled_btn(btn_row, "➕ Add mapping row",
                   command=self._add_mapping_row).pack(side="left")

        self.mapping_status_lbl = tk.Label(
            self.mapping_card, text="",
            bg=CARD_BG, fg=SUCCESS, font=FONT_SM,
            anchor="w", justify="left",
        )
        self.mapping_status_lbl.pack(fill="x", pady=(6, 0))

        self._add_mapping_row()

    def _add_mapping_row(self):
        lv = tk.StringVar(value="— select —")
        rv = tk.StringVar(value="— select —")
        self.mapping_rows.append((lv, rv))
        idx = len(self.mapping_rows) - 1

        row_f = tk.Frame(self.mapping_rows_frame, bg=CARD_BG)
        row_f.pack(fill="x", pady=2)

        left_cols  = list(self.df_left.columns)  if self.df_left  is not None else []
        right_cols = list(self.df_right.columns) if self.df_right is not None else []

        l_cb = ttk.Combobox(row_f, textvariable=lv,
                             values=["— select —"] + left_cols,
                             state="readonly", font=FONT, width=28)
        l_cb.pack(side="left", padx=(0, 4))
        l_cb.bind("<<ComboboxSelected>>", lambda e: self._validate_mappings())

        tk.Label(row_f, text="🔗", bg=CARD_BG, fg=PRIMARY,
                 font=FONT_LG, width=3).pack(side="left")

        r_cb = ttk.Combobox(row_f, textvariable=rv,
                             values=["— select —"] + right_cols,
                             state="readonly", font=FONT, width=28)
        r_cb.pack(side="left", padx=(4, 8))
        r_cb.bind("<<ComboboxSelected>>", lambda e: self._validate_mappings())

        if not hasattr(self, "_mapping_cbs"):
            self._mapping_cbs = []
        self._mapping_cbs.append((l_cb, r_cb))

        if idx > 0:
            styled_btn(row_f, "🗑",
                       command=lambda f=row_f, i=idx: self._delete_mapping_row(f, i),
                       danger=True).pack(side="left")

    def _delete_mapping_row(self, frame, idx):
        frame.destroy()
        self.mapping_rows[idx] = (None, None)
        self._validate_mappings()

    def _validate_mappings(self):
        valid = self._get_valid_mappings()
        if valid:
            self.mapping_status_lbl.config(
                fg=SUCCESS,
                text="✅ " + "  |  ".join(
                    [f"{m['left']} = {m['right']}" for m in valid]
                ),
            )
        else:
            self.mapping_status_lbl.config(fg=MUTED, text="No valid mappings yet.")

    def _get_valid_mappings(self):
        valid = []
        for lv, rv in self.mapping_rows:
            if lv is None or rv is None:
                continue
            l, r = lv.get(), rv.get()
            if l and r and l != "— select —" and r != "— select —":
                valid.append({"left": l, "right": r})
        return valid

    def _refresh_mapping_dropdowns(self):
        if not hasattr(self, "_mapping_cbs"):
            return
        left_cols  = list(self.df_left.columns)  if self.df_left  is not None else []
        right_cols = list(self.df_right.columns) if self.df_right is not None else []
        for l_cb, r_cb in self._mapping_cbs:
            try:
                l_cb["values"] = ["— select —"] + left_cols
                r_cb["values"] = ["— select —"] + right_cols
            except tk.TclError:
                pass

    # ── STEP 4 ────────────────────────────────────────────────────────────────
    def _build_step4(self):
        step_header(self.scroll_frame, 4, "Select Fields to Include in Joined File")

        self.fields_card = make_card(self.scroll_frame, text="  Column Selection  ")
        self.fields_card.pack(fill="x", padx=12, pady=4)

        tk.Label(
            self.fields_card,
            text="🔑 Join key columns are always included. "
                 "Tick the additional columns you want in the output.",
            bg=CARD_BG, fg=MUTED, font=FONT_SM, anchor="w",
        ).pack(fill="x", pady=(0, 8))

        cols_row = tk.Frame(self.fields_card, bg=CARD_BG)
        cols_row.pack(fill="x")
        cols_row.columnconfigure(0, weight=1, uniform="half")
        cols_row.columnconfigure(1, weight=1, uniform="half")

        # ── Left panel ────────────────────────────────────────────────────────
        lp = tk.Frame(cols_row, bg=CARD_BG)
        lp.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        tk.Label(lp, text="🔵 Left File Columns",
                 bg=CARD_BG, fg=PRIMARY, font=FONT_HEADER).pack(anchor="w")

        lbtn_row = tk.Frame(lp, bg=CARD_BG)
        lbtn_row.pack(fill="x", pady=4)
        styled_btn(lbtn_row, "✅ All",
                   command=lambda: self._select_all_cols("left")).pack(side="left", padx=(0, 4))
        styled_btn(lbtn_row, "❌ Clear",
                   command=lambda: self._clear_cols("left")).pack(side="left")

        self.left_col_canvas = tk.Canvas(lp, bg=CARD_BG, highlightthickness=0, height=160)
        lsb = ttk.Scrollbar(lp, orient="vertical", command=self.left_col_canvas.yview)
        self.left_col_canvas.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.left_col_canvas.pack(fill="x", expand=True)
        self.left_col_inner = tk.Frame(self.left_col_canvas, bg=CARD_BG)
        self.left_col_canvas.create_window((0, 0), window=self.left_col_inner, anchor="nw")
        self.left_col_inner.bind(
            "<Configure>",
            lambda e: self.left_col_canvas.configure(
                scrollregion=self.left_col_canvas.bbox("all")
            ),
        )
        self.left_col_count_lbl = tk.Label(lp, text="", bg=CARD_BG, fg=MUTED, font=FONT_SM)
        self.left_col_count_lbl.pack(anchor="w")

        # ── Right panel ───────────────────────────────────────────────────────
        rp = tk.Frame(cols_row, bg=CARD_BG)
        rp.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        tk.Label(rp, text="🟠 Right File Columns",
                 bg=CARD_BG, fg=PRIMARY, font=FONT_HEADER).pack(anchor="w")

        rbtn_row = tk.Frame(rp, bg=CARD_BG)
        rbtn_row.pack(fill="x", pady=4)
        styled_btn(rbtn_row, "✅ All",
                   command=lambda: self._select_all_cols("right")).pack(side="left", padx=(0, 4))
        styled_btn(rbtn_row, "❌ Clear",
                   command=lambda: self._clear_cols("right")).pack(side="left")

        self.right_col_canvas = tk.Canvas(rp, bg=CARD_BG, highlightthickness=0, height=160)
        rsb = ttk.Scrollbar(rp, orient="vertical", command=self.right_col_canvas.yview)
        self.right_col_canvas.configure(yscrollcommand=rsb.set)
        rsb.pack(side="right", fill="y")
        self.right_col_canvas.pack(fill="x", expand=True)
        self.right_col_inner = tk.Frame(self.right_col_canvas, bg=CARD_BG)
        self.right_col_canvas.create_window((0, 0), window=self.right_col_inner, anchor="nw")
        self.right_col_inner.bind(
            "<Configure>",
            lambda e: self.right_col_canvas.configure(
                scrollregion=self.right_col_canvas.bbox("all")
            ),
        )
        self.right_col_count_lbl = tk.Label(rp, text="", bg=CARD_BG, fg=MUTED, font=FONT_SM)
        self.right_col_count_lbl.pack(anchor="w")

    def _populate_col_checkboxes(self, side):
        if side == "left":
            inner  = self.left_col_inner
            df     = self.df_left
            cvars  = self.left_col_vars
            canvas = self.left_col_canvas
            clbl   = self.left_col_count_lbl
        else:
            inner  = self.right_col_inner
            df     = self.df_right
            cvars  = self.right_col_vars
            canvas = self.right_col_canvas
            clbl   = self.right_col_count_lbl

        for w in inner.winfo_children():
            w.destroy()
        cvars.clear()

        if df is None:
            return

        for col in df.columns:
            var = tk.BooleanVar(value=True)
            cvars[col] = var
            tk.Checkbutton(
                inner, text=col, variable=var,
                bg=CARD_BG, fg=TEXT, font=FONT_SM,
                anchor="w", activebackground=PRIMARY_LT,
                command=lambda s=side: self._update_col_count(s),
            ).pack(fill="x", anchor="w")

        canvas.configure(scrollregion=canvas.bbox("all"))
        self._update_col_count(side)

    def _update_col_count(self, side):
        if side == "left":
            cvars = self.left_col_vars
            lbl   = self.left_col_count_lbl
            total = len(self.df_left.columns)  if self.df_left  is not None else 0
        else:
            cvars = self.right_col_vars
            lbl   = self.right_col_count_lbl
            total = len(self.df_right.columns) if self.df_right is not None else 0

        selected = sum(1 for v in cvars.values() if v.get())
        lbl.config(text=f"{selected} of {total} columns selected")

    def _select_all_cols(self, side):
        cvars = self.left_col_vars if side == "left" else self.right_col_vars
        for v in cvars.values():
            v.set(True)
        self._update_col_count(side)

    def _clear_cols(self, side):
        cvars     = self.left_col_vars  if side == "left" else self.right_col_vars
        join_keys = self._get_join_keys(side)
        for col, v in cvars.items():
            v.set(col in join_keys)
        self._update_col_count(side)

    def _get_join_keys(self, side):
        valid = self._get_valid_mappings()
        return [m["left"] if side == "left" else m["right"] for m in valid]

    def _get_selected_cols(self, side):
        cvars     = self.left_col_vars  if side == "left" else self.right_col_vars
        join_keys = self._get_join_keys(side)
        selected  = [c for c, v in cvars.items() if v.get()]
        for k in join_keys:
            if k not in selected:
                selected.insert(0, k)
        return selected

    # ── STEP 5 ────────────────────────────────────────────────────────────────
    def _build_step5(self):
        step_header(self.scroll_frame, 5, "Name & Download Result")

        card = make_card(self.scroll_frame, text="  Export  ")
        card.pack(fill="x", padx=12, pady=4)

        name_row = tk.Frame(card, bg=CARD_BG)
        name_row.pack(fill="x", pady=(0, 8))
        tk.Label(name_row, text="Output file name:",
                 bg=CARD_BG, fg=TEXT, font=FONT).pack(side="left")
        tk.Entry(name_row, textvariable=self.output_name_var,
                 font=FONT, width=40, relief="solid", bd=1).pack(side="left", padx=8)
        tk.Label(name_row, text=".xlsx",
                 bg=CARD_BG, fg=MUTED, font=FONT).pack(side="left")

        self.join_btn = styled_btn(
            card, "🚀  Execute Join",
            command=self._execute_join_thread,
            primary=True,
        )
        self.join_btn.pack(fill="x", pady=(0, 8))

        self.progress = ttk.Progressbar(card, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 6))

        self.result_lbl = tk.Label(
            card, textvariable=self.result_label_var,
            bg=CARD_BG, fg=SUCCESS, font=FONT,
            anchor="w", justify="left",
        )
        self.result_lbl.pack(fill="x")

        separator(card)
        tk.Label(card, text="👁 Result Preview (first 20 rows)",
                 bg=CARD_BG, fg=TEXT, font=FONT_HEADER).pack(anchor="w")

        preview_frame = tk.Frame(card, bg=CARD_BG)
        preview_frame.pack(fill="both", expand=True, pady=4)

        self.result_tv = ttk.Treeview(preview_frame, height=8, show="headings")
        rv_sb_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.result_tv.xview)
        rv_sb_y = ttk.Scrollbar(preview_frame, orient="vertical",   command=self.result_tv.yview)
        self.result_tv.configure(xscrollcommand=rv_sb_x.set, yscrollcommand=rv_sb_y.set)
        rv_sb_y.pack(side="right",  fill="y")
        rv_sb_x.pack(side="bottom", fill="x")
        self.result_tv.pack(fill="both", expand=True)

        separator(card)
        self.download_btn = styled_btn(
            card, "⬇️  Save Excel File…",
            command=self._save_file,
            primary=True,
        )
        self.download_btn.pack(fill="x", pady=(4, 0))
        self.download_btn.config(state="disabled")

    # ─────────────────────────────────────────────────────────────────────────
    # FILE LOADING
    # ─────────────────────────────────────────────────────────────────────────
    def _browse_file(self, side):
        path = filedialog.askopenfilename(
            title=f"Select {side.capitalize()} Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return

        with open(path, "rb") as f:
            file_bytes = f.read()

        fname = os.path.basename(path)

        if side == "left":
            self.left_bytes = file_bytes
            self.left_file_label.set(fname)
        else:
            self.right_bytes = file_bytes
            self.right_file_label.set(fname)

        try:
            xf     = pd.ExcelFile(io.BytesIO(file_bytes))
            sheets = xf.sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
            return

        if side == "left":
            self.left_sheets = sheets
            self.left_sheet_cb["values"] = sheets
            self.left_sheet_var.set(sheets[0])
        else:
            self.right_sheets = sheets
            self.right_sheet_cb["values"] = sheets
            self.right_sheet_var.set(sheets[0])

        self._reload_file(side)

    def _reload_file(self, side):
        if side == "left":
            file_bytes = self.left_bytes
            sheet      = self.left_sheet_var.get()
            skip       = self.left_skip_var.get()
            info_lbl   = self.left_info_lbl
            preview_tv = self.left_preview_tv
        else:
            file_bytes = self.right_bytes
            sheet      = self.right_sheet_var.get()
            skip       = self.right_skip_var.get()
            info_lbl   = self.right_info_lbl
            preview_tv = self.right_preview_tv

        if file_bytes is None or not sheet:
            return

        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=sheet,
                skiprows=int(skip) if skip > 0 else None,
                dtype=str,
            )
        except Exception as e:
            messagebox.showerror("Read Error", str(e))
            return

        if side == "left":
            self.df_left = df
        else:
            self.df_right = df

        skip_note = f" | skipped {skip} row(s)" if skip > 0 else ""
        info_lbl.config(
            fg=SUCCESS,
            text=(
                f"✅  {df.shape[0]:,} rows × {df.shape[1]} cols"
                f" | sheet: {sheet}{skip_note}"
            ),
        )

        self._populate_preview(preview_tv, df)
        self._populate_col_checkboxes(side)
        self._refresh_mapping_dropdowns()
        self._validate_mappings()
        self._set_status(
            f"{side.capitalize()} file loaded — "
            f"{df.shape[0]:,} rows × {df.shape[1]} cols."
        )

    def _populate_preview(self, tv, df):
        tv.delete(*tv.get_children())
        cols = list(df.columns)
        tv["columns"] = cols
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=max(80, len(str(c)) * 9), stretch=False)
        for _, row in df.head(10).iterrows():
            tv.insert("", "end", values=list(row))

    # ─────────────────────────────────────────────────────────────────────────
    # JOIN EXECUTION
    # ─────────────────────────────────────────────────────────────────────────
    def _execute_join_thread(self):
        t = threading.Thread(target=self._execute_join, daemon=True)
        t.start()

    def _execute_join(self):
        if self.df_left is None or self.df_right is None:
            messagebox.showwarning("Missing Files", "Please load both Excel files first.")
            return

        valid = self._get_valid_mappings()
        if not valid:
            messagebox.showwarning("No Mappings",
                                   "Please configure at least one join key mapping.")
            return

        self.join_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.progress.start(12)
        self._set_status("Joining data…")

        try:
            left_keys  = [m["left"]  for m in valid]
            right_keys = [m["right"] for m in valid]
            how        = self.join_type.get()

            l_cols   = self._get_selected_cols("left")
            r_cols   = self._get_selected_cols("right")
            df_l_sub = self.df_left[l_cols]
            df_r_sub = self.df_right[r_cols]

            if left_keys == right_keys:
                df_result = pd.merge(
                    df_l_sub, df_r_sub,
                    on=left_keys, how=how,
                    suffixes=("_left", "_right"),
                )
            else:
                df_result = pd.merge(
                    df_l_sub, df_r_sub,
                    left_on=left_keys, right_on=right_keys,
                    how=how, suffixes=("_left", "_right"),
                )

            self.df_result = df_result
            self.after(0, self._on_join_success)

        except Exception as e:
            self.after(0, lambda: self._on_join_error(str(e)))

    def _on_join_success(self):
        self.progress.stop()
        self.join_btn.config(state="normal")
        self.download_btn.config(state="normal")
        df = self.df_result
        self.result_label_var.set(
            f"✅  Join complete — {df.shape[0]:,} rows × {df.shape[1]} columns  "
            f"({self.join_type.get().upper()} JOIN)"
        )
        self.result_lbl.config(fg=SUCCESS)
        self._populate_preview(self.result_tv, df.head(20))
        self._set_status(
            f"Join complete: {df.shape[0]:,} rows × {df.shape[1]} cols."
        )

    def _on_join_error(self, msg):
        self.progress.stop()
        self.join_btn.config(state="normal")
        self.result_label_var.set(f"❌  Join failed: {msg}")
        self.result_lbl.config(fg=DANGER)
        self._set_status(f"Join failed: {msg}")
        messagebox.showerror("Join Error", msg)

    # ─────────────────────────────────────────────────────────────────────────
    # SAVE
    # ─────────────────────────────────────────────────────────────────────────
    def _save_file(self):
        if self.df_result is None:
            messagebox.showwarning("No Result", "Please execute the join first.")
            return

        default_name = (
            self.output_name_var.get().strip().replace(" ", "_") or "joined_result"
        ) + ".xlsx"

        path = filedialog.asksaveasfilename(
            title="Save Joined Excel File",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            self._set_status("Saving file…")
            df     = self.df_result
            cols   = list(df.columns)
            buffer = io.BytesIO()

            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Joined Data")
                ws = writer.sheets["Joined Data"]
                for i, col in enumerate(cols):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col)),
                    ) + 2
                    ws.set_column(i, i, min(max_len, 50))

            with open(path, "wb") as f:
                f.write(buffer.getvalue())

            self._set_status(f"File saved: {path}")
            messagebox.showinfo(
                "Saved",
                f"✅ File saved successfully!\n\n{path}\n\n"
                f"{df.shape[0]:,} rows × {df.shape[1]} columns",
            )
        except Exception as e:
            self._set_status(f"Save failed: {e}")
            messagebox.showerror("Save Error", str(e))

    # ─────────────────────────────────────────────────────────────────────────
    # UTILITY
    # ─────────────────────────────────────────────────────────────────────────
    def _set_status(self, msg):
        self.status_var.set(f"  {msg}")


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = ExcelJoinApp()
    app.mainloop()
