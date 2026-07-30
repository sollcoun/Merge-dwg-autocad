# -*- coding: utf-8 -*-
"""
merge_dwg_gui.py
=================
Графическое приложение для объединения .dwg файлов.
"""

import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import merge_dwg_com as core

try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".dwg_merger_settings.json")

# --------------------------------------------------------------------------
# Цвета
# --------------------------------------------------------------------------
COLORS = {
    "bg":           "#1e1e1e",
    "bg_card":      "#2a2a2a",
    "bg_input":     "#333333",
    "bg_hover":     "#3a3a3a",
    "bg_tree":      "#252525",
    "border":       "#3d3d3d",
    "accent":       "#3b82f6",
    "accent_hover": "#60a5fa",
    "success":      "#4ade80",
    "warning":      "#f59e0b",
    "error":        "#f87171",
    "text":         "#e5e5e5",
    "text_dim":     "#a3a3a3",
    "text_muted":   "#737373",
    "white_btn":    "#f5f5f5",
    "white_btn_fg": "#1a1a1a",
}

def get_desktop_path() -> str:
    return os.path.join(os.path.expanduser("~"), "Desktop")

def load_settings() -> dict:
    defaults = {
        "last_source_dir": "",
        "last_output_path": "",
        "paperspace": False,
        "visible": True,
        "include_images": True,
        "show_log": True,
    }
    if os.path.isfile(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            defaults.update(saved)
        except Exception:
            pass
    return defaults

def save_settings(settings: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# --------------------------------------------------------------------------
# Подсказки
# --------------------------------------------------------------------------
class Tooltip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left",
                 bg="#2a2a2a", fg=COLORS["text"], font=("Segoe UI", 9),
                 padx=10, pady=6, wraplength=260).pack()

    def _hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

# --------------------------------------------------------------------------
# Главное окно
# --------------------------------------------------------------------------
class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DWG Merger")
        self.geometry("980x740")
        self.minsize(940, 720)
        self.configure(bg=COLORS["bg"])

        self.settings = load_settings()
        self.output_path = tk.StringVar(value=self.settings.get("last_output_path", ""))
        self.paperspace_var = tk.BooleanVar(value=self.settings.get("paperspace", False))
        self.visible_var = tk.BooleanVar(value=self.settings.get("visible", True))
        self.include_images_var = tk.BooleanVar(value=self.settings.get("include_images", True))
        self.show_log_var = tk.BooleanVar(value=self.settings.get("show_log", True))

        self.items = {}
        self.merge_thread = None
        self.stop_flag = threading.Event()
        self.log_queue = queue.Queue()
        self.drop_queue = queue.Queue()
        self.dropdown = None

        self._setup_styles()
        self._build_ui()
        self._setup_drag_drop()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_log_queue)

        last_dir = self.settings.get("last_source_dir", "")
        if last_dir and os.path.isdir(last_dir):
            self._add_files_from_folder(last_dir)

        self._toggle_log_visibility()
        self.bind("<Button-1>", self._on_root_click, add="+")
        self.bind("<Configure>", self._on_configure)

    def _on_configure(self, event):
        """Закрываем настройки при изменении размера/положения главного окна"""
        if event.widget != self:
            return
        if self.dropdown and self.dropdown.winfo_exists():
            self._close_dropdown()

    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        c = COLORS
        font = ("Segoe UI", 10)
        font_sm = ("Segoe UI", 9)

        style.configure(".", background=c["bg"], foreground=c["text"], font=font)
        style.configure("TFrame", background=c["bg"])
        style.configure("Card.TFrame", background=c["bg_card"])
        style.configure("TLabel", background=c["bg"], foreground=c["text"], font=font)
        style.configure("Dim.TLabel", background=c["bg"], foreground=c["text_dim"], font=font_sm)
        style.configure("Muted.TLabel", background=c["bg"], foreground=c["text_muted"], font=font_sm)
        style.configure("Title.TLabel", background=c["bg"], foreground=c["text"], font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", background=c["bg"], foreground=c["text_dim"], font=font_sm)
        style.configure("AccentSmall.TButton",
                background=c["accent"], foreground="#ffffff",
                bordercolor=c["accent"], lightcolor=c["accent"],
                darkcolor=c["accent"], font=("Segoe UI", 10, "bold"),
                padding=(16, 8))
        style.map("AccentSmall.TButton",
                background=[("active", c["accent_hover"]), ("disabled", "#555555")],
                foreground=[("disabled", "#999999")])

        style.configure("TButton",
                        background=c["bg_input"], foreground=c["text"],
                        bordercolor=c["border"], lightcolor=c["bg_input"],
                        darkcolor=c["bg_input"], font=font, padding=(16, 9))
        style.map("TButton",
                  background=[("active", c["bg_hover"]), ("disabled", c["bg_card"])],
                  foreground=[("disabled", c["text_muted"])])

        style.configure("White.TButton",
                        background=c["white_btn"], foreground=c["white_btn_fg"],
                        bordercolor=c["white_btn"], lightcolor=c["white_btn"],
                        darkcolor=c["white_btn"], font=("Segoe UI", 11, "bold"), padding=(28, 12))
        style.map("White.TButton",
                  background=[("active", "#e0e0e0"), ("disabled", "#555555")],
                  foreground=[("disabled", "#999999")])

        style.configure("Ghost.TButton",
                        background=c["bg"], foreground=c["text_dim"],
                        bordercolor=c["border"], font=font, padding=(14, 8))
        style.map("Ghost.TButton",
                  background=[("active", c["bg_hover"])],
                  foreground=[("active", c["text"])])

        style.configure("Danger.TButton",
                        background="#2a1f1f", foreground=c["error"],
                        bordercolor="#4a2a2a", font=font, padding=(14, 8))
        style.map("Danger.TButton",
                  background=[("active", "#3a2a2a")])

        style.configure("TEntry",
                        fieldbackground=c["bg_input"], foreground=c["text"],
                        insertcolor=c["text"], bordercolor=c["border"],
                        padding=8)
        style.map("TEntry",
                  fieldbackground=[("focus", c["bg_hover"])],
                  bordercolor=[("focus", c["accent"])])

        style.configure("TProgressbar",
                        background=c["accent"], troughcolor="#333333",
                        thickness=4)

        style.configure("Treeview",
                        background=c["bg_tree"], foreground=c["text"],
                        fieldbackground=c["bg_tree"], bordercolor=c["border"],
                        font=font, rowheight=42)
        style.configure("Treeview.Heading",
                        background=c["bg_input"], foreground=c["text_muted"],
                        font=font_sm, relief="flat")
        style.map("Treeview",
                  background=[("selected", "#1e3a5f")],
                  foreground=[("selected", c["text"])])

        style.configure("Vertical.TScrollbar",
                        background=c["bg_input"], troughcolor=c["bg_tree"],
                        arrowcolor=c["text_muted"], relief="flat")

    def _build_ui(self):
        c = COLORS

        # ========== HEADER ==========
        header = ttk.Frame(self, padding=(24, 18, 24, 6))
        header.pack(fill="x")

        left = ttk.Frame(header)
        left.pack(side="left")

        logo_frame = tk.Frame(left, bg=c["accent"], width=28, height=28)
        logo_frame.pack(side="left", padx=(0, 12))
        logo_frame.pack_propagate(False)
        tk.Label(logo_frame, text="⧉", bg=c["accent"], fg="white",
                font=("Segoe UI", 13, "bold")).place(relx=0.5, rely=0.5, anchor="center")

        title_box = ttk.Frame(left)
        title_box.pack(side="left")
        ttk.Label(title_box, text="DWG Merger", style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_box, text="Объединение чертежей", style="Sub.TLabel").pack(anchor="w")

        self.settings_btn = ttk.Button(header, text="  ⚙  Настройки", style="Ghost.TButton",
                                    command=self._toggle_dropdown)
        self.settings_btn.pack(side="right")
        Tooltip(self.settings_btn, "Открыть настройки")

        # ========== TOOLBAR ==========
        toolbar = ttk.Frame(self, padding=(24, 12, 24, 4))
        toolbar.pack(fill="x")

        for text, cmd, tip in [
            ("  📁  Папка", self._on_select_folder, "Добавить все .dwg из папки"),
            ("  📄  Файлы", self._on_add_files, "Выбрать отдельные файлы"),
            ("  ✕  Убрать", self._on_remove_selected, "Убрать выделенные"),
            ("  🗑  Очистить", self._on_clear_list, "Очистить весь список"),
        ]:
            b = ttk.Button(toolbar, text=text, command=cmd)
            b.pack(side="left", padx=(0, 8))
            Tooltip(b, tip)

        # ========== DRAG ZONE ==========
        self.drop_zone = tk.Frame(self, bg=c["bg"], highlightbackground=c["border"],
                                highlightthickness=1, highlightcolor=c["accent"])
        self.drop_zone.pack(fill="x", padx=24, pady=(10, 6))

        drop_inner = tk.Frame(self.drop_zone, bg=c["bg"], pady=12)
        drop_inner.pack(fill="x")
        tk.Label(drop_inner, text="↑  Перетащите файлы сюда, чтобы добавить в список",
                bg=c["bg"], fg=c["text_dim"], font=("Segoe UI", 10)).pack()

        # ========== НИЖНЯЯ ЧАСТЬ — ПАКУЕМ ПЕРВОЙ (side="bottom") ==========
        bottom = ttk.Frame(self)
        bottom.pack(side="bottom", fill="x")

        # Одна линия: кнопки + прогресс-бар
        run_bar = ttk.Frame(bottom, padding=(24, 10, 24, 6))
        run_bar.pack(fill="x")

        self.run_btn = ttk.Button(run_bar, text="  ▷  Объединить", style="AccentSmall.TButton",
                                command=self._on_run)
        self.run_btn.pack(side="left")
        Tooltip(self.run_btn, "Запустить объединение")

        self.stop_btn = ttk.Button(run_bar, text="  Стоп", style="Danger.TButton",
                                command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.open_folder_btn = ttk.Button(run_bar, text="  📂  Открыть папку",
                                        command=self._on_open_output_folder,
                                        state="disabled", style="Ghost.TButton")
        self.open_folder_btn.pack(side="left", padx=(8, 0))

        # Прогресс-бар справа на той же линии
        self.progress = ttk.Progressbar(run_bar, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(16, 0), ipady=1)

        # Log под ними
        self.log_card = ttk.Frame(bottom, style="Card.TFrame", padding=(12, 8))
        self.log_card.pack(fill="x", padx=24, pady=(2, 12))

        ttk.Label(self.log_card, text="Что происходит", style="Dim.TLabel").pack(anchor="w", pady=(0, 3))

        self.log_text = tk.Text(
            self.log_card, height=4, state="disabled", wrap="word",
            bg=c["bg_tree"], fg=c["text"], insertbackground=c["text"],
            selectbackground="#1e3a5f", font=("Consolas", 9),
            relief="flat", padx=8, pady=5)
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_configure("error", foreground=c["error"])
        self.log_text.tag_configure("warning", foreground=c["warning"])
        self.log_text.tag_configure("info", foreground=c["text_dim"])
        self.log_text.tag_configure("success", foreground=c["success"])

        # ========== FILE LIST (expand=True) — пакуем ПОСЛЕ bottom ==========
        list_header = ttk.Frame(self, padding=(24, 8, 24, 2))
        list_header.pack(fill="x")
        self.list_title = ttk.Label(list_header, text="Что объединяем", style="Dim.TLabel")
        self.list_title.pack(side="left")

        list_card = ttk.Frame(self, style="Card.TFrame")
        list_card.pack(fill="both", expand=True, padx=24, pady=(0, 4))

        tree_frame = ttk.Frame(list_card)
        tree_frame.pack(fill="both", expand=True, padx=2, pady=2)

        columns = ("included", "name", "type", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")

        self.tree.heading("included", text="")
        self.tree.heading("name", text="Файл")
        self.tree.heading("type", text="Тип")
        self.tree.heading("status", text="Статус")

        self.tree.column("included", width=40, anchor="center", stretch=False)
        self.tree.column("name", width=520, anchor="w")
        self.tree.column("type", width=90, anchor="center", stretch=False)
        self.tree.column("status", width=120, anchor="center", stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self._on_tree_click)

        self.tree.tag_configure("ok", foreground=c["success"])
        self.tree.tag_configure("fail", foreground=c["error"])
        self.tree.tag_configure("pending", foreground=c["text_dim"])
        self.tree.tag_configure("working", foreground=c["warning"])
        self.tree.tag_configure("excluded", foreground=c["text_muted"])
        self.tree.tag_configure("empty", foreground=c["text_muted"])

        self._refresh_empty_state()

    # ------------------------------------------------------------------
    # Выпадающие настройки с нормальными галочками ✓
    # ------------------------------------------------------------------
    def _toggle_dropdown(self):
        if self.dropdown and self.dropdown.winfo_exists():
            self._close_dropdown()
        else:
            self._open_dropdown()

    def _open_dropdown(self):
        if self.dropdown and self.dropdown.winfo_exists():
            return

        drop = tk.Toplevel(self)
        drop.wm_overrideredirect(True)
        drop.configure(bg=COLORS["border"])
        drop.attributes("-topmost", True)

        outer = tk.Frame(drop, bg=COLORS["border"], padx=1, pady=1)
        outer.pack(fill="both", expand=True)

        card = tk.Frame(outer, bg=COLORS["bg_card"], padx=22, pady=18)
        card.pack(fill="both", expand=True)

        # Заголовок
        tk.Label(card, text="Настройки объединения",
                bg=COLORS["bg_card"], fg=COLORS["text"],
                font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(card, text="Куда сохранить и что показывать в процессе",
                bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 14))

        # Путь
        tk.Label(card, text="КУДА СОХРАНЯТЬ РЕЗУЛЬТАТ",
                bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 5))

        path_row = tk.Frame(card, bg=COLORS["bg_card"])
        path_row.pack(fill="x", pady=(0, 2))

        entry = ttk.Entry(path_row, textvariable=self.output_path)
        entry.pack(side="left", fill="x", expand=True, ipady=3)

        ttk.Button(path_row, text="  Обзор", command=self._on_browse_output,
                style="Ghost.TButton").pack(side="left", padx=(8, 0))

        tk.Label(card, text="ⓘ  Если поле пустое — результат сохранится на рабочий стол",
                bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 14))

        # Параметры
        tk.Label(card, text="ПАРАМЕТРЫ",
                bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 8))

        def make_check(parent, var, title, desc, command=None):
            row = tk.Frame(parent, bg=COLORS["bg_card"])
            row.pack(fill="x", pady=(0, 9))

            def toggle():
                var.set(not var.get())
                update_mark()
                if command:
                    command()

            def update_mark():
                if var.get():
                    mark_lbl.config(text="✓", fg=COLORS["accent"])
                else:
                    mark_lbl.config(text="☐", fg=COLORS["text_muted"])

            mark_lbl = tk.Label(row, text="☐", bg=COLORS["bg_card"],
                                fg=COLORS["text_muted"], font=("Segoe UI", 13),
                                width=2, cursor="hand2")
            mark_lbl.pack(side="left", padx=(0, 6))
            mark_lbl.bind("<Button-1>", lambda e: toggle())

            text_frame = tk.Frame(row, bg=COLORS["bg_card"])
            text_frame.pack(side="left", fill="x")

            title_lbl = tk.Label(text_frame, text=title, bg=COLORS["bg_card"],
                                fg=COLORS["text"], font=("Segoe UI", 10), cursor="hand2")
            title_lbl.pack(anchor="w")
            title_lbl.bind("<Button-1>", lambda e: toggle())

            tk.Label(text_frame, text=desc, bg=COLORS["bg_card"],
                    fg=COLORS["text_muted"], font=("Segoe UI", 8)).pack(anchor="w")

            update_mark()

        make_check(card, self.paperspace_var,
                "Листы (Paper Space)",
                "Включить компоновки листов в итоговый файл")

        make_check(card, self.visible_var,
                "Показывать окно AutoCAD",
                "Видно, что программа делает в реальном времени")

        make_check(card, self.include_images_var,
                "Вставлять картинки (BMP+BPW)",
                "Переносить растровые вставки из исходных чертежей")

        make_check(card, self.show_log_var,
                "Показывать журнал «Что происходит»",
                "Подробный лог для диагностики ошибок",
                command=self._toggle_log_visibility)

        # Кнопки
        btn_row = tk.Frame(card, bg=COLORS["bg_card"])
        btn_row.pack(fill="x", pady=(14, 0))

        ttk.Button(btn_row, text="Отмена", style="Ghost.TButton",
                command=self._close_dropdown).pack(side="left")
        ttk.Button(btn_row, text="Готово", style="White.TButton",
                command=self._close_dropdown).pack(side="right")

    # ========== ПОЗИЦИОНИРОВАНИЕ ==========
        self.update_idletasks()
        drop.update_idletasks()

        drop_width = 360
        drop_height = card.winfo_reqheight() + 6

        # Точные координаты кнопки
        btn_x = self.settings_btn.winfo_rootx()
        btn_y = self.settings_btn.winfo_rooty()
        btn_w = self.settings_btn.winfo_width()
        btn_h = self.settings_btn.winfo_height()

        # Правый край выпадающего окна = правый край кнопки
        bx = btn_x + btn_w - drop_width
        by = btn_y + btn_h + 6

        # Небольшое ограничение, чтобы не уезжало слишком далеко влево
        if bx < btn_x - 300:
            bx = btn_x - 20

        drop.geometry(f"{drop_width}x{drop_height}+{bx}+{by}")
        self.dropdown = drop

    def _close_dropdown(self):
        if self.dropdown and self.dropdown.winfo_exists():
            self._save_current_settings()
            self.dropdown.destroy()
        self.dropdown = None

    def _on_root_click(self, event):
        if not self.dropdown or not self.dropdown.winfo_exists():
            return
        if event.widget == self.settings_btn:
            return
        try:
            dx = self.dropdown.winfo_rootx()
            dy = self.dropdown.winfo_rooty()
            dw = self.dropdown.winfo_width()
            dh = self.dropdown.winfo_height()
            if not (dx <= event.x_root <= dx + dw and dy <= event.y_root <= dy + dh):
                self._close_dropdown()
        except Exception:
            self._close_dropdown()

    def _toggle_log_visibility(self):
        if self.show_log_var.get():
            self.log_card.pack(fill="x", padx=24, pady=(2, 14))
        else:
            self.log_card.pack_forget()

    # ------------------------------------------------------------------
    # Drag-and-drop
    # ------------------------------------------------------------------
    def _setup_drag_drop(self):
        if not HAS_WINDND:
            return
        try:
            windnd.hook_dropfiles(self, func=self._on_drop)
        except Exception:
            pass

    def _on_drop(self, paths):
        normalized = []
        for p in paths:
            if isinstance(p, bytes):
                try:
                    p = p.decode(sys.getfilesystemencoding() or "mbcs")
                except UnicodeDecodeError:
                    try:
                        p = p.decode("utf-8")
                    except UnicodeDecodeError:
                        p = p.decode("mbcs", errors="replace")
            normalized.append(os.path.normpath(p))
        if normalized:
            try:
                self.drop_queue.put(normalized)
            except Exception:
                pass

    def _process_dropped_paths(self, paths: list):
        folders, dwg_files, bmp_files, ignored = [], [], [], []
        for path in paths:
            if os.path.isdir(path):
                folders.append(path)
            elif os.path.isfile(path):
                lower = path.lower()
                if lower.endswith(".dwg"):
                    dwg_files.append(path)
                elif lower.endswith(".bmp"):
                    bmp_files.append(path)
                else:
                    ignored.append(os.path.basename(path))
            else:
                ignored.append(path)

        for folder in folders:
            self._add_files_from_folder(folder)
        if dwg_files:
            added = self._add_paths(dwg_files)
            if added:
                self._log(f"Перетащено .dwg файлов: {added}", "info")
        if bmp_files:
            pairs, missing = [], []
            for bmp_path in bmp_files:
                bpw_path = core.find_bpw_for_bmp(bmp_path)
                if bpw_path:
                    pairs.append((bmp_path, bpw_path))
                else:
                    missing.append(bmp_path)
            img_added = self._add_image_pairs(pairs) if pairs else 0
            if img_added:
                self._log(f"Перетащено изображений BMP+BPW: {img_added}", "info")
            if missing:
                messagebox.showinfo("Привязка не найдена",
                    "Для этих изображений рядом нет файла .bpw:\n\n" +
                    "\n".join(os.path.basename(p) for p in missing))
        if ignored:
            self._log("Пропущены: " + ", ".join(ignored[:8]), "warning")

    # ------------------------------------------------------------------
    # Список файлов
    # ------------------------------------------------------------------
    EMPTY_ROW_ID = "__empty_hint__"

    def _refresh_empty_state(self):
        has_ph = self.tree.exists(self.EMPTY_ROW_ID)
        count = len(self.items)
        if count == 0:
            self.list_title.config(text="Что объединяем")
        elif count == 1:
            self.list_title.config(text="Что объединяем — 1 файл")
        else:
            self.list_title.config(text=f"Что объединяем — {count} файлов")

        if self.items:
            if has_ph:
                self.tree.delete(self.EMPTY_ROW_ID)
            return
        if not has_ph:
            self.tree.insert("", "end", iid=self.EMPTY_ROW_ID,
                values=("", "Список пуст — добавьте файлы кнопками выше или перетащите сюда", "", ""),
                tags=("empty",))

    def _existing_paths(self) -> set:
        return {os.path.normcase(os.path.abspath(v["path"])) for v in self.items.values()}

    def _add_paths(self, paths: list):
        existing = self._existing_paths()
        duplicates, added = [], 0
        for path in paths:
            key = os.path.normcase(os.path.abspath(path))
            if key in existing:
                duplicates.append(path)
                continue
            display = f"{os.path.basename(path)}\n{os.path.dirname(path)}"
            item_id = self.tree.insert("", "end",
                values=("✓", display, "DWG", "готов"), tags=("pending",))
            self.items[item_id] = {"path": path, "included": True, "kind": "dwg"}
            existing.add(key)
            added += 1
        if duplicates:
            messagebox.showinfo("Уже в списке",
                "Эти файлы уже добавлены:\n\n" + "\n".join(duplicates[:8]))
        self._refresh_empty_state()
        return added

    def _add_image_pairs(self, pairs: list):
        existing = self._existing_paths()
        duplicates, added = [], 0
        for bmp_path, bpw_path in pairs:
            key = os.path.normcase(os.path.abspath(bmp_path))
            if key in existing:
                duplicates.append(bmp_path)
                continue
            display = f"{os.path.basename(bmp_path)}\n{os.path.dirname(bmp_path)}"
            item_id = self.tree.insert("", "end",
                values=("✓", display, "BMP+BPW", "привязка найдена"), tags=("pending",))
            self.items[item_id] = {"path": bmp_path, "bpw": bpw_path, "included": True, "kind": "image"}
            existing.add(key)
            added += 1
        if duplicates:
            messagebox.showinfo("Уже в списке",
                "Эти изображения уже добавлены:\n\n" + "\n".join(duplicates[:8]))
        self._refresh_empty_state()
        return added

    def _add_files_from_folder(self, folder: str):
        files = core.find_dwg_files(folder)
        added = self._add_paths(files)
        self.settings["last_source_dir"] = folder
        self._log(f"Из папки «{folder}» добавлено .dwg: {added}", "info")
        pairs, missing_bpw = core.find_bmp_bpw_pairs_verbose(folder)
        img_added = self._add_image_pairs(pairs)
        if img_added:
            self._log(f"Добавлено изображений BMP+BPW: {img_added}", "info")
        for bmp in missing_bpw:
            self._log(f"Пропущено {os.path.basename(bmp)} — нет .bpw", "warning")

    def _on_select_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с .dwg",
            initialdir=self.settings.get("last_source_dir") or None)
        if folder:
            self._add_files_from_folder(folder)

    def _on_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите .dwg или .bmp",
            initialdir=self.settings.get("last_source_dir") or None,
            filetypes=[("DWG и BMP", "*.dwg;*.bmp"), ("DWG", "*.dwg"),
                       ("BMP", "*.bmp"), ("Все", "*.*")])
        if not paths:
            return
        dwg = [p for p in paths if p.lower().endswith(".dwg")]
        bmp = [p for p in paths if p.lower().endswith(".bmp")]
        if dwg:
            self._log(f"Добавлено .dwg вручную: {self._add_paths(dwg)}", "info")
        if bmp:
            pairs, missing = [], []
            for b in bmp:
                bpw = core.find_bpw_for_bmp(b)
                if bpw:
                    pairs.append((b, bpw))
                else:
                    missing.append(b)
            if pairs:
                self._log(f"Добавлено BMP+BPW: {self._add_image_pairs(pairs)}", "info")
            if missing:
                messagebox.showinfo("Привязка не найдена",
                    "Нет .bpw для:\n\n" + "\n".join(os.path.basename(p) for p in missing))

    def _on_remove_selected(self):
        for item_id in self.tree.selection():
            if item_id == self.EMPTY_ROW_ID:
                continue
            self.tree.delete(item_id)
            self.items.pop(item_id, None)
        self._refresh_empty_state()

    def _on_clear_list(self):
        for item_id in list(self.items):
            self.tree.delete(item_id)
        self.items.clear()
        self._refresh_empty_state()

    def _on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if not item_id or item_id not in self.items:
            return

        # Работаем только с колонкой галочки
        if col != "#1":
            return

        # Новое состояние берём от того файла, по которому кликнули
        info = self.items[item_id]
        new_state = not info["included"]

        # Какие строки нужно обновить:
        # если кликнули по уже выделенной строке — меняем все выделенные
        # если кликнули по невыделенной — меняем только её
        selected = self.tree.selection()
        if item_id in selected and len(selected) > 1:
            items_to_update = [i for i in selected if i in self.items]
        else:
            items_to_update = [item_id]

        for iid in items_to_update:
            self.items[iid]["included"] = new_state
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "✓" if new_state else "☐"
            self.tree.item(iid, values=vals,
                        tags=("pending",) if new_state else ("excluded",))

        # Не даём Treeview менять selection при клике на галочку
        return "break"

    # ------------------------------------------------------------------
    # Сохранение / запуск
    # ------------------------------------------------------------------
    def _on_browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            defaultextension=".dwg",
            initialdir=os.path.dirname(self.output_path.get()) or get_desktop_path(),
            initialfile=os.path.basename(self.output_path.get()) or "merged.dwg",
            filetypes=[("DWG", "*.dwg")])
        if path:
            self.output_path.set(path)

    def _get_output_path(self) -> str:
        path = self.output_path.get().strip()
        if not path:
            path = os.path.join(get_desktop_path(), "merged.dwg")
            self.output_path.set(path)
        return path

    def _on_open_output_folder(self):
        folder = os.path.dirname(os.path.abspath(self._get_output_path()))
        if os.path.isdir(folder):
            os.startfile(folder)

    def _on_run(self):
        dwg_paths = [v["path"] for v in self.items.values()
                     if v["included"] and v.get("kind") == "dwg"]
        image_pairs = [(v["path"], v["bpw"]) for v in self.items.values()
                       if v["included"] and v.get("kind") == "image"]
        if not self.include_images_var.get():
            image_pairs = []

        output_path = self._get_output_path()

        try:
            warnings = core.validate_before_merge(dwg_paths, output_path, image_pairs=image_pairs)
        except core.MergeError as e:
            messagebox.showwarning("Не готово", str(e))
            return

        if warnings and not messagebox.askyesno("Проверьте", "\n\n".join(warnings) + "\n\nПродолжить?"):
            return

        self._save_current_settings()
        total = len(dwg_paths) + len(image_pairs)
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.open_folder_btn.config(state="disabled")
        self.progress.config(maximum=max(total, 1), value=0)
        self._set_all_status("в очереди", "pending")
        self._log(f"Запуск: {len(dwg_paths)} DWG + {len(image_pairs)} изображений → {output_path}", "info")

        self.stop_flag.clear()
        self.merge_thread = threading.Thread(
            target=self._run_merge_worker, args=(dwg_paths, image_pairs, output_path), daemon=True)
        self.merge_thread.start()

    def _on_stop(self):
        self.stop_flag.set()
        self._log("Останавливаю…", "warning")
        self.stop_btn.config(state="disabled")

    def _run_merge_worker(self, dwg_paths, image_pairs, output_path):
        path_to_item = {v["path"]: k for k, v in self.items.items()}

        def log_cb(msg, level):
            self.log_queue.put(("log", msg, level))

        def progress_cb(done, total, filename, success=True):
            self.log_queue.put(("progress", done, total, filename))
            for path in dwg_paths + [b for b, _ in image_pairs]:
                if os.path.basename(path) == filename:
                    item_id = path_to_item.get(path)
                    if item_id:
                        self.log_queue.put(("status", item_id, success))
                    break

        try:
            summary = core.run_merge(
                dwg_paths, output_path,
                paperspace=self.paperspace_var.get(),
                visible=self.visible_var.get(),
                log_callback=log_cb, progress_callback=progress_cb,
                stop_flag=self.stop_flag, image_pairs=image_pairs)
            self.log_queue.put(("done", summary))
        except core.MergeError as e:
            self.log_queue.put(("fatal", str(e)))
        except Exception as e:
            self.log_queue.put(("fatal", core.describe_com_error(e)))

    def _poll_log_queue(self):
        try:
            while True:
                event = self.log_queue.get_nowait()
                kind = event[0]
                if kind == "log":
                    self._log(event[1], event[2])
                elif kind == "progress":
                    self.progress.config(maximum=event[2], value=event[1])
                elif kind == "status":
                    self._mark_item_status(event[1], event[2])
                elif kind == "done":
                    self._on_merge_done(event[1])
                elif kind == "fatal":
                    self._on_merge_fatal(event[1])
        except queue.Empty:
            pass
        try:
            while True:
                self._process_dropped_paths(self.drop_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _mark_item_status(self, item_id, success=True):
        vals = list(self.tree.item(item_id, "values"))
        vals[3] = "обработан" if success else "ошибка"
        self.tree.item(item_id, values=vals, tags=("ok" if success else "fail",))

    def _set_all_status(self, text, tag):
        for item_id, info in self.items.items():
            if info["included"]:
                vals = list(self.tree.item(item_id, "values"))
                vals[3] = text
                self.tree.item(item_id, values=vals, tags=(tag,))

    def _on_merge_done(self, summary):
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.open_folder_btn.config(state="normal")
        self._log("=" * 48, "info")
        self._log(f"ГОТОВО • DWG: {summary['ok']} успешно, {summary['fail']} ошибок • "
                  f"объектов: {summary['total_objects']}", "success")
        img_ok = summary.get("images_ok", 0)
        img_fail = summary.get("images_fail", 0)
        if img_ok or img_fail:
            self._log(f"Изображений: {img_ok} вставлено, {img_fail} с ошибкой", "info")
        if summary["errors"]:
            for path, msg in summary["errors"]:
                self._log(f"  {os.path.basename(path)}: {msg}", "error")
            messagebox.showwarning("Готово с ошибками",
                f"Успешно: {summary['ok']}, ошибок: {summary['fail']}\nСмотрите журнал.")
        else:
            messagebox.showinfo("Готово",
                f"Успешно объединено {summary['ok']} DWG"
                + (f" и {img_ok} изображений" if img_ok else ""))

    def _on_merge_fatal(self, msg):
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._log(f"Ошибка: {msg}", "error")
        messagebox.showerror("Ошибка", msg)

    def _log(self, msg: str, level: str = "info"):
        if not self.show_log_var.get():
            return
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n", level)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _save_current_settings(self):
        self.settings.update({
            "last_output_path": self.output_path.get(),
            "paperspace": self.paperspace_var.get(),
            "visible": self.visible_var.get(),
            "include_images": self.include_images_var.get(),
            "show_log": self.show_log_var.get(),
        })
        save_settings(self.settings)

    def _on_close(self):
        if self.merge_thread and self.merge_thread.is_alive():
            if not messagebox.askyesno("Идёт объединение", "Закрыть программу?"):
                return
            self.stop_flag.set()
        self._close_dropdown()
        self._save_current_settings()
        self.destroy()

def main():
    if sys.platform != "win32":
        print("Только Windows + AutoCAD")
        sys.exit(1)
    app = MergeApp()
    app.mainloop()

if __name__ == "__main__":
    main()