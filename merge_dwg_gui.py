# -*- coding: utf-8 -*-
"""
merge_dwg_gui.py
=================
Графическое приложение для объединения .dwg файлов (для сотрудников,
не работающих с командной строкой). Использует ту же проверенную логику
слияния, что и merge_dwg_com.py (функция run_merge), через AutoCAD COM.

Помимо .dwg, приложение автоматически находит в выбранной папке пары
изображений с привязкой "картинка.bmp" + "картинка.bpw" и вставляет их
в результирующий чертёж в правильном месте/масштабе/повороте — вручную
указывать координаты не нужно.

Поддерживается перетаскивание (drag-and-drop) папок и файлов .dwg / .bmp
прямо в окно приложения (требуется pip install windnd).

Требования: Windows, установленный AutoCAD, pip install pywin32
            (опционально: pip install windnd — для drag-and-drop).

Запуск:
    python merge_dwg_gui.py
    (или готовый merge_dwg_gui.exe, собранный через PyInstaller — см. README)
"""

import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Логика слияния берётся из уже отлаженного модуля
import merge_dwg_com as core

# Drag-and-drop файлов/папок (Windows). Если windnd не установлен —
# приложение работает как раньше, просто без перетаскивания.
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".dwg_merger_settings.json")

# --------------------------------------------------------------------------
# Цветовая схема (тёмный инженерный стиль)
# --------------------------------------------------------------------------
COLORS = {
    "bg":           "#1a1d23",
    "bg_card":      "#242830",
    "bg_input":     "#2c313a",
    "bg_hover":     "#353b47",
    "bg_tree":      "#1e2229",
    "border":       "#3a4150",
    "accent":       "#3d8bfd",
    "accent_hover": "#5a9fff",
    "accent_dim":   "#2a5a9e",
    "success":      "#3fb950",
    "warning":      "#d29922",
    "error":        "#f85149",
    "text":         "#e6edf3",
    "text_dim":     "#8b949e",
    "text_muted":   "#6e7681",
    "header_bg":    "#161b22",
    "progress":     "#3d8bfd",
    "progress_bg":  "#2c313a",
}


# --------------------------------------------------------------------------
# Хранение настроек между запусками
# --------------------------------------------------------------------------
def load_settings() -> dict:
    defaults = {
        "last_source_dir": "",
        "last_output_path": "",
        "paperspace": False,
        "visible": True,
        "include_images": True,
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
# Главное окно
# --------------------------------------------------------------------------
class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DWG Merger — объединение чертежей")
        self.geometry("960x700")
        self.minsize(820, 600)
        self.configure(bg=COLORS["bg"])

        self.settings = load_settings()
        self.output_path = tk.StringVar(value=self.settings.get("last_output_path", ""))
        self.paperspace_var = tk.BooleanVar(value=self.settings.get("paperspace", False))
        self.visible_var = tk.BooleanVar(value=self.settings.get("visible", True))
        self.include_images_var = tk.BooleanVar(value=self.settings.get("include_images", True))

        # items: {tree_id: {"path": str, "included": bool, "kind": "dwg"|"image",
        #                    "bpw": str (только для kind == "image")}}
        self.items = {}
        self.merge_thread = None
        self.stop_flag = threading.Event()
        self.log_queue = queue.Queue()
        self.drop_queue = queue.Queue()

        self._setup_styles()
        self._build_ui()
        self._setup_drag_drop()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_log_queue)

        # Восстанавливаем последнюю папку
        last_dir = self.settings.get("last_source_dir", "")
        if last_dir and os.path.isdir(last_dir):
            self._add_files_from_folder(last_dir)

    # ------------------------------------------------------------------
    # Стили ttk
    # ------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        c = COLORS
        font_ui = ("Segoe UI", 10)
        font_ui_sm = ("Segoe UI", 9)
        font_ui_bold = ("Segoe UI", 10, "bold")
        font_title = ("Segoe UI", 16, "bold")
        font_sub = ("Segoe UI", 9)

        style.configure(".", background=c["bg"], foreground=c["text"], font=font_ui)
        style.configure("TFrame", background=c["bg"])
        style.configure("Card.TFrame", background=c["bg_card"])
        style.configure("Header.TFrame", background=c["header_bg"])

        style.configure("TLabel", background=c["bg"], foreground=c["text"], font=font_ui)
        style.configure("Card.TLabel", background=c["bg_card"], foreground=c["text"])
        style.configure("Dim.TLabel", background=c["bg"], foreground=c["text_dim"], font=font_ui_sm)
        style.configure("CardDim.TLabel", background=c["bg_card"], foreground=c["text_dim"], font=font_ui_sm)
        style.configure("Title.TLabel", background=c["header_bg"], foreground=c["text"], font=font_title)
        style.configure("Sub.TLabel", background=c["header_bg"], foreground=c["text_dim"], font=font_sub)
        style.configure("Accent.TLabel", background=c["bg"], foreground=c["accent"], font=font_ui_sm)

        style.configure("TLabelframe", background=c["bg_card"], foreground=c["text"],
                         bordercolor=c["border"], relief="flat")
        style.configure("TLabelframe.Label", background=c["bg_card"], foreground=c["text_dim"],
                         font=font_ui_bold)

        # Кнопки
        style.configure("TButton", background=c["bg_input"], foreground=c["text"],
                         bordercolor=c["border"], lightcolor=c["bg_input"],
                         darkcolor=c["bg_input"], focuscolor=c["accent"],
                         font=font_ui, padding=(12, 6))
        style.map("TButton",
                  background=[("active", c["bg_hover"]), ("disabled", c["bg_card"])],
                  foreground=[("disabled", c["text_muted"])])

        style.configure("Accent.TButton", background=c["accent"], foreground="#ffffff",
                         bordercolor=c["accent"], lightcolor=c["accent"],
                         darkcolor=c["accent_dim"], font=font_ui_bold, padding=(16, 8))
        style.map("Accent.TButton",
                  background=[("active", c["accent_hover"]), ("disabled", c["accent_dim"])],
                  foreground=[("disabled", "#a0a8b0")])

        style.configure("Danger.TButton", background="#3d1f1f", foreground=c["error"],
                         bordercolor="#5a2a2a", font=font_ui, padding=(12, 6))
        style.map("Danger.TButton",
                  background=[("active", "#5a2a2a"), ("disabled", c["bg_card"])])

        # Чекбоксы
        style.configure("TCheckbutton", background=c["bg_card"], foreground=c["text"],
                         font=font_ui, focuscolor=c["accent"])
        style.map("TCheckbutton",
                  background=[("active", c["bg_card"])],
                  foreground=[("active", c["text"])])

        # Entry
        style.configure("TEntry", fieldbackground=c["bg_input"], foreground=c["text"],
                         insertcolor=c["text"], bordercolor=c["border"],
                         lightcolor=c["border"], darkcolor=c["border"],
                         padding=6)
        style.map("TEntry",
                  fieldbackground=[("focus", c["bg_hover"])],
                  bordercolor=[("focus", c["accent"])])

        # Progressbar
        style.configure("TProgressbar", background=c["progress"], troughcolor=c["progress_bg"],
                         bordercolor=c["border"], lightcolor=c["progress"],
                         darkcolor=c["progress"], thickness=8)

        # Treeview
        style.configure("Treeview",
                         background=c["bg_tree"], foreground=c["text"],
                         fieldbackground=c["bg_tree"], bordercolor=c["border"],
                         font=font_ui, rowheight=28)
        style.configure("Treeview.Heading",
                         background=c["bg_input"], foreground=c["text_dim"],
                         font=font_ui_bold, bordercolor=c["border"], relief="flat")
        style.map("Treeview",
                  background=[("selected", c["accent_dim"])],
                  foreground=[("selected", c["text"])])
        style.map("Treeview.Heading",
                  background=[("active", c["bg_hover"])])

        # Scrollbar
        style.configure("Vertical.TScrollbar", background=c["bg_input"],
                         troughcolor=c["bg_tree"], bordercolor=c["border"],
                         arrowcolor=c["text_dim"], relief="flat")
        style.map("Vertical.TScrollbar",
                  background=[("active", c["bg_hover"])])

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        c = COLORS

        # ---- Header ----
        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 14))
        header.pack(fill="x")

        left_h = ttk.Frame(header, style="Header.TFrame")
        left_h.pack(side="left")
        ttk.Label(left_h, text="DWG Merger", style="Title.TLabel").pack(anchor="w")
        ttk.Label(left_h, text="Объединение чертежей AutoCAD  ·  Model Space + Paper Space + BMP/BPW",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

        # ---- Toolbar ----
        toolbar = ttk.Frame(self, padding=(16, 12, 16, 4))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="  📁  Папка", command=self._on_select_folder).pack(side="left")
        ttk.Button(toolbar, text="  📄  Файлы", command=self._on_add_files).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="  ✕  Удалить", command=self._on_remove_selected).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="  🗑  Очистить", command=self._on_clear_list).pack(side="left", padx=(8, 0))

        if HAS_WINDND:
            ttk.Label(toolbar, text="  ↗  перетащите папку или файлы в окно",
                      style="Accent.TLabel").pack(side="left", padx=(16, 0))
        else:
            ttk.Label(toolbar, text="  (drag-and-drop: pip install windnd)",
                      style="Dim.TLabel").pack(side="left", padx=(16, 0))

        # ---- File list (card) ----
        list_card = ttk.Frame(self, style="Card.TFrame", padding=2)
        list_card.pack(fill="both", expand=True, padx=16, pady=(8, 4))

        list_inner = ttk.Frame(list_card, style="Card.TFrame", padding=(10, 8))
        list_inner.pack(fill="both", expand=True)

        ttk.Label(list_inner, text="ФАЙЛЫ ДЛЯ ОБЪЕДИНЕНИЯ", style="CardDim.TLabel").pack(anchor="w", pady=(0, 6))

        tree_frame = ttk.Frame(list_inner, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True)

        columns = ("included", "name", "type", "path", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("included", text="✓")
        self.tree.heading("name", text="Файл")
        self.tree.heading("type", text="Тип")
        self.tree.heading("path", text="Путь")
        self.tree.heading("status", text="Статус")
        self.tree.column("included", width=40, anchor="center", stretch=False)
        self.tree.column("name", width=180, anchor="w")
        self.tree.column("type", width=80, anchor="center", stretch=False)
        self.tree.column("path", width=420, anchor="w")
        self.tree.column("status", width=130, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.tag_configure("ok", foreground=c["success"])
        self.tree.tag_configure("fail", foreground=c["error"])
        self.tree.tag_configure("pending", foreground=c["text_dim"])
        self.tree.tag_configure("excluded", foreground=c["text_muted"])

        # ---- Output path ----
        out_card = ttk.Frame(self, style="Card.TFrame", padding=(14, 10))
        out_card.pack(fill="x", padx=16, pady=(6, 4))

        ttk.Label(out_card, text="РЕЗУЛЬТАТ", style="CardDim.TLabel").pack(anchor="w", pady=(0, 4))
        out_row = ttk.Frame(out_card, style="Card.TFrame")
        out_row.pack(fill="x")
        entry = ttk.Entry(out_row, textvariable=self.output_path)
        entry.pack(side="left", fill="x", expand=True, ipady=2)
        ttk.Button(out_row, text="  Обзор…", command=self._on_browse_output).pack(side="left", padx=(8, 0))

        # ---- Options ----
        opt_card = ttk.Frame(self, style="Card.TFrame", padding=(14, 8))
        opt_card.pack(fill="x", padx=16, pady=(4, 4))

        ttk.Checkbutton(opt_card, text="  Paper Space (листы)",
                         variable=self.paperspace_var).pack(side="left")
        ttk.Checkbutton(opt_card, text="  Показывать AutoCAD",
                         variable=self.visible_var).pack(side="left", padx=(20, 0))
        ttk.Checkbutton(opt_card, text="  Вставлять BMP+BPW",
                         variable=self.include_images_var).pack(side="left", padx=(20, 0))

        # ---- Run bar ----
        run_bar = ttk.Frame(self, padding=(16, 8, 16, 4))
        run_bar.pack(fill="x")

        self.run_btn = ttk.Button(run_bar, text="  ▶  Объединить", style="Accent.TButton",
                                   command=self._on_run)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(run_bar, text="  ■  Стоп", style="Danger.TButton",
                                    command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        self.open_folder_btn = ttk.Button(run_bar, text="  📂  Открыть папку",
                                           command=self._on_open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(run_bar, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(14, 0), ipady=1)

        # ---- Log ----
        log_card = ttk.Frame(self, style="Card.TFrame", padding=(10, 8))
        log_card.pack(fill="both", expand=False, padx=16, pady=(4, 12))

        ttk.Label(log_card, text="ЖУРНАЛ", style="CardDim.TLabel").pack(anchor="w", pady=(0, 4))

        self.log_text = tk.Text(
            log_card, height=9, state="disabled", wrap="word",
            bg=c["bg_tree"], fg=c["text"], insertbackground=c["text"],
            selectbackground=c["accent_dim"], selectforeground=c["text"],
            font=("Consolas", 9), relief="flat", borderwidth=0,
            padx=8, pady=6,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("error", foreground=c["error"])
        self.log_text.tag_configure("warning", foreground=c["warning"])
        self.log_text.tag_configure("info", foreground=c["text_dim"])
        self.log_text.tag_configure("success", foreground=c["success"])

    # ------------------------------------------------------------------
    # Drag-and-drop  (callback только кладёт в очередь — UI не трогает)
    # ------------------------------------------------------------------
    def _setup_drag_drop(self):
        if not HAS_WINDND:
            return
        try:
            windnd.hook_dropfiles(self, func=self._on_drop)
        except Exception:
            pass

    def _on_drop(self, paths):
        """Вызывается windnd из чужого потока — только очередь, без tkinter."""
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
        """Обработка drop — только из UI-потока."""
        folders = []
        dwg_files = []
        bmp_files = []
        ignored = []

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
            pairs = []
            missing = []
            for bmp_path in bmp_files:
                stem, _ = os.path.splitext(bmp_path)
                bpw_path = stem + ".bpw"
                if not os.path.isfile(bpw_path):
                    import glob as _glob
                    candidates = _glob.glob(stem + ".[bB][pP][wW]")
                    bpw_path = candidates[0] if candidates else None
                if bpw_path and os.path.isfile(bpw_path):
                    pairs.append((bmp_path, bpw_path))
                else:
                    missing.append(bmp_path)

            img_added = self._add_image_pairs(pairs) if pairs else 0
            if img_added:
                self._log(f"Перетащено изображений BMP+BPW: {img_added}", "info")
            if missing:
                messagebox.showwarning(
                    "Нет привязки",
                    "Для следующих изображений не найден соседний .bpw файл — "
                    "они не добавлены:\n\n"
                    + "\n".join(os.path.basename(p) for p in missing))

        if ignored:
            self._log(
                "Пропущены (не .dwg/.bmp и не папка): "
                + ", ".join(ignored[:10])
                + ("..." if len(ignored) > 10 else ""),
                "warning")

    # ------------------------------------------------------------------
    # Работа со списком файлов
    # ------------------------------------------------------------------
    def _existing_basenames(self) -> set:
        return {os.path.basename(v["path"]).lower() for v in self.items.values()}

    def _add_paths(self, paths: list):
        existing = self._existing_basenames()
        duplicates = []
        added = 0
        for path in paths:
            base = os.path.basename(path).lower()
            if base in existing:
                duplicates.append(os.path.basename(path))
                continue
            item_id = self.tree.insert(
                "", "end",
                values=("☑", os.path.basename(path), "DWG", path, "готов"),
                tags=("pending",))
            self.items[item_id] = {"path": path, "included": True, "kind": "dwg"}
            existing.add(base)
            added += 1

        if duplicates:
            messagebox.showwarning(
                "Дубликаты пропущены",
                "Следующие файлы уже есть в списке (по имени) и не были добавлены повторно:\n\n"
                + "\n".join(duplicates)
            )
        return added

    def _add_image_pairs(self, pairs: list):
        existing = self._existing_basenames()
        duplicates = []
        added = 0
        for bmp_path, bpw_path in pairs:
            base = os.path.basename(bmp_path).lower()
            if base in existing:
                duplicates.append(os.path.basename(bmp_path))
                continue
            item_id = self.tree.insert(
                "", "end",
                values=("☑", os.path.basename(bmp_path), "BMP+BPW", bmp_path,
                        "привязка найдена"),
                tags=("pending",))
            self.items[item_id] = {"path": bmp_path, "bpw": bpw_path,
                                    "included": True, "kind": "image"}
            existing.add(base)
            added += 1

        if duplicates:
            messagebox.showwarning(
                "Дубликаты пропущены",
                "Следующие изображения уже есть в списке (по имени) и не были добавлены повторно:\n\n"
                + "\n".join(duplicates)
            )
        return added

    def _add_files_from_folder(self, folder: str):
        files = core.find_dwg_files(folder)
        added = self._add_paths(files)
        self.settings["last_source_dir"] = folder
        self._log(f"Из папки '{folder}' добавлено файлов .dwg: {added}", "info")

        pairs, missing_bpw = core.find_bmp_bpw_pairs_verbose(folder)
        img_added = self._add_image_pairs(pairs)
        if img_added:
            self._log(f"Из папки '{folder}' добавлено изображений BMP+BPW: {img_added}", "info")
        for bmp_path in missing_bpw:
            self._log(f"Пропущено {os.path.basename(bmp_path)}: не найден "
                       f"соответствующий .bpw (нет привязки)", "warning")

    def _on_select_folder(self):
        folder = filedialog.askdirectory(
            title="Выберите папку с .dwg файлами",
            initialdir=self.settings.get("last_source_dir") or None)
        if not folder:
            return
        self._add_files_from_folder(folder)

    def _on_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите .dwg или .bmp файлы",
            initialdir=self.settings.get("last_source_dir") or None,
            filetypes=[("DWG и BMP файлы", "*.dwg;*.bmp"),
                       ("DWG файлы", "*.dwg"),
                       ("BMP файлы (нужен соседний .bpw)", "*.bmp"),
                       ("Все файлы", "*.*")])
        if not paths:
            return

        dwg_paths = [p for p in paths if p.lower().endswith(".dwg")]
        bmp_paths = [p for p in paths if p.lower().endswith(".bmp")]

        added = self._add_paths(dwg_paths) if dwg_paths else 0
        if added:
            self._log(f"Добавлено .dwg файлов вручную: {added}", "info")

        if bmp_paths:
            pairs = []
            missing = []
            for bmp_path in bmp_paths:
                stem, _ = os.path.splitext(bmp_path)
                bpw_path = stem + ".bpw"
                if not os.path.isfile(bpw_path):
                    import glob as _glob
                    candidates = _glob.glob(stem + ".[bB][pP][wW]")
                    bpw_path = candidates[0] if candidates else None
                if bpw_path and os.path.isfile(bpw_path):
                    pairs.append((bmp_path, bpw_path))
                else:
                    missing.append(bmp_path)

            img_added = self._add_image_pairs(pairs) if pairs else 0
            if img_added:
                self._log(f"Добавлено изображений BMP+BPW вручную: {img_added}", "info")
            if missing:
                messagebox.showwarning(
                    "Нет привязки",
                    "Для следующих изображений не найден соседний .bpw файл — "
                    "они не добавлены:\n\n" + "\n".join(os.path.basename(p) for p in missing))

    def _on_remove_selected(self):
        selected = self.tree.selection()
        for item_id in selected:
            self.tree.delete(item_id)
            self.items.pop(item_id, None)

    def _on_clear_list(self):
        for item_id in list(self.items.keys()):
            self.tree.delete(item_id)
        self.items.clear()

    def _on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.items:
            return
        if col == "#1":
            info = self.items[item_id]
            info["included"] = not info["included"]
            mark = "☑" if info["included"] else "☐"
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = mark
            self.tree.item(item_id, values=vals,
                            tags=("pending",) if info["included"] else ("excluded",))

    # ------------------------------------------------------------------
    # Вывод / сохранение
    # ------------------------------------------------------------------
    def _on_browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            defaultextension=".dwg",
            initialdir=os.path.dirname(self.output_path.get()) or None,
            initialfile=os.path.basename(self.output_path.get()) or "merged.dwg",
            filetypes=[("DWG файлы", "*.dwg")])
        if path:
            self.output_path.set(path)

    def _on_open_output_folder(self):
        path = self.output_path.get()
        folder = os.path.dirname(os.path.abspath(path))
        if os.path.isdir(folder):
            os.startfile(folder)

    # ------------------------------------------------------------------
    # Запуск / остановка
    # ------------------------------------------------------------------
    def _on_run(self):
        dwg_paths = [v["path"] for v in self.items.values()
                     if v["included"] and v.get("kind") == "dwg"]
        image_pairs = [(v["path"], v["bpw"]) for v in self.items.values()
                       if v["included"] and v.get("kind") == "image"]
        if not self.include_images_var.get():
            image_pairs = []

        if not dwg_paths and not image_pairs:
            messagebox.showerror("Нет файлов", "Список файлов для объединения пуст.")
            return
        output_path = self.output_path.get().strip()
        if not output_path:
            messagebox.showerror("Не указан результат", "Укажите путь к результирующему .dwg файлу.")
            return

        self._save_current_settings()

        total_items = len(dwg_paths) + len(image_pairs)
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.open_folder_btn.config(state="disabled")
        self.progress.config(maximum=total_items, value=0)
        self._set_all_status("в очереди", "pending")
        self._log(f"Запуск: {len(dwg_paths)} DWG, {len(image_pairs)} изображений → {output_path}", "info")

        self.stop_flag.clear()
        self.merge_thread = threading.Thread(
            target=self._run_merge_worker, args=(dwg_paths, image_pairs, output_path), daemon=True)
        self.merge_thread.start()

    def _on_stop(self):
        self.stop_flag.set()
        self._log("Останавливаю после текущего файла…", "warning")
        self.stop_btn.config(state="disabled")

    def _run_merge_worker(self, dwg_paths: list, image_pairs: list, output_path: str):
        path_to_item = {v["path"]: k for k, v in self.items.items()}

        def log_cb(msg, level):
            self.log_queue.put(("log", msg, level))

        def progress_cb(done, total, filename):
            self.log_queue.put(("progress", done, total, filename))
            all_paths = dwg_paths + [bmp for bmp, _bpw in image_pairs]
            for path in all_paths:
                if os.path.basename(path) == filename:
                    item_id = path_to_item.get(path)
                    if item_id:
                        self.log_queue.put(("status", item_id))
                    break

        try:
            summary = core.run_merge(
                dwg_paths, output_path,
                paperspace=self.paperspace_var.get(),
                visible=self.visible_var.get(),
                log_callback=log_cb, progress_callback=progress_cb,
                stop_flag=self.stop_flag, image_pairs=image_pairs)
            self.log_queue.put(("done", summary))
        except Exception as e:
            self.log_queue.put(("fatal", str(e)))

    # ------------------------------------------------------------------
    # Очереди UI
    # ------------------------------------------------------------------
    def _poll_log_queue(self):
        try:
            while True:
                event = self.log_queue.get_nowait()
                kind = event[0]
                if kind == "log":
                    _, msg, level = event
                    self._log(msg, level)
                elif kind == "progress":
                    _, done, total, filename = event
                    self.progress.config(maximum=total, value=done)
                elif kind == "status":
                    _, item_id = event
                    self._mark_item_status(item_id)
                elif kind == "done":
                    _, summary = event
                    self._on_merge_done(summary)
                elif kind == "fatal":
                    _, msg = event
                    self._on_merge_fatal(msg)
        except queue.Empty:
            pass

        try:
            while True:
                paths = self.drop_queue.get_nowait()
                self._process_dropped_paths(paths)
        except queue.Empty:
            pass

        self.after(100, self._poll_log_queue)

    def _mark_item_status(self, item_id):
        vals = list(self.tree.item(item_id, "values"))
        vals[4] = "обработан"
        self.tree.item(item_id, values=vals, tags=("ok",))

    def _set_all_status(self, text, tag):
        for item_id, info in self.items.items():
            if not info["included"]:
                continue
            vals = list(self.tree.item(item_id, "values"))
            vals[4] = text
            self.tree.item(item_id, values=vals, tags=(tag,))

    def _on_merge_done(self, summary):
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.open_folder_btn.config(state="normal")
        self._log("=" * 56, "info")
        self._log(f"ГОТОВО. DWG успешно: {summary['ok']}, ошибок: {summary['fail']}, "
                   f"объектов: {summary['total_objects']}", "success")
        images_ok = summary.get("images_ok", 0)
        images_fail = summary.get("images_fail", 0)
        if images_ok or images_fail:
            self._log(f"Изображений вставлено: {images_ok}, с ошибкой: {images_fail}", "info")
        if summary["errors"]:
            for path, msg in summary["errors"]:
                self._log(f"  Ошибка в {os.path.basename(path)}: {msg}", "error")
            messagebox.showwarning(
                "Завершено с ошибками",
                f"DWG успешно: {summary['ok']} (ошибок: {summary['fail']}), "
                f"изображений успешно: {images_ok} (ошибок: {images_fail}).\n"
                f"Подробности — в журнале.")
        else:
            messagebox.showinfo(
                "Готово",
                f"Успешно объединено: {summary['ok']} DWG-файлов"
                + (f" и вставлено {images_ok} изображений." if images_ok else "."))

    def _on_merge_fatal(self, msg):
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._log(f"КРИТИЧЕСКАЯ ОШИБКА: {msg}", "error")
        messagebox.showerror("Ошибка", f"Не удалось выполнить объединение:\n{msg}")

    def _log(self, msg: str, level: str = "info"):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n", level)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    # ------------------------------------------------------------------
    # Настройки / закрытие
    # ------------------------------------------------------------------
    def _save_current_settings(self):
        self.settings["last_output_path"] = self.output_path.get()
        self.settings["paperspace"] = self.paperspace_var.get()
        self.settings["visible"] = self.visible_var.get()
        self.settings["include_images"] = self.include_images_var.get()
        save_settings(self.settings)

    def _on_close(self):
        if self.merge_thread and self.merge_thread.is_alive():
            if not messagebox.askyesno("Идёт объединение",
                                        "Процесс объединения ещё выполняется. Всё равно закрыть?"):
                return
            self.stop_flag.set()
        self._save_current_settings()
        self.destroy()


def main():
    if sys.platform != "win32":
        print("Это приложение работает только на Windows (требуется AutoCAD).")
        sys.exit(1)
    app = MergeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
