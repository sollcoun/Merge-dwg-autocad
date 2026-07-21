# -*- coding: utf-8 -*-
"""
merge_dwg_gui.py
=================
Графическое приложение для объединения .dwg файлов (для сотрудников,
не работающих с командной строкой). Использует ту же проверенную логику
слияния, что и merge_dwg_com.py (функция run_merge), через AutoCAD COM.

Требования: Windows, установленный AutoCAD, pip install pywin32.

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

SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".dwg_merger_settings.json")


# --------------------------------------------------------------------------
# Хранение настроек между запусками
# --------------------------------------------------------------------------
def load_settings() -> dict:
    defaults = {
        "last_source_dir": "",
        "last_output_path": "",
        "paperspace": False,
        "visible": True,
    }
    if os.path.isfile(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            defaults.update(saved)
        except Exception:
            pass  # повреждённый файл настроек не должен ронять приложение
    return defaults


def save_settings(settings: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # сохранение настроек не критично для работы приложения


# --------------------------------------------------------------------------
# Главное окно
# --------------------------------------------------------------------------
class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Объединение DWG-файлов")
        self.geometry("880x640")
        self.minsize(760, 560)

        self.settings = load_settings()
        self.output_path = tk.StringVar(value=self.settings.get("last_output_path", ""))
        self.paperspace_var = tk.BooleanVar(value=self.settings.get("paperspace", False))
        self.visible_var = tk.BooleanVar(value=self.settings.get("visible", True))

        # items: {tree_id: {"path": str, "included": bool}}
        self.items = {}
        self.merge_thread = None
        self.stop_flag = threading.Event()
        self.log_queue = queue.Queue()

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_log_queue)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Выбрать папку...", command=self._on_select_folder).pack(side="left")
        ttk.Button(top, text="Добавить файлы...", command=self._on_add_files).pack(side="left", padx=6)
        ttk.Button(top, text="Удалить выбранные", command=self._on_remove_selected).pack(side="left")
        ttk.Button(top, text="Очистить список", command=self._on_clear_list).pack(side="left", padx=6)

        # Список файлов
        list_frame = ttk.Frame(self, padding=(10, 0))
        list_frame.pack(fill="both", expand=True)

        columns = ("included", "name", "path", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("included", text="✓")
        self.tree.heading("name", text="Файл")
        self.tree.heading("path", text="Путь")
        self.tree.heading("status", text="Статус")
        self.tree.column("included", width=36, anchor="center", stretch=False)
        self.tree.column("name", width=160, anchor="w")
        self.tree.column("path", width=440, anchor="w")
        self.tree.column("status", width=140, anchor="w")

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.tag_configure("ok", foreground="#1a7f37")
        self.tree.tag_configure("fail", foreground="#c9302c")
        self.tree.tag_configure("pending", foreground="#666666")
        self.tree.tag_configure("excluded", foreground="#aaaaaa")

        # Настройки вывода
        out_frame = ttk.LabelFrame(self, text="Результат", padding=10)
        out_frame.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Entry(out_frame, textvariable=self.output_path).pack(side="left", fill="x", expand=True)
        ttk.Button(out_frame, text="Обзор...", command=self._on_browse_output).pack(side="left", padx=(6, 0))

        # Опции
        opt_frame = ttk.Frame(self, padding=(10, 6))
        opt_frame.pack(fill="x")
        ttk.Checkbutton(opt_frame, text="Копировать также Paper Space (листы)",
                         variable=self.paperspace_var).pack(side="left")
        ttk.Checkbutton(opt_frame, text="Показывать окно AutoCAD",
                         variable=self.visible_var).pack(side="left", padx=20)

        # Кнопки запуска
        run_frame = ttk.Frame(self, padding=10)
        run_frame.pack(fill="x")
        self.run_btn = ttk.Button(run_frame, text="Объединить", command=self._on_run)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(run_frame, text="Остановить", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        self.open_folder_btn = ttk.Button(run_frame, text="Открыть папку с результатом",
                                           command=self._on_open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=6)

        self.progress = ttk.Progressbar(run_frame, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # Лог
        log_frame = ttk.LabelFrame(self, text="Журнал", padding=6)
        log_frame.pack(fill="both", expand=False, padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, height=10, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("error", foreground="#c9302c")
        self.log_text.tag_configure("warning", foreground="#b8860b")
        self.log_text.tag_configure("info", foreground="#222222")

        # Восстанавливаем последнюю папку, если она ещё существует
        last_dir = self.settings.get("last_source_dir", "")
        if last_dir and os.path.isdir(last_dir):
            self._add_files_from_folder(last_dir)

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
            item_id = self.tree.insert("", "end", values=("☑", os.path.basename(path), path, "готов"),
                                        tags=("pending",))
            self.items[item_id] = {"path": path, "included": True}
            existing.add(base)
            added += 1

        if duplicates:
            messagebox.showwarning(
                "Дубликаты пропущены",
                "Следующие файлы уже есть в списке (по имени) и не были добавлены повторно:\n\n"
                + "\n".join(duplicates)
            )
        return added

    def _add_files_from_folder(self, folder: str):
        files = core.find_dwg_files(folder)
        added = self._add_paths(files)
        self.settings["last_source_dir"] = folder
        self._log(f"Из папки '{folder}' добавлено файлов: {added}", "info")

    def _on_select_folder(self):
        folder = filedialog.askdirectory(
            title="Выберите папку с .dwg файлами",
            initialdir=self.settings.get("last_source_dir") or None)
        if not folder:
            return
        self._add_files_from_folder(folder)

    def _on_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите .dwg файлы",
            initialdir=self.settings.get("last_source_dir") or None,
            filetypes=[("DWG файлы", "*.dwg"), ("Все файлы", "*.*")])
        if not paths:
            return
        added = self._add_paths(list(paths))
        self._log(f"Добавлено файлов вручную: {added}", "info")

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
        # Клик по колонке чекбокса переключает включение/исключение файла
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if not item_id or item_id not in self.items:
            return
        if col == "#1":  # колонка "✓"
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
            os.startfile(folder)  # Windows-only, соответствует Windows-only природе AutoCAD COM

    # ------------------------------------------------------------------
    # Запуск / остановка слияния
    # ------------------------------------------------------------------
    def _on_run(self):
        included = [v["path"] for v in self.items.values() if v["included"]]
        if not included:
            messagebox.showerror("Нет файлов", "Список файлов для объединения пуст.")
            return
        output_path = self.output_path.get().strip()
        if not output_path:
            messagebox.showerror("Не указан результат", "Укажите путь к результирующему .dwg файлу.")
            return

        self._save_current_settings()

        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.open_folder_btn.config(state="disabled")
        self.progress.config(maximum=len(included), value=0)
        self._set_all_status("в очереди", "pending")
        self._log(f"Запуск объединения: {len(included)} файлов -> {output_path}", "info")

        self.stop_flag.clear()
        self.merge_thread = threading.Thread(
            target=self._run_merge_worker, args=(included, output_path), daemon=True)
        self.merge_thread.start()

    def _on_stop(self):
        self.stop_flag.set()
        self._log("Останавливаю после текущего файла...", "warning")
        self.stop_btn.config(state="disabled")

    def _run_merge_worker(self, included_paths: list, output_path: str):
        """Выполняется в фоновом потоке — не блокирует интерфейс."""
        path_to_item = {v["path"]: k for k, v in self.items.items()}

        def log_cb(msg, level):
            self.log_queue.put(("log", msg, level))

        def progress_cb(done, total, filename):
            self.log_queue.put(("progress", done, total, filename))
            # Определяем, какой файл только что обработан, по имени
            for path in included_paths:
                if os.path.basename(path) == filename:
                    item_id = path_to_item.get(path)
                    if item_id:
                        self.log_queue.put(("status", item_id))
                    break

        try:
            summary = core.run_merge(
                included_paths, output_path,
                paperspace=self.paperspace_var.get(),
                visible=self.visible_var.get(),
                log_callback=log_cb, progress_callback=progress_cb,
                stop_flag=self.stop_flag)
            self.log_queue.put(("done", summary))
        except Exception as e:
            self.log_queue.put(("fatal", str(e)))

    # ------------------------------------------------------------------
    # Обмен между фоновым потоком и интерфейсом через очередь
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
        self.after(100, self._poll_log_queue)

    def _mark_item_status(self, item_id):
        # Статус (ок/ошибка) виден по последней строке лога для этого файла;
        # здесь просто помечаем как "обработан" — подробности смотрите в журнале.
        vals = list(self.tree.item(item_id, "values"))
        vals[3] = "обработан"
        self.tree.item(item_id, values=vals)

    def _set_all_status(self, text, tag):
        for item_id, info in self.items.items():
            if not info["included"]:
                continue
            vals = list(self.tree.item(item_id, "values"))
            vals[3] = text
            self.tree.item(item_id, values=vals, tags=(tag,))

    def _on_merge_done(self, summary):
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.open_folder_btn.config(state="normal")
        self._log("=" * 60, "info")
        self._log(f"ГОТОВО. Успешно: {summary['ok']}, ошибок: {summary['fail']}, "
                   f"объектов скопировано: {summary['total_objects']}", "info")
        if summary["errors"]:
            for path, msg in summary["errors"]:
                self._log(f"  Ошибка в {os.path.basename(path)}: {msg}", "error")
            messagebox.showwarning(
                "Завершено с ошибками",
                f"Успешно: {summary['ok']}, с ошибками: {summary['fail']}.\n"
                f"Подробности — в журнале внизу окна.")
        else:
            messagebox.showinfo("Готово", f"Все {summary['ok']} файлов успешно объединены.")

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