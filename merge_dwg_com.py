# -*- coding: utf-8 -*-
"""
merge_dwg_com.py
=================
Объединение нескольких .dwg файлов в один результирующий .dwg
через ActiveX/COM API установленного AutoCAD (Windows).

Идея работы:
    Для каждого исходного файла AutoCAD открывает документ, забирает
    все объекты Model Space (и, опционально, Paper Space), после чего
    методом Database.CopyObjects() копирует их в результирующий документ.
    Этот метод — программный эквивалент вашего Ctrl+C -> "Вставить с
    исходными координатами": объекты переносятся БЕЗ смещения и БЕЗ
    трансформации, вместе со всеми зависимостями (слои, типы линий,
    веса линий, стили текста, блоки, атрибуты).

Требования:
    - Windows
    - Установленный AutoCAD (любая современная версия с COM-интерфейсом)
    - pip install pywin32 tqdm

Пример запуска:
    python merge_dwg_com.py --source "C:\\Projects\\All" --output "C:\\Result\\merged.dwg"
    python merge_dwg_com.py --source "C:\\Projects\\All" --output "C:\\Result\\merged.dwg" ^
        --template "C:\\Templates\\acadiso.dwt" --paperspace --visible

Важно:
    - Все исходные файлы должны использовать одну и ту же систему координат
      (единицы, UCS/WCS), иначе объекты "наложатся" некорректно — это
      естественное следствие копирования без смещения, точно как при
      ручной вставке с исходными координатами.
    - Внешние ссылки (XREF): CopyObjects копирует xref-блок как есть
      (ссылку на путь). Если нужно, чтобы геометрия xref физически
      "влилась" в результат, предварительно выполните XREF BIND в
      исходных файлах.

О надёжности COM-вызовов:
    AutoCAD — однопоточный (STA) COM-сервер: пока он занят обработкой одной
    команды, все входящие вызовы отклоняются с ошибкой RPC_E_CALL_REJECTED
    ("Вызов был отклонен", код -2147418111). Это нормальная ситуация при
    автоматизации, а не признак поломки — правильный способ справиться
    с ней — повторять отклонённый вызов через небольшую паузу, "прокачивая"
    очередь сообщений COM (pythoncom.PumpWaitingMessages()). Все обращения
    к объектам AutoCAD в этом скрипте обёрнуты функцией com_retry() именно
    по этой причине.
"""

import argparse
import glob
import logging
import os
import sys
import time

try:
    import win32com.client
    from win32com.client import VARIANT
    import pythoncom
    import pywintypes
except ImportError:
    print("Не найден модуль pywin32. Установите командой: pip install pywin32")
    sys.exit(1)

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable


# Максимальное число объектов в одном вызове CopyObjects.
# Большие массивы иногда приводят к сбоям маршаллинга COM-массивов,
# поэтому копируем крупные наборы объектов пачками.
COPY_CHUNK_SIZE = 3000

# Коды ошибок COM, при которых имеет смысл повторить вызов чуть позже —
# они означают "сервер занят прямо сейчас", а не "вызов некорректен".
RETRYABLE_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED — "Вызов был отклонен"
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE (изредка тоже временная ситуация)
}


# --------------------------------------------------------------------------
# Надёжный вызов COM-методов/свойств AutoCAD с автоповтором
# --------------------------------------------------------------------------
def com_retry(func, *args, retries=40, delay=0.5, logger=None, **kwargs):
    """
    Выполняет func(*args, **kwargs) с автоматическим повтором, если AutoCAD
    временно занят (RPC_E_CALL_REJECTED и т.п.). Между попытками "прокачивает"
    очередь COM-сообщений — это обязательно для STA-сервера вроде AutoCAD.

    Используется для КАЖДОГО обращения к объектам AutoCAD (Documents.Open,
    ModelSpace, CopyObjects, Close, SaveAs и т.д.), т.к. любое из них может
    временно "отклониться", пока AutoCAD дорабатывает предыдущую операцию.
    """
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            hresult = e.args[0]
            if hresult in RETRYABLE_HRESULTS:
                last_exc = e
                pythoncom.PumpWaitingMessages()
                time.sleep(delay)
                continue
            raise  # ошибка не из "временных" — пробрасываем сразу
    if logger:
        logger.warning(f"com_retry: превышено число попыток ({retries})")
    raise last_exc


# --------------------------------------------------------------------------
# Логирование
# --------------------------------------------------------------------------
def setup_logging(log_path: str) -> logging.Logger:
    logger = logging.getLogger("merge_dwg_com")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    fh = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# --------------------------------------------------------------------------
# Работа с AutoCAD через COM
# --------------------------------------------------------------------------
def get_or_start_acad(visible: bool, logger: logging.Logger):
    """Запускает новый экземпляр AutoCAD (или подключается к запущенному)."""
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    except Exception as e:
        logger.error(f"Не удалось запустить/подключиться к AutoCAD через COM: {e}")
        raise

    acad.Visible = visible

    # Даём AutoCAD полноценно инициализироваться, прежде чем слать команды.
    # Это дольше, чем может показаться нужным, но именно нехватка этой паузы
    # чаще всего и приводит к каскаду RPC_E_CALL_REJECTED в начале работы.
    logger.info("Ожидание готовности AutoCAD...")
    for _ in range(20):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.3)
    time.sleep(2)

    return acad


def find_dwg_files(source_folder: str) -> list:
    pattern = os.path.join(source_folder, "**", "*.dwg")
    return sorted(glob.glob(pattern, recursive=True))


def create_target_document(acad, template: str, logger: logging.Logger):
    """Создаёт новый (пустой либо по шаблону) документ — будущий результат."""
    try:
        if template:
            doc = com_retry(acad.Documents.Add, template, logger=logger)
            logger.info(f"Создан целевой документ по шаблону: {template}")
        else:
            doc = com_retry(acad.Documents.Add, logger=logger)
            logger.info("Создан новый пустой целевой документ (без шаблона)")
        # Пауза после создания документа — тому тоже нужно время "осесть"
        time.sleep(1)
        pythoncom.PumpWaitingMessages()
        return doc
    except Exception as e:
        logger.error(f"Ошибка создания целевого документа: {e}")
        raise


def collect_modelspace_objects(doc, logger: logging.Logger) -> list:
    """Собирает объекты только Model Space."""
    objects = []
    model_space = com_retry(lambda: doc.ModelSpace, logger=logger)
    for obj in model_space:
        objects.append(obj)
    return objects


def collect_paperspace_layouts(doc, logger: logging.Logger) -> list:
    """
    Возвращает список (имя_листа, [объекты]) для всех листов (Paper Space),
    кроме вкладки Model. Каждый лист обрабатывается отдельно, т.к. его
    объекты (например, видовые экраны) физически могут принадлежать
    ТОЛЬКО Paper Space, а не Model Space — их нельзя копировать вперемешку
    с объектами модели в один и тот же "владелец" (owner) при CopyObjects.
    """
    result = []
    layouts = com_retry(lambda: doc.Layouts, logger=logger)
    for layout in layouts:
        if layout.Name.lower() == "model":
            continue
        try:
            block = com_retry(lambda: layout.Block, logger=logger)
            objs = [obj for obj in block]
            result.append((layout.Name, objs))
        except Exception:
            pass  # пропускаем "битые" листы, не прерывая процесс
    return result


def make_unique_layout_name(target_doc, desired_name: str, logger: logging.Logger) -> str:
    """
    Возвращает имя листа, гарантированно не конфликтующее с уже существующими
    в результирующем документе (AutoCAD не разрешит два листа с одинаковым именем).
    """
    existing = set()
    for layout in com_retry(lambda: target_doc.Layouts, logger=logger):
        existing.add(layout.Name.lower())

    name = desired_name[:255]
    if name.lower() not in existing:
        return name

    counter = 2
    while True:
        candidate = f"{desired_name[:250]} ({counter})"
        if candidate.lower() not in existing:
            return candidate
        counter += 1


def chunked(lst, size):
    for i in range(0, len(lst), size):
        yield tuple(lst[i:i + size])


def copy_objects_batch(source_db, objects: list, target_owner, offset, logger: logging.Logger) -> int:
    """
    Копирует список объектов в target_owner (Model Space ИЛИ блок конкретного
    листа) пачками через CopyObjects, с опциональным сдвигом. Возвращает
    количество фактически скопированных объектов.
    """
    if not objects:
        return 0

    copied = 0
    new_objects = []
    for batch in chunked(objects, COPY_CHUNK_SIZE):
        variant_batch = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, batch)
        result = com_retry(source_db.CopyObjects, variant_batch, target_owner, logger=logger)
        if result:
            new_objects.extend(result)
            copied += len(result)

    if offset != (0.0, 0.0, 0.0) and new_objects:
        from_pt = (0.0, 0.0, 0.0)
        to_pt = offset
        for obj in new_objects:
            try:
                com_retry(obj.Move, from_pt, to_pt, logger=logger)
            except Exception:
                pass  # не все типы объектов поддерживают Move — пропускаем

    return copied


def merge_file_into_target(acad, source_path: str, target_doc,
                            include_paperspace: bool, offset,
                            logger: logging.Logger):
    """
    Открывает source_path, копирует его объекты в target_doc:
      - Model Space исходника -> Model Space результата (в общую "кучу", как и раньше)
      - каждый лист (Paper Space) исходника -> ОТДЕЛЬНЫЙ новый лист результата
        (создаётся автоматически, с именем вида "<файл>_<исходное_имя_листа>")
    Возвращает (успех: bool, кол-во объектов: int, сообщение: str).
    """
    src_doc = None
    try:
        # True = открыть в режиме "только чтение" (быстрее, безопаснее)
        src_doc = com_retry(acad.Documents.Open, source_path, True, logger=logger)
        # Небольшая пауза сразу после открытия — файл может ещё "дозагружаться"
        # (регенерация, подгрузка xref и т.п.), прежде чем отдаст ModelSpace.
        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()
    except Exception as e:
        return False, 0, f"Не удалось открыть файл: {e}"

    total_copied = 0
    file_stem = os.path.splitext(os.path.basename(source_path))[0]
    try:
        source_db = com_retry(lambda: src_doc.Database, logger=logger)

        # 1) Model Space -> Model Space результата
        model_objs = collect_modelspace_objects(src_doc, logger)
        target_space = com_retry(lambda: target_doc.ModelSpace, logger=logger)
        total_copied += copy_objects_batch(source_db, model_objs, target_space, offset, logger)

        # 2) Каждый лист (Paper Space) -> отдельный новый лист результата
        if include_paperspace:
            for layout_name, layout_objs in collect_paperspace_layouts(src_doc, logger):
                if not layout_objs:
                    continue
                new_name = make_unique_layout_name(
                    target_doc, f"{file_stem}_{layout_name}", logger)
                try:
                    new_layout = com_retry(target_doc.Layouts.Add, new_name, logger=logger)
                    new_block = com_retry(lambda: new_layout.Block, logger=logger)
                    total_copied += copy_objects_batch(
                        source_db, layout_objs, new_block, offset, logger)
                except Exception as e:
                    logger.warning(
                        f"Не удалось скопировать лист '{layout_name}' из "
                        f"{os.path.basename(source_path)}: {e}")

        if total_copied == 0:
            return True, 0, "файл не содержит объектов (пропущен)"

        return True, total_copied, ""
    except Exception as e:
        return False, total_copied, f"Ошибка копирования объектов: {e}"
    finally:
        try:
            if src_doc is not None:
                com_retry(src_doc.Close, False, logger=logger)  # закрыть без сохранения
                time.sleep(0.3)
                pythoncom.PumpWaitingMessages()
        except Exception as e:
            logger.warning(f"Не удалось закрыть исходный документ {source_path}: {e}")


# --------------------------------------------------------------------------
# Переиспользуемая функция слияния (вызывается из CLI и из GUI)
# --------------------------------------------------------------------------
def run_merge(dwg_files: list, output_path: str, template=None,
              paperspace=False, offset=(0.0, 0.0, 0.0), visible=True,
              logger=None, log_callback=None, progress_callback=None,
              stop_flag=None):
    """
    Основная функция слияния — не зависит от argparse/CLI, поэтому её же
    использует и GUI-приложение (merge_dwg_gui.py).

    Параметры:
        dwg_files          — список путей к исходным .dwg файлам (уже
                              отфильтрованный вызывающей стороной: GUI сам
                              решает, какие файлы включены и нет дублей)
        output_path         — путь к результирующему .dwg
        log_callback(msg, level)  — необязательный колбэк для вывода строк
                              лога в GUI (level: "info"/"warning"/"error")
        progress_callback(done, total, filename) — необязательный колбэк
                              для обновления прогресс-бара в GUI
        stop_flag            — необязательный объект с атрибутом .is_set()
                              (например, threading.Event) для отмены процесса
                              пользователем из GUI между файлами

    Возвращает dict:
        {"ok": int, "fail": int, "total_objects": int, "errors": [(path, msg), ...]}
    """
    def log(msg, level="info"):
        if logger:
            getattr(logger, level, logger.info)(msg)
        if log_callback:
            log_callback(msg, level)

    acad = get_or_start_acad(visible, logger or _NullLogger())
    target_doc = create_target_document(acad, template, logger or _NullLogger())

    ok_count = 0
    fail_count = 0
    total_objects = 0
    errors = []
    total = len(dwg_files)

    for i, path in enumerate(dwg_files, start=1):
        if stop_flag is not None and stop_flag.is_set():
            log(f"Остановлено пользователем перед файлом {os.path.basename(path)}", "warning")
            break

        success, count, msg = merge_file_into_target(
            acad, path, target_doc, paperspace, offset, logger or _NullLogger())

        if success:
            ok_count += 1
            total_objects += count
            log(f"OK   {os.path.basename(path)}  (+{count} объектов) {msg}", "info")
        else:
            fail_count += 1
            errors.append((path, msg))
            log(f"FAIL {os.path.basename(path)}: {msg}", "error")

        if progress_callback:
            progress_callback(i, total, os.path.basename(path))

    try:
        output_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        com_retry(target_doc.SaveAs, output_path, logger=logger)
        log(f"Результирующий файл сохранён: {output_path}", "info")
    except Exception as e:
        log(f"Ошибка сохранения результирующего файла: {e}", "error")
        raise

    return {"ok": ok_count, "fail": fail_count, "total_objects": total_objects, "errors": errors}


class _NullLogger:
    """Заглушка-логгер для случаев, когда run_merge() вызывают без logging.Logger."""
    def info(self, *a, **kw): pass
    def warning(self, *a, **kw): pass
    def error(self, *a, **kw): pass
    def debug(self, *a, **kw): pass


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def parse_point(s: str):
    try:
        x, y, z = (float(v) for v in s.split(","))
        return (x, y, z)
    except Exception:
        raise argparse.ArgumentTypeError('Точка должна быть в формате "X,Y,Z", например "10,5,0"')


def main():
    parser = argparse.ArgumentParser(
        description="Объединение .dwg файлов в один через AutoCAD COM API (Database.CopyObjects).")
    parser.add_argument("--source", required=True, help="Папка с исходными .dwg файлами")
    parser.add_argument("--output", required=True, help="Путь к результирующему .dwg файлу")
    parser.add_argument("--template", default=None, help="Путь к шаблону .dwt (необязательно)")
    parser.add_argument("--paperspace", action="store_true",
                         help="Копировать также объекты листов (Paper Space), не только Model Space")
    parser.add_argument("--offset", type=parse_point, default=(0.0, 0.0, 0.0),
                         help='Сдвиг для КАЖДОГО файла вида "dx,dy,dz". '
                              'По умолчанию "0,0,0" — исходные координаты сохраняются без изменений.')
    parser.add_argument("--visible", action="store_true", default=True,
                         help="Показывать окно AutoCAD во время работы (рекомендуется, включено по умолчанию)")
    parser.add_argument("--hidden", dest="visible", action="store_false",
                         help="Скрыть окно AutoCAD (может маскировать всплывающие диалоги — не рекомендуется)")
    parser.add_argument("--log", default="merge_dwg_com.log", help="Путь к файлу лога")
    args = parser.parse_args()

    logger = setup_logging(args.log)

    if not os.path.isdir(args.source):
        logger.error(f"Папка не найдена: {args.source}")
        sys.exit(1)

    dwg_files = find_dwg_files(args.source)
    if not dwg_files:
        logger.error("В указанной папке не найдено ни одного .dwg файла")
        sys.exit(1)
    logger.info(f"Найдено файлов для объединения: {len(dwg_files)}")

    pbar = tqdm(total=len(dwg_files), desc="Объединение", unit="файл")

    def _progress(done, total, filename):
        pbar.n = done
        pbar.set_postfix_str(filename)
        pbar.refresh()

    try:
        summary = run_merge(
            dwg_files, args.output, template=args.template,
            paperspace=args.paperspace, offset=args.offset, visible=args.visible,
            logger=logger, progress_callback=_progress)
    except Exception:
        pbar.close()
        sys.exit(1)
    pbar.close()

    logger.info("=" * 70)
    logger.info(f"ГОТОВО. Успешно: {summary['ok']}, ошибок: {summary['fail']}, "
                f"всего скопировано объектов: {summary['total_objects']}")
    if summary["errors"]:
        logger.info("Файлы, обработанные с ошибкой:")
        for path, msg in summary["errors"]:
            logger.info(f"  - {path}: {msg}")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()