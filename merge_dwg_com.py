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

Дополнительно (с этой версии):
    Скрипт также умеет автоматически подхватывать растровые изображения с
    привязкой в формате world file — пары файлов "картинка.bmp" +
    "картинка.bpw", лежащие в исходной папке. Для каждой такой пары:
        - читается .bpw (6 чисел — аффинные коэффициенты привязки);
        - определяется размер BMP в пикселях (из заголовка файла);
        - вычисляются координаты, поворот и масштаб изображения в мире;
        - изображение вставляется в Model Space результирующего чертежа
          сразу в правильном месте, с правильным размером/поворотом.
    Ничего вручную указывать не нужно — достаточно, чтобы .bmp и .bpw
    лежали рядом и назывались одинаково (кроме расширения).

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
import struct
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


# --------------------------------------------------------------------------
# Растровые изображения с привязкой BMP + BPW (world file)
# --------------------------------------------------------------------------
def find_bmp_bpw_pairs_verbose(source_folder: str):
    """
    Ищет пары (image.bmp, image.bpw) в указанной папке (рекурсивно).
    Возвращает (pairs, missing_bpw_paths):
        pairs             — список (bmp_path, bpw_path) для файлов,
                             у которых нашёлся .bpw;
        missing_bpw_paths — .bmp-файлы, для которых .bpw не найден
                             (просто картинка без привязки — пропускаем,
                             т.к. корректно разместить её нечем).
    """
    pattern = os.path.join(source_folder, "**", "*.bmp")
    bmp_files = sorted(glob.glob(pattern, recursive=True))

    pairs = []
    missing = []
    for bmp_path in bmp_files:
        stem, _ = os.path.splitext(bmp_path)
        bpw_path = stem + ".bpw"
        if not os.path.isfile(bpw_path):
            # На случай другого регистра расширения (BPW/Bpw и т.п.)
            candidates = glob.glob(stem + ".[bB][pP][wW]")
            bpw_path = candidates[0] if candidates else None
        if bpw_path and os.path.isfile(bpw_path):
            pairs.append((bmp_path, bpw_path))
        else:
            missing.append(bmp_path)
    return pairs, missing


def find_bmp_bpw_pairs(source_folder: str) -> list:
    """Короткая версия find_bmp_bpw_pairs_verbose — только найденные пары."""
    pairs, _missing = find_bmp_bpw_pairs_verbose(source_folder)
    return pairs


def parse_bpw(bpw_path: str):
    """
    Читает world-файл (.bpw) — 6 чисел по одной на строку, задающих
    аффинное преобразование "координата пикселя -> мировая координата":

        x = A*col + B*row + C
        y = D*col + E*row + F

    где col/row — номер столбца/строки пикселя (0 — первый пиксель).
    Порядок строк в файле — стандартный для world-файлов (.tfw/.jgw/.bpw):

        строка 1: A — размер пикселя по X
        строка 2: D — поворот (участвует в формуле Y)
        строка 3: B — поворот (участвует в формуле X)
        строка 4: E — размер пикселя по Y (обычно отрицательный)
        строка 5: C — X координата центра верхнего левого пикселя
        строка 6: F — Y координата центра верхнего левого пикселя

    Если поворота нет (самый частый случай, B = D = 0), порядок строк 2 и 3
    не влияет на результат — оба значения нулевые.
    """
    with open(bpw_path, "r", encoding="utf-8-sig") as f:
        values = [float(line.strip()) for line in f if line.strip()]
    if len(values) < 6:
        raise ValueError(
            f"Файл '{bpw_path}' содержит {len(values)} значений вместо 6")
    A, D, B, E, C, F = values[:6]
    return A, D, B, E, C, F


def read_bmp_pixel_size(bmp_path: str):
    """
    Определяет ширину/высоту BMP-файла в пикселях напрямую из заголовка
    файла, без внешних зависимостей (PIL и т.п.). Поддерживает и
    современный BITMAPINFOHEADER (40+ байт, ширина/высота — int32 со
    знаком), и старый OS/2 BITMAPCOREHEADER (12 байт, uint16). Отрицательная
    высота означает "top-down" BMP (строки сверху вниз) — для наших целей
    важно только количество пикселей, поэтому берём модуль.
    """
    with open(bmp_path, "rb") as f:
        header = f.read(26)
    if len(header) < 26 or header[0:2] != b"BM":
        raise ValueError(f"Файл '{bmp_path}' не похож на корректный BMP")

    dib_header_size = struct.unpack_from("<I", header, 14)[0]
    if dib_header_size == 12:
        width, height = struct.unpack_from("<HH", header, 18)
    else:
        width, height = struct.unpack_from("<ii", header, 18)
        height = abs(height)
    return width, height


def _bpw_pixel_to_world(A, B, C, D, E, F, col, row):
    x = A * col + B * row + C
    y = D * col + E * row + F
    return (x, y)


def compute_image_world_corners(A, D, B, E, C, F, width_px, height_px):
    """
    Возвращает мировые координаты трёх углов изображения, нужных для
    построения аффинной трансформации: нижний левый (BL), нижний правый
    (BR) и верхний левый (TL). Используется смещение в полпикселя, т.к.
    C/F в .bpw задают координаты ЦЕНТРА крайнего пикселя, а нам нужен
    внешний контур изображения целиком.
    """
    tl = _bpw_pixel_to_world(A, B, C, D, E, F, -0.5, -0.5)
    tr = _bpw_pixel_to_world(A, B, C, D, E, F, width_px - 0.5, -0.5)
    br = _bpw_pixel_to_world(A, B, C, D, E, F, width_px - 0.5, height_px - 0.5)
    bl = _bpw_pixel_to_world(A, B, C, D, E, F, -0.5, height_px - 0.5)
    return {"TL": tl, "TR": tr, "BR": br, "BL": bl}


def _build_affine_transform_matrix(native_min, native_max, world_bl, world_br, world_tl):
    """
    Строит 4x4 матрицу трансформации для AcadEntity.TransformBy, которая
    переводит только что вставленное (AddRaster, масштаб 1, поворот 0)
    изображение из его "естественного" положения (нижний левый угол —
    native_min, без поворота) в целевой прямоугольник, заданный тремя
    мировыми координатами углов. Так поддерживается полный аффинный
    случай — произвольный поворот и РАЗНЫЙ масштаб по X/Y (в отличие от
    свойства ScaleFactor у RasterImage, которое масштабирует только
    равномерно и не умеет поворот+неравномерный масштаб одновременно).
    """
    native_width = native_max[0] - native_min[0]
    native_height = native_max[1] - native_min[1]
    if abs(native_width) < 1e-12 or abs(native_height) < 1e-12:
        return None

    rel_br = (world_br[0] - world_bl[0], world_br[1] - world_bl[1])
    rel_tl = (world_tl[0] - world_bl[0], world_tl[1] - world_bl[1])

    l00 = rel_br[0] / native_width
    l10 = rel_br[1] / native_width
    l01 = rel_tl[0] / native_height
    l11 = rel_tl[1] / native_height

    nx0, ny0 = native_min[0], native_min[1]
    tx = world_bl[0] - (l00 * nx0 + l01 * ny0)
    ty = world_bl[1] - (l10 * nx0 + l11 * ny0)

    return (
        (l00, l01, 0.0, tx),
        (l10, l11, 0.0, ty),
        (0.0, 0.0, 1.0, 0.0),
        (0.0, 0.0, 0.0, 1.0),
    )


def insert_raster_image(target_doc, bmp_path: str, bpw_path: str, logger: logging.Logger):
    """
    Вставляет BMP-изображение в Model Space целевого документа с точной
    геопривязкой по данным .bpw:
        1. Читает .bpw -> аффинные коэффициенты A,B,C,D,E,F.
        2. Читает размер BMP в пикселях (из заголовка файла).
        3. Вычисляет мировые координаты нужных углов изображения.
        4. Вставляет растр через AddRaster (масштаб 1, поворот 0) во
           временную точку — нижний левый угол цели.
        5. Измеряет фактический ("естественный") размер вставленного
           изображения через GetBoundingBox и строит аффинную матрицу,
           переводящую его точно в целевой прямоугольник.
        6. Применяет матрицу через TransformBy.
    Возвращает (успех: bool, сообщение: str).
    """
    try:
        A, D, B, E, C, F = parse_bpw(bpw_path)
    except Exception as e:
        return False, f"Не удалось прочитать .bpw: {e}"

    try:
        width_px, height_px = read_bmp_pixel_size(bmp_path)
    except Exception as e:
        return False, f"Не удалось прочитать размер BMP: {e}"

    corners = compute_image_world_corners(A, D, B, E, C, F, width_px, height_px)
    bl, br, tl = corners["BL"], corners["BR"], corners["TL"]

    try:
        model_space = com_retry(lambda: target_doc.ModelSpace, logger=logger)
        insertion_variant = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                     (bl[0], bl[1], 0.0))
        img = com_retry(model_space.AddRaster, bmp_path, insertion_variant,
                         1.0, 0.0, logger=logger)
        time.sleep(0.2)
        pythoncom.PumpWaitingMessages()

        native_min, native_max = com_retry(img.GetBoundingBox, logger=logger)
        matrix = _build_affine_transform_matrix(native_min, native_max, bl, br, tl)
        if matrix is None:
            return False, ("изображение вставлено, но привязка не применена "
                            "(нулевой размер после вставки)")

        matrix_variant = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, matrix)
        com_retry(img.TransformBy, matrix_variant, logger=logger)
        return True, ""
    except Exception as e:
        return False, f"Ошибка вставки растра: {e}"


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
              stop_flag=None, image_pairs=None):
    """
    Основная функция слияния — не зависит от argparse/CLI, поэтому её же
    использует и GUI-приложение (merge_dwg_gui.py).

    Параметры:
        dwg_files          — список путей к исходным .dwg файлам (уже
                              отфильтрованный вызывающей стороной: GUI сам
                              решает, какие файлы включены и нет дублей)
        output_path         — путь к результирующему .dwg
        image_pairs         — необязательный список (bmp_path, bpw_path) —
                              растровые изображения с привязкой, которые
                              нужно вставить в результирующий чертёж после
                              объединения .dwg (см. find_bmp_bpw_pairs())
        log_callback(msg, level)  — необязательный колбэк для вывода строк
                              лога в GUI (level: "info"/"warning"/"error")
        progress_callback(done, total, filename) — необязательный колбэк
                              для обновления прогресс-бара в GUI
        stop_flag            — необязательный объект с атрибутом .is_set()
                              (например, threading.Event) для отмены процесса
                              пользователем из GUI между файлами

    Возвращает dict:
        {"ok": int, "fail": int, "total_objects": int, "errors": [(path, msg), ...],
         "images_ok": int, "images_fail": int}
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

    # ------------------------------------------------------------------
    # Растровые изображения (BMP + BPW) — вставляются последними, уже
    # в готовый результирующий документ, каждое сразу в свою привязку.
    # ------------------------------------------------------------------
    image_pairs = image_pairs or []
    images_ok = 0
    images_fail = 0
    total_images = len(image_pairs)

    for i, (bmp_path, bpw_path) in enumerate(image_pairs, start=1):
        if stop_flag is not None and stop_flag.is_set():
            log(f"Остановлено пользователем перед изображением "
                f"{os.path.basename(bmp_path)}", "warning")
            break

        success, msg = insert_raster_image(target_doc, bmp_path, bpw_path,
                                            logger or _NullLogger())
        if success:
            images_ok += 1
            log(f"OK   {os.path.basename(bmp_path)}  (вставлено по BPW-привязке)", "info")
        else:
            images_fail += 1
            errors.append((bmp_path, msg))
            log(f"FAIL {os.path.basename(bmp_path)}: {msg}", "error")

        if progress_callback:
            progress_callback(total + i, total + total_images, os.path.basename(bmp_path))

    try:
        output_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        com_retry(target_doc.SaveAs, output_path, logger=logger)
        log(f"Результирующий файл сохранён: {output_path}", "info")
    except Exception as e:
        log(f"Ошибка сохранения результирующего файла: {e}", "error")
        raise

    return {"ok": ok_count, "fail": fail_count, "total_objects": total_objects,
            "errors": errors, "images_ok": images_ok, "images_fail": images_fail}


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
    parser.add_argument("--no-images", dest="images", action="store_false", default=True,
                         help="Не искать и не вставлять пары BMP+BPW из папки --source "
                              "(по умолчанию они вставляются автоматически)")
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

    image_pairs = []
    if args.images:
        image_pairs, missing_bpw = find_bmp_bpw_pairs_verbose(args.source)
        if image_pairs:
            logger.info(f"Найдено изображений BMP+BPW для вставки: {len(image_pairs)}")
        for bmp_path in missing_bpw:
            logger.warning(f"Пропущен {os.path.basename(bmp_path)}: не найден "
                            f"соответствующий .bpw (нет привязки)")

    pbar = tqdm(total=len(dwg_files) + len(image_pairs), desc="Объединение", unit="файл")

    def _progress(done, total, filename):
        pbar.n = done
        pbar.set_postfix_str(filename)
        pbar.refresh()

    try:
        summary = run_merge(
            dwg_files, args.output, template=args.template,
            paperspace=args.paperspace, offset=args.offset, visible=args.visible,
            logger=logger, progress_callback=_progress, image_pairs=image_pairs)
    except Exception:
        pbar.close()
        sys.exit(1)
    pbar.close()

    logger.info("=" * 70)
    logger.info(f"ГОТОВО. Успешно: {summary['ok']}, ошибок: {summary['fail']}, "
                f"всего скопировано объектов: {summary['total_objects']}")
    if image_pairs:
        logger.info(f"Изображений вставлено: {summary['images_ok']}, "
                    f"с ошибкой: {summary['images_fail']}")
    if summary["errors"]:
        logger.info("Файлы, обработанные с ошибкой:")
        for path, msg in summary["errors"]:
            logger.info(f"  - {path}: {msg}")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()