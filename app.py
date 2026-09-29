# main_app.py
"""
GUI для редактирования форматирования .docx файлов.

Архитектура:
    - DocumentModel (gui_helper.py) хранит "карту" документа: список абзацев,
      внутри каждого — список прогонов, у каждого прогона есть ссылка
      на реальный lxml-элемент <w:r> в дереве документа.
    - WordXMLFormatter (formatter.py) содержит методы двух видов:
        change_*        -> применяют стиль ко ВСЕМ прогонам/абзацам документа
        apply_*_to_run  -> применяют стиль к ОДНОМУ конкретному прогону
        apply_*_to_para -> применяют стиль к ОДНОМУ конкретному абзацу
    - Это приложение использует apply_* для точечного редактирования
      (по ТЗ) и change_* для функции "реформатировать весь файл под один формат".

Undo/отмена реализована через сохранение глубокой копии (deepcopy) XML-узла
свойств (rPr/pPr) ДО открытия диалога редактирования. Если пользователь
нажимает "Отмена" - сохранённый узел возвращается на место. Так правки одного
прогона не задевают остальные (в отличие от отмены через copy.deepcopy всего
дерева на каждый чих, что было бы дорого и избыточно).

Для отмены "реформатирования всего файла" используется отдельный, более
тяжёлый снимок - deepcopy всего xml_root, потому что change_* методы
проходят по ВСЕМ прогонам/абзацам разом и точечных откатов тут не бывает.
"""

import os
import copy
import threading
from tkinter import filedialog, colorchooser, messagebox

import customtkinter as ctk

try:
    from PIL import Image, ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

# Бэкенд товарища и модель для GUI - из приложенных файлов
from gui_helper import DocumentModel, W_NS
from formatter import WordXMLFormatter, NS_MAP


# ============================== РЕЖИМ НАМЕРЕННЫХ БАГОВ (ДЛЯ QA) ==============================
#
# Этот блок существует ИСКЛЮЧИТЕЛЬНО для тестировщиков: он внедряет в
# приложение заранее известные, воспроизводимые дефекты, чтобы проверить
# способность QA-команды их находить. Ни один из этих багов не должен
# попасть в реальную поставку - переключатель ниже должен быть False
# в любой сборке, кроме тестовой.
#
# Архитектура: BUGGY_MODE - мастер-переключатель. Каждый конкретный баг
# управляется своим собственным флагом BUG_*, который по умолчанию
# наследует значение BUGGY_MODE, но может быть переопределён вручную
# (например, чтобы включить только ОДИН конкретный баг для точечного
# тест-кейса, не включая остальные). Чтобы переопределить один баг отдельно
# от мастер-переключателя, просто замените "BUGGY_MODE" на True/False
# в соответствующей строке ниже.
#
# Каждый баг реализован как ДУБЛИКАТ блока кода: есть "чистая" ветка
# (корректное поведение) и "багованная" ветка (дефект), выбор между
# ними происходит через if BUG_XXX: / else: в месте использования -
# ничего не удаляется и не патчится "на лету", обе версии логики видны
# в коде целиком.

BUGGY_MODE = True  # <-- ГЛАВНЫЙ ВЫКЛЮЧАТЕЛЬ. False = приложение работает корректно.

# --- Баг 1: диалог открытия файла принимает не только .docx ---
# Пользователь может выбрать .txt/.pdf/произвольный файл, что приведёт к
# необработанному исключению при попытке распаковать его как zip/docx.
BUG_ACCEPT_ANY_FILE = BUGGY_MODE

# --- Баг 2: кнопка "Сохранить как" активна сразу после загрузки файла ---
# Позволяет сохранить документ, не завершив (и даже не начав) форматирование -
# то есть, по сути, ничего не отредактировав.
BUG_SAVE_ALWAYS_ENABLED = BUGGY_MODE

# --- Баг 3: в диалоге редактирования прогона применяется не тот шрифт,
#     что выбран пользователем в выпадающем списке (циклический сдвиг) ---
BUG_SWAP_FONT_CHOICE = BUGGY_MODE

# --- Баг 4: "Курсив" применяет жирный, "Жирный" применяет курсив ---
BUG_SWAP_BOLD_ITALIC = BUGGY_MODE

# --- Баг 5: в свёрнутом виде абзаца показывается текст ПРОГОНА, а не
#     абзаца целиком (чисто визуальный дефект, данные не искажены) ---
BUG_PARA_PREVIEW_SHOWS_RUN = BUGGY_MODE

# --- Баг 6 (пасхалка): реформатирование всего файла с размером шрифта
#     ровно 67 показывает картинку вместо применения формата; реальное
#     форматирование в этом случае вообще не происходит ---
BUG_EASTER_EGG_SIZE_67 = BUGGY_MODE

# --- Баг 7: кнопка "Отмена" в редакторе прогона на самом деле ПОДТВЕРЖДАЕТ
#     изменения вместо их отката - инверсия функции "передумать" ---
BUG_CANCEL_ACTUALLY_CONFIRMS = BUGGY_MODE

# --- Баг 8: счётчик абзацев в статус-баре после загрузки файла показывает
#     число на единицу больше реального (off-by-one) ---
BUG_PARAGRAPH_COUNT_OFF_BY_ONE = BUGGY_MODE

# Путь к картинке для бага-пасхалки №6. Копируется вместе с main_app.py
# в подпапку assets/ - см. итоговую поставку.
EASTER_EGG_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "easter_egg.jpg")


# ============================== КОНСТАНТЫ ==============================

APP_WIDTH = 800
APP_HEIGHT = 600

# Список шрифтов для выпадающего списка. ТЗ не called конкретный набор,
# поэтому даю разумный набор часто используемых в русскоязычных документах
# шрифтов + подпись "Другой..." для ручного ввода, если нужного нет в списке.
FONT_CHOICES = [
    "Times New Roman", "Arial", "Calibri", "Verdana",
    "Georgia", "Tahoma", "Cambria", "Courier New", "Другой...",
]

# Диапазон разумных размеров шрифта для валидации ручного ввода (в пунктах)
FONT_SIZE_MIN = 1
FONT_SIZE_MAX = 400

ALIGNMENT_CHOICES = {
    "По левому краю": "left",
    "По центру": "center",
    "По правому краю": "right",
    "По ширине": "both",
}
ALIGNMENT_CHOICES_REVERSE = {v: k for k, v in ALIGNMENT_CHOICES.items()}

LINE_SPACING_CHOICES = {
    "Одинарный": "single",
    "Полуторный": "1.5",
    "Двойной": "double",
}
LINE_SPACING_REVERSE = {v: k for k, v in LINE_SPACING_CHOICES.items()}

STYLE_CHOICES = {
    "Обычный": (False, False),
    "Жирный": (True, False),
    "Курсив": (False, True),
    "Жирный курсив": (True, True),
}
# обратный маппинг (bold, italic) -> подпись, для восстановления состояния combobox
STYLE_CHOICES_REVERSE = {v: k for k, v in STYLE_CHOICES.items()}


def hex_to_ctk_color(hex_val: str) -> str:
    """
    Приводит цвет из XML (может быть 'По умолчанию', 'auto' или hex без #)
    к валидному tkinter-цвету для превью в GUI.
    """
    if not hex_val or hex_val in ("По умолчанию", "auto"):
        return "#000000"
    cleaned = hex_val.strip().lstrip("#")
    if len(cleaned) == 6:
        try:
            int(cleaned, 16)
            return f"#{cleaned}"
        except ValueError:
            return "#000000"
    return "#000000"


# ============================== ДИАЛОГ РЕДАКТИРОВАНИЯ ПРОГОНА ==============================

class RunEditDialog(ctk.CTkToplevel):
    """
    Модальное окно редактирования ОДНОГО прогона (run).
    Пользователь выбирает шрифт/начертание/цвет из выпадающих списков,
    размер вводит с клавиатуры. Подтверждение применяет изменения к XML
    немедленно (через apply_*_to_run); Отмена откатывает узел rPr
    к состоянию на момент открытия диалога.

    Диалог можно открывать повторно после подтверждения - пользователь
    может передумать и продолжить редактирование того же прогона
    (требование ТЗ).
    """

    def __init__(self, master, formatter: WordXMLFormatter, run_info: dict, on_close_callback):
        super().__init__(master)
        self.formatter = formatter
        self.run_info = run_info
        self.run_element = run_info["xml_element"]
        self.on_close_callback = on_close_callback

        # --- Снимок состояния ДО редактирования, для отмены ---
        rpr = self.run_element.find("w:rPr", NS_MAP)
        self._rpr_snapshot = copy.deepcopy(rpr) if rpr is not None else None
        self._rpr_existed = rpr is not None

        self.title("Редактирование прогона")
        self.geometry("420x430")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()  # модальность

        preview_text = run_info["text"] if run_info["text"].strip() else "(пустой текст)"
        preview_short = preview_text[:80] + ("..." if len(preview_text) > 80 else "")

        ctk.CTkLabel(
            self, text=f'Текст: "{preview_short}"',
            wraplength=380, justify="left", font=ctk.CTkFont(size=12, slant="italic"),
        ).pack(padx=16, pady=(16, 8), anchor="w")

        # --- Шрифт ---
        ctk.CTkLabel(self, text="Шрифт:").pack(padx=16, anchor="w")
        current_font = run_info.get("font", "Times New Roman")
        self.font_var = ctk.StringVar(
            value=current_font if current_font in FONT_CHOICES else "Другой..."
        )
        self.font_combo = ctk.CTkComboBox(
            self, values=FONT_CHOICES, variable=self.font_var,
            command=self._on_font_combo_change, width=360,
        )
        self.font_combo.pack(padx=16, pady=(0, 4))

        self.custom_font_entry = ctk.CTkEntry(self, placeholder_text="Введите название шрифта", width=360)
        if self.font_var.get() == "Другой...":
            self.custom_font_entry.insert(0, current_font)
            self.custom_font_entry.pack(padx=16, pady=(0, 8))
        else:
            # держим виджет созданным, но не показанным, чтобы не проверять None далее
            self.custom_font_entry.pack_forget()

        # --- Размер (клавиатурный ввод, как указано в ТЗ) ---
        ctk.CTkLabel(self, text="Размер шрифта (пт):").pack(padx=16, anchor="w")
        current_size = run_info.get("size", "14 pt")
        size_digits = "".join(ch for ch in str(current_size) if ch.isdigit())
        self.size_entry = ctk.CTkEntry(self, width=360, placeholder_text="Например, 14")
        self.size_entry.insert(0, size_digits if size_digits else "14")
        self.size_entry.pack(padx=16, pady=(0, 8))

        # --- Начертание (выпадающий список, как в ТЗ, вместо отдельных чекбоксов) ---
        ctk.CTkLabel(self, text="Начертание:").pack(padx=16, anchor="w")
        is_bold = run_info.get("is_bold", False)
        is_italic = run_info.get("is_italic", False)
        current_style_label = STYLE_CHOICES_REVERSE.get((is_bold, is_italic), "Обычный")
        self.style_var = ctk.StringVar(value=current_style_label)
        self.style_combo = ctk.CTkComboBox(
            self, values=list(STYLE_CHOICES.keys()), variable=self.style_var, width=360,
        )
        self.style_combo.pack(padx=16, pady=(0, 8))

        # --- Подчёркивание отдельным переключателем (в ТЗ упомянуто "и т.д." к начертанию) ---
        self.underline_var = ctk.BooleanVar(value=run_info.get("is_underline", False))
        self.underline_check = ctk.CTkCheckBox(
            self, text="Подчёркнутый", variable=self.underline_var,
        )
        self.underline_check.pack(padx=16, pady=(0, 8), anchor="w")

        # --- Цвет: кнопка + системный colorchooser (согласовано с пользователем;
        #     в ТЗ буквально написано "выпадающий список", но нативного
        #     color-picker с выпадающим списком в tkinter нет - кнопка с
        #     превью и полной палитрой даёт пользователю больше гибкости) ---
        ctk.CTkLabel(self, text="Цвет текста:").pack(padx=16, anchor="w")
        self.current_color_hex = hex_to_ctk_color(run_info.get("color", "000000"))
        color_row = ctk.CTkFrame(self, fg_color="transparent")
        color_row.pack(padx=16, pady=(0, 8), fill="x")
        self.color_preview = ctk.CTkLabel(
            color_row, text="", width=30, height=24, fg_color=self.current_color_hex, corner_radius=4,
        )
        self.color_preview.pack(side="left", padx=(0, 8))
        ctk.CTkButton(color_row, text="Выбрать цвет...", command=self._pick_color, width=200).pack(side="left")

        # --- Кнопки подтверждения / отмены ---
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=16, pady=(16, 16), fill="x", side="bottom")
        ctk.CTkButton(btn_row, text="Отмена", command=self._on_cancel, fg_color="gray40", width=170).pack(side="left")
        ctk.CTkButton(btn_row, text="Подтвердить", command=self._on_confirm, width=170).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_font_combo_change(self, choice):
        if choice == "Другой...":
            self.custom_font_entry.delete(0, "end")
            self.custom_font_entry.pack(padx=16, pady=(0, 8))
        else:
            self.custom_font_entry.pack_forget()

    def _pick_color(self):
        # askcolor возвращает ((r,g,b), '#rrggbb') или (None, None) при отмене
        _, hex_color = colorchooser.askcolor(color=self.current_color_hex, title="Выберите цвет текста")
        if hex_color:
            self.current_color_hex = hex_color
            self.color_preview.configure(fg_color=hex_color)

    def _validate_and_get_font_name(self):
        choice = self.font_var.get()
        if choice == "Другой...":
            custom = self.custom_font_entry.get().strip()
            if not custom:
                messagebox.showwarning("Проверка", "Введите название шрифта.", parent=self)
                return None
            return custom
        return choice

    def _validate_and_get_size(self):
        raw = self.size_entry.get().strip().replace(",", ".")
        try:
            size = float(raw)
        except ValueError:
            messagebox.showwarning("Проверка", "Размер шрифта должен быть числом.", parent=self)
            return None
        if not (FONT_SIZE_MIN <= size <= FONT_SIZE_MAX):
            messagebox.showwarning(
                "Проверка", f"Размер шрифта должен быть от {FONT_SIZE_MIN} до {FONT_SIZE_MAX} пт.", parent=self
            )
            return None
        return size

    def _on_confirm(self):
        font_name = self._validate_and_get_font_name()
        if font_name is None:
            return
        size = self._validate_and_get_size()
        if size is None:
            return

        is_bold, is_italic = STYLE_CHOICES[self.style_var.get()]
        is_underline = self.underline_var.get()

        # БАГ (BUG_SWAP_FONT_CHOICE): вместо выбранного шрифта применяется
        # СЛЕДУЮЩИЙ по списку FONT_CHOICES (циклический сдвиг на 1 позицию).
        # Если пользователь ввёл произвольное имя через "Другой..." (нет в
        # списке), подменяем на первый шрифт из списка - тоже неверно.
        if BUG_SWAP_FONT_CHOICE:
            plain_fonts = [f for f in FONT_CHOICES if f != "Другой..."]
            if font_name in plain_fonts:
                idx = plain_fonts.index(font_name)
                applied_font_name = plain_fonts[(idx + 1) % len(plain_fonts)]
            else:
                applied_font_name = plain_fonts[0]
        else:
            applied_font_name = font_name

        # БАГ (BUG_SWAP_BOLD_ITALIC): "Жирный" и "Курсив" меняются местами.
        # "Жирный курсив" и "Обычный" остаются как есть, т.к. их подмена
        # была бы неотличима от корректного поведения (симметричные значения).
        if BUG_SWAP_BOLD_ITALIC:
            applied_is_bold, applied_is_italic = is_italic, is_bold
        else:
            applied_is_bold, applied_is_italic = is_bold, is_italic

        # Применяем все свойства к конкретному прогону через point-wise API форматтера
        self.formatter.apply_font_to_run(self.run_element, applied_font_name)
        self.formatter.apply_size_to_run(self.run_element, size)
        self.formatter.apply_bold_to_run(self.run_element, applied_is_bold)
        self.formatter.apply_italic_to_run(self.run_element, applied_is_italic)
        self.formatter.apply_underline_to_run(self.run_element, is_underline)
        self.formatter.apply_color_to_run(self.run_element, self.current_color_hex)

        self.on_close_callback(confirmed=True)
        self.destroy()

    def _on_cancel(self):
        # БАГ (BUG_CANCEL_ACTUALLY_CONFIRMS): вместо отката "Отмена" на самом
        # деле применяет текущие значения из полей диалога - то есть ведёт
        # себя как "Подтвердить". Пользователь думает, что передумал, но
        # изменения всё равно фиксируются в документе.
        if BUG_CANCEL_ACTUALLY_CONFIRMS:
            font_name = self._validate_and_get_font_name()
            size = self._validate_and_get_size()
            # Если поля невалидны - тихо используем безопасные запасные
            # значения вместо предупреждения, т.к. "Отмена" не должна
            # выглядеть как блокирующее действие для пользователя.
            if font_name is None:
                font_name = self.run_info.get("font", "Times New Roman")
            if size is None:
                size = 14
            is_bold, is_italic = STYLE_CHOICES[self.style_var.get()]
            self.formatter.apply_font_to_run(self.run_element, font_name)
            self.formatter.apply_size_to_run(self.run_element, size)
            self.formatter.apply_bold_to_run(self.run_element, is_bold)
            self.formatter.apply_italic_to_run(self.run_element, is_italic)
            self.formatter.apply_underline_to_run(self.run_element, self.underline_var.get())
            self.formatter.apply_color_to_run(self.run_element, self.current_color_hex)

            self.on_close_callback(confirmed=True)
            self.destroy()
            return

        # Откатываем rPr к снимку, сделанному при открытии диалога.
        current_rpr = self.run_element.find("w:rPr", NS_MAP)
        if current_rpr is not None:
            self.run_element.remove(current_rpr)
        if self._rpr_snapshot is not None:
            self.run_element.insert(0, copy.deepcopy(self._rpr_snapshot))
        # если rPr не существовал изначально и не был создан - ничего вставлять не нужно

        self.on_close_callback(confirmed=False)
        self.destroy()


# ============================== БЛОК АБЗАЦА ==============================

class ParagraphFrame(ctk.CTkFrame):
    """
    Разворачиваемый (collapsible) блок одного абзаца.
    Показывает превью текста и кнопку разворота на прогоны (ТЗ: "функция
    разворачивания абзацев на прогоны внутри этого конкретного абзаца").
    Также содержит combobox'ы для выравнивания/интервала абзаца и поля
    для отступов (в см) - параметры уровня абзаца, отдельные от прогонов.
    """

    def __init__(self, master, app, formatter: WordXMLFormatter, para_info: dict):
        super().__init__(master, corner_radius=8, fg_color=("gray85", "gray17"))
        self.app = app
        self.formatter = formatter
        self.para_info = para_info
        self.p_element = para_info["xml_element"]
        self.expanded = False

        # --- снимок pPr для возможной отмены изменений уровня абзаца ---
        self._ppr_snapshot = None

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=8, pady=8)

        self.toggle_btn = ctk.CTkButton(
            header, text="▶", width=28, command=self._toggle_expand,
        )
        self.toggle_btn.pack(side="left", padx=(0, 8))

        # БАГ (BUG_PARA_PREVIEW_SHOWS_RUN): вместо текста самого абзаца
        # показывается текст ПЕРВОГО прогона внутри него. Чисто визуальный
        # дефект - данные (para_info, xml_element) не искажены, подменяется
        # только то, что видит пользователь в свёрнутом заголовке.
        if BUG_PARA_PREVIEW_SHOWS_RUN and para_info["runs"]:
            run_text = para_info["runs"][0]["text"]
            preview = run_text if run_text.strip() else "(пустой абзац)"
        else:
            preview = para_info["preview"] if para_info["preview"].strip() else "(пустой абзац)"
        ctk.CTkLabel(
            header, text=f"Абзац {para_info['id'] + 1}: {preview}",
            anchor="w", font=ctk.CTkFont(size=13),
        ).pack(side="left", fill="x", expand=True)

        # --- Панель свойств абзаца: выравнивание, интервал, отступы ---
        self.para_controls = ctk.CTkFrame(self, fg_color="transparent")
        # изначально скрыта - показывается вместе с прогонами при разворачивании

        row1 = ctk.CTkFrame(self.para_controls, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(row1, text="Выравнивание:", width=110, anchor="w").pack(side="left")
        current_align = self._read_current_alignment()
        self.align_var = ctk.StringVar(value=ALIGNMENT_CHOICES_REVERSE.get(current_align, "По левому краю"))
        ctk.CTkComboBox(
            row1, values=list(ALIGNMENT_CHOICES.keys()), variable=self.align_var,
            command=self._on_alignment_change, width=180,
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(row1, text="Интервал:", width=80, anchor="w").pack(side="left")
        current_spacing = self._read_current_spacing()
        self.spacing_var = ctk.StringVar(value=LINE_SPACING_REVERSE.get(current_spacing, "Одинарный"))
        ctk.CTkComboBox(
            row1, values=list(LINE_SPACING_CHOICES.keys()), variable=self.spacing_var,
            command=self._on_spacing_change, width=140,
        ).pack(side="left")

        row2 = ctk.CTkFrame(self.para_controls, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(row2, text="Отступы (см) - слева / справа / красная строка:", anchor="w").pack(side="left")
        left_cm, right_cm, first_line_cm = self._read_current_indents()
        self.indent_left_entry = ctk.CTkEntry(row2, width=60, placeholder_text="0")
        self.indent_left_entry.insert(0, str(left_cm))
        self.indent_left_entry.pack(side="left", padx=(8, 4))
        self.indent_right_entry = ctk.CTkEntry(row2, width=60, placeholder_text="0")
        self.indent_right_entry.insert(0, str(right_cm))
        self.indent_right_entry.pack(side="left", padx=4)
        self.indent_first_entry = ctk.CTkEntry(row2, width=60, placeholder_text="0")
        self.indent_first_entry.insert(0, str(first_line_cm))
        self.indent_first_entry.pack(side="left", padx=4)
        ctk.CTkButton(row2, text="Применить отступы", command=self._apply_indents, width=140).pack(side="left", padx=(8, 0))

        # --- Контейнер прогонов (скрыт до разворачивания) ---
        self.runs_container = ctk.CTkFrame(self, fg_color="transparent")
        self._build_run_rows()

    # -------------------- чтение текущих свойств абзаца из XML --------------------

    def _read_current_alignment(self) -> str:
        ppr = self.p_element.find("w:pPr", NS_MAP)
        if ppr is not None:
            jc = ppr.find("w:jc", NS_MAP)
            if jc is not None:
                return jc.get(f"{{{W_NS}}}val", "left")
        return "left"

    def _read_current_spacing(self) -> str:
        ppr = self.p_element.find("w:pPr", NS_MAP)
        if ppr is not None:
            spacing = ppr.find("w:spacing", NS_MAP)
            if spacing is not None:
                line_val = spacing.get(f"{{{W_NS}}}line")
                reverse_map = {"240": "single", "360": "1.5", "480": "double"}
                return reverse_map.get(line_val, "single")
        return "single"

    def _read_current_indents(self):
        TWIPS_PER_CM = 567.0
        ppr = self.p_element.find("w:pPr", NS_MAP)
        left = right = first = 0.0
        if ppr is not None:
            ind = ppr.find("w:ind", NS_MAP)
            if ind is not None:
                left_raw = ind.get(f"{{{W_NS}}}left")
                right_raw = ind.get(f"{{{W_NS}}}right")
                first_raw = ind.get(f"{{{W_NS}}}firstLine")
                if left_raw:
                    left = round(int(left_raw) / TWIPS_PER_CM, 2)
                if right_raw:
                    right = round(int(right_raw) / TWIPS_PER_CM, 2)
                if first_raw:
                    first = round(int(first_raw) / TWIPS_PER_CM, 2)
        return left, right, first

    # -------------------- разворот/сворот --------------------

    def _toggle_expand(self):
        self.expanded = not self.expanded
        if self.expanded:
            self.toggle_btn.configure(text="▼")
            self.para_controls.pack(fill="x", padx=(44, 8), pady=(0, 8))
            self.runs_container.pack(fill="x", padx=(44, 8), pady=(0, 8))
        else:
            self.toggle_btn.configure(text="▶")
            self.para_controls.pack_forget()
            self.runs_container.pack_forget()

    def _build_run_rows(self):
        for run_info in self.para_info["runs"]:
            row = ctk.CTkFrame(self.runs_container, fg_color=("gray75", "gray25"), corner_radius=6)
            row.pack(fill="x", pady=3)

            text_preview = run_info["text"] if run_info["text"].strip() else "(пусто)"
            text_preview = text_preview[:50] + ("..." if len(text_preview) > 50 else "")
            ctk.CTkLabel(row, text=f'"{text_preview}"', anchor="w").pack(
                side="left", padx=8, pady=6, fill="x", expand=True
            )

            style_bits = []
            if run_info.get("is_bold"):
                style_bits.append("Ж")
            if run_info.get("is_italic"):
                style_bits.append("К")
            if run_info.get("is_underline"):
                style_bits.append("П")
            style_label = " ".join(style_bits) if style_bits else "—"
            ctk.CTkLabel(row, text=f'{run_info.get("font", "?")}, {run_info.get("size", "?")}, {style_label}', width=200).pack(
                side="left", padx=8
            )

            ctk.CTkButton(
                row, text="Редактировать", width=110,
                command=lambda ri=run_info: self._open_run_dialog(ri),
            ).pack(side="right", padx=8, pady=6)

    def _open_run_dialog(self, run_info):
        def _on_dialog_close(confirmed: bool):
            # После подтверждения перечитываем свойства прогона из XML,
            # чтобы обновить подпись в строке (шрифт/размер/начертание),
            # и заново собираем модель GUI, т.к. gui_helper строит её один раз
            # при загрузке и не отслеживает последующие мутации сам.
            self.app.refresh_model_and_view()

        RunEditDialog(self.app, self.formatter, run_info, _on_dialog_close)

    # -------------------- обработчики свойств абзаца --------------------

    def _on_alignment_change(self, choice):
        align_val = ALIGNMENT_CHOICES[choice]
        self.formatter.apply_alignment_to_para(self.p_element, align_val)

    def _on_spacing_change(self, choice):
        spacing_val = LINE_SPACING_CHOICES[choice]
        self.formatter.apply_spacing_to_para(self.p_element, spacing_val)

    def _apply_indents(self):
        try:
            left = float(self.indent_left_entry.get().strip().replace(",", ".") or 0)
            right = float(self.indent_right_entry.get().strip().replace(",", ".") or 0)
            first = float(self.indent_first_entry.get().strip().replace(",", ".") or 0)
        except ValueError:
            messagebox.showwarning("Проверка", "Отступы должны быть числами.", parent=self.app)
            return
        if left < 0 or right < 0 or first < 0:
            messagebox.showwarning(
                "Проверка",
                "Отступы не могут быть отрицательными "
                "(движок форматирования игнорирует отрицательные значения).",
                parent=self.app,
            )
            return
        self.formatter.apply_indents_to_para(self.p_element, left, right, first)
        messagebox.showinfo("Готово", "Отступы применены к абзацу.", parent=self.app)


# ============================== ГЛАВНОЕ ОКНО ==============================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.title("Редактор форматирования DOCX")
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(APP_WIDTH, APP_HEIGHT)

        self.doc_model: DocumentModel | None = None
        self.formatter: WordXMLFormatter | None = None
        self.input_path: str | None = None

        # Снимок для отмены "реформатирования всего файла" (пункт ТЗ:
        # "Пользователь должен иметь возможность изменить решение и
        # продолжить форматирование файла")
        self._file_reformat_snapshot = None
        self._file_reformat_committed = False

        self._build_toolbar()
        self._build_scroll_area()
        self._build_statusbar()

        self.paragraph_frames: list[ParagraphFrame] = []

    # -------------------- построение статичных частей UI --------------------

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=50, corner_radius=0)
        toolbar.pack(fill="x", side="top")

        ctk.CTkButton(toolbar, text="Открыть DOCX...", command=self._on_open_file, width=140).pack(
            side="left", padx=6, pady=8
        )
        ctk.CTkButton(
            toolbar, text="Реформатировать весь файл...", command=self._on_reformat_whole_file, width=200
        ).pack(side="left", padx=6, pady=8)

        # Кнопка меняет подпись в зависимости от состояния "завершено/не завершено",
        # реализуя переключаемое подтверждение из ТЗ.
        self.finish_btn = ctk.CTkButton(
            toolbar, text="Завершить форматирование файла", command=self._on_toggle_finish, width=230
        )
        self.finish_btn.pack(side="left", padx=6, pady=8)

        # БАГ (BUG_SAVE_ALWAYS_ENABLED): кнопка создаётся сразу активной,
        # минуя обычное состояние "disabled до завершения форматирования".
        initial_save_state = "normal" if BUG_SAVE_ALWAYS_ENABLED else "disabled"
        self.save_btn = ctk.CTkButton(
            toolbar, text="Сохранить как...", command=self._on_save_as, width=140, state=initial_save_state
        )
        self.save_btn.pack(side="right", padx=6, pady=8)

    def _build_scroll_area(self):
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Абзацы документа")
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        self.placeholder_label = ctk.CTkLabel(
            self.scroll_frame, text="Откройте .docx файл, чтобы начать редактирование.",
            font=ctk.CTkFont(size=14),
        )
        self.placeholder_label.pack(pady=40)

    def _build_statusbar(self):
        status_frame = ctk.CTkFrame(self, height=40, corner_radius=0)
        status_frame.pack(fill="x", side="bottom")

        self.status_label = ctk.CTkLabel(status_frame, text="Готово", anchor="w")
        self.status_label.pack(side="left", padx=10, fill="x", expand=True)

        # Прогрессбар в indeterminate-режиме для длительных операций
        # (загрузка/сохранение), как требует ТЗ.
        self.progress_bar = ctk.CTkProgressBar(status_frame, mode="indeterminate", width=200)
        self.progress_bar.pack(side="right", padx=10, pady=8)
        self.progress_bar.set(0)

    def _set_busy(self, is_busy: bool, message: str = ""):
        """Включает/выключает индикатор прогресса и блокирует кнопки на время фоновой операции."""
        if is_busy:
            self.status_label.configure(text=message)
            self.progress_bar.start()
        else:
            self.progress_bar.stop()
            self.progress_bar.set(0)
            self.status_label.configure(text=message or "Готово")

    # -------------------- открытие файла --------------------

    def _on_open_file(self):
        if BUG_ACCEPT_ANY_FILE:
            # БАГ: фильтр типов файлов снят - можно выбрать что угодно.
            # Ошибка проявится позже, в DocumentModel(path), когда файл
            # окажется не валидным .docx/zip-архивом.
            path = filedialog.askopenfilename(
                title="Выберите DOCX файл",
                filetypes=[("Все файлы", "*.*")],
            )
        else:
            path = filedialog.askopenfilename(
                title="Выберите DOCX файл",
                filetypes=[("Word Document", "*.docx")],  # ТЗ: "только docx"
            )
        if not path:
            return

        self._set_busy(True, "Загрузка документа...")

        def worker():
            try:
                model = DocumentModel(path)
                error = None
            except Exception as exc:  # noqa: BLE001 - показываем пользователю любую ошибку загрузки
                model = None
                error = exc
            self.after(0, lambda: self._on_document_loaded(path, model, error))

        threading.Thread(target=worker, daemon=True).start()

    def _on_document_loaded(self, path, model, error):
        self._set_busy(False)
        if error is not None:
            messagebox.showerror(
                "Ошибка загрузки",
                f"Не удалось открыть файл как DOCX:\n{error}",
                parent=self,
            )
            return

        self.doc_model = model
        self.formatter = model.formatter
        self.input_path = path
        self._file_reformat_snapshot = None
        self._file_reformat_committed = False
        self.finish_btn.configure(text="Завершить форматирование файла")
        # БАГ (BUG_SAVE_ALWAYS_ENABLED): при загрузке НОВОГО файла кнопка
        # "Сохранить как" не блокируется заново, остаётся доступной.
        if not BUG_SAVE_ALWAYS_ENABLED:
            self.save_btn.configure(state="disabled")

        self._rebuild_paragraph_view()

        real_paragraph_count = len(self.doc_model.get_model())
        # БАГ (BUG_PARAGRAPH_COUNT_OFF_BY_ONE): отображаемое число абзацев
        # на единицу больше реального (классический off-by-one).
        displayed_count = real_paragraph_count + 1 if BUG_PARAGRAPH_COUNT_OFF_BY_ONE else real_paragraph_count
        self.status_label.configure(
            text=f"Загружен: {os.path.basename(path)} ({displayed_count} абз.)"
        )

    # -------------------- построение / перестроение списка абзацев --------------------

    def _rebuild_paragraph_view(self):
        for child in self.scroll_frame.winfo_children():
            child.destroy()
        self.paragraph_frames = []

        if self.doc_model is None:
            self.placeholder_label = ctk.CTkLabel(
                self.scroll_frame, text="Откройте .docx файл, чтобы начать редактирование.",
            )
            self.placeholder_label.pack(pady=40)
            return

        for para_info in self.doc_model.get_model():
            frame = ParagraphFrame(self.scroll_frame, self, self.formatter, para_info)
            frame.pack(fill="x", pady=4, padx=2)
            self.paragraph_frames.append(frame)

    def refresh_model_and_view(self):
        """
        Пересобирает данные модели (gui_helper строит их один раз при загрузке
        и не отслеживает последующие мутации XML сам) и перерисовывает список
        абзацев, сохраняя текущее состояние развёрнутости там, где это возможно.
        """
        if self.doc_model is None:
            return
        expanded_ids = {f.para_info["id"] for f in self.paragraph_frames if f.expanded}
        self.doc_model._build_gui_model()
        self._rebuild_paragraph_view()
        for frame in self.paragraph_frames:
            if frame.para_info["id"] in expanded_ids:
                frame._toggle_expand()

    # -------------------- реформатирование всего файла --------------------

    def _on_reformat_whole_file(self):
        if self.doc_model is None:
            messagebox.showwarning("Нет документа", "Сначала откройте .docx файл.", parent=self)
            return
        WholeFileReformatDialog(self, self.formatter, self._on_whole_file_reformatted)

    def _on_whole_file_reformatted(self):
        # Снимок для потенциальной последующей отмены на уровне "завершения файла"
        self._file_reformat_snapshot = copy.deepcopy(self.doc_model.xml_root)
        self.refresh_model_and_view()
        self.status_label.configure(text="Применено единое форматирование ко всему файлу.")

    # -------------------- завершение форматирования файла (с возможностью отмены) --------------------

    def _on_toggle_finish(self):
        if self.doc_model is None:
            messagebox.showwarning("Нет документа", "Сначала откройте .docx файл.", parent=self)
            return

        if not self._file_reformat_committed:
            # Переход в состояние "завершено": делаем снимок ТЕКУЩЕГО состояния
            # дерева, чтобы при "Изменить решение" вернуться именно сюда.
            self._file_reformat_snapshot = copy.deepcopy(self.doc_model.xml_root)
            self._file_reformat_committed = True
            self.finish_btn.configure(text="Изменить решение (продолжить форматирование)")
            self.save_btn.configure(state="normal")
            self.status_label.configure(text="Форматирование файла завершено. Можно сохранить.")
        else:
            # Откат к снимку и возврат в режим редактирования
            if self._file_reformat_snapshot is not None:
                self.doc_model.xml_root[:] = copy.deepcopy(self._file_reformat_snapshot)[:]
                # обновляем formatter, т.к. он хранит ссылку на self.root, а не на его детей -
                # ссылка на сам объект xml_root не меняется, поэтому formatter продолжит работать корректно
            self._file_reformat_committed = False
            self.finish_btn.configure(text="Завершить форматирование файла")
            self.save_btn.configure(state="disabled")
            self.refresh_model_and_view()
            self.status_label.configure(text="Возврат к редактированию форматирования.")

    # -------------------- сохранение --------------------

    def _on_save_as(self):
        if self.doc_model is None:
            return
        # БАГ (BUG_SAVE_ALWAYS_ENABLED): проверка завершённости форматирования
        # снята - можно сохранить документ, даже не нажав "Завершить
        # форматирование файла" (то есть не подтвердив вообще никаких правок).
        if not BUG_SAVE_ALWAYS_ENABLED and not self._file_reformat_committed:
            messagebox.showwarning(
                "Форматирование не завершено",
                'Сначала нажмите "Завершить форматирование файла".',
                parent=self,
            )
            return

        output_path = filedialog.asksaveasfilename(
            title="Сохранить как...",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
        )
        if not output_path:
            return

        self._set_busy(True, "Сохранение документа...")

        def worker():
            try:
                self.doc_model.apply_changes_and_save(output_path)
                error = None
            except Exception as exc:  # noqa: BLE001
                error = exc
            self.after(0, lambda: self._on_save_finished(output_path, error))

        threading.Thread(target=worker, daemon=True).start()

    def _on_save_finished(self, output_path, error):
        self._set_busy(False)
        if error is not None:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл:\n{error}", parent=self)
            return
        messagebox.showinfo("Готово", f"Документ сохранён:\n{output_path}", parent=self)
        self.status_label.configure(text=f"Сохранено: {os.path.basename(output_path)}")


# ============================== ДИАЛОГ-ПАСХАЛКА (BUG_EASTER_EGG_SIZE_67) ==============================

class EasterEggImageDialog(ctk.CTkToplevel):
    """
    Всплывающее окно с картинкой - срабатывает, когда при реформатировании
    всего файла введён размер шрифта 67 (см. BUG_EASTER_EGG_SIZE_67).
    Существует ИСКЛЮЧИТЕЛЬНО для тестирования; при BUGGY_MODE = False
    этот путь кода никогда не вызывается.
    """

    def __init__(self, master):
        super().__init__(master)
        self.title("?!")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        if not _PIL_AVAILABLE:
            # Pillow не установлен - показываем понятную ошибку вместо
            # падения всего приложения с ModuleNotFoundError.
            self.geometry("380x140")
            ctk.CTkLabel(
                self, text="Секретная картинка не может быть показана:\n"
                           "не установлена библиотека Pillow (pip install Pillow).",
                wraplength=340, justify="center",
            ).pack(padx=16, pady=(24, 12))
            ctk.CTkButton(self, text="Закрыть", command=self.destroy, width=120).pack(pady=(0, 16))
            return

        if not os.path.isfile(EASTER_EGG_IMAGE_PATH):
            self.geometry("380x140")
            ctk.CTkLabel(
                self, text=f"Секретная картинка не найдена:\n{EASTER_EGG_IMAGE_PATH}",
                wraplength=340, justify="center",
            ).pack(padx=16, pady=(24, 12))
            ctk.CTkButton(self, text="Закрыть", command=self.destroy, width=120).pack(pady=(0, 16))
            return

        pil_image = Image.open(EASTER_EGG_IMAGE_PATH)
        # Ограничиваем размер окна разумным максимумом, сохраняя пропорции,
        # на случай если картинку в будущем заменят на что-то крупное.
        max_dim = 420
        w, h = pil_image.size
        scale = min(1.0, max_dim / max(w, h))
        display_size = (int(w * scale), int(h * scale))
        if scale < 1.0:
            pil_image = pil_image.resize(display_size, Image.LANCZOS)

        self._tk_image = ImageTk.PhotoImage(pil_image)  # ссылка сохранена на self - иначе GC соберёт картинку

        self.geometry(f"{display_size[0] + 40}x{display_size[1] + 90}")

        image_label = ctk.CTkLabel(self, image=self._tk_image, text="")
        image_label.pack(padx=16, pady=(16, 8))

        ctk.CTkButton(self, text="Закрыть", command=self.destroy, width=140).pack(pady=(0, 16))


class WholeFileReformatDialog(ctk.CTkToplevel):
    """
    Отдельная опция реформатирования ВСЕГО файла под один выбранный
    пользователем формат (пункт ТЗ). Использует глобальные change_* методы
    форматтера, которые проходят по каждому прогону/абзацу документа.

    Внимание: применение здесь перезаписывает точечные правки, сделанные
    ранее для отдельных прогонов/абзацов - это ожидаемое поведение
    (согласовано отдельно, т.к. в ТЗ явно не описан приоритет между
    точечным и глобальным форматированием).
    """

    def __init__(self, master, formatter: WordXMLFormatter, on_applied_callback):
        super().__init__(master)
        self.formatter = formatter
        self.on_applied_callback = on_applied_callback

        self.title("Реформатировать весь файл")
        self.geometry("420x600")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        ctk.CTkLabel(
            self, text="Единый формат будет применён ко ВСЕМ абзацам и прогонам документа.\n"
                       "Ранее сделанные точечные правки могут быть перезаписаны.",
            wraplength=380, justify="left", text_color=("gray20", "gray70"),
        ).pack(padx=16, pady=(16, 12))

        ctk.CTkLabel(self, text="Шрифт:").pack(padx=16, anchor="w")
        self.font_var = ctk.StringVar(value=FONT_CHOICES[0])
        self.font_combo = ctk.CTkComboBox(
            self, values=FONT_CHOICES, variable=self.font_var, command=self._on_font_combo_change, width=380,
        )
        self.font_combo.pack(padx=16, pady=(0, 4))
        self.custom_font_entry = ctk.CTkEntry(self, placeholder_text="Введите название шрифта", width=380)
        self.custom_font_entry.pack_forget()

        ctk.CTkLabel(self, text="Размер шрифта (пт):").pack(padx=16, anchor="w")
        self.size_entry = ctk.CTkEntry(self, width=380, placeholder_text="Например, 14")
        self.size_entry.insert(0, "14")
        self.size_entry.pack(padx=16, pady=(0, 8))

        ctk.CTkLabel(self, text="Начертание:").pack(padx=16, anchor="w")
        self.style_var = ctk.StringVar(value="Обычный")
        ctk.CTkComboBox(self, values=list(STYLE_CHOICES.keys()), variable=self.style_var, width=380).pack(
            padx=16, pady=(0, 8)
        )
        self.underline_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self, text="Подчёркнутый", variable=self.underline_var).pack(padx=16, pady=(0, 8), anchor="w")

        ctk.CTkLabel(self, text="Цвет текста:").pack(padx=16, anchor="w")
        self.current_color_hex = "#000000"
        color_row = ctk.CTkFrame(self, fg_color="transparent")
        color_row.pack(padx=16, pady=(0, 8), fill="x")
        self.color_preview = ctk.CTkLabel(
            color_row, text="", width=30, height=24, fg_color=self.current_color_hex, corner_radius=4
        )
        self.color_preview.pack(side="left", padx=(0, 8))
        ctk.CTkButton(color_row, text="Выбрать цвет...", command=self._pick_color, width=200).pack(side="left")

        ctk.CTkLabel(self, text="Выравнивание:").pack(padx=16, anchor="w")
        self.align_var = ctk.StringVar(value="По левому краю")
        ctk.CTkComboBox(self, values=list(ALIGNMENT_CHOICES.keys()), variable=self.align_var, width=380).pack(
            padx=16, pady=(0, 8)
        )

        ctk.CTkLabel(self, text="Межстрочный интервал:").pack(padx=16, anchor="w")
        self.spacing_var = ctk.StringVar(value="Одинарный")
        ctk.CTkComboBox(self, values=list(LINE_SPACING_CHOICES.keys()), variable=self.spacing_var, width=380).pack(
            padx=16, pady=(0, 8)
        )

        ctk.CTkLabel(self, text="Отступы (см) - слева / справа / красная строка:").pack(padx=16, anchor="w")
        indent_row = ctk.CTkFrame(self, fg_color="transparent")
        indent_row.pack(padx=16, pady=(0, 8), fill="x")
        self.indent_left_entry = ctk.CTkEntry(indent_row, width=110, placeholder_text="0")
        self.indent_left_entry.insert(0, "0")
        self.indent_left_entry.pack(side="left", padx=(0, 8))
        self.indent_right_entry = ctk.CTkEntry(indent_row, width=110, placeholder_text="0")
        self.indent_right_entry.insert(0, "0")
        self.indent_right_entry.pack(side="left", padx=8)
        self.indent_first_entry = ctk.CTkEntry(indent_row, width=110, placeholder_text="0")
        self.indent_first_entry.insert(0, "0")
        self.indent_first_entry.pack(side="left", padx=8)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=16, pady=(16, 16), fill="x", side="bottom")
        ctk.CTkButton(btn_row, text="Отмена", command=self.destroy, fg_color="gray40", width=170).pack(side="left")
        ctk.CTkButton(btn_row, text="Применить ко всему файлу", command=self._on_apply, width=190).pack(side="right")

    def _on_font_combo_change(self, choice):
        if choice == "Другой...":
            self.custom_font_entry.delete(0, "end")
            self.custom_font_entry.pack(padx=16, pady=(0, 8))
        else:
            self.custom_font_entry.pack_forget()

    def _pick_color(self):
        _, hex_color = colorchooser.askcolor(color=self.current_color_hex, title="Выберите цвет текста")
        if hex_color:
            self.current_color_hex = hex_color
            self.color_preview.configure(fg_color=hex_color)

    def _on_apply(self):
        font_choice = self.font_var.get()
        font_name = self.custom_font_entry.get().strip() if font_choice == "Другой..." else font_choice
        if not font_name:
            messagebox.showwarning("Проверка", "Введите название шрифта.", parent=self)
            return

        raw_size = self.size_entry.get().strip().replace(",", ".")
        try:
            size = float(raw_size)
        except ValueError:
            messagebox.showwarning("Проверка", "Размер шрифта должен быть числом.", parent=self)
            return
        if not (FONT_SIZE_MIN <= size <= FONT_SIZE_MAX):
            messagebox.showwarning(
                "Проверка", f"Размер шрифта должен быть от {FONT_SIZE_MIN} до {FONT_SIZE_MAX} пт.", parent=self
            )
            return

        # БАГ-ПАСХАЛКА (BUG_EASTER_EGG_SIZE_67): реформатирование всего файла
        # с размером шрифта РОВНО 67 не применяет форматирование вообще -
        # вместо этого показывается картинка, и диалог закрывается без
        # каких-либо изменений в документе (остальные поля формы игнорируются).
        if BUG_EASTER_EGG_SIZE_67 and size == 67:
            EasterEggImageDialog(self)

            return

        try:
            left_cm = float(self.indent_left_entry.get().strip().replace(",", ".") or 0)
            right_cm = float(self.indent_right_entry.get().strip().replace(",", ".") or 0)
            first_cm = float(self.indent_first_entry.get().strip().replace(",", ".") or 0)
        except ValueError:
            messagebox.showwarning("Проверка", "Отступы должны быть числами.", parent=self)
            return
        if left_cm < 0 or right_cm < 0 or first_cm < 0:
            messagebox.showwarning(
                "Проверка",
                "Отступы не могут быть отрицательными "
                "(движок форматирования игнорирует отрицательные значения).",
                parent=self,
            )
            return

        is_bold, is_italic = STYLE_CHOICES[self.style_var.get()]
        is_underline = self.underline_var.get()
        align_val = ALIGNMENT_CHOICES[self.align_var.get()]
        spacing_val = LINE_SPACING_CHOICES[self.spacing_var.get()]

        # Применяем ГЛОБАЛЬНО - через change_* методы, проходящие по всем прогонам/абзацам
        self.formatter.change_font_name(font_name)
        self.formatter.change_font_size(size)
        self.formatter.toggle_bold(is_bold)
        self.formatter.toggle_italic(is_italic)
        self.formatter.toggle_underline(is_underline)
        # ВАЖНО: change_text_color (в отличие от apply_color_to_run) не убирает
        # ведущий '#' из hex-значения перед записью в XML - Word требует hex без
        # '#' (например, val="333333", а не val="#333333"). Нормализуем здесь на
        # стороне вызова, чтобы не полагаться на исправление в formatter.py.
        self.formatter.change_text_color(self.current_color_hex.lstrip("#"))
        self.formatter.change_paragraph_alignment(align_val)
        self.formatter.change_line_spacing(spacing_val)
        self.formatter.change_paragraph_indents(left_cm, right_cm, first_cm)

        self.on_applied_callback()
        self.destroy()


# ============================== ТОЧКА ВХОДА ==============================

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()