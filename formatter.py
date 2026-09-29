# используем lxml для сохранения неймспейсов и целостности xml word
from turtle import color

import lxml.etree as ET

# основной адрес пространства имен wordprocessingml
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
# словарь для удобного поиска тегов с префиксом 'w:'
NS_MAP = {'w': W_NS}
# регистрируем префикс 'w', чтобы избежать мусорных ns0, ns1 при сохранении
ET.register_namespace('w', W_NS)

class WordXMLFormatter:
    # класс-форматтер для изменения стилей в xml-дереве документа
    def __init__(self, xml_root):
        # инициализация класса корневым элементом xml
        self.root = xml_root # сохраняем корень дерева для последующего обхода

    def _get_or_create_rpr(self, run_element):
        # вспомогательный метод: получить или создать свойства текста (w:rPr)
        rpr = run_element.find('w:rPr', NS_MAP) # ищем существующий тег свойств внутри фрагмента
        if rpr is None: # если тега нет, его нужно создать
            rpr = ET.Element(f'{{{W_NS}}}rPr') # создаем новый пустой элемент свойств
            run_element.insert(0, rpr) # вставляем строго в начало (жесткое требование схемы word)
        return rpr # возвращаем найденный или созданный тег

    def _get_or_create_ppr(self, paragraph_element):
        # вспомогательный метод: получить или создать свойства абзаца (w:pPr)
        ppr = paragraph_element.find('w:pPr', NS_MAP) # ищем существующий тег свойств абзаца
        if ppr is None: # если тега нет, его нужно создать
            ppr = ET.Element(f'{{{W_NS}}}pPr') # создаем новый пустой элемент свойств
            paragraph_element.insert(0, ppr) # вставляем строго в начало абзаца
        return ppr # возвращаем найденный или созданный тег

    # ==================== 1. шрифт и размер ====================

    def change_font_name(self, font_name):
        # метод изменения гарнитуры шрифта во всем документе
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты (runs)
            rpr = self._get_or_create_rpr(run) # получаем блок свойств для текущего фрагмента
            rFonts = rpr.find('w:rFonts', NS_MAP) # ищем тег настроек шрифтов
            if rFonts is None: # если настроек нет, создаем тег
                rFonts = ET.SubElement(rpr, f'{{{W_NS}}}rFonts') # добавляем тег шрифтов в свойства
            
            # задаем шрифт для всех категорий, чтобы кириллица не осталась старой
            rFonts.set(f'{{{W_NS}}}ascii', font_name) # задаем шрифт для латиницы
            rFonts.set(f'{{{W_NS}}}hAnsi', font_name) # задаем шрифт для high-ansi символов
            rFonts.set(f'{{{W_NS}}}cs', font_name) # задаем шрифт для кириллицы и сложных скриптов
            rFonts.set(f'{{{W_NS}}}eastAsia', font_name) # задаем шрифт для азиатских иероглифов

    def change_font_size(self, size_in_points):
        # метод изменения размера шрифта
        size_half_points = str(int(size_in_points * 2)) # word хранит размер в полу-пунктах, умножаем на 2
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты
            rpr = self._get_or_create_rpr(run) # получаем блок свойств фрагмента
            
            sz = rpr.find('w:sz', NS_MAP) # ищем тег основного размера
            if sz is None: sz = ET.SubElement(rpr, f'{{{W_NS}}}sz') # создаем тег, если его нет
            sz.set(f'{{{W_NS}}}val', size_half_points) # устанавливаем значение размера
            
            szCs = rpr.find('w:szCs', NS_MAP) # ищем тег размера для сложных скриптов (кириллицы)
            if szCs is None: szCs = ET.SubElement(rpr, f'{{{W_NS}}}szCs') # создаем тег, если его нет
            szCs.set(f'{{{W_NS}}}val', size_half_points) # устанавливаем значение размера для кириллицы

    # ==================== 2. начертание ====================

    def toggle_bold(self, is_bold=True):
        # метод включения или выключения жирного шрифта
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты
            rpr = self._get_or_create_rpr(run) # получаем блок свойств
            b = rpr.find('w:b', NS_MAP) # ищем тег жирного начертания
            if is_bold and b is None: # если нужно включить, а тега нет
                ET.SubElement(rpr, f'{{{W_NS}}}b') # добавляем пустой тег (в word это означает "включено")
            elif not is_bold and b is not None: # если нужно выключить, а тег есть
                rpr.remove(b) # физически удаляем тег из xml для чистоты

    def toggle_italic(self, is_italic=True):
        # метод включения или выключения курсива
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты
            rpr = self._get_or_create_rpr(run) # получаем блок свойств
            i = rpr.find('w:i', NS_MAP) # ищем тег курсива
            if is_italic and i is None: # если нужно включить, а тега нет
                ET.SubElement(rpr, f'{{{W_NS}}}i') # добавляем пустой тег
            elif not is_italic and i is not None: # если нужно выключить, а тег есть
                rpr.remove(i) # физически удаляем тег из xml

    def toggle_underline(self, is_underlined=True, style='single'):
        # метод включения или выключения подчеркивания заданного стиля
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты
            rpr = self._get_or_create_rpr(run) # получаем блок свойств
            u = rpr.find('w:u', NS_MAP) # ищем тег подчеркивания
            if is_underlined: # если нужно включить
                if u is None: u = ET.SubElement(rpr, f'{{{W_NS}}}u') # создаем тег, если его нет
                u.set(f'{{{W_NS}}}val', style) # задаем тип подчеркивания (single, double, wave и т.д.)
            elif u is not None: # если нужно выключить, а тег есть
                rpr.remove(u) # физически удаляем тег из xml

    def change_text_color(self, hex_color):
        # метод изменения цвета текста
        for run in self.root.iter(f'{{{W_NS}}}r'): # перебираем все текстовые фрагменты
            rpr = self._get_or_create_rpr(run) # получаем блок свойств
            color = rpr.find('w:color', NS_MAP) # ищем тег цвета
            if color is None: # если тега нет, создаем его
                color = ET.SubElement(rpr, f'{{{W_NS}}}color') # добавляем тег цвета в свойства
            color.set(f'{{{W_NS}}}val', hex_color.upper()) # задаем цвет (word требует hex без # и в верхнем регистре)

    # ==================== 3. абзац и интервалы ====================

    def change_paragraph_alignment(self, alignment='left'):
        # метод изменения выравнивания абзацев
        # словарь перевода понятных слов в стандарты xml word ('both' = по ширине)
        align_map = {'left': 'left', 'center': 'center', 'right': 'right', 'both': 'both'}
        val = align_map.get(alignment, 'left') # получаем значение xml, по умолчанию 'left'

        for p in self.root.iter(f'{{{W_NS}}}p'): # перебираем все абзацы в документе
            ppr = self._get_or_create_ppr(p) # получаем блок свойств абзаца
            jc = ppr.find('w:jc', NS_MAP) # ищем тег выравнивания
            if jc is None: # если тега нет, создаем его
                jc = ET.SubElement(ppr, f'{{{W_NS}}}jc') # добавляем тег выравнивания
            jc.set(f'{{{W_NS}}}val', val) # устанавливаем выбранное выравнивание

    def change_line_spacing(self, spacing_type='single'):
        # метод изменения междустрочного интервала
        # word использует доли строки: 240=1.0, 360=1.5, 480=2.0
        spacing_map = {'single': '240', '1.5': '360', 'double': '480'}
        val = spacing_map.get(str(spacing_type), '240') # получаем числовое значение, по умолчанию 240

        for p in self.root.iter(f'{{{W_NS}}}p'): # перебираем все абзацы
            ppr = self._get_or_create_ppr(p) # получаем блок свойств абзаца
            spacing = ppr.find('w:spacing', NS_MAP) # ищем тег интервалов
            if spacing is None: # если тега нет, создаем его
                spacing = ET.SubElement(ppr, f'{{{W_NS}}}spacing') # добавляем тег интервалов
            spacing.set(f'{{{W_NS}}}line', val) # задаем множитель строки
            spacing.set(f'{{{W_NS}}}lineRule', 'auto') # указываем правило "авто" (множитель, а не фиксированная высота)

    # ==================== 4. отступы ====================

    def change_paragraph_indents(self, left_cm=0.0, right_cm=0.0, first_line_cm=0.0):
        # метод управления отступами абзаца в сантиметрах
        TWIPS_PER_CM = 567.0 # константа: 1 сантиметр равен 567 твипам (единица измерения word)
        
        # переводим сантиметры в твипы и округляем до целых чисел
        left_twips = str(int(round(left_cm * TWIPS_PER_CM))) # левый отступ в твипах
        right_twips = str(int(round(right_cm * TWIPS_PER_CM))) # правый отступ в твипах
        first_line_twips = str(int(round(first_line_cm * TWIPS_PER_CM))) # красная строка в твипах

        for p in self.root.iter(f'{{{W_NS}}}p'): # перебираем все абзацы
            ppr = self._get_or_create_ppr(p) # получаем блок свойств абзаца
            ind = ppr.find('w:ind', NS_MAP) # ищем тег отступов
            
            # если нужны отступы, а тега нет, создаем его
            if (left_cm > 0 or right_cm > 0 or first_line_cm > 0) and ind is None:
                ind = ET.SubElement(ppr, f'{{{W_NS}}}ind') # добавляем тег отступов
            elif ind is None: # если отступы не нужны и тега нет
                continue # пропускаем этот абзац, чтобы не засорять xml пустыми тегами

            # обработка левого отступа: задаем или удаляем атрибут для чистоты xml
            if left_cm > 0: # если задан левый отступ
                ind.set(f'{{{W_NS}}}left', left_twips) # устанавливаем значение левого отступа
            elif f'{{{W_NS}}}left' in ind.attrib: # если отступ равен 0, но атрибут существует
                del ind.attrib[f'{{{W_NS}}}left'] # удаляем атрибут, чтобы сбросить форматирование

            # обработка правого отступа: задаем или удаляем атрибут
            if right_cm > 0: # если задан правый отступ
                ind.set(f'{{{W_NS}}}right', right_twips) # устанавливаем значение правого отступа
            elif f'{{{W_NS}}}right' in ind.attrib: # если отступ равен 0, но атрибут существует
                del ind.attrib[f'{{{W_NS}}}right'] # удаляем атрибут

            # обработка красной строки (отступ первой строки): задаем или удаляем
            if first_line_cm > 0: # если задана красная строка
                ind.set(f'{{{W_NS}}}firstLine', first_line_twips) # устанавливаем значение отступа первой строки
            elif f'{{{W_NS}}}firstLine' in ind.attrib: # если отступ равен 0, но атрибут существует
                del ind.attrib[f'{{{W_NS}}}firstLine'] # удаляем атрибут
        # === ЭТИ МЕТОДЫ ДОЛЖНЫ БЫТЬ ВНУТРИ КЛАССА WordXMLFormatter ===
    
    def apply_font_to_run(self, run_element, font_name):
        rpr = self._get_or_create_rpr(run_element)
        rFonts = rpr.find('w:rFonts', NS_MAP)
        if rFonts is None:
            rFonts = ET.SubElement(rpr, f'{{{W_NS}}}rFonts')
        for attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
            rFonts.set(f'{{{W_NS}}}{attr}', font_name)

    def apply_size_to_run(self, run_element, size_in_points):
        rpr = self._get_or_create_rpr(run_element)
        half_points = str(int(size_in_points * 2))
        for tag in ['sz', 'szCs']:
            sz = rpr.find(f'w:{tag}', NS_MAP)
            if sz is None: sz = ET.SubElement(rpr, f'{{{W_NS}}}{tag}')
            sz.set(f'{{{W_NS}}}val', half_points)

    def apply_bold_to_run(self, run_element, is_bold=True):
        rpr = self._get_or_create_rpr(run_element)
        b = rpr.find('w:b', NS_MAP)
        if is_bold and b is None:
            ET.SubElement(rpr, f'{{{W_NS}}}b')
        elif not is_bold and b is not None:
            rpr.remove(b)

    def apply_alignment_to_para(self, p_element, alignment='left'):
        align_map = {'left': 'left', 'center': 'center', 'right': 'right', 'both': 'both'}
        val = align_map.get(alignment, 'left')
        ppr = self._get_or_create_ppr(p_element)
        jc = ppr.find('w:jc', NS_MAP)
        if jc is None:
            jc = ET.SubElement(ppr, f'{{{W_NS}}}jc')
        jc.set(f'{{{W_NS}}}val', val)
    def apply_italic_to_run(self, run_element, is_italic=True):
        rpr = self._get_or_create_rpr(run_element)
        i = rpr.find('w:i', NS_MAP)
        if is_italic and i is None:
            ET.SubElement(rpr, f'{{{W_NS}}}i')
        elif not is_italic and i is not None:
            rpr.remove(i)
    def apply_underline_to_run(self, run_element, is_underlined=True, style='single'):
        rpr = self._get_or_create_rpr(run_element)
        u = rpr.find('w:u', NS_MAP)
        if is_underlined:
            if u is None: u = ET.SubElement(rpr, f'{{{W_NS}}}u')
            u.set(f'{{{W_NS}}}val', style)
        elif u is not None:
            rpr.remove(u)
    def apply_color_to_run(self, run_element, hex_color):
        rpr = self._get_or_create_rpr(run_element)
        color = rpr.find('w:color', NS_MAP)
        if color is None:
            color = ET.SubElement(rpr, f'{{{W_NS}}}color')
        color.set(f'{{{W_NS}}}val', hex_color.upper().replace('#', ''))
    def apply_spacing_to_para(self, p_element, spacing_type='single'):
        spacing_map = {'single': '240', '1.5': '360', 'double': '480'}
        val = spacing_map.get(str(spacing_type), '240')
        ppr = self._get_or_create_ppr(p_element)
        spacing = ppr.find('w:spacing', NS_MAP)
        if spacing is None:
            spacing = ET.SubElement(ppr, f'{{{W_NS}}}spacing')
        spacing.set(f'{{{W_NS}}}line', val)
        spacing.set(f'{{{W_NS}}}lineRule', 'auto')
    def apply_indents_to_para(self, p_element, left_cm=0.0, right_cm=0.0, first_line_cm=0.0):
        TWIPS_PER_CM = 567.0
        left_twips = str(int(round(left_cm * TWIPS_PER_CM)))
        right_twips = str(int(round(right_cm * TWIPS_PER_CM)))
        first_line_twips = str(int(round(first_line_cm * TWIPS_PER_CM)))
    
        ppr = self._get_or_create_ppr(p_element)
        ind = ppr.find('w:ind', NS_MAP)
    
        if (left_cm > 0 or right_cm > 0 or first_line_cm > 0) and ind is None:
            ind = ET.SubElement(ppr, f'{{{W_NS}}}ind')
        elif ind is None:
            return
    
        if left_cm > 0: ind.set(f'{{{W_NS}}}left', left_twips)
        elif f'{{{W_NS}}}left' in ind.attrib: del ind.attrib[f'{{{W_NS}}}left']
    
        if right_cm > 0: ind.set(f'{{{W_NS}}}right', right_twips)
        elif f'{{{W_NS}}}right' in ind.attrib: del ind.attrib[f'{{{W_NS}}}right']
    
        if first_line_cm > 0: ind.set(f'{{{W_NS}}}firstLine', first_line_twips)
        elif f'{{{W_NS}}}firstLine' in ind.attrib: del ind.attrib[f'{{{W_NS}}}firstLine']