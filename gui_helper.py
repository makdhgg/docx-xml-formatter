# gui_helper.py
import zipfile
import lxml.etree as ET
from formatter import WordXMLFormatter # Класс твоего товарища

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

class DocumentModel:
    """
    Этот класс хранит "карту" документа для GUI.
    Он не меняет XML напрямую, а хранит удобные данные и ссылки на оригинальные XML-элементы.
    """
    def __init__(self, input_docx_path):
        self.input_path = input_docx_path
        self.formatter = None
        self.xml_root = None
        self.paragraphs_data = [] # Здесь будет лежать всё для GUI
        
        self._load_document()

    def _load_document(self):
        """Загружает и парсит документ, создавая модель для GUI"""
        # 1. Читаем XML из архива
        with zipfile.ZipFile(self.input_path, 'r') as docx_zip:
            xml_bytes = docx_zip.read('word/document.xml')
            
        # 2. Парсим XML (обязательно remove_blank_text=False для Word)
        parser = ET.XMLParser(remove_blank_text=False)
        self.xml_root = ET.fromstring(xml_bytes, parser=parser)
        
        # 3. Инициализируем форматтер твоего товарища
        self.formatter = WordXMLFormatter(self.xml_root)
        
        # 4. Строим "карту" для GUI
        self._build_gui_model()

    def _build_gui_model(self):
        """Собирает данные из XML в удобный список словарей для GUI"""
        self.paragraphs_data = []
        
        # Находим все абзацы в документе
        xml_paragraphs = list(self.xml_root.iter(f'{{{W_NS}}}p'))
        
        for para_index, p_elem in enumerate(xml_paragraphs):
            # Собираем весь текст абзаца для предпросмотра в GUI
            # (ищем все текстовые теги <w:t> внутри абзаца и склеиваем их)
            full_text = "".join(t.text or "" for t in p_elem.iter(f'{{{W_NS}}}t'))
            
            para_info = {
                'id': para_index, # Уникальный номер абзаца для GUI
                'preview': full_text[:60] + ("..." if len(full_text) > 60 else ""), # Короткий текст для списка
                'full_text': full_text, # Полный текст (если нужен)
                'xml_element': p_elem, # ВАЖНО: Ссылка на оригинальный XML абзаца!
                'runs': [] # Сюда положим информацию о прогонах
            }
            
            # Находим все прогоны внутри этого абзаца
            xml_runs = list(p_elem.iter(f'{{{W_NS}}}r'))
            
            for run_index, r_elem in enumerate(xml_runs):
                # Собираем текст конкретного прогона
                run_text = "".join(t.text or "" for t in r_elem.iter(f'{{{W_NS}}}t'))
                
                # Пытаемся узнать текущий шрифт и размер (для отображения в GUI)
                # Это упрощенная проверка, реальное форматирование может быть в стилях
                                # Пытаемся узнать текущее форматирование прогона
                rpr = r_elem.find('w:rPr', {'w': W_NS})
                
                current_font = "По умолчанию"
                current_size = "По умолчанию"
                is_bold = False
                is_italic = False
                is_underline = False
                text_color = "По умолчанию" # Или "000000" (черный)
                
                if rpr is not None:
                    # 1. Шрифт
                    rFonts = rpr.find('w:rFonts', {'w': W_NS})
                    if rFonts is not None and rFonts.get(f'{{{W_NS}}}ascii'):
                        current_font = rFonts.get(f'{{{W_NS}}}ascii')
                    
                    # 2. Размер
                    sz = rpr.find('w:sz', {'w': W_NS})
                    if sz is not None and sz.get(f'{{{W_NS}}}val'):
                        current_size = f"{int(sz.get(f'{{{W_NS}}}val')) // 2} pt"
                    
                    # 3. Жирный (наличие тега <w:b/> означает True)
                    if rpr.find('w:b', {'w': W_NS}) is not None:
                        is_bold = True
                        
                    # 4. Курсив (наличие тега <w:i/> означает True)
                    if rpr.find('w:i', {'w': W_NS}) is not None:
                        is_italic = True
                        
                    # 5. Подчеркивание
                    u = rpr.find('w:u', {'w': W_NS})
                    if u is not None and u.get(f'{{{W_NS}}}val') not in ['none', None]:
                        is_underline = True
                        
                    # 6. Цвет текста
                    color = rpr.find('w:color', {'w': W_NS})
                    if color is not None and color.get(f'{{{W_NS}}}val'):
                        text_color = color.get(f'{{{W_NS}}}val')

                run_info = {
                    'id': run_index,
                    'text': run_text,
                    'font': current_font,
                    'size': current_size,
                    'is_bold': is_bold,          # <--- ДОБАВЛЕНО
                    'is_italic': is_italic,      # <--- ДОБАВЛЕНО
                    'is_underline': is_underline,# <--- ДОБАВЛЕНО
                    'color': text_color,         # <--- ДОБАВЛЕНО
                    'xml_element': r_elem        # ГЛАВНОЕ: Ссылка на оригинальный XML
                }
                para_info['runs'].append(run_info)
                
            self.paragraphs_data.append(para_info)

    def get_model(self):
        """Возвращает готовые данные для отображения в GUI"""
        return self.paragraphs_data

    def apply_changes_and_save(self, output_docx_path):
        """
        Сохраняет все изменения, которые пользователь сделал через GUI, в новый файл.
        Вызывать только когда пользователь нажал кнопку "Сохранить".
        """
        if self.xml_root is None:
            raise ValueError("Документ не загружен")
            
        # Превращаем измененное XML-дерево обратно в байты
        modified_xml_bytes = ET.tostring(self.xml_root, encoding='UTF-8', xml_declaration=True)
        
        # Создаем новый архив, копируя всё из старого, но заменяя document.xml
        with zipfile.ZipFile(self.input_path, 'r') as zin:
            with zipfile.ZipFile(output_docx_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, modified_xml_bytes)
                    else:
                        zout.writestr(item, zin.read(item.filename))
                        
        print(f"Документ успешно сохранен в: {output_docx_path}")