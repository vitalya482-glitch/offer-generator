# HVAC module for offer-generator

Комплект добавляет простую вкладку HVAC:

- выбирает HVAC calculation Excel;
- автоматически читает позиции из первого листа или выбранного листа;
- по умолчанию берёт строку суммы `DDP Almaty`, если её нет — `TOTAL`, затем `Total per quantity`;
- позволяет галочками включить позиции в Вариант 1 и/или Вариант 2;
- ищет шаблон КП внутри модуля: `brands/hvac/templates/HVAC_offer_template_TAGS.docx`;
- если шаблон не найден, поле шаблона остаётся пустым и можно выбрать файл вручную;
- формирует DOCX без прикладывания спецификаций.

## Куда положить файлы

Скопировать папки из архива в корень проекта:

```text
brands/hvac/__init__.py
brands/hvac/excel_reader.py
brands/hvac/offer_builder.py
brands/hvac/template_finder.py
brands/hvac/templates/HVAC_offer_template_TAGS.docx
gui/pages/hvac_page.py
```

## Подключение вкладки в `gui/main_window.py`

Вверху файла добавить импорт:

```python
from gui.pages.hvac_page import HVACPage
```

Там, где создаются вкладки `QTabWidget`, добавить:

```python
self.tabs.addTab(HVACPage(self), "HVAC")
```

Если в проекте объект вкладок называется не `self.tabs`, добавь аналогично рядом со Stulz/Riello.

## Зависимости

Для модуля нужны:

```text
openpyxl
python-docx
PySide6
```

Если `openpyxl` или `python-docx` уже есть в проекте — повторно добавлять не нужно.

## PyInstaller / spec

Чтобы встроенный шаблон попадал в exe, добавь в `datas`:

```python
("brands/hvac/templates/HVAC_offer_template_TAGS.docx", "brands/hvac/templates"),
```

Если сборка идёт через команду PyInstaller:

```bat
--add-data "brands/hvac/templates/HVAC_offer_template_TAGS.docx;brands/hvac/templates"
```

## Логика чтения Excel

Модуль ищет структуру вида:

```text
row 2: Model / % / Q-ty / Duct ... / % / Q-ty / AHU ...
col A: Quantity, DDP Almaty, TOTAL, Total per quantity
```

Количество берётся из строки `Quantity`, из колонки перед названием позиции.
Сумма берётся из строки `DDP Almaty` в колонке позиции.
