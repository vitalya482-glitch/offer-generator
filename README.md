# SAM Offer Generator

Генератор коммерческих предложений из Excel-калькуляции и Word-шаблона.

## Новая логика

1. Пользователь выбирает папку проекта на сервере.
2. Программа сканирует папку и находит:
   - Excel-файлы: `.xlsx`, `.xlsm`
   - Word-шаблоны: `.docx`
3. Пользователь выбирает направление:
   - Stulz
   - Riello
   - DC Eltek
   - Generator
4. Пользователь выбирает Excel, Word-шаблон и лист Excel.
5. Программа формирует Word КП в выбранную папку результата.

## Структура

```text
app.py                  # точка входа
core/                   # общая логика
  models.py             # структуры данных
  excel_reader.py       # чтение Excel Stulz
  docx_renderer.py      # заполнение Word и таблиц
  project_scanner.py    # поиск файлов в папке проекта
  utils.py              # общие функции
brands/                 # логика направлений
  stulz.py              # рабочая логика Stulz
  riello.py             # заготовка
  dc_eltek.py           # заготовка
  generator.py          # заготовка
gui/
  main_window.py        # главное окно
config/
  managers.json         # задел под менеджеров
  signers.json          # задел под подписантов
```

## Запуск GUI

```bash
python app.py
```

или

```bash
python app.py --gui
```

## CLI пример

```bash
python app.py \
  --brand Stulz \
  --project-dir "C:/Projects/Test" \
  --calc "C:/Projects/Test/Calc.xlsx" \
  --template "C:/Projects/Test/Offer_Template.docx" \
  --client "ТОО Example"
```

## Важно

Шаблоны КП и калькуляции больше не хранятся внутри проекта. Программа работает с файлами, которые пользователь выбирает из папки проекта.
