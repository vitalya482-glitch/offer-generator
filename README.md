# Генератор коммерческого предложения в Word

Программа формирует КП `.docx` на базе:

- Word-шаблона КП: `templates/kp_template.docx`
- Excel-калькуляции: `samples/Calc_23-12-24 PAC.xlsx`
- PDF-файлов с техописанием: папка `pdf/`

## Что уже умеет MVP

1. Запускаться в красивом GUI на PySide6 в фирменном стиле SAM Group.
2. Читать модель, количество, цену и итог из Excel.
2. Подставлять клиента, дату, версию и сумму в Word.
3. Заполнять таблицу оборудования в КП.
4. Заполнять таблицу опций в разделе «СПЕЦИФИКАЦИЯ».
5. Если в папке `pdf/` есть файл с названием модели, например `CCU121A.pdf`, вставлять текст из PDF в конец КП.
6. Собираться в Windows `.exe` через GitHub Actions.

## Быстрый запуск на компьютере

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py --gui
```

Или без окна:

```bash
python app.py --client "ТОО Ромашка" --calc "samples/Calc_23-12-24 PAC.xlsx" --template "templates/kp_template.docx"
```

Готовый файл появится в папке `output/`.

## Как выложить на GitHub

```bash
git init
git add .
git commit -m "Initial KP generator"
git branch -M main
git remote add origin https://github.com/USERNAME/kp-generator.git
git push -u origin main
```

После push откройте вкладку **Actions**. Workflow `Build Windows EXE` соберет файл `kp-generator.exe`.

## Как работать с PDF

Положите PDF-файлы в папку `pdf/`. Название файла должно содержать модель из Excel:

```text
pdf/CCU121A.pdf
pdf/ASU211AL.pdf
```

Если модель из Excel `CCU121A`, программа найдет `CCU121A.pdf` и вставит текст из него в КП.

## Что надо доработать на следующем этапе

- точные правила чтения всех итоговых сумм из Excel;
- вставка картинок из PDF;
- экспорт готового КП в PDF;
- отдельный справочник товаров и описаний.

## Для нового чата ChatGPT

Если работа продолжается в новом чате, сначала покажите ChatGPT файл:

```text
CHATGPT_CONTEXT.md
```

Там описаны цель проекта, исходные файлы, текущее состояние, GitHub-репозиторий и следующие задачи.

Также полезен файл:

```text
FILE_NOTES.md
```

В нем есть пометки по каждому файлу проекта.


## GUI в стиле SAM Group

Интерфейс переведен с tkinter на PySide6. В окне есть:

- темная бренд-панель SAM Group;
- красная акцентная кнопка формирования КП;
- карточки с полями проекта;
- выбор Word-шаблона, Excel-расчета, папки PDF и папки результата;
- автоматический список листов Excel;
- предварительная проверка: модель, количество, версия, условия поставки, итог и количество опций.

Для запуска:

```bash
pip install -r requirements.txt
python app.py --gui
```

Для сборки Windows `.exe` GitHub Actions использует PyInstaller и добавляет в сборку папки `templates/`, `samples/` и `config.example.json`.
