# Контекст проекта для нового чата ChatGPT

Этот файл нужен, чтобы в новом чате быстро объяснить, что уже сделано и что нужно дальше.

## Цель проекта

Сделать Windows-программу **Offer Generator** для автоматического формирования коммерческого предложения в Word.

Программа должна:

1. Брать расчет и цены из Excel-файла.
2. Брать техническое описание из PDF-файлов.
3. Вставлять данные в Word-шаблон коммерческого предложения.
4. Формировать готовый `.docx`.
5. Позже собираться в `.exe` через GitHub Actions.

## Пользовательские исходные файлы

### Word-шаблон КП

Файл:

```text
Offer_Company_20-01-21(v1) PAC.docx
```

В проекте он лежит как:

```text
templates/kp_template.docx
```

В нем есть:

- шапка КП;
- клиент `ТОО «[Организация]»`;
- таблица оборудования;
- условия оплаты, поставки, монтажа и ПНР;
- раздел `СПЕЦИФИКАЦИЯ`;
- технические характеристики.

### Excel-шаблон расчета

Файл:

```text
Calc_23-12-24 PAC.xlsx
```

В проекте он лежит как:

```text
samples/Calc_23-12-24 PAC.xlsx
```

Это не просто прайс, а инженерный калькулятор. Из него надо брать:

- модель оборудования;
- количество;
- опции;
- цены;
- итоговые суммы;
- EXW/DDP или условия поставки.

## Текущее состояние проекта

Создан стартовый проект `kp-generator-project.zip`.

В проекте есть:

```text
app.py
requirements.txt
README.md
CHATGPT_CONTEXT.md
FILE_NOTES.md
config.example.json
templates/kp_template.docx
samples/Calc_23-12-24 PAC.xlsx
pdf/
output/
.github/workflows/build-windows.yml
```

## Как запускать локально

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py --gui
```

Без окна:

```bash
python app.py --client "ТОО Ромашка" --calc "samples/Calc_23-12-24 PAC.xlsx" --template "templates/kp_template.docx"
```

Готовый файл должен появиться в папке:

```text
output/
```

## GitHub

Пользователь решил создать отдельный репозиторий:

```text
offer-generator
```

Описание репозитория:

```text
Windows application for generating commercial offers from Excel, PDF and Word templates
```

При создании репозитория рекомендовано:

- `Public` или `Private` — по желанию;
- `Add README` — OFF;
- `.gitignore` — No .gitignore;
- `License` — No license.

Потом нужно загрузить файлы проекта в GitHub и открыть вкладку `Actions`.

## Важно для следующего чата

Если пользователь продолжит в новом чате, попросить его загрузить:

1. этот архив проекта `kp-generator-project-annotated.zip` или весь репозиторий;
2. Word-шаблон КП;
3. Excel-калькулятор;
4. при наличии — PDF-файлы техописаний.

После этого нужно продолжить с проверки:

1. правильно ли программа читает Excel;
2. какие ячейки отвечают за модель, количество и цену;
3. как лучше вставлять данные в Word;
4. как добавить PDF-описания;
5. как собрать `.exe` через GitHub Actions.

## Ближайшие задачи

1. Загрузить проект на GitHub.
2. Проверить, что GitHub Actions видит workflow.
3. Исправить возможные ошибки сборки `.exe`.
4. Сделать точное чтение Excel по реальным ячейкам.
5. Заменить статические места в Word-шаблоне на переменные.
6. Добавить PDF-техописания.
7. Добавить нормальный Windows-интерфейс.
