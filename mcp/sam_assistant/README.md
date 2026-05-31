# MCP Server for Offer Generator

## Подключение к Continue MCP

1. Установите зависимости:
   ```bash
   pip install -r requirements.txt
   ```

2. Запустите сервер:
   ```bash
   python server.py
   ```

3. Подключитесь к серверу через Continue MCP.

## Используемые инструменты MCP:
- `scan_project(project_path: str)` - показывает дерево проекта, исключая:
  `.git`, `.venv`, `build`, `dist`, `__pycache__`
- `search_code(project_path: str, query: str)` - ищет текст по файлам:
  `.py`, `.md`, `.json`, `.txt`, `.yaml`, `.yml`, `.spec`
- `read_project_file(path: str)` - читает текстовый файл проекта.
- `list_new_files(project_path: str)` - показывает `git status --short`, чтобы видеть новые незакоммиченные файлы.
- `find_specification_logic(project_path: str)` - ищет в проекте слова:
  `specification`, `specifications`, `spec`, `спецификац`, `spec_preview`, `spec_block`
- `read_pdf_text(path: str)` - извлекает текст из PDF через PyMuPDF.
- `analyze_price_pdf(path: str)` - извлекает текст из PDF и возвращает:
  - количество страниц;
  - первые 3000 символов;
  - найденные слова `price`, `model`, `EUR`, `USD`, `cooling`, `capacity`, `kW`;
  - краткое предположение, является ли файл прайс-листом.