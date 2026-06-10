# Project analysis: SAM Offer Generator

## Что уже есть в проекте

Проект уже разделен на логические части:

```text
app.py                  # CLI/GUI входная точка
core/                   # общая бизнес-логика
brands/                 # логика направлений/брендов
gui/                    # PySide6 интерфейс
config/                 # JSON-справочники и настройки
mcp/sam_assistant/      # отдельный MCP-сервер
prices/                 # референсные прайсы/PDF
.github/workflows/      # сборка GitHub Actions
```

Основная рабочая линия сейчас — `Stulz`. Остальные бренды (`Riello`, `DC Eltek`, `Generator`) подключены через общий registry, но их генерация пока оставлена как заготовка с `NotImplementedError`.

## Что было проблемой для релиза

В исходном `sam_offer_generator.spec` сборка была похожа на монолитный one-file EXE: `a.binaries` и `a.datas` передавались прямо в `EXE(...)`. Для твоего требования это неудобно, потому что пользователь получает один большой исполняемый файл вместо переносимой папки.

Также workflow загружал только один общий artifact `dist/`, без отдельных архивов модулей.

## Что изменено

1. `sam_offer_generator.spec` переведен на PyInstaller `onedir`:

```text
EXE(..., exclude_binaries=True)
COLLECT(exe, a.binaries, a.zipfiles, a.datas, name='SAM-Offer-Generator')
```

Результат сборки:

```text
dist/SAM-Offer-Generator/
  SAM-Offer-Generator.exe
  _internal/
```

2. Добавлен `tools/prepare_portable_release.py`.

Он дополняет PyInstaller-папку:

```text
SAM-Offer-Generator/
  config/                  # редактируемые JSON рядом с EXE
  modules/source/          # копии исходных модулей
  prices/                  # прайсы, если есть
  README_RELEASE.txt
  run_gui.cmd
  release_info.json
```

3. Добавлен `MODULES_MANIFEST.json`.

Он описывает отдельные скачиваемые модули:

- `entrypoint`
- `core`
- `brands`
- `gui`
- `config`
- `mcp-sam-assistant`
- `prices`
- `github-build`

4. Добавлен `tools/package_modules.py`.

Скрипт создает отдельные модульные ZIP:

```text
offer-generator-core.zip
offer-generator-brands.zip
offer-generator-gui.zip
offer-generator-config.zip
...
```

5. Обновлен `.github/workflows/main.yml`.

Теперь GitHub Actions:

- проверяет синтаксис Python;
- собирает one-dir Windows app;
- готовит portable-папку;
- собирает отдельные ZIP модулей;
- публикует artifacts по отдельности;
- при push тега `v*` создает GitHub Release с portable ZIP и модульными ZIP.

6. Добавлен `core/runtime_paths.py`.

Он решает важную проблему PyInstaller: редактируемые JSON должны жить рядом с EXE в `config/`, а не внутри `_internal` или временной папки PyInstaller.

## Проверено локально

В текущей среде проверены:

```bash
python -m compileall app.py brands core gui config tools
python tools/package_modules.py --output /mnt/data/source-modules-final
```

Windows PyInstaller-сборку в этой среде я не запускал, потому что полноценная целевая сборка должна выполняться на `windows-latest` в GitHub Actions. Workflow для этого подготовлен.

## Какой вариант использовать

Для пользователя Windows:

```text
SAM-Offer-Generator-windows-portable.zip
```

Для разработки или частичного скачивания:

```text
offer-generator-core.zip
offer-generator-brands.zip
offer-generator-gui.zip
offer-generator-config.zip
```

Для GitHub-релиза:

```bash
git tag v0.2.0
git push origin v0.2.0
```
