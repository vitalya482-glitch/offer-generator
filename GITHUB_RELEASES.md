# GitHub release layout

Проект теперь собирается и публикуется не как один монолитный EXE, а как переносимая папка и набор отдельных модульных архивов.

## Что появится в GitHub Actions

После push в `main` или ручного запуска workflow `Build Windows portable release` во вкладке **Actions** будут доступны артефакты:

| Артефакт | Что внутри | Когда скачивать |
|---|---|---|
| `SAM-Offer-Generator-windows-portable-folder` | Готовая папка `SAM-Offer-Generator/` | Когда нужна программа папкой, без ZIP |
| `SAM-Offer-Generator-windows-portable-zip` | `SAM-Offer-Generator-windows-portable.zip` | Основной вариант для пользователя Windows |
| `offer-generator-entrypoint-module` | `app.py`, README, requirements и базовые файлы | Когда нужно обновить входную точку или документацию |
| `offer-generator-core-module` | Папка `core/` | Когда нужно обновить общую бизнес-логику |
| `offer-generator-brands-module` | Папка `brands/` | Когда нужно обновить логику брендов |
| `offer-generator-gui-module` | Папка `gui/` | Когда нужно обновить интерфейс |
| `offer-generator-config-module` | Папка `config/` и `config.example.json` | Когда нужно обновить справочники и настройки |
| `offer-generator-mcp-sam-assistant-module` | Папка `mcp/sam_assistant/` | Когда нужен только MCP-ассистент |
| `offer-generator-prices-module` | Папка `prices/` | Когда нужны только прайсы/референсы |
| `offer-generator-github-build-module` | `.github/`, spec, scripts, tools | Когда нужно обновить сборку GitHub Actions |
| `offer-generator-source-modules-all` | Все модульные ZIP внутри одного архива | Когда нужны все исходные модули, но не Windows-сборка |

## Что будет в GitHub Release

Если создать тег версии, например `v0.2.0`, workflow автоматически создаст Release и прикрепит файлы:

```text
SAM-Offer-Generator-windows-portable.zip
MODULES_INDEX.md
SHA256SUMS.txt
offer-generator-entrypoint.zip
offer-generator-core.zip
offer-generator-brands.zip
offer-generator-gui.zip
offer-generator-config.zip
offer-generator-mcp-sam-assistant.zip
offer-generator-prices.zip
offer-generator-github-build.zip
offer-generator-source-modules.zip
```

## Как сделать релиз

```bash
git add .
git commit -m "Build portable modular release"
git push origin main

git tag v0.2.0
git push origin v0.2.0
```

После этого открыть GitHub -> **Actions** или GitHub -> **Releases**.

## Структура переносимой Windows-папки

После сборки основной ZIP содержит папку:

```text
SAM-Offer-Generator/
  SAM-Offer-Generator.exe
  run_gui.cmd
  README_RELEASE.txt
  release_info.json
  config/
    managers.json
    signers.json
    stulz_options.json
    stulz_winplan.json
  modules/
    source/
      brands/
      core/
      gui/
      config/
  prices/
  _internal/
```

Важно: нельзя переносить отдельно только `SAM-Offer-Generator.exe`. Это `onedir`-сборка, поэтому рядом должны оставаться `_internal/`, `config/` и остальные файлы.

## Локальная сборка на Windows

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build-windows-portable.ps1
```

Отдельно собрать только модульные исходные архивы:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build-source-modules.ps1
```

## Как это устроено

- `sam_offer_generator.spec` переведен на PyInstaller `onedir` через `COLLECT(...)`.
- `tools/prepare_portable_release.py` добавляет в папку релиза редактируемый `config/`, README, manifest и копии исходных модулей.
- `tools/package_modules.py` читает `MODULES_MANIFEST.json` и создает отдельные ZIP по модулям.
- `core/runtime_paths.py` делает так, чтобы в собранной программе редактируемые JSON-настройки жили рядом с EXE, а не внутри внутренней PyInstaller-папки.
