# Установка патча LVKUpdater для Offer Generator

Распаковать архив в корень репозитория `offer-generator` с заменой файлов.

## Файлы

```text
.github/workflows/main.yml
app.update.json
core/lvk_updater_launcher.py
gui/main_window.py
INSTALL_LVK_UPDATER_PATCH.md
```

## Что изменилось

- В сборку Offer Generator добавляется `LVKUpdater.exe` из релиза `LVK-Updater v0.2.0`.
- В app-модуль добавляется `app.update.json`.
- Кнопка `Обновления` больше не использует старый Python-updater.
- GitHub Release теперь публикует третий asset: `offer-generator.json`.
- `app.update.json` смотрит на:

```text
https://github.com/vitalya482-glitch/offer-generator/releases/latest/download/offer-generator.json
```

Поэтому отдельный ручной апдейт `LVK-Update-Feed` для Offer Generator больше не нужен.

## Проверка после сборки

В архиве `SAM-Offer-Generator-App-No-Runtime.zip` должны быть:

```text
SAM-Offer-Generator/
  SAM-Offer-Generator.exe
  LVKUpdater.exe
  app.update.json
  release_info.json
  config/
  assets/
  prices/
  templates/
```

В GitHub Release должны быть:

```text
SAM-Offer-Generator-App-No-Runtime.zip
SAM-Offer-Generator-Runtime-Win64.zip
offer-generator.json
```

## Важно

При нажатии кнопки `Обновления` программа закроется, потому что внешний C++ updater должен получить возможность заменить файлы. Если обновлений нет, LVKUpdater покажет сообщение. В этом случае программу нужно открыть заново вручную.
