# Runtime module patch

Этот патч ничего не меняет в бизнес-логике проекта.

Заменяется только файл GitHub Actions:

```text
.github/workflows/main.yml
```

После установки патча GitHub Actions дополнительно создаст два новых артефакта:

```text
SAM-Offer-Generator-runtime-win64-module
  содержит SAM-Offer-Generator-Runtime-Win64.zip
  внутри архива: SAM-Offer-Generator/_internal/

SAM-Offer-Generator-app-no-runtime-module
  содержит SAM-Offer-Generator-App-No-Runtime.zip
  внутри архива: SAM-Offer-Generator/ без папки _internal/
```

Полная portable-сборка `SAM-Offer-Generator-windows-portable-zip` остаётся как была.
Все source-модули `core`, `gui`, `brands`, `config`, `prices` остаются как были.

## Как использовать модульную установку

1. Скачать `SAM-Offer-Generator-App-No-Runtime.zip`.
2. Скачать `SAM-Offer-Generator-Runtime-Win64.zip`.
3. Распаковать оба архива в одно место с объединением папок.
4. Итоговая структура должна быть такой:

```text
SAM-Offer-Generator/
  SAM-Offer-Generator.exe
  run_gui.cmd
  config/
  prices/
  _internal/
```

При обычном обновлении программы можно обновлять только `SAM-Offer-Generator-App-No-Runtime.zip`, если зависимости Python/PySide6/Qt не менялись.
