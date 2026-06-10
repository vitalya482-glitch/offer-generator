# Auto Release Patch

Скопируйте содержимое архива в корень проекта с заменой файлов.

## Заменяется

```text
.github/workflows/main.yml
```

## Что делает workflow

1. При каждом `push` в `main` собирает Windows portable-сборку.
2. Создаёт автоматический GitHub Release с тегом вида:

```text
vYYYY.MM.DD.RUN_NUMBER
```

Например:

```text
v2026.06.10.67
```

3. При `push` обычного тега `v*`, например `v0.2.1`, создаёт Release именно с этим тегом.
4. Прикрепляет к Release файлы:

```text
SAM-Offer-Generator-App-No-Runtime.zip
SAM-Offer-Generator-Runtime-Win64.zip
SAM-Offer-Generator-windows-portable.zip
RELEASE_SHA256SUMS.txt
source-modules ZIP/MD/TXT
```

## Почему это нужно для updater

Updater читает:

```text
https://api.github.com/repos/OWNER/REPO/releases/latest
```

Теперь latest release будет появляться автоматически, без ручного создания Release в интерфейсе GitHub.

## Важно для config/update.json

Убедитесь, что в программе указано:

```json
{
  "github_owner": "vitalya482-glitch",
  "github_repo": "offer-generator",
  "app_asset_name": "SAM-Offer-Generator-App-No-Runtime.zip",
  "runtime_asset_name": "SAM-Offer-Generator-Runtime-Win64.zip"
}
```

Имена `app_asset_name` и `runtime_asset_name` должны совпадать с именами файлов в Release, не с названиями Actions artifacts.

## После установки патча

1. Сделайте commit и push в `main`.
2. Дождитесь успешного workflow.
3. Откройте GitHub → Releases.
4. Там должен появиться новый Release.
5. После этого кнопка `Обновления` в программе сможет найти latest release.
