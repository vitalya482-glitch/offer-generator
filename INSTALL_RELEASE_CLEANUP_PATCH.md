# Release cleanup patch

Скопируйте содержимое архива в корень проекта с заменой файла.

## Заменяется

```text
.github/workflows/main.yml
```

## Что добавлено

После создания GitHub Release workflow запускает шаг:

```text
Cleanup old GitHub Releases and tags
```

Он оставляет только последние 3 GitHub Releases, отсортированные по дате создания, а более старые удаляет вместе с их git-тегами.

## Что важно

- `latest release` не удаляется, потому что это один из трёх последних релизов.
- Updater использует только `latest release`, поэтому старые релизы ему не нужны.
- Actions artifacts по-прежнему хранятся по `retention-days: 30`; этот патч чистит именно GitHub Releases и tags.
- Если релизов меньше или ровно 3, workflow ничего не удалит.

## Проверка

1. Сделайте commit и push в `main`.
2. Дождитесь успешного workflow.
3. Откройте GitHub -> Releases.
4. Должно остаться максимум 3 последних релиза.
