# Патч диагностики обновлений

Скопируйте содержимое этой папки в корень проекта с заменой файлов.

## Заменяются файлы

```text
core/update_client.py
updater.py
```

## Что изменилось

Кнопка "Обновления" теперь показывает не общий текст `HTTP Error 404`, а понятную причину и подсказку, что делать.

Добавлена диагностика для типовых случаев:

1. Не найден `config/update.json`.
2. Ошибка JSON в `config/update.json`.
3. Нет прав на чтение `config/update.json`.
4. Проверка обновлений выключена через `enabled: false`.
5. Не заполнены `github_owner` / `github_repo`.
6. Некорректно заполнены `github_owner` / `github_repo`.
7. Некорректный `latest_release_api_url`.
8. Нет интернета / GitHub недоступен.
9. Таймаут подключения.
10. Ошибка SSL-сертификата.
11. GitHub latest release не найден: нет Release, есть только Actions artifacts.
12. Репозиторий private или указан неверный repo.
13. GitHub вернул 401 Unauthorized.
14. GitHub вернул 403 Forbidden / rate limit.
15. GitHub вернул 429 Too Many Requests.
16. GitHub вернул 5xx.
17. Ответ GitHub не JSON.
18. Latest release найден, но нет `tag_name`.
19. Latest release найден, но нет assets.
20. App-модуль не найден среди файлов Release; теперь показывается список доступных файлов.
21. Asset найден, но нет `browser_download_url`.
22. Нет прав на создание папки `updates`.
23. Ошибка скачивания app-модуля.
24. Не найден `updater.exe`.
25. Windows не разрешил запустить updater.
26. ZIP обновления повреждён.
27. Файл приложения занят и не заменяется.

## Важно

Этот патч не меняет логику обновления. Он только делает ошибки понятнее.

После применения патча нужно пересобрать приложение и скачать/заменить `SAM-Offer-Generator-app-no-runtime-module`.
`_internal/` обновлять не нужно.
