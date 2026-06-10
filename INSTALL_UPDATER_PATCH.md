# Updater patch

Патч добавляет обновление через отдельный `updater.exe`.

## Что заменить

```text
.github/workflows/main.yml
gui/main_window.py
sam_offer_generator.spec
```

## Что добавить

```text
core/update_client.py
updater.py
config/update.json
UPDATE_SYSTEM.md
FILE_LIST.txt
```

## Настройка GitHub repository

После копирования патча откройте:

```text
config/update.json
```

И замените:

```json
"github_owner": "OWNER",
"github_repo": "REPO"
```

на реальные значения вашего репозитория.

Например:

```json
"github_owner": "Vitaliy-Litvinov",
"github_repo": "offer-generator"
```

## Как работает

1. В программе появится кнопка `Обновления`.
2. Программа проверит latest GitHub Release.
3. Если версия новее, скачает `SAM-Offer-Generator-App-No-Runtime.zip`.
4. Запустит `updater.exe`.
5. Основная программа закроется.
6. `updater.exe` заменит файлы в текущей portable-папке.
7. Программа запустится снова.

Папка `_internal/` при App-обновлении не трогается.

## Без UAC

Updater не просит права администратора. Но программа должна лежать в папке, куда текущий пользователь может писать, например:

```text
C:\Users\<User>\Apps\SAM-Offer-Generator\
```

или:

```text
%LOCALAPPDATA%\SAM-Offer-Generator\
```

Не ставьте portable-папку в `C:\Program Files`, иначе обновление может потребовать права администратора.
