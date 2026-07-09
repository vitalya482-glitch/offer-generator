Offer Generator — файлы для замены

Сначала выпустить LVKUpdater v0.3.0.

Затем заменить в репозитории offer-generator:
  core/lvk_updater_launcher.py
  gui/main_window.py
  .github/workflows/main.yml

Workflow уже переключён на LVKUpdater 0.3.0.

После push сборщик скачает LVKUpdater-win-x64.zip из Release v0.3.0 и положит LVKUpdater.exe в App-No-Runtime пакет.

Ожидаемое поведение:
- кнопка «Обновления» не закрывает Offer Generator;
- пустая консоль не показывается;
- если обновлений нет, программа остаётся открытой;
- если обновление подтверждено, updater штатно закрывает программу, устанавливает файлы и запускает её снова.
