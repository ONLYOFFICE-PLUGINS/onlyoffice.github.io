effects plugin
==============
Плагин для применения различных эффектов к изображению.

СТРУКТУРА
------------

      effects-core/                 код плагина
          effects-core/cpp          C++ код
          effects-core/js           JavaScript код
          effects-core/test         все необходимое для теста
          effects-core/deploy       файлы сборки
      
      effects-build/                скрипты для сборки

УСТАНОВКА
---------
Для установки плагина просто склонируйте репозиторий:
      
      git clone https://github.com/SanDiegoss/plugin.git
      
СБОРКА
------
Для сборки нам потребуется python (скачать и установить можно [тут](https://www.python.org/)), а также два репозитория, которые следует склонировать
рядом с текущим:
      
      # скрипт для сборки
      git clone https://github.com/ONLYOFFICE/core.git
      
      # необходимые инструменты для скрипта
      git clone https://github.com/ONLYOFFICE/build_tools.git

Открываем командную строку (терминал в Linux) и переходим в каталог со скриптом:

      cd <папка с репозиториями>/core/Common/js/
      
Для компиляции и сборки используется скрипт make.py, вводим команду:
      
      python make.py <папка с репозиториями>/plugin/effects-build/effects.json
      
Скрипт автоматически скачает все необходимые инструменты и компоненты, а также скомпилирует и соберет плагин.

ЗАПУСК
------
Для запуска и теста плагина нам потребуется веб-сервер, например локальный http.server в python3:

      python3 -m http.server 8080
      
После запуска в браузере переходим на http://localhost:8080/. Готово!

