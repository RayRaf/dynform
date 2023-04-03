# dynform

Как запусить и редактировать:

Способ 1. Через командную строку:
1. Установить на компьютер [Python](https://www.python.org/). Не забыть поставить галочку напротив "Add to PATH;
2. Установить читалку .xls файлов xlrd командой "pip install xlrd";
3. Открыть командную строку и прейти к папке со скачанными файлами. Команда формата "cd C:\Новая папка\"
4. Выполнить команду "python main.py"

Способ 2. С помощью IDE Pycharm
1. Установить [IDE PyCharm](https://www.jetbrains.com/ru-ru/pycharm/download/#section=windows/);
2. Открыть проект в PyCharm;
3. Перед запуском пройтись по коду: везде, где выводятся красные подчеркивания щелкнуть и установить предлагаемые пакеты;
4. Запустить, нажав кнопку Run или Debug.

Как собрать исполняемый exe файл:
1. Убедиться, что установлен [Python](https://www.python.org/). Например командй проверки версии: python -V (Прописная V). Если не установлен, то установить;
2. Убедиться, что установлен [Pyinstaller](https://pyinstaller.org/en/stable/index.html). Например командой проверки версии: pyinstaller -v (Строчная v). Если не установлен, то установить командой pip install -U pyinstaller;
3. Выполнить сборку exe файла командой pyinstaller --onefile --noconsole main.py. --noconsole - опция отключения отображения консоли, --onefile - опция сборки в один единый exe файл.


Использована [эта](https://xlrd.readthedocs.io/en/latest/ "Go to Web Site XLRD") читалка файлов Excel. Поддерживается только формат .xls. Есть [Другие](https://pythonist.ru/kak-chitat-excel-fajl-xlsx-v-python/ "На сайт") читалки.
