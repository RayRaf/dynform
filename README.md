# dynform
 
## Как запусить и редактировать (Windows):

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
1. Убедиться, что установлен [Python](https://www.python.org/). Например командой проверки версии: python -V (Прописная V). Если не установлен, то установить;
2. Убедиться, что установлен [Pyinstaller](https://pyinstaller.org/en/stable/index.html). Например командой проверки версии: pyinstaller -v (Строчная v). Если не установлен, то установить командой pip install -U pyinstaller;
3. Выполнить сборку exe файла командой pyinstaller --onefile --noconsole main.py. --noconsole - опция отключения отображения консоли, --onefile - опция сборки в один единый exe файл.


## Как запусить и редактировать (Ubuntu):
1. Проверить, установлен ли Python. Например командой python3 -V (Прописная V).  Если отобразится версия - установлена. Скорее всего установлена, т.к. в Ubuntu идет предустановленной;
2. Проверить, установлен ли xlrd: <br>
  2.1. Зайти в интерактивный режим python командой python3;<br>
  2.2. Выполнить команду help('modules') - отобразится список всех модулей;<br>
  2.3. Если не установлен - установить командой pip install xlrd;<br>
3. Проверить, установлен ли tkinter. См. п.2 для проверки. Для установки выполнить команду sudo apt install python3-tk;<br>
4. Перейти в папку с файлами проекта в терминале. В проводнике выбрать пункт "Open in Terminal" или из любого места указав путь командой CD ~путь~;<br>
5. Выполнить команду python3 main.py.



Использована [эта](https://xlrd.readthedocs.io/en/latest/ "Go to Web Site XLRD") читалка файлов Excel. Поддерживается только формат .xls. Есть [Другие](https://pythonist.ru/kak-chitat-excel-fajl-xlsx-v-python/ "На сайт") читалки.
