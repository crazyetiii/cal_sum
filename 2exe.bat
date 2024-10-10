pyinstaller --onefile --noconsole --icon=logo.ico --add-data "logo.ico;." --add-data "BCompare;BCompare" compare.py

copy .\dist\compare.exe .

