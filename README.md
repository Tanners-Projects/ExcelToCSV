# ExcelToCSV
Convert Excel files into CSV files of a set amount of lines.


Dependencies:

- dearpygui: pip3 install dearpygui, pip install dearpygui
- math: default
- pandas: pip3 install pandas
- numpy: pip3 install numpy (might not need anymore, need to check names per file then remove)
- os: default
- Arial font: File found in C:\Windows\Fonts
- BEO image: Currently needs to be in same directory, need to see what
             to do when turning file into exe

- Directory: File needs to be in same directory as program

To build:
pyinstaller --onefile --noconsole ExcelToCSV2.py
pyinstaller ExcelToCSV2.spec
