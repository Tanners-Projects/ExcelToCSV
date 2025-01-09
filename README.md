# ExcelToCSV
Convert Excel files into CSV files of a set amount of lines.

To run via command line, only ExcelToCSV2.py and BEO.jpg are needed. The rest are used when building the .exe from the .py file.

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


To Do:

Convert python to C++

Add error checking
- Incorrect file name/type
- Incorrect input/output jpath
- No input
- Missing images, non-critical dependencies (as of now, only image)
- Edge cases such as too few/many names

Include Windows File Explorer for input/output selction

Build .exe such that Windows Defender/firewall does not block by default
