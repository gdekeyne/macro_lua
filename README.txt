LUA AUTOCAD COMPARATOR

Created by Gautier Dekeyne from Extia
Contact: gdekeyne@extia.be

A small graphical interface allowing the user to compare the output of an autocad Excel sheet
with the LUA Excel sheet.

Installation:
To install, download archive from:
```
https://github.com/gdekeyne/macro_lua/archive/master.zip
```

Download python from:
```
https://www.python.org/downloads/
```
Before following the installation, tick the box "add Python to path".

In a windows prompt, go to the folder where you unzipped the archive and enter the following lines:
```
pip install -r requirements.txt
pyinstaller --onefile --icon=static\YAd_icon.ico lua_autocad_comparator.py
mv dist\* .
rm *.spec
rmdir dist
rmdir __pycache__ /s
rmdir build /s
```

Then, simply run the application in a command prompt or by a double-click on the output-file.