pip install -r requirements.txt
pyinstaller --onefile --windowed --icon=static\YAd_icon.ico lua_autocad_comparator.py
move dist\* .
del *.spec
rmdir dist
rmdir __pycache__ /s /q
rmdir build /s /q
