pip install -r requirements.txt
pyinstaller --onefile --icon=static\YAd_icon.ico lua_autocad_comparator.py
mv dist\* .
rm *.spec
rmdir dist
rmdir __pycache__ /s
rmdir build /s