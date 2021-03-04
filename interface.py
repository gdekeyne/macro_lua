
"""
Definition of the Graphical User Interface
"""

from PySimpleGUI import theme, Text, FileBrowse, FolderBrowse, Column, Window, Multiline, VSeperator, Button, In

theme('DarkGrey9')
lua_file = [Text("LUA file\t\t"), In(size=(25, 1), key="LUA"), FileBrowse()]
autocad_file = [Text("Autocad file\t"), In(size=(25, 1), key="AUTOCAD"), FileBrowse()]
box_name = [Text("Box name\t"), In(size=(25, 1), key="BOX", default_text='IX1-2')]
output_path = [Text("Output folder\t"), In(size=(25, 1), key="PATH"), FolderBrowse()]
output_name = [Text("Output file name\t"), In(size=(25, 1), key="NAME")]
ok_button = [Button("Compare", enable_events=True, key="COMPARE")]
line_split = VSeperator()
left_column = [lua_file, autocad_file, box_name, output_path, output_name, ok_button]

box_title = [Text("Logs:")]
log_box = [Multiline(size=(35, 10), key="LOGBOX")]
right_column = [box_title, log_box]

layout = [[Column(left_column), line_split, Column(right_column)]]
window = Window("LUA autocad comparator", layout)
