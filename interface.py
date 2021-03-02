
"""
Definition of the Graphical User Interface
"""

import PySimpleGUI as sg

sg.theme('DarkGrey9')
lua_file = [sg.Text("LUA file\t\t"), sg.In(size=(25, 1), key="LUA"), sg.FileBrowse()]
autocad_file = [sg.Text("Autocad file\t"), sg.In(size=(25, 1), key="AUTOCAD"), sg.FileBrowse()]
box_name = [sg.Text("Box name\t"), sg.In(size=(25, 1), key="BOX", default_text='IX1-2')]
output_path = [sg.Text("Output folder\t"), sg.In(size=(25, 1), key="PATH"), sg.FolderBrowse()]
output_name = [sg.Text("Output file name\t"), sg.In(size=(25, 1), key="NAME")]
ok_button = [sg.Button("Compare", enable_events=True, key="COMPARE")]
line_split = sg.VSeperator()
left_column = [lua_file, autocad_file, box_name, output_path, output_name, ok_button]

box_title = [sg.Text("Logs:")]
log_box = [sg.Multiline(size=(35, 10), key="LOGBOX")]
right_column = [box_title, log_box]

layout = [[sg.Column(left_column), line_split, sg.Column(right_column)]]
window = sg.Window("LUA autocad comparator", layout)
