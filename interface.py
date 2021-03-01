
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
input_list = [lua_file, autocad_file, box_name, output_path, output_name, ok_button]

layout = [[sg.Column(input_list)]]
window = sg.Window("LUA autocad comparator", layout)
