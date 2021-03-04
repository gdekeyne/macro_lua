
"""
Entry point of the program
"""

from os.path import join
from warnings import filterwarnings
from macro import Macro
from interface import window
from PySimpleGUI import WIN_CLOSED

# Ignore warnings
filterwarnings("ignore")


while True:
    event, values = window.read()
    if (event == "Exit") or (event == WIN_CLOSED):
        break
    # Trigger the comparison only if all the fields are filled
    if (event == "COMPARE") and all([values[key] for key in ['LUA', 'AUTOCAD', 'BOX', 'PATH', 'NAME']]):
        window['LOGBOX'].reroute_stderr_to_here()
        window['LOGBOX'].reroute_stdout_to_here()
        file_name = values['NAME']
        if '.' not in file_name:
            file_name += '.xlsx'
        output = join(values['PATH'], file_name)
        try:
            macro = Macro(values['LUA'], values['AUTOCAD'], values['BOX'])
            macro.get_indexes()
            macro.compare()
            macro.adjust_columns()
            macro.save_result(output)
            print('File successfully generated at: {}\n'.format(output))
        except Exception as e:
            print(e, '\n')

window.close()
