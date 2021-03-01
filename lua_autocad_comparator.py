
"""
Entry point of the program
"""

import os
from macro import Macro
from interface import window, sg


while True:
    event, values = window.read()
    if (event == "Exit") or (event == sg.WIN_CLOSED):
        break

    # Trigger the comparison only if all the fields are filled
    if (event == "COMPARE") and all([values[key] for key in ['LUA', 'AUTOCAD', 'BOX', 'PATH', 'NAME']]):
        output = os.path.join(values['PATH'], values['NAME'])
        try:
            macro = Macro(values['LUA'], values['AUTOCAD'], values['BOX'])
            macro.get_indexes()
            macro.compare()
            macro.adjust_columns()
            macro.save_result(output)
            print('File successfully generated at: {}'.format(output))
        except Exception as e:
            print(e)

window.close()
