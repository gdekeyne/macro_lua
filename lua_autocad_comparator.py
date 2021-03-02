
"""
Entry point of the program
"""

import os
import warnings
from contextlib import redirect_stderr, redirect_stdout
from macro import Macro
from interface import window, sg
from io import StringIO

# Ignore warnings
warnings.filterwarnings("ignore")

out = StringIO()
err = StringIO()

while True:
    event, values = window.read()
    if (event == "Exit") or (event == sg.WIN_CLOSED):
        break
    # Trigger the comparison only if all the fields are filled
    if (event == "COMPARE") and all([values[key] for key in ['LUA', 'AUTOCAD', 'BOX', 'PATH', 'NAME']]):
        file_name = values['NAME']
        if '.' not in file_name:
            file_name += '.xlsx'
        output = os.path.join(values['PATH'], file_name)
        try:
            with redirect_stderr(err):
                macro = Macro(values['LUA'], values['AUTOCAD'], values['BOX'])
                macro.get_indexes()
                macro.compare()
                macro.adjust_columns()
                macro.save_result(output)
            with redirect_stdout(out):
                print('File successfully generated at: {}'.format(output))
        except Exception as e:
            with redirect_stdout(out):
                print(e)

        window['LOGBOX'].print(err.getvalue(), out.getvalue())

window.close()
