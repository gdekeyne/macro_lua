import sys

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

import pandas as pd
import numpy as np


class Macro:
    def __init__(self, lua, autocad):
        self.lua = pd.read_excel(lua, skiprows=1)
        self.autocad = pd.read_excel(autocad)
        self.results = openpyxl.load_workbook(autocad)
        self.sheet = self.results.worksheets[0]
        self.arrange_autocad()
        self.arrange_lua()
        self.n_col = len(self.autocad.columns)
        self.arrange_sheet()
        self.index_offset = 3   # Sheets starts counts at 1, 1st line ignored, 2nd line is column names
        self.green = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        self.yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.red = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

    def arrange_sheet(self):
        for i in range(6):
            self.sheet.insert_cols(self.n_col)

        for n, key in enumerate(['index', 'dutch', 'french', 'signal', 'polarity']):
            list(self.sheet.columns)[n + self.n_col][0].value = key

        # Inserting columns at the end of the file inserts it at the second to last place
        # Hence the configuration is at the last column and needs to be moved back
        for cell_config, cell_last_column in zip(list(self.sheet.columns)[self.n_col-1], list(self.sheet.columns)[-1]):
            cell_config.value = cell_last_column.value
            cell_last_column.value = ''

    def arrange_autocad(self):
        for key in self.autocad.columns:
            self.autocad.rename(columns={key: key.lower().replace(' ', '_')}, inplace=True)

        if 'standard_configuration_id' not in self.autocad.columns:
            self.autocad['standard_configuration_id'] = [None] * len(self.autocad)

    def arrange_lua(self):
        for key in self.lua.columns:
            self.lua.rename(columns={key: key.lower().replace(' ', '_')}, inplace=True)

        self.lua['full_name'] = [io_type[-1] + str(io_name) if str(io_name).replace('.', '').isdigit() else io_name
                                 for io_name, io_type in zip(self.lua['i/o_name'], self.lua['i/o_type'])]

    def get_indexes(self):
        for row, (eq_name, io_name, conf_id) in enumerate(zip(self.autocad['eq_name'],
                                                              self.autocad['i/o_name'],
                                                              self.autocad['standard_configuration_id'])):
            index_eq = self.get_index_eq(eq_name)
            index_io = self.get_index_io(io_name)
            index = np.intersect1d(index_io, index_eq)

            if conf_id is not None:
                index_conf = self.get_index_conf(conf_id)
                index = np.intersect1d(index, index_conf)
                if len(index) != 1:
                    raise ValueError('{} match found instead of 1 for EQ name {}, I/O name {} and configuration id {}'
                                     .format(len(index), eq_name, io_name, conf_id))

            if len(index) != 1:
                raise ValueError('{} match found instead of 1 for EQ name {} and I/O name {}'
                                 .format(len(index), eq_name, io_name))

            list(self.sheet.columns)[self.n_col][row + 1].value = index[0] + self.index_offset

    def get_index_eq(self, eq_name):
        return np.where(self.lua['interface_device_name'] == eq_name)[0]

    def get_index_io(self, io_name):
        return np.where(self.lua['full_name'] == io_name)[0]

    def get_index_conf(self, conf_id):
        return np.where(self.lua['standard_configuration_id'] == conf_id)[0]

    def get_difference(self, reference_key, compare, index, cross=''):
        str_compare = str(compare).replace('\P', '')
        # If there is no cross in the reference table
        if not isinstance(cross, str):
            return '', None

        # If the reference and the autocad output are identical
        reference = str(self.lua[reference_key][index])
        if reference == str_compare:
            return 'ok', self.green

        # If there is a difference between the tables
        if str_compare == 'nan':
            return 'unfilled', self.yellow

        return reference, self.red

    def fill_cell(self, reference_key, compare, index, column, row, cross=''):
        content, color = self.get_difference(reference_key, compare, index, cross)
        list(self.sheet.columns)[column][row].value = content
        if color is not None:
            list(self.sheet.columns)[column][row].fill = color

    def compare(self):
        index_list = list(self.sheet.columns)[self.n_col]
        for row, (cell, dutch, french, signal, polarity) in \
            enumerate(zip(index_list[1:],
                          self.autocad['i/o_descr_nl'],
                          self.autocad['i/o_descr_fr'],
                          self.autocad['i/o_conn_02_function'],
                          self.autocad['i/o_conn_01_function'])):
            index = int(cell.value) - self.index_offset
            cross = self.get_cross(index)
            self.fill_cell('infolist/princ.diagr._name_nl', dutch, index, self.n_col + 1, row + 1)
            self.fill_cell('infolist/princ.diagr._name_fr', french, index, self.n_col + 2, row + 1)
            self.fill_cell('signal_code', signal, index, self.n_col + 3, row + 1, cross)
            self.fill_cell('polarity', polarity, index, self.n_col + 4, row + 1, cross)

    def get_cross(self, index):
        return self.lua['ix1-2'][index]

    def adjust_columns(self):
        worksheet = self.results.active
        for n, col in enumerate(worksheet.columns):
            column = get_column_letter(col[0].column)  # Get the column name
            worksheet.column_dimensions[column].width = 25

    def save_result(self, output_path):
        try:
            self.results.save(output_path)
        except FileNotFoundError:
            raise ValueError('Path to results file {} not found'.format(output_path))


if __name__ == '__main__':
    path_lua, path_autocad = sys.argv[1], sys.argv[2]
    if len(sys.argv) > 3:
        output = sys.argv[3]
    else:
        output = path_autocad

    macro = Macro(path_lua, path_autocad)
    macro.get_indexes()
    macro.compare()
    macro.adjust_columns()
    macro.save_result(output)
