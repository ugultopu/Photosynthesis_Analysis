__author__ = "Utku Gultopu"
__copyright__ = """

    Copyright (c) 2017 Utku Gultopu

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

import os
import re
import csv
import openpyxl
import configparser

config = configparser.ConfigParser()
config.read('config.ini')
config = config['DEFAULT']

wb = openpyxl.load_workbook(config['InputSpreadsheetPath'])
sheet = wb.get_sheet_by_name(config['SpreadsheetPageName'])

def get_variable_indices(label_row):
    labels_of_interest = ['Y(II)', 'Y(NPQ)', 'Y(NO)', 'qP', 'ETR', 'NPQ']
    all_indices = {}
    for label in labels_of_interest:
        all_indices[label] = [i for i, j in enumerate(label_row) if label in j]
    return all_indices

def get_averages(path):
    with open(path, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=';')
        rows = []

        all_indices = get_variable_indices(next(reader))

        first_row = next(reader)
        rows.append(first_row)

        # BEGIN: Get last row of reader
        count = 2
        for row in reader:
            count += 1
            pass
        rows.append(row)
        # END: Get last row of reader

        avg = {}
        for label in all_indices:
            avg[label] = []
            indices = all_indices[label]
            for row in rows:
                values = [float(row[i]) for i in indices]
                avg[label].append(sum(values) / len(values))
        return avg

def get_row_and_column_displacement(path):
    width = 11
    height = 22
    row = 0
    column = 0
    if '15.Day' in path:
        column += width
    if '30.Day' in path:
        column += 2 * width
    if '2.Leaf' in path:
        row += 4
    if '3.Leaf' in path:
        row += 2 * 4
    if '1100' in path:
        row += 1
    p = re.compile('(\d)\.{}'.format(config['ConcentrationText']))
    m = p.search(path)
    if m is not None:
        row += int(m.group(1)) * height
    return row, column

base_path = config['BasePath']
base_row = 6    # 0 index
base_column = 1 # 0 index
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.endswith('.csv'):
            row_disp, column_disp = get_row_and_column_displacement(os.path.relpath(root, base_path))
            avg = get_averages(os.path.join(root, file))
            cur_base_row = base_row + row_disp
            cur_base_column = base_column + column_disp
            for label in avg:
                avgs = avg[label]
                if label == 'Y(II)':
                    sheet[cur_base_row][cur_base_column].value = round(avgs[0], 3)
                    sheet[cur_base_row][cur_base_column + 1].value = round(avgs[1], 3)
                if label == 'Y(NPQ)':
                    sheet[cur_base_row][cur_base_column + 2].value = round(avgs[1], 3)
                if label == 'Y(NO)':
                    sheet[cur_base_row][cur_base_column + 3].value = round(avgs[1], 3)
                if label == 'NPQ':
                    sheet[cur_base_row][cur_base_column + 4].value = round(avgs[1], 3)
                if label == 'qP':
                    sheet[cur_base_row][cur_base_column + 6].value = round(avgs[1], 3)
                if label == 'ETR':
                    sheet[cur_base_row][cur_base_column + 7].value = round(avgs[1], 3)
wb.save(os.path.join(config['BasePath'], config['OutputSpreadsheetName']))

