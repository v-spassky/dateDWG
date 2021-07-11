"""
This script is developed to check, whether drawings of a Solidworks
project are up-to-date with their parent models.

Supposed use: put all the parts` and assemblies` names inside Excel
spreadsheet column. Provide the script with information about the
spreadsheet, as well with information about a directory inside your
file system, where parts, asemblies and drawings are.

Output: next to each cell, on the right side, there will be a conclusion:
OK - if the drawing was modified later, then the model;
OUTDATED - if the opposite;
ERROR - if the script wasn`t able to check.

WARNING: Close the workbook before running the script, or it 
won`t save changes.
WARNING: It is assumed that each part/assembly has unique name, distinct
from all th other files in a directory of search.
"""

import os
import time
import sys
import openpyxl

start_time = time.time()

# --------------------------------------------------------------------------
# Absolute path to an Excel spreadsheet containig parts`/assemblies` names.
WORKBOOK_PATH = r'Z:\pyscripts\list_test.xlsx'
# Name of a sheet inside the spreadsheet containig parts`/assemblies` names.
WORKSHEET_NAME = 'Sheet2'
# Column containig parts`/assemblies` names.
TARGETED_COLUMN = 'A'
# Directory where parts/assemblies are.
DIRECTORY_OF_SEARCH = r'C:\PDM\Архив_ОГК\\'

# Setting up an eye-catching style for cells which show some kind of warning.
warning_colour = openpyxl.styles.colors.Color(rgb='00FF0000')
warning_fill = openpyxl.styles.fills.PatternFill(
    patternType='solid', fgColor=warning_colour)
# --------------------------------------------------------------------------

workbook = openpyxl.load_workbook(WORKBOOK_PATH)
worksheet = workbook[WORKSHEET_NAME]

# Assembling everything in the targeted column in a list.
decimals = []
for cell in worksheet[TARGETED_COLUMN]:
    decimals.append(cell.value)

print(f'\nInitial decimals: {decimals}')

# Filter for non-string values.
for decimal in decimals:
    if not isinstance(decimal, str):
        decimals.remove(decimal)

# Custom filter for values that are incoherent
# with projets`s notation convention.
for decimal in decimals:
    if (decimal[0] or decimal[1]) not in ['A', 'T', 'А', 'Т', 'А', 'Т']:
        decimals.remove(decimal)

print(f'Filtered decimals: {decimals}')


def find_path_by_name(name, path):
    """
    The function searches for a file`s path based on it`s name.

    WARNING: It is assumed that each part/assembly has unique name, distinct 
    from all th other files in a directory of search.
    """

    result = []
    for root, dirs, files in os.walk(path):
        if name in files:
            result.append(os.path.join(root, name))
            break
    return result


for decimal in decimals:

    try:

        # Finding a path for both model and drawing files.
        dwg_path = find_path_by_name(f'{decimal}.slddrw', DIRECTORY_OF_SEARCH)
        if not dwg_path:
            dwg_path = find_path_by_name(
                f'{decimal}.SLDDRW', DIRECTORY_OF_SEARCH)
        if decimal[3] == '3':
            model_path = find_path_by_name(
                f'{decimal}.sldasm', DIRECTORY_OF_SEARCH)
            if not model_path:
                model_path = find_path_by_name(
                    f'{decimal}.SLDASM', DIRECTORY_OF_SEARCH)
        else:
            model_path = find_path_by_name(
                f'{decimal}.sldprt', DIRECTORY_OF_SEARCH)
            if not model_path:
                model_path = find_path_by_name(
                    f'{decimal}.SLDPRT', DIRECTORY_OF_SEARCH)

        # Getting a 'last modified' time for model and drawing.
        lastmodified_dwg = os.path.getmtime(dwg_path[0])
        lastmodified_model = os.path.getmtime(model_path[0])

        # Comparing and writing a conclustin to the Excel spreadsheet.
        if lastmodified_model > lastmodified_dwg:
            print(f'{decimal} found to be not up to date!')
            for cell in worksheet[TARGETED_COLUMN]:
                if cell.value == decimal:
                    worksheet[f'B{cell.row}'] = 'OUTDATED'
                    worksheet[f'B{cell.row}'].fill = warning_fill
        else:
            print(f'{decimal} is up to date.')
            for cell in worksheet[TARGETED_COLUMN]:
                if cell.value == decimal:
                    worksheet[f'B{cell.row}'] = 'OK'

    except Exception:

        print(f'Oops! {sys.exc_info()[0]} occurred.')
        for cell in worksheet[TARGETED_COLUMN]:
            if cell.value == decimal:
                worksheet[f'B{cell.row}'] = 'ERROR'
                worksheet[f'B{cell.row}'].fill = warning_fill

workbook.save(WORKBOOK_PATH)

end_time = time.time()
print(
    f'''Estimated time: {round((end_time - start_time), 2)} 
    seconds({round((end_time - start_time)/60, 2)} minutes.)\n''')
