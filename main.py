import openpyxl
import os
import time
import sys

start = time.time()

# --------------------------------------------------------------------------------
# Absolute path to an Excel spreadsheet containig parts`/assemblies` names.
workbook_path = 'Z:\pyscripts\list_test.xlsx'
# Name of a sheet inside the spreadsheet containig parts`/assemblies` names.
worksheet_name = 'Sheet2'
# Column containig parts`/assemblies` names.
targeted_column = 'A'
# Directory where parts/assemblies are.
directory_of_search = 'C:\PDM\Архив_ОГК\\'

# Setting up an eye-catching style for cells which show some kind of warning.
warning_colour = openpyxl.styles.colors.Color(rgb='00FF0000')
warning_fill = openpyxl.styles.fills.PatternFill(
    patternType='solid', fgColor=warning_colour)
# --------------------------------------------------------------------------------

# WARNING: Close the workbook before running the script, or it won`t save changes.
workbook = openpyxl.load_workbook(workbook_path)
worksheet = workbook[worksheet_name]

# Assembling everything in the targeted column in a list.
decimals = []
for cell in worksheet[targeted_column]:
    decimals.append(cell.value)

print(f'\nInitial decimals: {decimals}')

# Filter for non-string values.
for decimal in decimals:
    if type(decimal) != str:
        decimals.remove(decimal)

# Custom filter for values that are incoherent with projets`s notation convention.
for decimal in decimals:
    if (decimal[0] or decimal[1]) not in ['A', 'T', 'А', 'Т', 'А', 'Т']:
        decimals.remove(decimal)

print(f'Filtered decimals: {decimals}')


def find_all_paths(name, path):
    """
    The function searches for a file`s path based on it`s name 

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
        dwg_path = find_all_paths(f'{decimal}.slddrw', directory_of_search)
        if not dwg_path:
            dwg_path = find_all_paths(
                f'{decimal}.SLDDRW', directory_of_search)
        if decimal[3] == '3':
            model_path = find_all_paths(
                f'{decimal}.sldasm', directory_of_search)
            if not model_path:
                model_path = find_all_paths(
                    f'{decimal}.SLDASM', directory_of_search)
        else:
            model_path = find_all_paths(
                f'{decimal}.sldprt', directory_of_search)
            if not model_path:
                model_path = find_all_paths(
                    f'{decimal}.SLDPRT', directory_of_search)

        # Getting a 'last modified' time for model and drawing.
        lastmodified_dwg = os.path.getmtime(dwg_path[0])
        lastmodified_model = os.path.getmtime(model_path[0])

        # Comparing and writing a conclustin to the Excel spreadsheet.
        if lastmodified_model > lastmodified_dwg:
            print(f'{decimal} found to be not up to date!')
            for cell in worksheet[targeted_column]:
                if cell.value == decimal:
                    worksheet[f'B{cell.row}'] = 'OUTDATED'
                    worksheet[f'B{cell.row}'].fill = warning_fill
        else:
            print(f'{decimal} is up to date.')
            for cell in worksheet[targeted_column]:
                if cell.value == decimal:
                    worksheet[f'B{cell.row}'] = 'OK'

    except Exception:

        print(f'Oops! {sys.exc_info()[0]} occurred.')
        for cell in worksheet[targeted_column]:
            if cell.value == decimal:
                worksheet[f'B{cell.row}'] = 'ERROR'
                worksheet[f'B{cell.row}'].fill = warning_fill

workbook.save(workbook_path)

end = time.time()
print(
    f'Estimated time: {round((end - start), 2)} seconds ({round((end - start)/60, 2)} minutes.)\n')

# if __name__ == '__main__':
