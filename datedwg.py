"""
This script is developed to check, whether drawings of a Solidworks
project are up-to-date with their parent models.

Supposed use: put all the parts` and assemblies` names inside Excel
spreadsheet column. Provide the script with information about the
spreadsheet, as well with information about a directory inside your
file system, where parts, asemblies and drawings are.

Output: next to each cell in the conclusion column there will be a conclusion:
OK - if the drawing was modified later, than the model;
OUTDATED - if the opposite;
ERROR - if the script wasn`t able to check.

WARNING: Close the workbook before running the script, or it
won`t save changes.
WARNING: It is assumed that each part/assembly has unique name, distinct
from all th other files in a directory of search.

Usage example:
python datedwg.py \
    --workbook /mnt/d/registries/projectAA53.xlsx \
    --sheet Sheet1 \
    --target-column A \
    --result-column B \
    --directory /mnt/d/projectAA53
"""

import os
import time
import sys
import argparse
import openpyxl

# Setting up an eye-catching style for cells which show some kind of warning.
WARNING_COLOUR = openpyxl.styles.colors.Color(rgb='00FF0000')
WARNING_FILL = openpyxl.styles.fills.PatternFill(
    patternType='solid',
    fgColor=WARNING_COLOUR,
)

parser = argparse.ArgumentParser()
parser.add_argument(
    '--workbook',
    help='Absolute path to an Excel spreadsheet containig parts`/assemblies` names.',
)
parser.add_argument(
    '--sheet',
    help='Name of a sheet inside the spreadsheet containig parts`/assemblies` names.',
)
parser.add_argument(
    '--target-column',
    help='Column containig parts`/assemblies` names.',
)
parser.add_argument(
    '--result-column',
    help='Column, into which relults will be put.',
)
parser.add_argument(
    '--directory',
    help='Directory where parts/assemblies are.',
)


def find_path_by_decimal(
    targeted_path: str,
    decimal: str,
    file_format: str,
) -> str:
    """
    The function searches for a file`s path based on it`s name (decimal) and format.

    Input:
    path - directory, where to search (string);
    decimal - name of the file without .format (string);
    file_format - the file`s format without dot in front (string);

    Output:
    res - path of the sought-fir file (string);

    Usage example:
    model_path = find_path_by_decimal(
        pathj='D:\parts_test\\',
        decimal='АТ.301243.432',
        file_format='SLDASM',
    )
    # model_path -> 'D:\parts_test\assmbl\АТ.301243.432.SLDASM'

    WARNING: It is assumed that each part/assembly has unique name, distinct
    from all th other files in a directory of search.
    """

    res = ''
    for root, _, files in os.walk(targeted_path, topdown=False):

        for file in files:
            # In case if the file has lowercase format.
            if file.startswith(f'{decimal}.{file_format.lower()}'):
                res = os.path.join(root, f'{decimal}.{file_format.lower()}')
                # As long as the file has unique name, it is safe to break
                # the loop once a single instance is met for the purpose of swiftness.
                break
            # In case if the file has uppercase format.
            if file.startswith(f'{decimal}.{file_format.upper()}'):
                res = os.path.join(root, f'{decimal}.{file_format.upper()}')
                # Same.
                break
    return res


def write_conclusion_to_worksheet(
    worksheet: openpyxl.Workbook,
    decimal: str, conclusion: str,
    warning: bool = False,
) -> None:
    """
    The writes a specified conclusion next to a cell containing
    specified file name (decimal).

    Input:
    worksheet - the worksheet, containing the column fith files` names ();
    decimal - name of the file without .format (string);
    conclusion - what to write next to the cell (string);
    warning - flag (False by default) which indicates whether the conclusion
    should be marked as outstanding (boolean);

    Output: None;

    Usecase example:
    write_conclusion_to_worksheet(worksheet, decimal, conclusion='ERROR', warning=True)
    """

    for cell in worksheet[TARGETED_COLUMN]:
        if cell.value == decimal:
            worksheet[f'{CONCLUSION_COLUMN}{cell.row}'] = conclusion
            if warning:
                worksheet[f'{CONCLUSION_COLUMN}{cell.row}'].fill = WARNING_FILL


def main() -> None:
    """
    Main method.
    """
    start_time = time.time()

    # Connecting to the workbook and the worksheet inside it.
    workbook = openpyxl.load_workbook(WORKBOOK_PATH)
    worksheet = workbook[WORKSHEET_NAME]

    # Assembling everything in the targeted column in a list.
    decimals = [worksheet[TARGETED_COLUMN][num].value
                for num in range(len(worksheet[TARGETED_COLUMN]))]

    print(f'\nInitial decimals: {decimals}')

    # Filter for non-string values.
    decimals = [decimal for decimal in decimals if isinstance(decimal, str)]

    # Custom filter for values that are incoherent
    # with projets`s notation convention.
    decimals = [decimal for decimal in decimals
                if (decimal[0] or decimal[1])
                in ['A', 'T', 'А', 'Т', 'А', 'Т']
                ]           # list includes cyrilic russian and ukrainian characters

    print(f'Filtered decimals: {decimals}')

    for decimal in decimals:
        try:
            # Finding a path for a drawing file.
            dwg_path = find_path_by_decimal(
                targeted_path=DIRECTORY_OF_SEARCH,
                decimal=decimal, file_format='slddrw')

            # Finding a path for a assembly file.
            if decimal[3] == '3':
                model_path = find_path_by_decimal(
                    targeted_path=DIRECTORY_OF_SEARCH,
                    decimal=decimal, file_format='sldasm')

            # Finding a path for a part file.
            else:
                model_path = find_path_by_decimal(
                    targeted_path=DIRECTORY_OF_SEARCH,
                    decimal=decimal, file_format='sldprt')

            # Getting a 'last modified' time for model and drawing.
            lastmodified_dwg = os.path.getmtime(dwg_path)
            lastmodified_model = os.path.getmtime(model_path)

            # Comparing and writing a conclustin to the Excel spreadsheet.
            if lastmodified_model > lastmodified_dwg:
                print(f'{decimal} found to be not up to date!')
                write_conclusion_to_worksheet(
                    worksheet, decimal, conclusion='OUTDATED', warning=True)

            else:
                print(f'{decimal} is up to date.')
                write_conclusion_to_worksheet(
                    worksheet, decimal, conclusion='OK', warning=False)

        except Exception:
            print(f'Oops! {sys.exc_info()[0]} occurred.')
            write_conclusion_to_worksheet(
                worksheet, decimal, conclusion='ERROR', warning=True)

    workbook.save(WORKBOOK_PATH)

    end_time = time.time()
    seconds = round((end_time - start_time), 2)
    minutes = round((seconds/60), 2)
    print(f'Estimated time: {seconds} seconds (about {minutes} minutes).\n')


if __name__ == '__main__':
    args = parser.parse_args()

    WORKBOOK_PATH = args.workbook
    WORKSHEET_NAME = args.sheet
    TARGETED_COLUMN = args.target_column
    CONCLUSION_COLUMN = args.result_column
    DIRECTORY_OF_SEARCH = args.directory

    main()
