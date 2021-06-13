from os import listdir, path
from pathlib import Path
import pandas as pd

from reader import Reader, ContractSorter
from const import (
    HIDDEN_COLUMNS,
    SHARED_DEPARTMENTS,
    REGIONAL_DEPARTMENTS,
    MISSING_FOLDER,
    RESULT_FILE_NAME,
    MONTH_PROMPT,
    REGISTRY_HEADER,
    MONTH_HEADER,
    TABLE_HEADERS,
    TABLE_TOTALS,
    TABLE_FOOTER,
    DATA_START_ROW,
    FINISH_PROMPT,
    WAIT_MESSAGE,
    ERROR_MESSAGE,
    SHEET_NAME,
    LAST_TABLE_COLUMN,
    TABLE_START_ROW,
    ADDRESS_COLUMN,
    SECTION_TEMPLATE,
    HEADER_TO_MERGE,
    COLUMN_WIDTHS,
    HIDDEN_COLUMNS,
    FOOTER_TO_MERGE,
)


def find_unique_filename(month):

    counter = 0
    last_attempted_name = RESULT_FILE_NAME.format(
        month=month, number='', format='xlsx'
    )
    while path.exists(last_attempted_name):
        counter += 1
        last_attempted_name = RESULT_FILE_NAME.format(
            month=month, number=counter, format='xlsx'
        )
    return last_attempted_name


cwd = Path.cwd()

month = input(MONTH_PROMPT)
dept_names = [*SHARED_DEPARTMENTS.values(), *REGIONAL_DEPARTMENTS.values()]

files_to_process = []
add_file = files_to_process.append

for folder in dept_names:

    try:
        for filename in listdir(folder):
            if filename.endswith('.xls') or filename.endswith('.xlsx'):
                add_file(cwd / folder / filename)
    except FileNotFoundError:
        print(MISSING_FOLDER.format(folder=folder))

print(WAIT_MESSAGE)

contracts = []
add_contracts = contracts.extend

for filepath in files_to_process:

    try:
        reader = Reader(pd.read_excel(filepath, header=None).iterrows(), month)
        add_contracts(reader.get_contracts())
    except Exception:
        print(ERROR_MESSAGE.format(filepath=filepath, row=reader.current_row))

data = [[registry_part] for registry_part in REGISTRY_HEADER]
data.append([MONTH_HEADER.format(month=month)])
data.extend([[], []])
data.append(TABLE_HEADERS)
data.append(range(1, len(TABLE_HEADERS) + 1))
data.extend(ContractSorter(contracts).get_dataframe_data())
data.append([])
data.append([
    cell.format(start=DATA_START_ROW, table_size=len(data)
) for cell in TABLE_TOTALS])
data.extend([[], []])
data.extend(TABLE_FOOTER)
last_row = len(data) - len(TABLE_FOOTER) - 2

dataframe = pd.DataFrame(data)

filename = find_unique_filename(month)

writer = pd.ExcelWriter(filename, engine='xlsxwriter')

# region TABLE FORMATTING
workbook  = writer.book
default_format = workbook.formats[0]
default_format.set_font_size(10)
default_format.set_align('center')
default_format.set_align('vcenter')
default_format.set_text_wrap()
worksheet = workbook.add_worksheet(SHEET_NAME)
writer.sheets[SHEET_NAME] = worksheet


def simple_format(**params):

    return {
        'type': 'no_errors',
        'format': workbook.add_format(params),
    }


def text_format(criteria, value, **params):

    return {
        'type': 'text',
        'criteria': criteria,
        'value': value,
        'format': workbook.add_format(params),
    }

table_range = f'A{TABLE_START_ROW}:{LAST_TABLE_COLUMN}{last_row}'
worksheet.conditional_format(table_range, simple_format(border=True))

header_range = f'A1:{LAST_TABLE_COLUMN}{TABLE_START_ROW+1}'
worksheet.conditional_format(header_range, simple_format(bold=True))

dept_range = f'{ADDRESS_COLUMN}{TABLE_START_ROW}:{ADDRESS_COLUMN}{last_row}'

for keyword in [*dept_names, SECTION_TEMPLATE[:2]]:
    worksheet.conditional_format(dept_range, text_format(
        'begins with', keyword, bold=True
    ))

for _range in HEADER_TO_MERGE:

    worksheet.merge_range(_range, "")

for row, _range in enumerate(FOOTER_TO_MERGE, last_row+3):

    worksheet.merge_range(_range.format(row=row), "")

for i, width in enumerate(COLUMN_WIDTHS):

    if i in HIDDEN_COLUMNS:
        worksheet.set_column_pixels(i, i, width+1, None, {'hidden': True})
    else:
        worksheet.set_column_pixels(i, i, width+1)

# endregion TABLE FORMATTING

dataframe.to_excel(
    writer,
    header=None,
    index=None,
    sheet_name=SHEET_NAME
)
writer.save()

input(FINISH_PROMPT)
