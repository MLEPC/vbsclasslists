# functions for create-child-info.py
from copy import copy

def format_info(df, writer, sheet_name):
    #print(len(df))
    # get the xlsxwriter workbook and worksheet objects for formatting
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # add a header format
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'align': 'left',
    'border': 1})
    # write the column headers with the defined format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    # add column formating
    col_format = workbook.add_format({
    'text_wrap': True,
    'text_wrap': True,
    'align': 'left',
    'valign': 'top',
    'border': 1})
    worksheet.set_column('A:A', 12, col_format)
    worksheet.set_column('B:B', 12, col_format)
    worksheet.set_column('C:C', 12, col_format)
    worksheet.set_column('D:D', 20, col_format)
    worksheet.set_column('E:E', 20, col_format)
    worksheet.set_column('F:F', 20, col_format)
    worksheet.set_column('G:G', 12, col_format)
    
    return writer
    
def format_str(df, column_list):
    for c in copy(column_list):
        df[c] = df[c].str.capitalize()
        df[c] = df[c].str.strip()
    return df