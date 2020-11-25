import openpyxl
import datetime
import os
import time


def str_to_dt(dt_str):
    """
        convert time string to datetime object.

        parameters:
            dt_str (str): a date & time string in the format %Y/%m/%d %H:%M:%S
        return:
            dt (datetime.datetime): a datetime.datetime object
    """
    FORMAT = "%Y/%m/%d %H:%M:%S"
    dt = datetime.datetime.strptime(dt_str, FORMAT)
    return dt


def find_rows_by_time(filename, start_datetime, end_datetime):
    START_ROW = 3
    COL_RECEIVE_TIME = 1

    work_book = openpyxl.load_workbook(filename, read_only=True)
    sheet = work_book.active

    rows = []
    for row in sheet.iter_rows(min_row=START_ROW, values_only=True):
        receive_time_dt = str_to_dt(row[COL_RECEIVE_TIME].strip())
        if receive_time_dt >= start_datetime and receive_time_dt <= end_datetime:
            rows.append(row)

    work_book.close()
    return rows


def write_rows_to_output_template(filename, rows):
    """
    docstring
    """
    if not filename.endswith(".xlsx"):
        return False

    COL_DIFFERENCE = 6
    TEMPLATE_FILENAME = "rx_12345_import_template.xlsx"
    TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", TEMPLATE_FILENAME)
    
    output_template = openpyxl.load_workbook(TEMPLATE_PATH)
    sheet = output_template["Sheet1"]
    
    row_count = len(rows)
    for row_index in range(0, row_count):
        col_index = 1 + COL_DIFFERENCE
        for cell_value in rows[row_index]:
            sheet.cell(row=row_index+2, column=col_index).value = cell_value
            col_index += 1

    output_template.save(filename)
    return True