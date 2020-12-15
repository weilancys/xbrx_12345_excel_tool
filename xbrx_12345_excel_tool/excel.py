import openpyxl
import datetime
import os
import time
from .utils import str_to_dt, get_now_report_str, dt_to_datetime_str, make_config_dirs


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


def get_all_ids_from_12345_banlidan_total_file(filename):
    ID_COL_INDEX = 0
    work_book = openpyxl.load_workbook(filename, read_only=True)
    sheet = work_book["Sheet0"]

    ids = []
    rows = sheet.iter_rows(min_row=3, values_only=True)
    for row in rows:
        if row[ID_COL_INDEX] is None or row[ID_COL_INDEX] == "":
            raise ValueError("empty id")
        ids.append(int(row[ID_COL_INDEX])) # ids should be ints

    work_book.close()
    return ids


def get_all_ids_from_xbrx_export_total_file(filename):
    ID_COL_INDEX = 1
    work_book = openpyxl.load_workbook(filename) # can't use read_only here, don't know why yet
    sheet = work_book.active

    ids = []
    rows = sheet.iter_rows(min_row=3, values_only=True)
    for row in rows:
        if row[ID_COL_INDEX] is None or row[ID_COL_INDEX] == "":
            continue # some rows have this column value as None
        ids.append(int(row[ID_COL_INDEX])) # ids should be ints

    work_book.close()

    ids_dict = {}
    for _id in ids:
        if _id not in ids_dict:
            ids_dict[_id] = 1
        else:
            ids_dict[_id] += 1

    return ids_dict


def generate_validation_report(banlidan_total_filename, xbrx_export_total_filename):
    banlidan_total_ids = get_all_ids_from_12345_banlidan_total_file(banlidan_total_filename)
    xbrx_export_total_ids = get_all_ids_from_xbrx_export_total_file(xbrx_export_total_filename)

    missing_ids = []
    validation_conclusion = None
    for banlidan_id in banlidan_total_ids:
        if banlidan_id not in xbrx_export_total_ids:
            missing_ids.append(banlidan_id)

    if missing_ids == []:
        validation_conclusion = "未发现漏单，12345办理单汇总表内全部工单号已在热线系统内登记。"
    else:
        validation_conclusion = "发现漏单!\n以下工单为漏单，请进一步核实:\n"
        for missing_id in missing_ids:
            validation_conclusion += str(missing_id) + "\n"

    attentions = ""
    for xbrx_export_total_id in xbrx_export_total_ids:
        if xbrx_export_total_ids[xbrx_export_total_id] > 1:
            attentions += str(xbrx_export_total_id) + "\n"

    to_be_used_1 = "起止时间：{banlidan_total_start_time} - {banlidan_total_end_time}"
    to_be_used_2 = "起止时间：{xbrx_export_total_start_time} - {xbrx_export_total_end_time}"


    validation_report_text = """
    复核时间：{validation_datetime}

    12345办理单汇总表(政府表)信息：
    文件名: {banlidan_total_filename}
    办理单条数：{banlidan_total_count}


    小白热线系统汇总表(三高表)信息：
    文件名: {xbrx_export_total_filename}
    办理单条数: {xbrx_export_total_count}


    复核结论：
    {validation_conclusion}
    
    
    存在重办或退回的工单：
    {attentions}
    """.format(
        validation_datetime = get_now_report_str(),
        banlidan_total_filename = banlidan_total_filename,
        banlidan_total_count = len(banlidan_total_ids),
        banlidan_total_start_time = 1,
        banlidan_total_end_time = 1,

        xbrx_export_total_filename = xbrx_export_total_filename,
        xbrx_export_total_count = len(xbrx_export_total_ids),
        xbrx_export_total_start_time = 1,
        xbrx_export_total_end_time = 1,

        validation_conclusion = validation_conclusion,
        attentions = attentions
    )

    return validation_report_text


def save_validation_report_file(validation_report_text):
    # todo return None if failed here.
    now = datetime.datetime.now()
    validation_logs_dir = make_config_dirs()[2]
    filename = f"复核记录_{dt_to_datetime_str(now)}.txt"
    file_path = os.path.join(validation_logs_dir, filename)
    with open(file_path, "w") as f:
        f.write(validation_report_text)
    return validation_logs_dir