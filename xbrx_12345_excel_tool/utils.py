import os
import datetime
import subprocess


def make_config_dirs():
    #home_dir = os.path.expanduser("~")
    home_dir = "D:\\"
    config_dir = os.path.join(home_dir, ".xbrx_12345_excel_tool")
    exported_templates_dir =  os.path.join(config_dir, "导出记录")
    validation_logs_dir =  os.path.join(config_dir, "复核记录")

    os.makedirs(config_dir, exist_ok=True)
    os.makedirs(exported_templates_dir, exist_ok=True)
    os.makedirs(validation_logs_dir, exist_ok=True)

    return config_dir, exported_templates_dir, validation_logs_dir


def make_today_dir():
    today = datetime.date.today()
    dirname = f"{today.year}-{today.month}-{today.day}"
    exported_templates_dir = make_config_dirs()[1]
    today_dir = os.path.join(exported_templates_dir, dirname)
    os.makedirs(today_dir, exist_ok=True)
    return today_dir


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


def dt_to_date_str(dt):
    FORMAT = "%Y%m%d"
    return dt.strftime(FORMAT)


def dt_to_datetime_str(dt):
    FORMAT = "%Y%m%d%H%M%S"
    return dt.strftime(FORMAT)


def get_now_report_str():
    now = datetime.datetime.now()
    FORMAT = "%Y-%m-%d %H:%M:%S"
    return now.strftime(FORMAT)


def open_folder(path, filename=None):
    """
    if filename is not None, file is selected after folder is displayed.
    """
    if filename is not None:
        r'explorer.exe /select, "D:\games\emu\arcade\mame0184\mame64.exe"'
        subprocess.Popen(r'explorer.exe /select, "{filename}"'.format(filename=filename))
    else:
        subprocess.Popen(f"explorer.exe {path}")


def load_config():
    """
    docstring
    """
    pass