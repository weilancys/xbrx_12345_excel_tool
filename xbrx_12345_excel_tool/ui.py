import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
from tkinter import ttk
import datetime
import os
from .excel import find_rows_by_time, write_rows_to_output_template, generate_validation_report, save_validation_report_file
from .utils import make_config_dirs, make_today_dir, dt_to_date_str, dt_to_datetime_str, open_folder


class xbrx_12345_excel_tool(tk.Tk):
    def __init__(self):
        super().__init__()
        make_config_dirs()
        self.__init_ui()

    def __init_ui(self):
        self.title("小白热线 12345 EXCEL 工具")
        self.geometry("490x200")

        # date and time ranges
        YEAR_RANGE = [year for year in range(2015, 2031)]
        MONTH_RANGE = [month for month in range(1, 13)]
        DAY_RANGE = [day for day in range(1, 32)]
        HOUR_RANGE = [hour for hour in range(0, 25)]
        MINUTE_RANGE = [minute for minute in range(0, 61)]
        SECOND_RANGE = [second for second in range(0, 61)]

        # tabs
        tab_parent = ttk.Notebook(self)
        tab_split = ttk.Frame(tab_parent)
        tab_validation = ttk.Frame(tab_parent)
        # tab_logs = ttk.Frame(tab_parent)

        tab_parent.add(tab_split, text="导出")
        tab_parent.add(tab_validation, text="复核")
        # tab_parent.add(tab_logs, text="日志")
        tab_parent.pack(expand=1, fill=tk.BOTH)


        # controls
        # frame 12345 source workbook
        frame_12345_source_workbook = ttk.LabelFrame(tab_split, text="12345源表：")
        btn_choose_12345_source_workbook = ttk.Button(frame_12345_source_workbook, text="选择12345源表", command=self.on_btn_choose_12345_source_workbook_click)
        btn_choose_12345_source_workbook.grid(row=0, column=0)
        self.str_12345_source_workbook = tk.StringVar()
        self.entry_12345_source_workbook_path = ttk.Entry(frame_12345_source_workbook, state="readonly", textvariable=self.str_12345_source_workbook)
        self.entry_12345_source_workbook_path.grid(row=0, column=1)
        frame_12345_source_workbook.pack(fill=tk.X)

        # frame set time
        frame_set_time = ttk.LabelFrame(tab_split, text="设置起止接收时间：")

        # start time input widgets
        now = datetime.datetime.now()

        label_start = ttk.Label(frame_set_time, text="起：")
        label_start.grid(row=0, column=0)

        self.cb_start_year = ttk.Combobox(frame_set_time, width=5, values=YEAR_RANGE)
        self.cb_start_year.set(now.year)
        self.cb_start_year.grid(row=0, column=1)
        label_start_year = ttk.Label(frame_set_time, text="年")
        label_start_year.grid(row=0, column=2)

        self.cb_start_month = ttk.Combobox(frame_set_time, width=3, values=MONTH_RANGE)
        self.cb_start_month.set(now.month)
        self.cb_start_month.grid(row=0, column=3)
        label_start_month = ttk.Label(frame_set_time, text="月")
        label_start_month.grid(row=0, column=4)

        self.cb_start_day = ttk.Combobox(frame_set_time, width=3, values=DAY_RANGE)
        self.cb_start_day.set(now.day)
        self.cb_start_day.grid(row=0, column=5)
        label_start_day = ttk.Label(frame_set_time, text="日")
        label_start_day.grid(row=0, column=6)

        self.cb_start_hour = ttk.Combobox(frame_set_time, width=3, values=HOUR_RANGE)
        self.cb_start_hour.set(now.hour)
        self.cb_start_hour.grid(row=0, column=7)
        label_start_hour = ttk.Label(frame_set_time, text="时")
        label_start_hour.grid(row=0, column=8)

        self.cb_start_minute = ttk.Combobox(frame_set_time, width=3, values=MINUTE_RANGE)
        self.cb_start_minute.set(now.minute)
        self.cb_start_minute.grid(row=0, column=9)
        label_start_minute = ttk.Label(frame_set_time, text="分")
        label_start_minute.grid(row=0, column=10)

        self.cb_start_second = ttk.Combobox(frame_set_time, width=3, values=SECOND_RANGE)
        self.cb_start_second.set(now.second)
        self.cb_start_second.grid(row=0, column=11)
        label_start_second = ttk.Label(frame_set_time, text="秒")
        label_start_second.grid(row=0, column=12)


        # end time input widgets
        label_end = ttk.Label(frame_set_time, text="止：")
        label_end.grid(row=1, column=0)

        self.cb_end_year = ttk.Combobox(frame_set_time, width=5, values=YEAR_RANGE)
        self.cb_end_year.set(now.year)
        self.cb_end_year.grid(row=1, column=1)
        label_end_year = ttk.Label(frame_set_time, text="年")
        label_end_year.grid(row=1, column=2)

        self.cb_end_month = ttk.Combobox(frame_set_time, width=3, values=MONTH_RANGE)
        self.cb_end_month.set(now.month)
        self.cb_end_month.grid(row=1, column=3)
        label_end_month = ttk.Label(frame_set_time, text="月")
        label_end_month.grid(row=1, column=4)

        self.cb_end_day = ttk.Combobox(frame_set_time, width=3, values=DAY_RANGE)
        self.cb_end_day.set(now.day)
        self.cb_end_day.grid(row=1, column=5)
        label_end_day = ttk.Label(frame_set_time, text="日")
        label_end_day.grid(row=1, column=6)

        self.cb_end_hour = ttk.Combobox(frame_set_time, width=3, values=HOUR_RANGE)
        self.cb_end_hour.set(now.hour)
        self.cb_end_hour.grid(row=1, column=7)
        label_end_hour = ttk.Label(frame_set_time, text="时")
        label_end_hour.grid(row=1, column=8)

        self.cb_end_minute = ttk.Combobox(frame_set_time, width=3, values=MINUTE_RANGE)
        self.cb_end_minute.set(now.minute)
        self.cb_end_minute.grid(row=1, column=9)
        label_end_minute = ttk.Label(frame_set_time, text="分")
        label_end_minute.grid(row=1, column=10)

        self.cb_end_second = ttk.Combobox(frame_set_time, width=3, values=SECOND_RANGE)
        self.cb_end_second.set(now.second)
        self.cb_end_second.grid(row=1, column=11)
        label_end_second = ttk.Label(frame_set_time, text="秒")
        label_end_second.grid(row=1, column=12)

        frame_set_time.pack(fill=tk.X)

        # frame export actions
        frame_export_actions = ttk.Frame(tab_split)
        btn_export = ttk.Button(frame_export_actions, text="导出热线系统模板", command=self.on_btn_export_click)
        btn_export.grid(row=0, column=0)
        btn_open_export_folder = ttk.Button(frame_export_actions, text="打开导出模板文件夹", command=self.on_btn_open_export_folder_click)
        btn_open_export_folder.grid(row=0, column=1)
        frame_export_actions.pack()


        # 12345 banlidan total frame
        frame_12345_banlidan_total = ttk.LabelFrame(tab_validation, text="12345办理单汇总表(政府表):")
        btn_choose_12345_banlidan_total = ttk.Button(frame_12345_banlidan_total, text="选择政府表", command=self.on_btn_choose_12345_banlidan_total_click)
        btn_choose_12345_banlidan_total.grid(row=0, column=0)
        self.str_12345_banlidan_total_path = tk.StringVar()
        self.entry_12345_banlidan_total_path = ttk.Entry(frame_12345_banlidan_total, state="readonly", textvariable=self.str_12345_banlidan_total_path)
        self.entry_12345_banlidan_total_path.grid(row=0, column=1)
        frame_12345_banlidan_total.pack(fill=tk.X)


        # xbrx system export total frame
        frame_xbrx_export_total = ttk.LabelFrame(tab_validation, text="小白热线系统汇总表(三高表):")
        btn_choose_xbrx_export_total = ttk.Button(frame_xbrx_export_total, text="选择三高表", command=self.on_btn_choose_xbrx_export_total_click)
        btn_choose_xbrx_export_total.grid(row=0, column=0)
        self.str_xbrx_export_total_path = tk.StringVar()
        self.entry_xbrx_export_total_path = ttk.Entry(frame_xbrx_export_total, state="readonly", textvariable=self.str_xbrx_export_total_path)
        self.entry_xbrx_export_total_path.grid(row=0, column=1)
        frame_xbrx_export_total.pack(fill=tk.X)


        # frame validation actions
        frame_validation_actions = ttk.Frame(tab_validation)
        btn_valid = ttk.Button(frame_validation_actions, text="生成复核报告", command=self.on_btn_valid_click)
        btn_valid.grid(row=0, column=0)
        btn_open_valid_folder = ttk.Button(frame_validation_actions, text="打开复核报告文件夹", command=self.on_btn_open_valid_folder_click)
        btn_open_valid_folder.grid(row=0, column=1)
        frame_validation_actions.pack()

    
    def on_btn_choose_12345_source_workbook_click(self):
        filename = tk.filedialog.askopenfilename()
        if filename == "":
            return
        if filename.endswith(".xls") or filename.endswith(".xlsx"):    
            self.str_12345_source_workbook.set(filename)
        else:
            tk.messagebox.showerror("错误", "仅支持excel文件")


    def on_btn_choose_12345_banlidan_total_click(self):
        filename = tk.filedialog.askopenfilename()
        if filename == "":
            return
        if filename.endswith(".xls") or filename.endswith(".xlsx"):    
            self.str_12345_banlidan_total_path.set(filename)
        else:
            tk.messagebox.showerror("错误", "仅支持excel文件")

    
    def on_btn_choose_xbrx_export_total_click(self):
        filename = tk.filedialog.askopenfilename()
        if filename == "":
            return
        if filename.endswith(".xls") or filename.endswith(".xlsx"):    
            self.str_xbrx_export_total_path.set(filename)
        else:
            tk.messagebox.showerror("错误", "仅支持excel文件")


    def on_btn_open_export_folder_click(self):
        exported_templates_dir = make_config_dirs()[1]
        open_folder(exported_templates_dir)

    
    def on_btn_export_click(self):
        source_excel_filename = self.str_12345_source_workbook.get()
        if source_excel_filename.strip() == "":
            tk.messagebox.showerror("错误", "请先选择Excel文件")
            return

        start_time = datetime.datetime(
            int(self.cb_start_year.get()),
            int(self.cb_start_month.get()),
            int(self.cb_start_day.get()),
            int(self.cb_start_hour.get()),
            int(self.cb_start_minute.get()),
            int(self.cb_start_second.get())
        )

        end_time = datetime.datetime(
            int(self.cb_end_year.get()),
            int(self.cb_end_month.get()),
            int(self.cb_end_day.get()),
            int(self.cb_end_hour.get()),
            int(self.cb_end_minute.get()),
            int(self.cb_end_second.get())
        )
        
        rows = find_rows_by_time(source_excel_filename, start_time, end_time)

        # make today dir -> save file in today dir
        # filename example: 
        # 导出模板_xx条_20201210_20201210041211_20201210061530.xlsx
        today_dir = make_today_dir()
        today = datetime.date.today()
        save_filename = f"导出模板_{len(rows)}条_{dt_to_date_str(today)}_{dt_to_datetime_str(start_time)}_{dt_to_datetime_str(end_time)}.xlsx"
        save_path = os.path.join(today_dir, save_filename)

        if write_rows_to_output_template(save_path, rows):
            if tk.messagebox.askyesno("成功", "保存成功！是否要查看导出模板？"):
                open_folder(today_dir, save_path)
        else:
            tk.messagebox.showerror("错误", "保存失败")

    
    def on_btn_valid_click(self):
        banlidan_total_filename = self.str_12345_banlidan_total_path.get()
        xbrx_export_total_filename = self.str_xbrx_export_total_path.get()

        if banlidan_total_filename == "" or xbrx_export_total_filename == "":
            tk.messagebox.showerror("错误", "请先选择Excel文件")
            return

        validation_report_text = generate_validation_report(banlidan_total_filename, xbrx_export_total_filename)
        validation_logs_dir = save_validation_report_file(validation_report_text)
        if validation_logs_dir is not None:
            if tk.messagebox.askyesno("成功", "复核报告已生成，是否要查看？"):
                open_folder(validation_logs_dir)
        else:
            tk.messagebox.showerror("错误", "保存失败")

    
    def on_btn_open_valid_folder_click(self):
        validation_logs_dir = make_config_dirs()[2]
        open_folder(validation_logs_dir)


    def run(self):
        self.mainloop()


class row_select_dialog(tk.Toplevel):
    """
    things that might make me happy now:
        1. coding
        2. partner
        3. racing & trucking
        4. reading
    """
    pass

# to be deleted
if __name__ == "__main__":
    app = xbrx_12345_excel_tool()
    app.run()