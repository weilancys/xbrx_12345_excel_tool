import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
from tkinter import ttk


class xbrx_12345_excel_tool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.__init_ui()

    def __init_ui(self):
        self.title("小白热线 12345 EXCEL 工具")
        self.geometry("800x600")

        # tabs
        tab_parent = ttk.Notebook(self)
        tab_split = ttk.Frame(tab_parent)
        tab_validation = ttk.Frame(tab_parent)
        tab_logs = ttk.Frame(tab_parent)

        tab_parent.add(tab_split, text="导出")
        tab_parent.add(tab_validation, text="复核")
        tab_parent.add(tab_logs, text="日志")
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
        frame_set_time = ttk.LabelFrame(tab_split, text="设置起始时间：")

        # start time input widgets
        label_start = ttk.Label(frame_set_time, text="起：")
        label_start.grid(row=0, column=0)

        cb_start_year = ttk.Combobox(frame_set_time, width=5)
        cb_start_year.grid(row=0, column=1)
        label_start_year = ttk.Label(frame_set_time, text="年")
        label_start_year.grid(row=0, column=2)

        cb_start_month = ttk.Combobox(frame_set_time, width=3)
        cb_start_month.grid(row=0, column=3)
        label_start_month = ttk.Label(frame_set_time, text="月")
        label_start_month.grid(row=0, column=4)

        cb_start_day = ttk.Combobox(frame_set_time, width=3)
        cb_start_day.grid(row=0, column=5)
        label_start_day = ttk.Label(frame_set_time, text="日")
        label_start_day.grid(row=0, column=6)

        cb_start_hour = ttk.Combobox(frame_set_time, width=3)
        cb_start_hour.grid(row=0, column=7)
        label_start_hour = ttk.Label(frame_set_time, text="时")
        label_start_hour.grid(row=0, column=8)

        cb_start_minute = ttk.Combobox(frame_set_time, width=3)
        cb_start_minute.grid(row=0, column=9)
        label_start_minute = ttk.Label(frame_set_time, text="分")
        label_start_minute.grid(row=0, column=10)

        cb_start_second = ttk.Combobox(frame_set_time, width=3)
        cb_start_second.grid(row=0, column=11)
        label_start_second = ttk.Label(frame_set_time, text="秒")
        label_start_second.grid(row=0, column=12)


        # end time input widgets
        label_end = ttk.Label(frame_set_time, text="止：")
        label_end.grid(row=1, column=0)

        cb_end_year = ttk.Combobox(frame_set_time, width=5)
        cb_end_year.grid(row=1, column=1)
        label_end_year = ttk.Label(frame_set_time, text="年")
        label_end_year.grid(row=1, column=2)

        cb_end_month = ttk.Combobox(frame_set_time, width=3)
        cb_end_month.grid(row=1, column=3)
        label_end_month = ttk.Label(frame_set_time, text="月")
        label_end_month.grid(row=1, column=4)

        cb_end_day = ttk.Combobox(frame_set_time, width=3)
        cb_end_day.grid(row=1, column=5)
        label_end_day = ttk.Label(frame_set_time, text="日")
        label_end_day.grid(row=1, column=6)

        cb_end_hour = ttk.Combobox(frame_set_time, width=3)
        cb_end_hour.grid(row=1, column=7)
        label_end_hour = ttk.Label(frame_set_time, text="时")
        label_end_hour.grid(row=1, column=8)

        cb_end_minute = ttk.Combobox(frame_set_time, width=3)
        cb_end_minute.grid(row=1, column=9)
        label_end_minute = ttk.Label(frame_set_time, text="分")
        label_end_minute.grid(row=1, column=10)

        cb_end_second = ttk.Combobox(frame_set_time, width=3)
        cb_end_second.grid(row=1, column=11)
        label_end_second = ttk.Label(frame_set_time, text="秒")
        label_end_second.grid(row=1, column=12)

        frame_set_time.pack(fill=tk.X)

        # button export
        btn_export = ttk.Button(tab_split, text="导出热线系统模板")
        btn_export.pack()

    
    def on_btn_choose_12345_source_workbook_click(self):
        filename = tk.filedialog.askopenfilename()
        if filename == "":
            return
        # tk.messagebox.showinfo("open", filename)
        # self.entry_12345_source_workbook_path.insert(0, filename)
        self.str_12345_source_workbook.set(filename)

    
    def run(self):
        self.mainloop()


if __name__ == "__main__":
    tool = xbrx_12345_excel_tool()
    tool.run()