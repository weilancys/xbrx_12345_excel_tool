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
        label = ttk.Label(frame_set_time, text="1234")
        label.pack()
        frame_set_time.pack(fill=tk.X)

        # button export
        btn_export = ttk.Button(tab_split, text="导出小白热线模板")
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