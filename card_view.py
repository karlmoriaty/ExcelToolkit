import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from openpyxl.utils import get_column_letter
import os
from spreadsheet_editor import SpreadsheetEditor


class CardView(tk.Canvas):
    def __init__(self, parent, editor, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.editor = editor
        self.cards = []
        self.scroll_y = 0
        
        self.configure(bg="#f5f5f5")
        self.bind("<MouseWheel>", self.on_mousewheel)
        
        self.draw_cards()

    def draw_cards(self):
        self.delete("all")
        self.cards = []
        
        sheet_names = self.editor.get_sheet_names()
        y_offset = 20 - self.scroll_y
        
        for sheet_name in sheet_names:
            self.editor.select_sheet(sheet_name)
            
            card_width = 700
            card_height = self.calculate_card_height(sheet_name)
            
            x = 50
            y = y_offset
            
            self.create_round_rect(x, y, x + card_width, y + card_height, 
                                 radius=15, fill="white", outline="#e0e0e0")
            
            self.create_text(x + 30, y + 30, text=sheet_name, 
                           font=("微软雅黑", 24, "bold"), anchor="nw", fill="#333333")
            
            used_range = self.editor.get_used_range()
            if used_range:
                min_row, max_row = used_range["min_row"], used_range["max_row"]
                min_col, max_col = used_range["min_col"], used_range["max_col"]
                
                col_widths = self.calculate_column_widths(min_row, max_row, min_col, max_col)
                
                table_y = y + 80
                row_height = 35
                header_bg = "#f0f7ff"
                
                for row in range(min_row, max_row + 1):
                    table_x = x + 30
                    bg = header_bg if row == min_row else "white"
                    
                    for col in range(min_col, max_col + 1):
                        col_idx = col - min_col
                        width = col_widths[col_idx]
                        
                        self.create_rectangle(table_x, table_y, table_x + width, table_y + row_height,
                                             fill=bg, outline="#e8e8e8")
                        
                        value = self.editor.get_cell_value(row, col) or ""
                        self.create_text(table_x + 10, table_y + row_height / 2,
                                        text=str(value), font=("微软雅黑", 11 if row == min_row else 10),
                                        anchor="w", fill="#333333")
                        
                        table_x += width
                    
                    table_y += row_height
            
            self.cards.append((x, y, x + card_width, y + card_height))
            y_offset += card_height + 30

    def calculate_card_height(self, sheet_name):
        base_height = 100
        
        used_range = self.editor.get_used_range()
        if used_range:
            rows = used_range["max_row"] - used_range["min_row"] + 1
            base_height += rows * 35
        
        return max(base_height, 200)

    def calculate_column_widths(self, min_row, max_row, min_col, max_col):
        widths = []
        padding = 30
        
        for col in range(min_col, max_col + 1):
            max_len = 0
            for row in range(min_row, max_row + 1):
                value = self.editor.get_cell_value(row, col) or ""
                text = str(value)
                max_len = max(max_len, len(text))
            
            width = max(max_len * 14 + padding, 100)
            widths.append(width)
        
        return widths

    def create_round_rect(self, x1, y1, x2, y2, radius=25, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def on_mousewheel(self, event):
        self.scroll_y -= event.delta
        self.scroll_y = max(0, self.scroll_y)
        self.draw_cards()

    def refresh(self):
        self.draw_cards()


class CardViewApp:
    def __init__(self, root):
        self.root = root
        self.root.title("卡片视图 - 表格处理程序")
        self.root.geometry("1200x900")
        self.root.state("zoomed")
        
        self.editor = SpreadsheetEditor()
        self.current_file = None
        
        self.create_menu()
        self.create_toolbar()
        self.create_main_content()
        
        self.load_sample_data()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件(F)", menu=file_menu)
        file_menu.add_command(label="新建(N)", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="打开(O)...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="保存(S)", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="另存为(A)...", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出(X)", command=self.root.quit)
        
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="编辑(E)", menu=edit_menu)
        edit_menu.add_command(label="添加人员卡片", command=self.add_person_card)
        edit_menu.add_command(label="编辑当前卡片", command=self.edit_current_card)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="视图(V)", menu=view_menu)
        view_menu.add_command(label="刷新视图", command=self.refresh_view)

    def create_toolbar(self):
        toolbar = tk.Frame(self.root, bg="#f3f3f3", height=40)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)
        
        btn_frame = tk.Frame(toolbar, bg="#f3f3f3")
        btn_frame.pack(side=tk.LEFT, padx=5, pady=5)
        
        buttons = [
            ("📄 新建", self.new_file),
            ("📂 打开", self.open_file),
            ("💾 保存", self.save_file),
            ("➕ 添加人员", self.add_person_card),
            ("🔄 刷新", self.refresh_view),
        ]
        
        for text, cmd in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd,
                          bg="#ffffff", relief="raised", padx=10, pady=3)
            btn.pack(side=tk.LEFT, padx=2)

    def create_main_content(self):
        self.card_canvas = CardView(self.root, self.editor)
        self.card_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.card_canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_sample_data(self):
        sample_data = [
            {
                "name": "孙波",
                "data": [
                    ["姓名", "旷课课时", "旷课科目", "旷课日期节数", "累计旷课课时"],
                    ["孙波", "2", "毛泽东思想概论", "2026年3月27日第7,8节课", "2"]
                ]
            },
            {
                "name": "杨晓然",
                "data": [
                    ["姓名", "旷课课时", "旷课科目", "旷课日期节数", "累计旷课课时"],
                    ["杨晓然", "2", "算法分析与设计", "2026年3月25日第7,8节课", "2"]
                ]
            },
            {
                "name": "张海洋",
                "data": [
                    ["姓名", "旷课课时", "旷课科目", "旷课日期节数", "累计旷课课时"],
                    ["张海洋", "2", "毛泽东思想概论", "2026年3月27日第7,8节课", "2"]
                ]
            }
        ]
        
        self.editor.delete_sheet("Sheet")
        
        for person in sample_data:
            self.editor.create_sheet(person["name"])
            self.editor.set_range_data(1, 1, person["data"])
            self.editor.auto_fit_columns()
        
        self.editor.select_sheet(sample_data[0]["name"])
        self.refresh_view()

    def add_person_card(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("添加人员卡片")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="姓名:").pack(pady=(10, 5))
        name_entry = ttk.Entry(dialog, width=40)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        ttk.Label(dialog, text="旷课课时:").pack(pady=(10, 5))
        hours_entry = ttk.Entry(dialog, width=40)
        hours_entry.pack(pady=5)
        
        ttk.Label(dialog, text="旷课科目:").pack(pady=(10, 5))
        subject_entry = ttk.Entry(dialog, width=40)
        subject_entry.pack(pady=5)
        
        ttk.Label(dialog, text="旷课日期节数:").pack(pady=(10, 5))
        date_entry = ttk.Entry(dialog, width=40)
        date_entry.pack(pady=5)
        
        ttk.Label(dialog, text="累计旷课课时:").pack(pady=(10, 5))
        total_entry = ttk.Entry(dialog, width=40)
        total_entry.pack(pady=5)
        
        def on_ok():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("错误", "请输入姓名")
                return
            
            data = [
                ["姓名", "旷课课时", "旷课科目", "旷课日期节数", "累计旷课课时"],
                [name, hours_entry.get(), subject_entry.get(), date_entry.get(), total_entry.get()]
            ]
            
            self.editor.create_sheet(name)
            self.editor.set_range_data(1, 1, data)
            self.editor.auto_fit_columns()
            
            dialog.destroy()
            self.refresh_view()
            messagebox.showinfo("成功", f"已添加 {name} 的卡片")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def edit_current_card(self):
        messagebox.showinfo("提示", "编辑功能开发中...")

    def new_file(self):
        self.editor = SpreadsheetEditor()
        self.current_file = None
        self.refresh_view()
        self.root.title("卡片视图 - 文档1")

    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.editor = SpreadsheetEditor(file_path)
            self.current_file = file_path
            self.refresh_view()
            self.root.title(f"卡片视图 - {os.path.basename(file_path)}")

    def save_file(self):
        if self.current_file:
            self.editor.save_workbook(self.current_file)
            messagebox.showinfo("成功", "文件已保存")
        else:
            self.save_as_file()

    def save_as_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.editor.save_workbook(file_path)
            self.current_file = file_path
            self.root.title(f"卡片视图 - {os.path.basename(file_path)}")
            messagebox.showinfo("成功", "文件已保存")

    def refresh_view(self):
        self.card_canvas.refresh()


def main():
    root = tk.Tk()
    app = CardViewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
