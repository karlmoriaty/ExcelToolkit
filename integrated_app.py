import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl.utils import get_column_letter
import os
from spreadsheet_editor import SpreadsheetEditor
from card_view import CardView


class CellCanvas(tk.Canvas):
    def __init__(self, parent, editor, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.editor = editor
        self.cell_width = 80
        self.cell_height = 25
        self.row_header_width = 50
        self.col_header_height = 25
        
        self.selected_row = 1
        self.selected_col = 1
        self.selection_start = None
        self.selection_end = None
        self.is_selecting = False
        
        self.edit_entry = None
        self.editing = False
        
        self.scroll_x = 0
        self.scroll_y = 0
        
        self.configure(bg="white")
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Double-Button-1>", self.on_double_click)
        self.bind("<Key>", self.on_key)
        self.bind("<FocusIn>", self.on_focus_in)
        self.bind("<FocusOut>", self.on_focus_out)
        self.bind("<MouseWheel>", self.on_mousewheel)
        
        self.draw_grid()

    def draw_grid(self):
        self.delete("all")
        
        max_rows = 100
        max_cols = 50
        
        x_offset = self.row_header_width - self.scroll_x
        y_offset = self.col_header_height - self.scroll_y
        
        for col in range(max_cols):
            x = x_offset + col * self.cell_width
            col_letter = get_column_letter(col + 1)
            
            bg_color = "#e6e6e6"
            if col + 1 == self.selected_col:
                bg_color = "#d0e0ff"
            
            self.create_rectangle(
                x, 0, x + self.cell_width, self.col_header_height,
                fill=bg_color, outline="#c0c0c0"
            )
            self.create_text(
                x + self.cell_width / 2, self.col_header_height / 2,
                text=col_letter, font=("微软雅黑", 9)
            )
        
        for row in range(max_rows):
            y = y_offset + row * self.cell_height
            
            bg_color = "#e6e6e6"
            if row + 1 == self.selected_row:
                bg_color = "#d0e0ff"
            
            self.create_rectangle(
                0, y, self.row_header_width, y + self.cell_height,
                fill=bg_color, outline="#c0c0c0"
            )
            self.create_text(
                self.row_header_width / 2, y + self.cell_height / 2,
                text=str(row + 1), font=("微软雅黑", 9)
            )
        
        for row in range(max_rows):
            for col in range(max_cols):
                x = x_offset + col * self.cell_width
                y = y_offset + row * self.cell_height
                
                is_selected = False
                if self.selection_start and self.selection_end:
                    min_row = min(self.selection_start[0], self.selection_end[0])
                    max_row = max(self.selection_start[0], self.selection_end[0])
                    min_col = min(self.selection_start[1], self.selection_end[1])
                    max_col = max(self.selection_start[1], self.selection_end[1])
                    is_selected = (min_row <= row + 1 <= max_row and 
                                   min_col <= col + 1 <= max_col)
                else:
                    is_selected = (row + 1 == self.selected_row and col + 1 == self.selected_col)
                
                bg_color = "white"
                if is_selected:
                    bg_color = "#e8f0ff"
                
                self.create_rectangle(
                    x, y, x + self.cell_width, y + self.cell_height,
                    fill=bg_color, outline="#e0e0e0"
                )
                
                value = self.editor.get_cell_value(row + 1, col + 1) or ""
                if value is not None:
                    self.create_text(
                        x + 5, y + self.cell_height / 2,
                        text=str(value), anchor="w", font=("微软雅黑", 9)
                    )
        
        if self.selection_start and self.selection_end:
            min_row = min(self.selection_start[0], self.selection_end[0])
            max_row = max(self.selection_start[0], self.selection_end[0])
            min_col = min(self.selection_start[1], self.selection_end[1])
            max_col = max(self.selection_start[1], self.selection_end[1])
            
            x1 = x_offset + (min_col - 1) * self.cell_width
            y1 = y_offset + (min_row - 1) * self.cell_height
            x2 = x_offset + max_col * self.cell_width
            y2 = y_offset + max_row * self.cell_height
            
            self.create_rectangle(
                x1, y1, x2, y2,
                outline="#0078d4", width=2, fill=""
            )
        else:
            x = x_offset + (self.selected_col - 1) * self.cell_width
            y = y_offset + (self.selected_row - 1) * self.cell_height
            
            self.create_rectangle(
                x, y, x + self.cell_width, y + self.cell_height,
                outline="#0078d4", width=2, fill=""
            )

    def get_cell_from_pos(self, x, y):
        x_offset = self.row_header_width - self.scroll_x
        y_offset = self.col_header_height - self.scroll_y
        
        col = int((x - x_offset) / self.cell_width) + 1
        row = int((y - y_offset) / self.cell_height) + 1
        
        return max(1, row), max(1, col)

    def on_click(self, event):
        self.focus_set()
        row, col = self.get_cell_from_pos(event.x, event.y)
        self.selected_row = row
        self.selected_col = col
        self.selection_start = (row, col)
        self.selection_end = None
        self.is_selecting = True
        self.draw_grid()
        self.event_generate("<<SelectionChanged>>")

    def on_drag(self, event):
        if self.is_selecting:
            row, col = self.get_cell_from_pos(event.x, event.y)
            self.selection_end = (row, col)
            self.draw_grid()

    def on_release(self, event):
        self.is_selecting = False

    def on_double_click(self, event):
        row, col = self.get_cell_from_pos(event.x, event.y)
        self.start_editing(row, col)

    def on_key(self, event):
        if self.editing:
            return
        
        if event.keysym in ["Left", "Right", "Up", "Down", "Return", "Tab"]:
            if event.keysym == "Left" and self.selected_col > 1:
                self.selected_col -= 1
            elif event.keysym == "Right":
                self.selected_col += 1
            elif event.keysym == "Up" and self.selected_row > 1:
                self.selected_row -= 1
            elif event.keysym == "Down" or event.keysym == "Return":
                self.selected_row += 1
            elif event.keysym == "Tab":
                self.selected_col += 1
            
            self.selection_start = (self.selected_row, self.selected_col)
            self.selection_end = None
            self.draw_grid()
            self.event_generate("<<SelectionChanged>>")
            return "break"
        
        if event.char and event.char.isprintable():
            self.start_editing(self.selected_row, self.selected_col, event.char)
            return "break"
        
        if event.keysym == "Delete" or event.keysym == "BackSpace":
            self.editor.set_cell_value(self.selected_row, self.selected_col, None)
            self.draw_grid()
            return "break"

    def start_editing(self, row, col, initial_text=""):
        if self.edit_entry:
            self.edit_entry.destroy()
        
        x_offset = self.row_header_width - self.scroll_x
        y_offset = self.col_header_height - self.scroll_y
        
        x = x_offset + (col - 1) * self.cell_width
        y = y_offset + (row - 1) * self.cell_height
        
        value = self.editor.get_cell_value(row, col) or ""
        self.edit_entry = tk.Entry(self, font=("微软雅黑", 9))
        self.edit_entry.place(x=x + 1, y=y + 1, width=self.cell_width - 2, height=self.cell_height - 2)
        
        if initial_text:
            self.edit_entry.insert(0, initial_text)
        else:
            self.edit_entry.insert(0, str(value))
        
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus_set()
        self.editing = True
        self.edit_entry.bind("<Return>", lambda e: self.finish_editing(row, col))
        self.edit_entry.bind("<Escape>", lambda e: self.cancel_editing())
        self.edit_entry.bind("<FocusOut>", lambda e: self.finish_editing(row, col))

    def finish_editing(self, row, col):
        if self.edit_entry:
            value = self.edit_entry.get()
            self.editor.set_cell_value(row, col, value if value else None)
            self.edit_entry.destroy()
            self.edit_entry = None
            self.editing = False
            self.draw_grid()

    def cancel_editing(self):
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
            self.editing = False

    def on_focus_in(self, event):
        self.draw_grid()

    def on_focus_out(self, event):
        if not self.editing:
            self.draw_grid()

    def on_mousewheel(self, event):
        self.scroll_y -= event.delta
        self.scroll_y = max(0, self.scroll_y)
        self.draw_grid()

    def get_selection(self):
        if self.selection_start and self.selection_end:
            min_row = min(self.selection_start[0], self.selection_end[0])
            max_row = max(self.selection_start[0], self.selection_end[0])
            min_col = min(self.selection_start[1], self.selection_end[1])
            max_col = max(self.selection_start[1], self.selection_end[1])
            return (min_row, min_col, max_row, max_col)
        return (self.selected_row, self.selected_col, self.selected_row, self.selected_col)

    def refresh(self):
        self.draw_grid()


class IntegratedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格处理程序 - 集成版")
        self.root.geometry("1400x900")
        self.root.state("zoomed")
        
        self.editor = SpreadsheetEditor()
        self.current_file = None
        self.current_view = "grid"
        
        self.create_menu()
        self.create_toolbar()
        self.create_view_container()
        self.create_sheet_tabs()
        self.create_status_bar()
        
        self.load_sample_data()
        self.refresh_sheet_tabs()

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
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="视图(V)", menu=view_menu)
        view_menu.add_command(label="表格视图", command=lambda: self.switch_view("grid"))
        view_menu.add_command(label="卡片视图", command=lambda: self.switch_view("card"))

    def create_toolbar(self):
        toolbar = tk.Frame(self.root, bg="#f3f3f3", height=50)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)
        
        left_frame = tk.Frame(toolbar, bg="#f3f3f3")
        left_frame.pack(side=tk.LEFT, padx=5, pady=10)
        
        buttons = [
            ("📄 新建", self.new_file),
            ("📂 打开", self.open_file),
            ("💾 保存", self.save_file),
        ]
        
        for text, cmd in buttons:
            btn = tk.Button(left_frame, text=text, command=cmd,
                          bg="#ffffff", relief="raised", padx=10, pady=3)
            btn.pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(left_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        tk.Button(left_frame, text="📊 表格视图", command=lambda: self.switch_view("grid"),
                 bg="#ffffff", relief="raised", padx=10, pady=3).pack(side=tk.LEFT, padx=2)
        tk.Button(left_frame, text="🃏 卡片视图", command=lambda: self.switch_view("card"),
                 bg="#ffffff", relief="raised", padx=10, pady=3).pack(side=tk.LEFT, padx=2)
        
        right_frame = tk.Frame(toolbar, bg="#f3f3f3")
        right_frame.pack(side=tk.RIGHT, padx=5, pady=10)
        
        tk.Button(right_frame, text="➕ 添加人员", command=self.add_person_card,
                 bg="#ffffff", relief="raised", padx=10, pady=3).pack(side=tk.LEFT, padx=2)
        tk.Button(right_frame, text="🔄 刷新", command=self.refresh_view,
                 bg="#ffffff", relief="raised", padx=10, pady=3).pack(side=tk.LEFT, padx=2)

    def create_view_container(self):
        self.view_container = tk.Frame(self.root)
        self.view_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.switch_view("grid")

    def switch_view(self, view_type):
        for widget in self.view_container.winfo_children():
            widget.destroy()
        
        self.current_view = view_type
        
        if view_type == "grid":
            self.create_grid_view()
        else:
            self.create_card_view()

    def create_grid_view(self):
        main_container = tk.Frame(self.view_container)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        self.cell_canvas = CellCanvas(main_container, self.editor)
        self.cell_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=self.cell_canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_card_view(self):
        main_container = tk.Frame(self.view_container)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        self.card_canvas = CardView(main_container, self.editor)
        self.card_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=self.card_canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_sheet_tabs(self):
        tab_frame = tk.Frame(self.root, bg="#e6e6e6", height=30)
        tab_frame.pack(side=tk.BOTTOM, fill=tk.X, before=self.view_container)
        tab_frame.pack_propagate(False)
        
        left_panel = tk.Frame(tab_frame, bg="#e6e6e6")
        left_panel.pack(side=tk.LEFT)
        
        tk.Button(left_panel, text="◀", width=3, bg="#e6e6e6", relief="flat").pack(side=tk.LEFT)
        tk.Button(left_panel, text="▶", width=3, bg="#e6e6e6", relief="flat").pack(side=tk.LEFT)
        
        self.tab_container = tk.Frame(tab_frame, bg="#e6e6e6")
        self.tab_container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.sheet_tabs = []
        
        add_btn = tk.Button(tab_frame, text="+", width=3, bg="#e6e6e6", relief="flat", command=self.add_person_card)
        add_btn.pack(side=tk.RIGHT)

    def create_status_bar(self):
        status_bar = tk.Frame(self.root, bg="#e6e6e6", height=22)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        status_bar.pack_propagate(False)
        
        left_status = tk.Frame(status_bar, bg="#e6e6e6")
        left_status.pack(side=tk.LEFT, padx=10)
        
        self.status_label = tk.Label(left_status, text="就绪", bg="#e6e6e6", font=("微软雅黑", 8))
        self.status_label.pack(side=tk.LEFT)
        
        right_status = tk.Frame(status_bar, bg="#e6e6e6")
        right_status.pack(side=tk.RIGHT, padx=10)
        
        self.view_label = tk.Label(right_status, text="表格视图", bg="#e6e6e6", font=("微软雅黑", 8))
        self.view_label.pack(side=tk.LEFT, padx=5)

    def refresh_sheet_tabs(self):
        for widget in self.tab_container.winfo_children():
            widget.destroy()
        
        self.sheet_tabs = []
        sheet_names = self.editor.get_sheet_names()
        
        for i, name in enumerate(sheet_names):
            is_active = self.editor.current_sheet and self.editor.current_sheet.title == name
            bg = "white" if is_active else "#e6e6e6"
            fg = "black" if is_active else "#666666"
            relief = "raised" if is_active else "flat"
            
            tab = tk.Button(
                self.tab_container, 
                text=f"  {name}  ", 
                bg=bg, 
                fg=fg,
                relief=relief,
                font=("微软雅黑", 9),
                command=lambda n=name: self.select_sheet(n)
            )
            tab.pack(side=tk.LEFT, padx=1, pady=2)
            self.sheet_tabs.append(tab)

    def select_sheet(self, sheet_name):
        self.editor.select_sheet(sheet_name)
        self.refresh_sheet_tabs()
        self.refresh_view()

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
            self.refresh_sheet_tabs()
            self.refresh_view()
            messagebox.showinfo("成功", f"已添加 {name} 的卡片")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def new_file(self):
        self.editor = SpreadsheetEditor()
        self.current_file = None
        self.refresh_sheet_tabs()
        self.refresh_view()
        self.root.title("表格处理程序 - 集成版 - 文档1")

    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.editor = SpreadsheetEditor(file_path)
            self.current_file = file_path
            self.refresh_sheet_tabs()
            self.refresh_view()
            self.root.title(f"表格处理程序 - 集成版 - {os.path.basename(file_path)}")

    def save_file(self):
        if self.current_file:
            self.editor.save_workbook(self.current_file)
            self.status_label.config(text="已保存")
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
            self.root.title(f"表格处理程序 - 集成版 - {os.path.basename(file_path)}")
            self.status_label.config(text="已保存")

    def refresh_view(self):
        if self.current_view == "grid" and hasattr(self, 'cell_canvas'):
            self.cell_canvas.refresh()
        elif self.current_view == "card" and hasattr(self, 'card_canvas'):
            self.card_canvas.refresh()
        
        view_text = "表格视图" if self.current_view == "grid" else "卡片视图"
        self.view_label.config(text=view_text)


def main():
    root = tk.Tk()
    app = IntegratedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
