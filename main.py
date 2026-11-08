import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, font
from openpyxl.utils import get_column_letter, column_index_from_string
import os
from spreadsheet_editor import SpreadsheetEditor
from format_templates import TemplateManager
from command_parser import CommandParser


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
                
                value = self.editor.get_cell_value(row + 1, col + 1)
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


class SpreadsheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WPS表格 - 文档1")
        self.root.geometry("1400x900")
        self.root.state("zoomed")
        
        self.editor = SpreadsheetEditor()
        self.template_manager = TemplateManager()
        self.command_parser = CommandParser(self.editor, self.template_manager)
        
        self.current_file = None
        
        self.create_menu()
        self.create_toolbar()
        self.create_main_content()
        self.create_sheet_tabs()
        self.create_status_bar()
        
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
        edit_menu.add_command(label="撤销(U)", accelerator="Ctrl+Z")
        edit_menu.add_command(label="恢复(R)", accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="剪切(T)", accelerator="Ctrl+X")
        edit_menu.add_command(label="复制(C)", accelerator="Ctrl+C")
        edit_menu.add_command(label="粘贴(P)", accelerator="Ctrl+V")
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="视图(V)", menu=view_menu)
        
        insert_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="插入(I)", menu=insert_menu)
        
        format_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="格式(O)", menu=format_menu)
        
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="数据(D)", menu=data_menu)
        
        window_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="窗口(W)", menu=window_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助(H)", menu=help_menu)
        help_menu.add_command(label="使用帮助", command=self.show_help)

    def create_toolbar(self):
        toolbar = tk.Frame(self.root, bg="#f3f3f3", height=80)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)
        
        top_toolbar = tk.Frame(toolbar, bg="#f3f3f3")
        top_toolbar.pack(side=tk.TOP, fill=tk.X, padx=2, pady=2)
        
        self.font_var = tk.StringVar(value="微软雅黑")
        font_combo = ttk.Combobox(top_toolbar, textvariable=self.font_var, width=15, state="readonly")
        font_combo["values"] = ["微软雅黑", "宋体", "黑体", "Arial", "Times New Roman"]
        font_combo.pack(side=tk.LEFT, padx=2)
        
        self.font_size_var = tk.StringVar(value="11")
        size_combo = ttk.Combobox(top_toolbar, textvariable=self.font_size_var, width=5, state="readonly")
        size_combo["values"] = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "28", "32", "36", "48", "72"]
        size_combo.pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(top_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        btn_frame = tk.Frame(top_toolbar, bg="#f3f3f3")
        btn_frame.pack(side=tk.LEFT)
        
        buttons = [
            ("B", self.apply_bold),
            ("I", self.apply_italic),
            ("U", self.apply_underline),
        ]
        
        for text, cmd in buttons:
            btn = tk.Button(btn_frame, text=text, width=4, command=cmd, 
                           bg="#f3f3f3", relief="flat", font=("微软雅黑", 9, "bold"))
            btn.pack(side=tk.LEFT, padx=1)
        
        ttk.Separator(top_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        color_frame = tk.Frame(top_toolbar, bg="#f3f3f3")
        color_frame.pack(side=tk.LEFT)
        
        tk.Button(color_frame, text="A", width=4, command=self.apply_text_color,
                 bg="#f3f3f3", relief="flat").pack(side=tk.LEFT, padx=1)
        tk.Button(color_frame, text="🪣", width=4, command=self.apply_fill_color,
                 bg="#f3f3f3", relief="flat").pack(side=tk.LEFT, padx=1)
        
        ttk.Separator(top_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        align_frame = tk.Frame(top_toolbar, bg="#f3f3f3")
        align_frame.pack(side=tk.LEFT)
        
        align_buttons = [
            ("⬅", "left"),
            ("⬌", "center"),
            ("➡", "right"),
        ]
        
        for text, align in align_buttons:
            btn = tk.Button(align_frame, text=text, width=4, 
                           command=lambda a=align: self.apply_alignment(a),
                           bg="#f3f3f3", relief="flat")
            btn.pack(side=tk.LEFT, padx=1)
        
        ttk.Separator(top_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        border_frame = tk.Frame(top_toolbar, bg="#f3f3f3")
        border_frame.pack(side=tk.LEFT)
        
        tk.Button(border_frame, text="▦", width=4, command=self.apply_border,
                 bg="#f3f3f3", relief="flat").pack(side=tk.LEFT, padx=1)
        
        bottom_toolbar = tk.Frame(toolbar, bg="#f3f3f3")
        bottom_toolbar.pack(side=tk.TOP, fill=tk.X, padx=2, pady=2)
        
        formula_bar = tk.Frame(bottom_toolbar, bg="white", relief="sunken", bd=1)
        formula_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(formula_bar, text="fx", bg="#f3f3f3", width=4).pack(side=tk.LEFT)
        
        self.address_label = tk.Label(formula_bar, text="A1", width=8, bg="white", relief="flat")
        self.address_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(formula_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y)
        
        self.formula_entry = tk.Entry(formula_bar, bg="white", relief="flat")
        self.formula_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    def create_main_content(self):
        main_container = tk.Frame(self.root)
        main_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.cell_canvas = CellCanvas(main_container, self.editor)
        self.cell_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=self.cell_canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.cell_canvas.bind("<<SelectionChanged>>", self.on_selection_changed)

    def create_sheet_tabs(self):
        tab_frame = tk.Frame(self.root, bg="#e6e6e6", height=30)
        tab_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tab_frame.pack_propagate(False)
        
        left_panel = tk.Frame(tab_frame, bg="#e6e6e6")
        left_panel.pack(side=tk.LEFT)
        
        tk.Button(left_panel, text="◀", width=3, bg="#e6e6e6", relief="flat").pack(side=tk.LEFT)
        tk.Button(left_panel, text="▶", width=3, bg="#e6e6e6", relief="flat").pack(side=tk.LEFT)
        
        self.tab_container = tk.Frame(tab_frame, bg="#e6e6e6")
        self.tab_container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.sheet_tabs = []
        
        add_btn = tk.Button(tab_frame, text="+", width=3, bg="#e6e6e6", relief="flat", command=self.create_new_sheet)
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
        
        tk.Label(right_status, text="📊", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        tk.Label(right_status, text="👁", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        tk.Label(right_status, text="⚙", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(right_status, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        tk.Label(right_status, text="🗖", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        tk.Label(right_status, text="🗗", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        tk.Label(right_status, text="🗔", bg="#e6e6e6").pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(right_status, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        self.zoom_label = tk.Label(right_status, text="100%", bg="#e6e6e6", font=("微软雅黑", 8))
        self.zoom_label.pack(side=tk.LEFT, padx=5)

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
        self.cell_canvas.refresh()

    def on_selection_changed(self, event):
        row, col = self.cell_canvas.selected_row, self.cell_canvas.selected_col
        address = f"{get_column_letter(col)}{row}"
        self.address_label.config(text=address)
        
        value = self.editor.get_cell_value(row, col)
        self.formula_entry.delete(0, tk.END)
        if value:
            self.formula_entry.insert(0, str(value))

    def new_file(self):
        self.editor = SpreadsheetEditor()
        self.command_parser = CommandParser(self.editor, self.template_manager)
        self.current_file = None
        self.refresh_sheet_tabs()
        self.cell_canvas.refresh()
        self.root.title("WPS表格 - 文档1")

    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.editor = SpreadsheetEditor(file_path)
            self.command_parser = CommandParser(self.editor, self.template_manager)
            self.current_file = file_path
            self.refresh_sheet_tabs()
            self.cell_canvas.refresh()
            self.root.title(f"WPS表格 - {os.path.basename(file_path)}")

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
            self.root.title(f"WPS表格 - {os.path.basename(file_path)}")
            self.status_label.config(text="已保存")

    def create_new_sheet(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("新建工作表")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="工作表名称:").pack(pady=10)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.pack(pady=5)
        name_entry.insert(0, f"Sheet{len(self.editor.get_sheet_names()) + 1}")
        name_entry.focus()
        
        def on_ok():
            sheet_name = name_entry.get().strip()
            if sheet_name:
                if self.editor.create_sheet(sheet_name):
                    self.refresh_sheet_tabs()
                    self.cell_canvas.refresh()
                    dialog.destroy()
                else:
                    messagebox.showerror("错误", "工作表已存在")
        
        ttk.Button(dialog, text="确定", command=on_ok).pack(pady=10)

    def apply_bold(self):
        selection = self.cell_canvas.get_selection()
        self.editor.apply_range_font_style(*selection, bold=True)
        self.cell_canvas.refresh()

    def apply_italic(self):
        selection = self.cell_canvas.get_selection()
        self.editor.apply_range_font_style(*selection, italic=True)
        self.cell_canvas.refresh()

    def apply_underline(self):
        pass

    def apply_text_color(self):
        pass

    def apply_fill_color(self):
        pass

    def apply_alignment(self, align_type):
        selection = self.cell_canvas.get_selection()
        self.editor.apply_range_alignment_style(*selection, horizontal=align_type)
        self.cell_canvas.refresh()

    def apply_border(self):
        selection = self.cell_canvas.get_selection()
        self.editor.apply_table_border(*selection, border_style="thin")
        self.cell_canvas.refresh()

    def show_help(self):
        help_text = """
WPS表格 - 使用帮助

主要功能:
1. 表格编辑：直接点击单元格进行编辑
2. 工作表管理：点击底部标签切换工作表，点击+新建
3. 格式设置：使用工具栏快速设置格式
4. 文件操作：新建、打开、保存Excel文件

快捷键:
  Ctrl+N - 新建文件
  Ctrl+O - 打开文件
  Ctrl+S - 保存文件
"""
        messagebox.showinfo("使用帮助", help_text)


def main():
    root = tk.Tk()
    app = SpreadsheetApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
