from openpyxl.utils import column_index_from_string


class CommandParser:
    def __init__(self, spreadsheet_editor, template_manager):
        self.editor = spreadsheet_editor
        self.template_manager = template_manager

    def parse_command(self, command):
        command = command.strip()
        if not command:
            return {"success": False, "message": "请输入命令"}

        parts = command.split()
        if not parts:
            return {"success": False, "message": "无效命令"}

        cmd = parts[0].lower()

        if cmd == "select":
            return self._parse_select(parts[1:])
        elif cmd == "create":
            return self._parse_create(parts[1:])
        elif cmd == "delete":
            return self._parse_delete(parts[1:])
        elif cmd == "set":
            return self._parse_set(parts[1:])
        elif cmd == "fill":
            return self._parse_fill(parts[1:])
        elif cmd == "align":
            return self._parse_align(parts[1:])
        elif cmd == "border":
            return self._parse_border(parts[1:])
        elif cmd == "table":
            return self._parse_table(parts[1:])
        elif cmd == "autofit":
            return self._parse_autofit(parts[1:])
        elif cmd == "template":
            return self._parse_template(parts[1:])
        elif cmd == "help":
            return self._show_help()
        else:
            return {"success": False, "message": f"未知命令: {cmd}"}

    def _parse_cell_range(self, cell_str):
        if ":" in cell_str:
            start, end = cell_str.split(":")
            start_row, start_col = self._parse_cell(start)
            end_row, end_col = self._parse_cell(end)
            return {
                "type": "range",
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col
            }
        else:
            row, col = self._parse_cell(cell_str)
            return {
                "type": "cell",
                "row": row,
                "col": col
            }

    def _parse_cell(self, cell_str):
        import re
        match = re.match(r'([A-Za-z]+)(\d+)', cell_str)
        if match:
            col_letter = match.group(1).upper()
            row = int(match.group(2))
            col = column_index_from_string(col_letter)
            return row, col
        raise ValueError(f"无效的单元格格式: {cell_str}")

    def _parse_select(self, args):
        if len(args) < 1:
            return {"success": False, "message": "用法: select <工作表名>"}
        sheet_name = " ".join(args)
        if self.editor.select_sheet(sheet_name):
            return {"success": True, "message": f"已选择工作表: {sheet_name}"}
        return {"success": False, "message": f"工作表不存在: {sheet_name}"}

    def _parse_create(self, args):
        if len(args) < 1:
            return {"success": False, "message": "用法: create <工作表名>"}
        sheet_name = " ".join(args)
        if self.editor.create_sheet(sheet_name):
            return {"success": True, "message": f"已创建工作表: {sheet_name}"}
        return {"success": False, "message": f"工作表已存在: {sheet_name}"}

    def _parse_delete(self, args):
        if len(args) < 1:
            return {"success": False, "message": "用法: delete <工作表名>"}
        sheet_name = " ".join(args)
        if self.editor.delete_sheet(sheet_name):
            return {"success": True, "message": f"已删除工作表: {sheet_name}"}
        return {"success": False, "message": f"无法删除工作表: {sheet_name}"}

    def _parse_set(self, args):
        if len(args) < 2:
            return {"success": False, "message": "用法: set <单元格> <值>"}
        cell_str = args[0]
        value = " ".join(args[1:])
        try:
            cell_info = self._parse_cell_range(cell_str)
            if cell_info["type"] == "cell":
                self.editor.set_cell_value(cell_info["row"], cell_info["col"], value)
                return {"success": True, "message": f"已设置 {cell_str} = {value}"}
            else:
                return {"success": False, "message": "set命令仅支持单个单元格"}
        except Exception as e:
            return {"success": False, "message": f"错误: {str(e)}"}

    def _parse_fill(self, args):
        if len(args) < 2:
            return {"success": False, "message": "用法: fill <单元格范围> <颜色>"}
        cell_str = args[0]
        color = args[1]
        try:
            cell_info = self._parse_cell_range(cell_str)
            if cell_info["type"] == "cell":
                self.editor.apply_fill_style(cell_info["row"], cell_info["col"], color)
            else:
                self.editor.apply_range_fill_style(
                    cell_info["start_row"], cell_info["start_col"],
                    cell_info["end_row"], cell_info["end_col"],
                    color
                )
            return {"success": True, "message": f"已填充颜色: {cell_str} -> {color}"}
        except Exception as e:
            return {"success": False, "message": f"错误: {str(e)}"}

    def _parse_align(self, args):
        if len(args) < 2:
            return {"success": False, "message": "用法: align <单元格范围> <left/center/right>"}
        cell_str = args[0]
        align_type = args[1]
        if align_type not in ["left", "center", "right"]:
            return {"success": False, "message": "对齐方式必须是 left, center, 或 right"}
        try:
            cell_info = self._parse_cell_range(cell_str)
            if cell_info["type"] == "cell":
                self.editor.apply_alignment_style(cell_info["row"], cell_info["col"], horizontal=align_type)
            else:
                self.editor.apply_range_alignment_style(
                    cell_info["start_row"], cell_info["start_col"],
                    cell_info["end_row"], cell_info["end_col"],
                    horizontal=align_type
                )
            return {"success": True, "message": f"已设置对齐: {cell_str} -> {align_type}"}
        except Exception as e:
            return {"success": False, "message": f"错误: {str(e)}"}

    def _parse_border(self, args):
        if len(args) < 1:
            return {"success": False, "message": "用法: border <单元格范围> [thin/medium/thick]"}
        cell_str = args[0]
        border_style = args[1] if len(args) > 1 else "thin"
        try:
            cell_info = self._parse_cell_range(cell_str)
            if cell_info["type"] == "cell":
                self.editor.apply_border_style(cell_info["row"], cell_info["col"], border_style=border_style)
            else:
                self.editor.apply_table_border(
                    cell_info["start_row"], cell_info["start_col"],
                    cell_info["end_row"], cell_info["end_col"],
                    border_style=border_style
                )
            return {"success": True, "message": f"已设置边框: {cell_str} -> {border_style}"}
        except Exception as e:
            return {"success": False, "message": f"错误: {str(e)}"}

    def _parse_table(self, args):
        if len(args) < 2:
            return {"success": False, "message": "用法: table <单元格范围> <模板名>"}
        cell_str = args[0]
        template_name = args[1]
        try:
            template = self.template_manager.get_template(template_name)
            if not template:
                return {"success": False, "message": f"模板不存在: {template_name}"}
            
            cell_info = self._parse_cell_range(cell_str)
            if cell_info["type"] == "cell":
                return {"success": False, "message": "table命令需要单元格范围"}
            
            start_row, start_col = cell_info["start_row"], cell_info["start_col"]
            end_row, end_col = cell_info["end_row"], cell_info["end_col"]
            
            if template.font_style:
                self.editor.apply_range_font_style(start_row, start_col, end_row, end_col, **template.font_style)
            if template.fill_style:
                self.editor.apply_range_fill_style(start_row, start_col, end_row, end_col, template.fill_style["color"])
            if template.alignment_style:
                self.editor.apply_range_alignment_style(start_row, start_col, end_row, end_col, **template.alignment_style)
            if template.border_style:
                self.editor.apply_table_border(start_row, start_col, end_row, end_col, **template.border_style)
            
            return {"success": True, "message": f"已应用模板: {template_name} 到 {cell_str}"}
        except Exception as e:
            return {"success": False, "message": f"错误: {str(e)}"}

    def _parse_autofit(self, args):
        self.editor.auto_fit_columns()
        return {"success": True, "message": "已自动调整列宽"}

    def _parse_template(self, args):
        if len(args) < 1:
            templates = self.template_manager.get_all_templates()
            template_list = "\n".join([f"- {t.name}: {t.description}" for t in templates])
            return {"success": True, "message": f"可用模板:\n{template_list}"}
        return {"success": False, "message": "用法: template (查看所有模板)"}

    def _show_help(self):
        help_text = """
可用命令:
  select <工作表名>    - 选择工作表
  create <工作表名>    - 创建新工作表
  delete <工作表名>    - 删除工作表
  set <单元格> <值>     - 设置单元格值
  fill <范围> <颜色>    - 填充背景色
  align <范围> <left/center/right> - 设置对齐
  border <范围> [样式]  - 设置边框
  table <范围> <模板名>  - 应用表格模板
  autofit               - 自动调整列宽
  template              - 查看所有模板
  help                  - 显示帮助
"""
        return {"success": True, "message": help_text}
