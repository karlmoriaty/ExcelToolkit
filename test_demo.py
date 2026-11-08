from spreadsheet_editor import SpreadsheetEditor
from format_templates import TemplateManager
from command_parser import CommandParser


def test_basic_operations():
    print("=== 测试基础操作 ===\n")
    
    editor = SpreadsheetEditor()
    template_manager = TemplateManager()
    parser = CommandParser(editor, template_manager)
    
    print("1. 创建工作表 '测试数据'")
    editor.create_sheet("测试数据")
    editor.select_sheet("测试数据")
    
    print("2. 设置表头数据")
    headers = ["姓名", "年龄", "性别", "职业", "城市"]
    for i, header in enumerate(headers):
        editor.set_cell_value(1, i + 1, header)
    
    print("3. 设置示例数据")
    sample_data = [
        ["张三", 28, "男", "工程师", "北京"],
        ["李四", 32, "女", "设计师", "上海"],
        ["王五", 25, "男", "产品经理", "深圳"],
        ["赵六", 30, "女", "数据分析师", "杭州"],
    ]
    editor.set_range_data(2, 1, sample_data)
    
    print("4. 应用表头模板到 A1:E1")
    header_template = template_manager.get_template("header")
    if header_template:
        editor.apply_range_font_style(1, 1, 1, 5, **header_template.font_style)
        editor.apply_range_fill_style(1, 1, 1, 5, header_template.fill_style["color"])
        editor.apply_range_alignment_style(1, 1, 1, 5, **header_template.alignment_style)
        editor.apply_table_border(1, 1, 1, 5, **header_template.border_style)
    
    print("5. 应用数据模板到 A2:E5")
    data_template = template_manager.get_template("data")
    if data_template:
        editor.apply_range_font_style(2, 1, 5, 5, **data_template.font_style)
        editor.apply_range_alignment_style(2, 1, 5, 5, **data_template.alignment_style)
        editor.apply_table_border(2, 1, 5, 5, **data_template.border_style)
    
    print("6. 自动调整列宽")
    editor.auto_fit_columns()
    
    print("7. 保存文件到 'demo_output.xlsx'")
    editor.save_workbook("demo_output.xlsx")
    
    print("\n=== 测试完成！文件已保存为 demo_output.xlsx ===")


def test_command_parser():
    print("\n=== 测试命令解析器 ===\n")
    
    editor = SpreadsheetEditor()
    template_manager = TemplateManager()
    parser = CommandParser(editor, template_manager)
    
    commands = [
        "create 命令测试",
        "select 命令测试",
        "set A1 产品名称",
        "set B1 价格",
        "set C1 库存",
        "set A2 笔记本电脑",
        "set B2 5999",
        "set C2 100",
        "set A3 无线鼠标",
        "set B3 129",
        "set C3 500",
        "table A1:C1 header",
        "table A2:C3 data",
        "autofit",
    ]
    
    for cmd in commands:
        print(f"> {cmd}")
        result = parser.parse_command(cmd)
        print(f"  {result['message']}")
    
    editor.save_workbook("command_demo.xlsx")
    print("\n=== 命令测试完成！文件已保存为 command_demo.xlsx ===")


if __name__ == "__main__":
    test_basic_operations()
    test_command_parser()
    print("\n所有测试完成！")
