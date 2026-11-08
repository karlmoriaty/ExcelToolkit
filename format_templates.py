import config


class FormatTemplate:
    def __init__(self, name, description=""):
        self.name = name
        self.description = description
        self.font_style = {}
        self.fill_style = {}
        self.alignment_style = {}
        self.border_style = {}

    def set_font(self, font_name=None, font_size=None, bold=False, italic=False, color=None):
        self.font_style = {
            "font_name": font_name,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": color
        }

    def set_fill(self, color):
        self.fill_style = {"color": color}

    def set_alignment(self, horizontal="center", vertical="center", wrap_text=False):
        self.alignment_style = {
            "horizontal": horizontal,
            "vertical": vertical,
            "wrap_text": wrap_text
        }

    def set_border(self, border_style="thin", apply_top=True, apply_bottom=True, apply_left=True, apply_right=True):
        self.border_style = {
            "border_style": border_style,
            "apply_top": apply_top,
            "apply_bottom": apply_bottom,
            "apply_left": apply_left,
            "apply_right": apply_right
        }


class TemplateManager:
    def __init__(self):
        self.templates = {}
        self._load_default_templates()

    def _load_default_templates(self):
        header = FormatTemplate("header", "表头样式")
        header.set_font(font_name=config.DEFAULT_FONT, font_size=config.HEADER_FONT_SIZE, bold=True)
        header.set_fill(color="D3D3D3")
        header.set_alignment(horizontal="center", vertical="center")
        header.set_border(border_style="thin")
        self.add_template(header)

        data_cell = FormatTemplate("data", "数据单元格样式")
        data_cell.set_font(font_name=config.DEFAULT_FONT, font_size=config.DEFAULT_FONT_SIZE)
        data_cell.set_alignment(horizontal="left", vertical="center")
        data_cell.set_border(border_style="thin")
        self.add_template(data_cell)

        number_cell = FormatTemplate("number", "数字单元格样式")
        number_cell.set_font(font_name=config.DEFAULT_FONT, font_size=config.DEFAULT_FONT_SIZE)
        number_cell.set_alignment(horizontal="right", vertical="center")
        number_cell.set_border(border_style="thin")
        self.add_template(number_cell)

        title = FormatTemplate("title", "标题样式")
        title.set_font(font_name=config.DEFAULT_FONT, font_size=16, bold=True)
        title.set_alignment(horizontal="center", vertical="center")
        self.add_template(title)

        highlight = FormatTemplate("highlight", "高亮样式")
        highlight.set_font(font_name=config.DEFAULT_FONT, font_size=config.DEFAULT_FONT_SIZE, bold=True)
        highlight.set_fill(color="FFFF00")
        highlight.set_alignment(horizontal="center", vertical="center")
        self.add_template(highlight)

    def add_template(self, template):
        self.templates[template.name] = template

    def get_template(self, name):
        return self.templates.get(name)

    def get_all_templates(self):
        return list(self.templates.values())

    def delete_template(self, name):
        if name in self.templates:
            del self.templates[name]
            return True
        return False
