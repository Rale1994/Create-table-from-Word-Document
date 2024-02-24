from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.oxml.ns import qn


def oboji_tekst_celija_tabele(table):
    for row in table.rows:
        for cell in row.cells:
            if cell.text:
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "FFFF00")
                cell._element.get_or_add_tcPr().append(shd)


def set_table_font(table, font_name, font_size, font_color):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.font.color.rgb = font_color


def color_filled_cells(table, filled_color):
    for row in table.rows:
        for cell in row.cells:
            text = cell.text.strip()
            if text:
                cell.paragraphs[0].runs[0].font.color.rgb = filled_color
            else:
                continue


def set_cell_shading(cell, fill_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill_color)
    cell._element.get_or_add_tcPr().append(shd)


def set_table_color(table):
    columns_to_color = set()
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if cell.text:
                columns_to_color.add(i)
    for i in columns_to_color:
        for row in table.rows:
            set_cell_shading(row.cells[i], "FFFF00")
