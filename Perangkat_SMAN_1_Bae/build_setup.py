import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_styled_cell(ws, row, col, value, font=None, fill=None, alignment=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    if font: c.font = font
    if fill: c.fill = fill
    if alignment: c.alignment = alignment
    if border: c.border = border
    return c

# Common Styles
thin_side = Side(border_style="thin", color="CBD5E1")
med_side = Side(border_style="medium", color="1B365D")
double_bottom = Side(border_style="double", color="1B365D")

table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
header_border = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

font_title = Font(name="Segoe UI", size=15, bold=True, color="1B365D")
font_subtitle = Font(name="Segoe UI", size=11, bold=True, color="475569")
font_meta = Font(name="Segoe UI", size=10, italic=True, color="64748B")

font_sec_hdr = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_steel = PatternFill(start_color="243B53", end_color="243B53", fill_type="solid")
fill_slate = PatternFill(start_color="334E68", end_color="334E68", fill_type="solid")
fill_blue_light = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_key = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
fill_pts = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

font_tbl_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=10, color="1E293B")
font_data_bold = Font(name="Segoe UI", size=10, bold=True, color="1E293B")
font_key = Font(name="Segoe UI", size=10, bold=True, color="166534")
font_score = Font(name="Segoe UI", size=14, bold=True, color="1B365D")

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
align_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

print("Base setup ready.")
