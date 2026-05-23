import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


companies_data = []
for i in range(1, 101):
    companies_data.append({
        "Company Name": f"NY Tech Solutions Inc {i}",
        "Industry": "Information Technology" if i % 2 == 0 else "Corporate Services",
        "Website": f"https://www.nytechsolutions{i}.com",
        "Email": f"contact@nytechsolutions{i}.com",
        "Phone": f"+1 (555) 019-{i:03d}",
        "Status": "Active" if i % 4 != 0 else "Pending"
    })


df = pd.DataFrame(companies_data)


wb = Workbook()
ws = wb.active
ws.title = "Scraped Companies"


ws.views.sheetView[0].showGridLines = True

# (Corporate Deep Navy Blue Theme)
headers = list(df.columns)
ws.append(headers)

header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Deep Navy
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
center_alignment = Alignment(horizontal="center", vertical="center")
left_alignment = Alignment(horizontal="left", vertical="center")


for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment

 
thin_border = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)


zebra_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid") 
active_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") 
pending_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") 

active_font = Font(name="Calibri", size=11, color="155724", bold=True)
pending_font = Font(name="Calibri", size=11, color="856404", bold=True)
regular_font = Font(name="Calibri", size=11, color="000000")


for row_idx, row_data in enumerate(companies_data, start=2):
    row_values = list(row_data.values())
    ws.append(row_values)
    
    
    is_even = (row_idx % 2 == 0)
    
    for col_idx in range(1, len(row_values) + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.font = regular_font
        cell.border = thin_border
        
        
            cell.alignment = center_alignment
        else:
            cell.alignment = left_alignment
            
        if is_even:
            cell.fill = zebra_fill
            
         
        if col_idx == 6:  # Status Column
            if cell.value == "Active":
                cell.fill = active_fill
                cell.font = active_font
            else:
                cell.fill = pending_fill
                cell.font = pending_font


last_row = len(companies_data) + 2
ws.cell(row=last_row, column=1, value="Total Companies:").font = Font(name="Calibri", size=11, bold=True)
ws.cell(row=last_row, column=1).alignment = Alignment(horizontal="right")


formula_cell = ws.cell(row=last_row, column=2, value=f"=COUNTA(A2:A{last_row-2})")
formula_cell.font = Font(name="Calibri", size=11, bold=True)
formula_cell.alignment = center_alignment


for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    # সামান্য অতিরিক্ত স্পেসিং যোগ করা
    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


file_name = "ny_scraped_companies_list.xlsx"
wb.save(file_name)
print(f'{file_name}'