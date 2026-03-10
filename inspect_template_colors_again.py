import openpyxl

file_path = "Event_Performance Management-History Query-2G_KPI_Reporting_Template-DFBG6870-20251204091043.xlsx"
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

dxfs = wb._differential_styles

print("Found conditional formatting rules:")
for cf in sheet.conditional_formatting:
    for rule in cf.cfRule:
        if rule.dxfId is not None:
            dxf = dxfs[rule.dxfId]
            if dxf.fill and dxf.fill.fgColor:
                if hasattr(dxf.fill.fgColor, 'rgb') and dxf.fill.fgColor.rgb is not None:
                    print(f"Cells: {cf.sqref}, Formula: {rule.formula}, Operator: {rule.operator}, RGB Color: {dxf.fill.fgColor.rgb}")
                elif hasattr(dxf.fill.fgColor, 'theme') and dxf.fill.fgColor.theme is not None:
                    print(f"Cells: {cf.sqref}, Formula: {rule.formula}, Operator: {rule.operator}, Theme Color: {dxf.fill.fgColor.theme}, Tint: {dxf.fill.fgColor.tint}")
