import openpyxl

file_path = "Event_Performance Management-History Query-2G_KPI_Reporting_Template-DFBG6870-20251204091043.xlsx"
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

dxfs = wb._differential_styles

print("Found conditional formatting rules details:")
for cf in sheet.conditional_formatting:
    for rule in cf.cfRule:
        if rule.dxfId is not None:
            dxf = dxfs[rule.dxfId]
            if dxf.fill:
                print(f"Cells: {cf.sqref}, Formula: {rule.formula}, Operator: {rule.operator}")
                print(f"  BgColor: {dxf.fill.bgColor.rgb if hasattr(dxf.fill.bgColor, 'rgb') else dxf.fill.bgColor.theme}")
                print(f"  FgColor: {dxf.fill.fgColor.rgb if hasattr(dxf.fill.fgColor, 'rgb') else dxf.fill.fgColor.theme}")
