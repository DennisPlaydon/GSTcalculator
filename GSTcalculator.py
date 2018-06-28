import openpyxl

wb = openpyxl.load_workbook('C:\\Users\\Dennis\\Desktop\\taxprac.xlsx')
sheet = wb['Sheet1']
for i in range(2, sheet.max_row):
    try:
        if 'Z Ormiston R' in sheet['F'+str(i)].value:
            #sheet['K'+str(i)] = 1
            sheet['K'+str(i)] = 'WORKS'
           
        elif (sheet['H'+str(i)].value) > 0:
            #sheet['K'+str(i)] = -1
            sheet['K'+str(i)] = 'WORKS'
    except Exception as exc:
        continue

wb.save('C:\\Users\\Dennis\\Desktop\\gst_copy.xlsx')
