import openpyxl, re

multiRegex = [r'Z\s\w+', r'Mobil\s\w+', r'Bunnings', r'Caltex\s\w+']
#r'Z\s\w+' matches Z fuel stations
#r'Mobil\s\w+ matches MOBIL fuel stations
#r'Bunnings' matches BUNNINGS
#r'Caltex\s\w+' matches CALTEX

wb = openpyxl.load_workbook('C:\\Users\\Dennis\\Desktop\\taxprac.xlsx')
sheet = wb['Sheet1']
for i in range(2, sheet.max_row):
    try:
        if (sheet['H'+str(i)].value) > 0:
            #sheet['K'+str(i)] = -1
            sheet['K'+str(i)] = 'WORKS'
        else:
            for k in range(len(multiRegex)):
                haRegex = re.compile(multiRegex[k])
                #mo1 = haRegex.search(sheet['F'+str(i)].value)
                mo1 = haRegex.search(sheet['F'+str(i)].value)
                try:
                    print(mo1.group())
                    #sheet['K'+str(i)] = 'WORKS'
                except:
                     continue
           
    except Exception as exc:
        continue

#wb.save('C:\\Users\\Dennis\\Desktop\\gst_copy.xlsx')
