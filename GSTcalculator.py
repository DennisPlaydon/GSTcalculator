import openpyxl, re

multiRegex = [r'Z\s\w+', r'Mobil\s\w+', r'Bunnings', r'Caltex\s\w+', r'Countdown', r'Bp\sConnect', r'Edl\s\w+', r'Pack\sN\sSave', r'New\sWorld', r'Inex\s']
#r'Z\s\w+' matches Z fuel stations
#r'Mobil\s\w+ matches MOBIL fuel stations
#r'Bunnings' matches BUNNINGS
#r'Caltex\s\w+' matches CALTEX
#r'Countdown' matches COUNTDOWN
#r'Edl\s\w+' matches EDL fasteners

wb = openpyxl.load_workbook('C:\\Users\\Dennis\\Desktop\\taxprac.xlsx')
sheet = wb['Sheet1']
for i in range(2, sheet.max_row):
    try:
        if (sheet['H'+str(i)].value) > 0:
            continue
            #sheet['K'+str(i)] = -1
            #sheet['K'+str(i)] = 'WORKS'
        else:
            for k in range(len(multiRegex)):
                haRegex = re.compile(multiRegex[k])
                #mo1 = haRegex.search(sheet['F'+str(i)].value)
                for rowOfCellObjects in sheet['D'+str(i):'F'+str(i)]:
                    for cellObj in rowOfCellObjects:
                        mo1 = haRegex.search(cellObj.value)
                        try:
                            mo1.group()
                            sheet['K'+str(i)] = 'WORKS'
                        except:
                            continue    
 
                '''
                try:
                    mo1 = haRegex.search(sheet['D'+str(i): 'F'+str(i)])
                    print(mo1.group())
                    sheet['K'+str(i)] = 'WORKS'
                except:
                     continue
                '''
    except Exception as exc:
        continue

wb.save('C:\\Users\\Dennis\\Desktop\\gst_copy.xlsx')

'''
TODO Fix error with INEX metals. Only some of the data changes

'''