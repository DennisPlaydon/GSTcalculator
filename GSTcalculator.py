import openpyxl, re
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell

multiRegex = [
    r'Z\s\w+', r'Mobil\s\w+', r'Bunnings', r'Caltex\s\w+', r'Countdown', r'Bp\sConnect', r'Edl\s\w+', \
     r'Pak\s\w+', r'New\sWorld', r'Inex\s', r'\w+\sEnergy', r'Pallet\sPackaging', r'Vulcan', r'Woodmart',\
      r'Workstore', r'Spraywell\s\w+', r'Tga\s\w+', r'Kmart', r'Spark\sNz']
#r'Z\s\w+' matches Z fuel stations
#r'Mobil\s\w+ matches MOBIL fuel stations
#r'Bunnings' matches BUNNINGS
#r'Caltex\s\w+' matches CALTEX
#r'Countdown' matches COUNTDOWN
#r'Edl\s\w+' matches EDL fasteners

zeroList = ['Runchun Wang', 'Cplaydon', 'Candice Playdon', 'Dennis Playdon']
redFill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

total_row = 0
wb = openpyxl.load_workbook('C:\\Users\\Dennis\\Desktop\\taxprac.xlsx')
sheet = wb['Sheet1']
sheet['L1'] = 'Payments'
sheet['M1'] = 'Receipts'
for i in range(2, sheet.max_row):
    try:
        if (sheet['H'+str(i)].value) > 0:
            #sheet['K'+str(i)] = -1
            sheet['K'+str(i)] = 'RECEIPT'
            sheet['M'+str(i)] = sheet['H'+str(i)].value
        elif sheet['D'+str(i)].value in zeroList:
            sheet['K'+str(i)] = 0
        else:
            for k in range(len(multiRegex)):
                haRegex = re.compile(multiRegex[k])
                #mo1 = haRegex.search(sheet['F'+str(i)].value)
                for rowOfCellObjects in sheet['D'+str(i):'F'+str(i)]:
                    for cellObj in rowOfCellObjects:
                        mo1 = haRegex.search(cellObj.value)
                        try:
                            mo1.group()
                            sheet['K'+str(i)] = 'PAYMENT'
                            sheet['L'+str(i)] = sheet['H'+str(i)].value
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
        if sheet['K'+str(i)].value == None:
            sheet['K'+str(i)].fill = redFill
        total_row += 1
        
    except Exception as exc:
        continue

sheet['M' + str(total_row+3)] = '=SUM(M2:M{0})'.format(str(total_row+1))
sheet['L' + str(total_row+3)] = '=SUM(L2:L{0})'.format(str(total_row+1))
print(sheet['M191'].value)
wb.save('C:\\Users\\Dennis\\Desktop\\gst_copy.xlsx')
'''
TODO 

'''