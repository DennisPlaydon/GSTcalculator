import re

list1 = ['Z batny', 'C crabs', 'D news', 'Z ormiston']

haRegex = re.compile(r'Z\s\w+')
for i in list1:
    try:
        mo1 = haRegex.search(i)
        print(mo1.group())
    except Exception as exc:
        continue
