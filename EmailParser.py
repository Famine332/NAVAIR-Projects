import os
path = (os.path.dirname(os.path.abspath(__file__)))
print(path)
xpath = path.replace('EmailParser.py', ' ')
xpath = r'{}'.format(xpath)
print(xpath)
xfiles = os.listdir(path)
print(xfiles)

for s in xfiles:
    if 'txt' in s:
        theFile = s
thexFile = r'{}'.format(theFile)

if theFile[4:6] == '01':
    month = 'January'
if theFile[4:6] == '02':
    month = 'February'
if theFile[4:6] == '03':
    month = 'March'
if theFile[4:6] == '04':
    month = 'April'
if theFile[4:6] == '05':
    month = 'May'
if theFile[4:6] == '06':
    month = 'June'
if theFile[4:6] == '07':
    month = 'July'
if theFile[4:6] == '08':
    month = 'August'
if theFile[4:6] == '09':
    month = 'September'
if theFile[4:6] == '10':
    month = 'October'
if theFile[4:6] == '11':
    month = 'November'
if theFile[4:6] == '12':
    month = 'December'
month = str(month)
print(month)

day = theFile[6:8]
day = str(day)
if '0' in day:
    day = day.replace('0', '')
print(day)

check = month+' '+day
print(check)

namefile = xpath + '\\' + thexFile
print('thes is the namefile: ', namefile)

with open(namefile, "r") as txtfile:
    data = txtfile.read()
    emailList = data.split('From:')
    theList = []
    for email in emailList:
        if check in email:
            theList.append(email)

theList = list(dict.fromkeys(theList))

with open(xpath+'\\'+'Reading File for {}.txt'.format(check), "w") as wfile:
    for email in theList:
        wfile.write(email)
        wfile.write('\n')
        wfile.write("********* NEXT EMAIL **********")
        wfile.write('\n')
        wfile.write('\n')
