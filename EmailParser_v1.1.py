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

check = month+' '+day
print(check)

namefile = xpath + '\\' + thexFile
print('thes is the namefile: ', namefile)

with open(namefile, "r") as txtfile:
    data = txtfile.read()

emails = data.split('From:')
final_emails = []
FINAL_emails = []
email_dict = {}

for email in emails:
    final_emails.append(email)

# This next line removes duplicate emails by leveraging dictionary key
# functionality.
final_emails = list(dict.fromkeys(final_emails))

for email in final_emails:
    # We'll check to see if any blank emails made it in; if so remove
    try:
        if len(email) < 20:
            final_emails = final_emails.remove(email)
    except:
        pass
    # The following lines are to clean up spacing
    email = email.replace('\n\n\n\n', '\n')
    email = email.replace('\n\n\n', '\n')
    email = email.replace('\n \n \n \n', '\n')
    email = email.replace('\n \n \n', '\n')
    email = email.replace('\n\n \n\n \n\n', '\n\n')
    email = email.replace('\n\n \n\n', '\n\n')
    # We previously split the string on 'From:', so we're cleaning it up and
    # adding it back in here
    if len(email) > 20:
        email = '\nFrom: '+email
    # The following lines parse out the email subject
    # First find the index of `Subject:`
    subj_index = email.find('Subject:')
    # Find the index of the colon after `Subject`
    colon_index = email.find(':', subj_index)
    # Find the index of the newline call at the end of the Subject
    new_line_index = email.find('\n', colon_index)
    # Define the substring for the email Subject
    Subject = email[colon_index+1:new_line_index]
    # Now I'll clean up RE:, FW:, [Non-DoD Source], etc prefixes so we can bin
    # same Subjects
    if 'RE:' in Subject:
        Subject = Subject.replace('RE: ', '')
    if 'FW:' in Subject:
        Subject = Subject.replace('FW: ', '')
    if '[Non-DoD Source]' in Subject:
        Subject = Subject.replace('[Non-DoD Source] ', '')
    if '[External]' in Subject:
        Subject = Subject.replace('[External] ', '')
    if '\t' in Subject:
        Subject = Subject.replace('\t', '')
    Subject = Subject.strip()

    # Now that Subjects are cleaned, we'll group the emails by Subject through
    # a dictionary
    if Subject in email_dict:
        email_dict[Subject].append(email)
    else:
        email_dict[Subject] = [email]

with open(xpath+'\\'+'Reading File for {}.txt'.format(check), "w") as wfile:
    # This section cycles through the dictionary to write the emails in order
    # by Subject
    for every in email_dict.values():
        for each in every:
            wfile.write(each)
            wfile.write('\n')
            wfile.write("********* NEXT EMAIL **********")
            wfile.write('\n')
            wfile.write('\n')
print('done')
