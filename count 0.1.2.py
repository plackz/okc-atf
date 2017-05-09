import os
from pprint import pprint
from operator import itemgetter
import win32com.client
import pandas as pd

path = 'I:/Quality Control/After the Fact Documentation/ATF Sort'
mn = 0
custnameList=[]
custnumList=[]
custList=[]

folders = ([name for name in os.listdir(path)
            if os.path.isdir(os.path.join(path, name)) and name[0].isalpha()]) # get all directories 
for folder in folders:
    contents = os.listdir(os.path.join(path,folder)) # get list of contents
    if len(contents) > mn: # if greater than the limit, print folder and number of contents
            custList.append(folder)
            custnumList.append(sum([len(files) for r, d, files in os.walk(path+'/'+folder)]))
total=sum(custnumList)
custList=list(zip(custList, custnumList))
custList=sorted(custList, key=itemgetter(1),reverse=True)
print ('Total: '+ str(total))
pprint(custList)


# email portion
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)

mailToList = 'pzaffina@slb.com;jwheeler7@slb.com;jdodson4@slb.com'

newMail.Subject = "ATF Customer Count"

# use pandas to convert list to html table
labels = ['Customer', 'Count of ATF']
df = pd.DataFrame.from_records(custList, columns=labels)

body = '<html><body>' + 'Total: '+ str(total) + df.to_html(index=False) + '</body></html>'

newMail.HTMLBody = body
newMail.To  = mailToList
newMail.Send()
