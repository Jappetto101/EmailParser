# import openpyxl

import pandas as pd

filename = "emailoutput.txt"

contents = open(filename,'r').read().split('\n')

eDict = {'msgID':[], 'From':[],'Sent':[], 'To':[], 'Cc':[], 'Subject':[], 'Attachments':[], 'Body':[]}

i=0
msgCount = 0
body = ""
while i < len(contents):
    if contents[i].strip() == "":
        i+=1
        continue
    elif contents[i].strip()[0:5]=="From:":
        eDict['msgID'].append(msgCount+1)
        eDict['Body'].append("")
        if contents[i].strip()[0:5] =="From:":
            eDict['From'].append(contents[i].strip()[6:].strip())
            i+=1
        else:
            eDict['From'].append("")
        if contents[i].strip()[0:5] in ["Sent:", "Date:"]:
            eDict['Sent'].append(contents[i].strip()[6:].strip())
            i+=1
        else:
            eDict['Sent'].append("")
        if contents[i].strip()[0:3] =="To:":
            eDict['To'].append(contents[i].strip()[4:].strip())
            i+=1
        else:
            eDict['To'].append("")
        if contents[i].strip()[0:3] =="Cc:":
            eDict['Cc'].append(contents[i].strip()[4:].strip())
            i+=1
        else:
            eDict['Cc'].append("")
        if contents[i].strip()[0:8] =="Subject:":
            eDict['Subject'].append(contents[i].strip()[9:].strip())
            i+=1
        else:
            eDict['Subject'].append("")
        if contents[i].strip()[0:12] =="Attachments:":
            eDict['Attachments'].append(contents[i].strip()[13:].strip())
            i+=1
        else:
            eDict['Attachments'].append("")
        if msgCount>0:
            eDict['Body'][msgCount-1] = body        
        msgCount += 1
        body = ""
        continue
    elif contents[i].strip()[0:5] != "From:":
        body += contents[i]
        body += "\n"
        i+=1
        continue

df=pd.DataFrame(eDict)
df.to_excel('output.xlsx', index= 0)