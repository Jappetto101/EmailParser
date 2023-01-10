# pandas and openpyxl libraries required in order to run
import pandas as pd

# txt file should be placed in the root folder where the emailParser.py
# file is located. It should be named "emailoutput.txt"
filename = "emailoutput.txt"

contents = open(filename,'r').read().split('\n')

# These are the columns being output to an excel file
eDict = {'msgID':[], 'From':[],'Sent':[], 'To':[], 'Cc':[], 'Subject':[], 'Attachments':[], 'Body':[]}

i=0
msgCount = 0
body = ""

# Looping through each line of the txt file to figure out where to put it
while i < len(contents):
    # removing all empty lines
    if contents[i].strip() == "":
        i+=1
        continue
    # The "From:" header initializes a new message ID and checks for
    # other header information like "Sent", "To", "Cc", "Attachments".
    # If no header information is found, an empty string is inserted.
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
    # Once header information is found, the body of the message is collated
    elif contents[i].strip()[0:5] != "From:":
        body += contents[i]
        body += "\n"
        i+=1
        continue
    
# Outputs result of code into xlsx
df=pd.DataFrame(eDict)
df.to_excel('output.xlsx', index= 0)