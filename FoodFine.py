#Wheee
from __future__ import print_function
from tabulate import tabulate
from tabula.io import read_pdf
import pandas as pd
import numpy as np
import datetime
import calendar
import os
import pygsheets
import tqdm
import time
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import subprocess

    
#reads table from pdf file
list_1 = read_pdf("/Users/domi_bryan/Food_Fine_Records/Current.pdf",pages="1") #address of pdf 
df = pd.DataFrame(np.concatenate(list_1))

#list_2 = read_pdf("/Users/domi_bryan/Food_Fine_Records/Current.pdf",pages="2") #address of pdf file
#df_2 = pd.DataFrame(np.concatenate(list_2))
#df_2.drop(df_2.head(1).index,inplace=True)
#df=pd.concat([df_1,df_2],ignore_index=True)

#Cleaning and Reformating
df = df.fillna('')
df.columns = df.iloc[0]
df = df[1:]
df.columns = df.columns.str.replace(' ', '_')
print(tabulate(df, showindex="false", tablefmt="fancy_grid"))  #Content Debunker
#print(df.columns) #Column Debunker
total_cost= df['Hall'].iloc[-1]
total_cost = total_cost.replace("$", "")
total_cost = float(total_cost)
#print(total_cost)

#Tail checker
tail_checker = df['S/N'].iloc[-1]
df.drop(df.tail(1).index,inplace=True)

df['Cost'] = df['Cost'].replace({'\$':''}, regex = True)
df['Cost'] = df['Cost'].astype(float)
df['Reason_for_Excuse/Remarks'].astype('string')
names = df.Full_Name.unique()
#print (names)

#Total Sum and Final Check
total_sum = 0
print("Food Fine per Person:")
for name in names:
    temp_sum = df.loc[df['Full_Name'] == name, 'Cost'].sum()
    print(" • {}: ${}".format(name, temp_sum)) #Yikes
    total_sum = total_sum + temp_sum
flagsend = False
if(df['S/N'].iloc[0]!='1'):
    print(df['S/N'][0])
    print("The table head got cut off. Readjust code.")
elif(tail_checker!='Total Hall 8'):
    #print(tail_checker)
    print("The table tail got cut off. Readjust code.")
elif(total_sum>total_cost):
    print("The total sum is ${}, ABOVE the total cost of ${} provided by the office. Manually check the issue.".format(total_sum, total_cost))
elif(total_sum<total_cost):
    print("The total sum is ${}, BELOW the total cost of ${} provided by the office. Manually check the issue.".format(total_sum, total_cost))
else:
    print("The total sum is ${}, the same as the total cost of ${}.\n\n".format(total_sum, total_cost))
    flagsend = True
if(flagsend==False):
    override = input("Override Anyway?[Y/N]")
    if(override == 'Y'):
        flagsend = True

#Waivers
while True:
    waiver_index = input("Any waivers? [Type 'N' without the quotation marks if no] ")
    if (waiver_index == 'N'):
        break
    temp_data = df.loc[df['S/N'] == waiver_index]
    df.drop(df.loc[df['S/N'] == waiver_index].index, inplace=True)
    print(tabulate(temp_data, showindex="false", tablefmt="fancy_grid") + "\nhas been removed from the record.\n\n")

names = df.Full_Name.unique()

tomorrow = datetime.date.today() + datetime.timedelta(days=1)
if(tomorrow.day<10):
    renameIndex = "0"+str(tomorrow.day) + "-" + calendar.month_abbr[tomorrow.month]
else:
    renameIndex = str(tomorrow.day) + "-" + calendar.month_abbr[tomorrow.month]
rfc_collect = datetime.datetime(tomorrow.year, tomorrow.month, tomorrow.day, 22, 0, 0, 000).isoformat() + 'Z'
rfc_pay = datetime.datetime(tomorrow.year, tomorrow.month, tomorrow.day+1, 22, 0, 0, 000).isoformat() + 'Z'
flagrename = input("Store files in archives?[Y/N] ")
if (flagrename == 'Y'):
    os.rename("/Users/domi_bryan/Food_Fine_Records/Current.pdf", "/Users/domi_bryan/Food_Fine_Records/Previous_Records/{}.pdf".format(renameIndex))
    print("File has been stored in the Previous_Records as {}.pdf".format(renameIndex))
    
    gc = pygsheets.authorize(service_file= "/Users/domi_bryan/pythonStuff/foodfinecredentials.json")
    sh = gc.open_by_key('redacted')
    try:
        sh.add_worksheet(renameIndex)
    except:
        pass
    wks_write = sh.worksheet_by_title(renameIndex)
    wks_write.clear('A1',None,'*')
    wks_write.set_dataframe(df, (1,1), encoding='utf-8', fit=True)
    wks_write.frozen_rows = 1

    wks_total = sh.worksheet_by_title('Total_Cost_Per_Person')
    wks_total.insert_cols(3, number=1, values = None, inherit=False)
    wks_total.update_value('D1', renameIndex)

    print("Resetting Weekly_Summary...")
    wks_checker = sh.worksheet_by_title('Weekly_Summary')
    wks_checker.update_value('E6', renameIndex)
    
    for i in tqdm.tqdm(range(2,16)):
        wks_checker.update_value('C{}'.format(i), False)

    print("Uploading data frame to Google Sheets...")
    for i in tqdm.tqdm(range(2,50)):
       formula = "=SUMIF(INDIRECT(\"\'\" & TEXT(D$1,\"dd-mmm\") & \"\'!$C$2:$C\"), $B{}, INDIRECT(\"\'\" & TEXT(D$1,\"dd-mmm\") & \"\'!$F$2:$F\"))".format(i)
       wks_total.update_value('D{}'.format(i), formula)

    wks_total.update_value('D51', "=SUM(D$2:D$49)")

    print("A copy of the dataframe has been uploaded to Google Sheets as '{}'.\n".format(renameIndex))

    print("Creating Reminders...\n")
    SCOPES = ['https://www.googleapis.com/auth/tasks']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('tasks', 'v1', credentials=creds)
        # Call the Tasks API
        collect = service.tasks().insert(tasklist='redacted', body={'title': 'Collect Food Fine', 'due': rfc_collect}).execute()
        collect = service.tasks().insert(tasklist='redacted', body={'title': 'Pay Food Fine to Boarding School', 'due': rfc_pay}).execute()
    except HttpError as err:
        print(err)
    print("Reminders Created.")

if(flagsend==True):
    for name in names:
        temp_sum= df.loc[df['Full_Name'] == name, 'Cost'].sum()
        message = "\n\nHi %s, I'm letting you know that you have *$%.2f* worth of food fine outstanding. Please have the exact change ready/paynow me by *Tomorrow, %d %s*. I will collect at night. \n\nThanks." % (name,temp_sum,tomorrow.day, calendar.month_abbr[tomorrow.month])
        print(message)
