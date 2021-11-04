import pandas as pd
import re
import json

f = open('../task1/task_input_list.json',)
 
# returns JSON object as
# a dictionary
data = json.load(f)

phone_num = []
emails = []
deposits_trans = []
withdrawal_trans = []

deposits = ["Deposits and other additions", "Counter Credit", "Total deposits and other additions"]
 
# Iterating through the json
# list
for idx,line in enumerate(data):
    # parsing the string, removing the characters
    line = re.sub('[!@#$,]', '', line)
    # finding the phone numbers using regular expression
    lis = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', line)
    # finding email match
    email_match = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', line)
    #appending mobile numbers to the list
    for num in lis:
        phone_num.append(num)
    # if there is any mail match, then append emails to the list
    if email_match == None:
        email_na = 1
    else:
        for email in email_match:
            emails.append(email)
    
    # finding withdrawal amounts list
    # since the withdrawal amount is shown by a minus sign, check for the
    # minus sign in the first index
    if line[0] == '-':
        withdrawal_trans.append(line)
    
    # if line is one of the defined parameters, then add the next value
    # to the deposit transaction list
    if line in deposits:
        dep = re.sub('[!@#$,]', '', data[idx+1])
        deposits_trans.append(dep)
    

if len(emails) == 0:
    emails.append("NA")

if "Date" in deposits_trans:
    deposits_trans.remove("Date")

#removing duplicates
deposit_list = list(set(deposits_trans))
withdrawal_list = list(set(withdrawal_trans))

#converting the string list to int list
deposit_list = [float(i) for i in deposit_list]
withdrawal_list = [float(i) for i in withdrawal_list]

print("List of phone numbers :",phone_num)
print("list of emails :",emails)

print("List of Deposit transaction", deposit_list)

print("List of withdrawal transaction", withdrawal_list)

max_amount = max(deposit_list)
min_amount = min(withdrawal_list)
website = "bankingservices.com"

output_lists = []
website = ['website', website]
output_lists.append(website)
emails = ['emails', emails]
output_lists.append(emails)
phone_num = ['phone_number', phone_num]
output_lists.append(phone_num)
minimum = ['min amount', min_amount]
maximum = ['max amount', max_amount]
output_lists.append(minimum)
output_lists.append(maximum)


df = pd.DataFrame(output_lists, columns=['Key', 'Value'])
print(df)

#convert data frame to excel
df.to_excel("task1_output.xlsx")





