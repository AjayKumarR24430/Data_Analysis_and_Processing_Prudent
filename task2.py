import pandas as pd 
import json

f1= pd.read_excel('../task2_data_processing.xlsx', 'Conflicting Roles')
f2= pd.read_excel('../task2_data_processing.xlsx', 'Access Listing - Validation')
f3= pd.read_excel('../task2_data_processing.xlsx', 'Invoice Validation')

# get role ids
role_id = list(f1['Conflict Role ID'])
print(role_id)

# get users
users_in_access_list = list(set(f2['User']))

users_in_invoice_list = list(set(f3['User']))

a = set(users_in_access_list)  
b = set(users_in_invoice_list)  

print(users_in_access_list)
  
if a == b:  
    print("The list1 and list2 are equal")  
else:  
    print("The list1 and list2 are not equal") 

user_difference = list(a-b)

print(user_difference)
# all users in sheet 3 are present in sheet 2
# only user Q is missing in sheet 3 which is present in sheet 2

# # checking authorization in sheet 2
# f2['Authorised Yes/No'] = (f2['A/P Invoice'] == 'Full Authorization' and f2['Adding Customer/Vendor Master Data'] == 'Full Authorization' and
#             f2['Adding Lead BP'] == 'Full Authorization' and f2['Remove Business Partner'] == 'Full Authorization')

f2.loc[(f2['A/P Invoice'] == 'Full Authorization') & (f2['Adding Customer/Vendor Master Data'] == 'Full Authorization') &
            (f2['Adding Lead BP'] == 'Full Authorization') & (f2['Remove Business Partner'] == 'Full Authorization'), 'Authorised Yes/No'] = 'Yes'
f2['Authorised Yes/No'] = f2['Authorised Yes/No'].fillna('No')

role_dict = {"Adding Customer/Vendor Master Data": "CRID001",
            "Adding Lead BP" : "CRID002",
            "Remove Business Partner" : "CRID003"
}

f2['Conflicting Rules'] = ''

for index, row in f2.iterrows():
    if row['Adding Customer/Vendor Master Data'] == 'No Authorization':
        row['Conflicting Rules'] += role_dict["Adding Customer/Vendor Master Data"] + ', '
    if row['Adding Lead BP'] == 'No Authorization':
        row['Conflicting Rules'] += role_dict["Adding Lead BP"] + ', '
    if row['Remove Business Partner'] == 'No Authorization':
        row['Conflicting Rules'] += role_dict["Remove Business Partner"] + ', '
    else:
        row['Conflicting Rules'] = 'None'

f2 = f2[f2.User != 'User Q']
f2 = f2.reset_index(drop=True)

new_df =f2[['User','Department', 'Authorised Yes/No', 'Conflicting Rules']]

task2_output = pd.merge(f3, new_df, on='User')

#convert data frame to excel
task2_output.to_excel("task2_output.xlsx")



