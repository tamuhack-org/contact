import pandas as pd
import vobject

#Load the excel file
excel_file_path = r'contacts_23_4.xlsx'  # CHANGE TO YOUR FILE PATH
df = pd.read_excel(excel_file_path)

#Strip leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

#Create a list to store the vCard entries
vcard_entries = []

#Loop through each row in the DataFrame
for index, row in df.iterrows():
    vcard = vobject.vCard()

    #Extract data from the DataFrame
    full_name = row['Name'].strip()  #Assuming 'Name' contains both first and last name
    phone_number = str(row['Phone Number']).replace('-', '')  #Remove hyphens
    email = row['Personal Email'].strip()
    birthday = str(row['Birthday(Required)'].date())  #Convert to string in format YYYY-MM-DD

    #Split full name into first and last names
    first_name, last_name = full_name.split(maxsplit=1)

    #Add data to the vCard
    vcard.add('VERSION').value = '3.0'
    vcard.add('FN').value = full_name
    vcard.add('N').value = vobject.vcard.Name(family=last_name, given=first_name)

    #Add TEL component without type argument
    tel_component = vcard.add('TEL')
    tel_component.value = phone_number
    tel_component.type_param = ['CELL', 'PREF', 'VOICE']

    vcard.add('EMAIL').value = email
    vcard.add('BDAY').value = birthday

    #Append the formatted vCard entry to the list
    vcard_entries.append(vcard.serialize())

#Save the vCard entries to a file
output_vcard_path = 'contacts.vcf'
with open(output_vcard_path, 'w') as f:
    f.write('\n'.join(vcard_entries))

print(f"vCard file '{output_vcard_path}' created successfully.")
