# Insert your python code here

#imports 
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# question 1
while True:
    directory = single_input("Please input the directory where the file is located.")
    directory = directory.strip().strip("\"'") + '/'

    if os.path.isdir(directory):
        filenames = os.listdir(directory)

        has_files = ('ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames
            and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames)

        if has_files:
            break
        else:
            show_message("Missing Files", "The folder doesn't contain all the required files.")
    else:
        show_message("Invalid Folder","The folder doesn't exist. Please give another.")


# question 2
df = pd.read_excel('ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for i, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df1 = pd.read_excel('ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for i, row in df1.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

# question 3
def get_order_data(invoice_num, filename, directory):
    path = directory + filename
    workbook = openpyxl.load_workbook(path)
    sheet = workbook['OrderForm']

    my_dict = {}
    my_dict['InvoiceNum'] = str(invoice_num)
    my_dict['Date'] = sheet['C9'].value.strftime('%m/%d/%Y')
    my_dict['CustomerAddress1'] = sheet['C16'].value
    my_dict['CustomerName'] = sheet['C14'].value
    my_dict['CustomerCompany'] = sheet['C15'].value
    my_dict['CustomerAddress2'] = f"{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}"
    
    
    if not pd.isna(sheet['B24'].value):
        my_dict['Description0'] = str(product_descriptions[sheet['B24'].value])
        my_dict['Price0'] = product_prices[sheet['B24'].value]
        my_dict['Quantity0'] = str(sheet['C24'].value)
        my_dict['Amount0'] = f"${product_prices[sheet['B24'].value] * sheet['C24'].value:,.2f}"

    if not pd.isna(sheet['B25'].value):    
        my_dict['Description1'] = str(product_descriptions[sheet['B25'].value])
        my_dict['Price1'] = str(product_prices[sheet['B25'].value])
        my_dict['Quantity1'] = str(sheet['C25'].value)
        my_dict['Amount1'] = f"${product_prices[sheet['B25'].value] * sheet['C25'].value:,.2f}"
    
    if not pd.isna(sheet['B26'].value):
        my_dict['Description2'] = str(product_descriptions[sheet['B26'].value])
        my_dict['Price2'] = str(product_prices[sheet['B26'].value])
        my_dict['Quantity2'] = str(sheet['C26'].value)
        my_dict['Amount2'] = f"${product_prices[sheet['B26'].value] * sheet['C26'].value:,.2f}"

    if not pd.isna(sheet['B27'].value):
        my_dict['Description3'] = str(product_descriptions[sheet['B27'].value])
        my_dict['Price3'] = str(product_prices[sheet['B27'].value])
        my_dict['Quantity3'] = str(sheet['C27'].value)
        my_dict['Amount3'] = f"${product_prices[sheet['B27'].value] * sheet['C27'].value:,.2f}"
    
    if not pd.isna(sheet['B28'].value):
        my_dict['Description4'] = str(product_descriptions[sheet['B28'].value])
        my_dict['Price4'] = str(product_prices[sheet['B28'].value])
        my_dict['Quantity4'] = str(sheet['C28'].value)
        my_dict['Amount4'] = f"${product_prices[sheet['B28'].value] * sheet['C28'].value:,.2f}"
    
    name = my_dict['CustomerName']
    email = sheet['C13'].value
    date = my_dict['Date']


    total = 0

    for row in [24,25,26,27,28]:
        unit_num = sheet[f'B{row}'].value

        if pd.isna(unit_num) == False:
            quantity = sheet[f'C{row}'].value
            total = total + product_prices[unit_num] * quantity

        my_dict['Total'] = f'${total:,.2f}'

    
    return my_dict, name, email, date

# question 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + "MasterInvoice copy.pdf", directory + f"Invoice {invoice_num}.pdf", order_data)

#question 5
gmail_user = "boston13billings@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    recipients = email
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
        
Attached is your invoice dated {date}

Thank you for your business.
    
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f"Invoice {invoice_num}.pdf"
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

# question 6
filenames = os.listdir(directory)
filenames = sorted(filenames)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1