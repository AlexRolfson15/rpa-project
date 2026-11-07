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

while True:
    directory = os.getcwd()+'/'
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message('ERROR, Please enter a valid directory.')

print(filenames)





df = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}

for index, row in df.iterrows():
    description = row['Description']
    number = row['Unit number']
    product_descriptions[number] = description
print(product_descriptions)

df1 = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Prices')
product_prices = {}

for index, row in df1.iterrows():
    price = row['Price per unit']
    number = row['Unit number']
    product_prices[number] = price
print(product_prices)





def get_order_data(invoice_num, filename, directory):
    order_details = pd.read_excel(f'{directory}{filename}')
    order_template = openpyxl.load_workbook(f'{directory}{filename}')
    sheet = order_template['OrderForm']
    name = sheet['C14'].value
    email = 'alex.rolfson15@gmail.com'
    date = sheet['C9'].value.strftime('%m/%d/%y')

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = sheet['C15'].value
    order_data['CustomerAddress2'] = f'{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}'

    total = 0

    unit = sheet['B24'].value
    price = product_prices[unit]
    order_data['Description0'] = product_descriptions[unit]
    order_data['Price0'] = f'${price:,.2f}'
    order_data['Quantity0'] = str(sheet['C24'].value)
    order_data['Amount0'] = f'{price * sheet['C24'].value:,.2f}'
    total = total + (price * sheet['C24'].value) 

    if pd.isna(sheet['B25'].value) == False:
        unit = sheet['B25'].value
        price = product_prices[unit]
        order_data['Description1'] = product_descriptions[unit]
        order_data['Price1'] = f'${price:,.2f}'
        order_data['Quantity1'] = str(sheet['C25'].value)
        order_data['Amount1'] = f'{price * sheet['C25'].value:,.2f}'
        total = total + (price * sheet['C25'].value) 

    if pd.isna(sheet['B26'].value) == False:
        unit = sheet['B26'].value
        price = product_prices[unit]
        order_data['Description2'] = product_descriptions[unit]
        order_data['Price2'] = f'${price:,.2f}'
        order_data['Quantity2'] = str(sheet['C26'].value)
        order_data['Amount2'] = f'{price * sheet['C26'].value:,.2f}'
        total = total + (price * sheet['C26'].value) 

    if pd.isna(sheet['B27'].value) == False:
        unit = sheet['B27'].value
        price = product_prices[unit]
        order_data['Description3'] = product_descriptions[unit]
        order_data['Price3'] = f'${price:,.2f}'
        order_data['Quantity3'] = str(sheet['C27'].value)
        order_data['Amount3'] = f'{price * sheet['C27'].value:,.2f}'
        total = total + (price * sheet['C27'].value) 

    if pd.isna(sheet['B28'].value) == False:
        unit = sheet['B28'].value
        price = product_prices[unit]
        order_data['Description4'] = product_descriptions[unit]
        order_data['Price4'] = f'${price:,.2f}'
        order_data['Quantity4'] = str(sheet['C28'].value)
        order_data['Amount4'] = f'{price * sheet['C28'].value:,.2f}'
        total = total + (price * sheet['C28'].value) 

    order_data['Total'] = f'{total:,.2f}'
    return order_data, name, email, date

get_order_data('8515', 'Order1.xlsx', directory)




def create_invoice(invoice_num, order_data, directory):
    update_dict = {}
    update_dict['InvoiceNum'] = order_data['InvoiceNum']
    update_dict['Date'] = order_data['Date']
    update_dict['Total'] = order_data['Total']
    update_dict['CustomerAddress1'] = order_data['CustomerAddress1']
    update_dict['CustomerName'] = order_data['CustomerName']
    update_dict['CustomerCompany'] = order_data['CustomerCompany']
    update_dict['CustomerAddress2'] = order_data['CustomerAddress2']
    update_dict['Description0'] = order_data['Description0']
    update_dict['Price0'] = order_data['Price0']
    update_dict['Quantity0'] = order_data['Quantity0']
    update_dict['Amount0'] = order_data['Amount0']
    if 'Description1' in order_data:
        update_dict['Description1'] = order_data['Description1']
        update_dict['Price1'] = order_data['Price1']
        update_dict['Quantity1'] = order_data['Quantity1']
        update_dict['Amount1'] = order_data['Amount1']
    if 'Description2' in order_data:
        update_dict['Description2'] = order_data['Description2']
        update_dict['Price2'] = order_data['Price2']
        update_dict['Quantity2'] = order_data['Quantity2']
        update_dict['Amount2'] = order_data['Amount2']
    if 'Description3' in order_data:
        update_dict['Description3'] = order_data['Description3']
        update_dict['Price3'] = order_data['Price3']
        update_dict['Quantity3'] = order_data['Quantity3']
        update_dict['Amount3'] = order_data['Amount3']
    if 'Description4' in order_data:
        update_dict['Description4'] = order_data['Description4']
        update_dict['Price4'] = order_data['Price4']
        update_dict['Quantity4'] = order_data['Quantity4']
        update_dict['Amount4'] = order_data['Amount4']
    create_pdf('MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', update_dict)
create_invoice(8515, order_data, directory)




gmail_user = "alex.rolfson15@gmail.com"

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTeech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.
    
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = {email}
    message['To'] = {email}
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory+f'/Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)  

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()   

send_email('8515', 'Thomas Holt', 'alex.rolfson15@gmail.com', '11-13-2023', directory)




filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and '.xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, order_data['CustomerName'], 'alex.rolfson15@gmail.com', order_data['Date'], directory)
        invoice_num = invoice_num + 1





