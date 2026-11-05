# Insert your python code here


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

# question 1 *******************************************

while True:
    directory = single_input('Please paste the file path to the folder with the product data, master invoice, and orders are located.')
    if directory[-1] != '\\':
        directory = directory + '\\'
    try:
        filenames = os.listdir(directory)
    except:
        filenames = []
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        show_message('Success','Good job! Your directory is valid and ready to go <3')
        break
    else:
        show_message('Error','Invalid directory input. Please try again <3')
    
directory

# question 2 *******************************************

product_descriptions = {}
product_prices = {} 

descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index,row in descriptions_df.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description

prices_df = pd.read_excel(directory + 'ProductData.xlsx','Prices')
for index,row in prices_df.iterrows():
    unit_number2 = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_number2] = price
    

# question 3 *******************************************

def get_order_data(invoice_num,filename,directory):

    orderform = openpyxl.load_workbook(directory+filename) 
    selected_sheet = orderform['OrderForm']

    
    orderdate = selected_sheet['C9'].value
    orderdate = orderdate.strftime('%m/%d/%y')
    custname = selected_sheet['C14'].value
    custemail = selected_sheet['C13'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = orderdate 
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip = str(selected_sheet['C19'].value)
    order_data['CustomerAddress2'] = f'{city}, {state} {zip}'

    total = 0

    if selected_sheet['B24'].value is not None:
        
        unitnum0 = selected_sheet['B24'].value
        description0 = product_descriptions[unitnum0]
        order_data['Description0'] = description0
        price0 = f'${product_prices[unitnum0]:,.2f}'
        order_data['Price0'] = price0
        quantity0 = str(selected_sheet['C24'].value)
        order_data['Quantity0'] = quantity0
        amount0 = round(selected_sheet['C24'].value * product_prices[unitnum0],2)
        total = total + amount0
        order_data['Amount0'] = f'${amount0:,.2f}' 


    if selected_sheet['B25'].value is not None:

        unitnum1 = selected_sheet['B25'].value
        description1 = product_descriptions[unitnum1]
        order_data['Description1'] = description1
        price1 = f'${product_prices[unitnum1]:,.2f}'
        order_data['Price1'] = price1
        quantity1 = str(selected_sheet['C25'].value)
        order_data['Quantity1'] = quantity1
        amount1 = round(selected_sheet['C25'].value * product_prices[unitnum1],2)
        total = total + amount1
        order_data['Amount1'] = f'${amount1:,.2f}' 
        

    if selected_sheet['B26'].value is not None:
        unitnum2 = selected_sheet['B26'].value
        description2 = product_descriptions[unitnum2]
        order_data['Description2'] = description2
        price2 = f'${product_prices[unitnum2]:,.2f}'
        order_data['Price2'] = price2
        quantity2 = str(selected_sheet['C26'].value)
        order_data['Quantity2'] = quantity2
        amount2 = round(selected_sheet['C26'].value * product_prices[unitnum2],2)
        total = total + amount2
        order_data['Amount2'] = f'${amount2:,.2f}' 
        
    
    order_data['Total'] = f'${total:,.2f}'

    return order_data, custname, custemail, orderdate

# question 4 *******************************************

def create_invoice(invoice_num,order_data,directory):

    masterpdf = directory+'MasterInvoice.pdf'
    topath = directory+f'Invoice {invoice_num}.pdf'
    
    create_pdf(masterpdf, topath, order_data[0]) 


create_invoice(8515,order_data,directory)

# question 5 *******************************************

def send_email(invoice_num,name,email,date,directory):
    
    # Enter your gmail user name and password
    gmail_user = "amygwarnick@gmail.com"
    # gmail_password_path = select_file('Please select the gmail password file')
    # gmail_password = open(gmail_password_path).read()

    # List of recipients
    # inserted by the function

    # Email subect

    subject = f'Your order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'\\Invoice {invoice_num}.pdf'
    

    message = add_attachment(path_to_file, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

# ******** test per instructions: *************

test = send_email(8515,'Thomas Holt','amygwarnick@gmail.com','11/13/2023',directory)
        

# question 6 *******************************************

filenames = os.listdir(directory)
print(filenames)

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

invoice_num = 8515

for file in filenames:
    
    if 'Order' in file and 'xlsx' in file:
        
        orderdata = get_order_data(invoice_num,file,directory)
        create_invoice(invoice_num, orderdata, directory)
        name = orderdata[1]
        email = order_data[2]
        date = order_data[3]
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num + 1
