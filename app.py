import os
import smtplib
from email.message import EmailMessage

import pandas as pd
from pypdf import PdfReader, PdfWriter
from analyticskit import single_input, show_message, select_file, add_attachment


# ---------------------------------------------------------------
# Question 1: Get and validate the directory
# ---------------------------------------------------------------
required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx',
                  'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input('Please tell me the directory where the files are located').strip().strip('"\'')
    if directory[-1] != '/':
        directory = directory + '/'

    if os.path.isdir(directory):
        files_in_dir = os.listdir(directory)
        missing = [f for f in required_files if f not in files_in_dir]
        if not missing:
            break
        show_message('Error', 'Missing files: ' + ', '.join(missing) + '. Pick another directory.')
    else:
        show_message('Error', 'That directory does not exist. Pick another one.')


# ---------------------------------------------------------------
# Question 2: Product description and price dictionaries
# ---------------------------------------------------------------
descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for index, row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']


# ---------------------------------------------------------------
# Question 3: Read an order form
# ---------------------------------------------------------------
def get_order_data(invoice_num, filename, directory):
    order_df = pd.read_excel(directory + filename, header=None)

    date = pd.to_datetime(order_df.iloc[8, 2]).strftime('%m/%d/%Y')
    email = str(order_df.iloc[12, 2]).strip()
    name = str(order_df.iloc[13, 2]).strip()
    company = str(order_df.iloc[14, 2]).strip()
    address = str(order_df.iloc[15, 2]).strip()
    city = str(order_df.iloc[16, 2]).strip()
    state = str(order_df.iloc[17, 2]).strip()
    zipcode = str(order_df.iloc[18, 2]).strip()

    order_data = {
        'InvoiceNum': str(invoice_num),
        'Date': date,
        'CustomerName': name,
        'CustomerCompany': company,
        'CustomerAddress1': address,
        'CustomerAddress2': city + ', ' + state + ' ' + zipcode,
    }

    total = 0
    row = 23
    item = 0
    while row < len(order_df) and pd.notna(order_df.iloc[row, 1]):
        unit = order_df.iloc[row, 1]
        quantity = int(order_df.iloc[row, 2])
        price = product_prices[unit]
        amount = price * quantity
        total += amount

        order_data['Description' + str(item)] = product_descriptions[unit]
        order_data['Quantity' + str(item)] = str(quantity)
        order_data['Price' + str(item)] = f'${price:,.2f}'
        order_data['Amount' + str(item)] = f'${amount:,.2f}'

        row += 1
        item += 1

    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date


# ---------------------------------------------------------------
# Question 4: Fill out the invoice PDF
# ---------------------------------------------------------------
def create_invoice(invoice_num, order_data, directory):
    reader = PdfReader(directory + 'MasterInvoice.pdf')
    writer = PdfWriter()
    writer.append(reader)

    writer.update_page_form_field_values(writer.pages[0], order_data)

    with open(directory + 'Invoice ' + str(invoice_num) + '.pdf', 'wb') as output_file:
        writer.write(output_file)


# ---------------------------------------------------------------
# Question 5: Email the invoice
# ---------------------------------------------------------------
gmail_user = 'caleb.hilton@gmail.com'
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read().strip()


def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()


# ---------------------------------------------------------------
# Question 6: Process every order
# ---------------------------------------------------------------
filenames = os.listdir(directory)
filenames.sort()

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1

show_message('Done', 'All invoices have been created and emailed.')
