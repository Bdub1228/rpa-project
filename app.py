# INSERT ANSWER FOR QUESTION 1 HERE
# directory = os.getcwd()
# 2nd Part of Question 1
while True:    
    directory = select_file('Please select one of the files `ProductData.xlsx`, `MasterInvoice.pdf`, `Order1.xlsx`, `Order2.xlsx`, `Order3.xlsx`, or `Order4.xlsx`.')
    directory = directory.split('/')
    del directory[-1]
    directory = '\\'.join(directory)
    if directory.endswith('\\') == True:
        pass
    else:
        directory = directory + '\\'
# 3rd part of question
    listeddirectory = os.listdir(directory)
    if 'ProductData.xlsx' in listeddirectory and 'MasterInvoice.pdf' in listeddirectory and 'Order1.xlsx' in listeddirectory and 'Order2.xlsx' in listeddirectory and 'Order3.xlsx' in listeddirectory and 'Order4.xlsx' in listeddirectory:
        break
    else:
        show_message('Error', 'Please re-enter the directory of the files `ProductData.xlsx`, `MasterInvoice.pdf`, `Order1.xlsx`, `Order2.xlsx`, `Order3.xlsx`, and `Order4.xlsx` files.')

print(directory)
# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
descdf = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')


for index,row in descdf.iterrows():
    description = row['Description']
    unit_num = row['Unit number']
    product_descriptions[unit_num] = description


pricedf = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')



for index,row in pricedf.iterrows():
    price = row['Price per unit']
    unit_num = row['Unit number']
    product_prices[unit_num] = price

print(product_prices)
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    invoice_num = int(invoice_num)
    order_data = {}
    orderwkbook = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet = orderwkbook['OrderForm']
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    CustomerAddress1 = selected_sheet['C16'].value
    CustomerName = selected_sheet['C14'].value
    CustomerCompany = selected_sheet['C15'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = int(selected_sheet['C19'].value)
    if zipcode < 10000:
        zipcode = '0' + str(zipcode)
    # zipcode = int(zipcode)
    CustomerAddress2 = f'{city}, {state} {zipcode}'
    CustomerEmail = selected_sheet['C13'].value
    ##ITEM 1--------------------------------------------------
    amount0=0
    if selected_sheet['C24'].value is not None:
        quantity0 = selected_sheet['C24'].value
        unit_number0 = selected_sheet['B24'].value
        price0 = product_prices[unit_number0]
        amount0 = (price0) * (quantity0)
        amount0str = f'${amount0:,.2f}'
        price0str = f'${price0:,.2f}'
        description0 = product_descriptions[unit_number0]
        order_data['Price0'] = price0str
        order_data['Quantity0'] = quantity0
        order_data['Amount0'] = amount0str
        order_data['Description0'] = description0
    ##ITEM 2--------------------------------------------------
    amount1=0
    if selected_sheet['C25'].value is not None:
        quantity1 = selected_sheet['C25'].value
        unit_number1 = selected_sheet['B25'].value
        price1 = product_prices[unit_number1]
        amount1 = (price1) * (quantity1)
        amount1str = f'${amount1:,.2f}'
        price1str = f'${price1:,.2f}'
        description1 = product_descriptions[unit_number1]
        order_data['Price1'] = price1str
        order_data['Quantity1'] = quantity1
        order_data['Amount1'] = amount1str
        order_data['Description1'] = description1
    ##ITEM 3--------------------------------------------------
    amount2=0
    if selected_sheet['C26'].value is not None:
        quantity2 = selected_sheet['C26'].value
        unit_number2 = selected_sheet['B26'].value
        price2 = product_prices[unit_number2]
        amount2 = (price2) * (quantity2)
        amount2str = f'${amount2:,.2f}'
        price2str = f'${price2:,.2f}'
        description2 = product_descriptions[unit_number2]
        order_data['Price2'] = price2str
        order_data['Quantity2'] = quantity2
        order_data['Amount2'] = amount2str
        order_data['Description2'] = description2
    ##ITEM 4--------------------------------------------------
    amount3=0
    if selected_sheet['C27'].value is not None:
        quantity3 = selected_sheet['C27'].value
        unit_number3 = selected_sheet['B27'].value
        price3 = product_prices[unit_number3]
        amount3 = (price3) * (quantity3)
        amount3str = f'${amount3:,.2f}'
        price3str = f'${price3:,.2f}'
        description3 = product_descriptions[unit_number3]
        order_data['Price3'] = price3str
        order_data['Quantity3'] = quantity3
        order_data['Amount3'] = amount3str
        order_data['Description3'] = description3
    ##ITEM 5--------------------------------------------------
    amount4=0
    if selected_sheet['C28'].value is not None:
        quantity4 = selected_sheet['C28'].value
        unit_number4 = selected_sheet['B28'].value
        price4 = product_prices[unit_number4]
        amount4 = int(price4) * int(quantity4)
        amount4str = f'${amount4:,.2f}'
        price4str = f'${price4:,.2f}'
        description4 = product_descriptions[unit_number4]
        order_data['Price4'] = price4str
        order_data['Quantity4'] = quantity4
        order_data['Amount4'] = amount4str
        order_data['Description4'] = description4
    
    ## PUT INFO INTO DICTIONARY ---------------
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date
    order_data['Total'] = f'${(amount0) + (amount1) + (amount2) + (amount3) + (amount4):,.2f}'
    order_data['CustomerAddress1'] = CustomerAddress1
    order_data['CustomerName'] = CustomerName
    order_data['CustomerCompany'] = CustomerCompany
    order_data['CustomerAddress2'] = CustomerAddress2
    return order_data, CustomerName, CustomerEmail, date
    
    
order_data, CustomerName, CustomerEmail, date = get_order_data(8515,'Order3.xlsx', directory)    
print(order_data)
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    invoice_num = int(invoice_num)
    create_pdf(directory + 'MasterInvoice.pdf',directory + f'Invoice {invoice_num}.pdf',order_data)

create_invoice(8515,order_data,directory)
# INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    invoice_num = int(invoice_num)
    # Enter your gmail user name and password

    gmail_user = "brockwright28@gmail.com"


    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name

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
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+ f'/Invoice {invoice_num}.pdf' 
    # path_to_file = directory + f'/Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
        
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,CustomerName, CustomerEmail, date, directory)
# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, CustomerName, CustomerEmail, date = get_order_data(invoice_num,filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, CustomerName, CustomerEmail, date, directory)
        invoice_num = invoice_num + 1