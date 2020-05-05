#imports
import time
from datetime import date
from reportlab.platypus import tables, SimpleDocTemplate,Table,Paragraph,Spacer,TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import inch
from num2words import num2words
import pandas as pd

#calculate tax_table and total amount
amount=0
tax_5 = 0
tax_18 = 0
total_pieces = 0
total_amount = 0
def calculate_table(bill_data):

    #tax_percentage = {6307:0.05,3921:0.18,40159100}
    global amount
    global tax_5
    global tax_18
    global total_pieces
    global total_amount
    for i, row in enumerate(bill_data):
        if i!= 0:
            amount = amount +row[7]
            total_pieces = total_pieces + row[3]
            if row[6]==5 :
                tax_5 = tax_5 + 0.05*row[7]
            elif row[6]==18 :
                tax_18 = tax_18 + 0.18*row[7]

    total_amount = amount + tax_18 + tax_5

# Python program to convert a list
# of character
def convert(s):
    # initialization of string to ""
    new = ""
    # traverse in the string
    for x in s:
        new += x
    # return string
    return new

#calculate amount in words
def amount_in_words(n):
    amount_word = num2words(n, lang='en_IN').title()
    amount_word = list(amount_word)

    for i, letter in enumerate(amount_word):
        if amount_word[i] == '-':
            amount_word[i] = ' '
        elif amount_word[i] == ',':
            amount_word[i] = ''

    amount_word_final = 'INR '+convert(amount_word)+' Rupees Only'
    return amount_word_final


#assign data
#static data
invoice_type = 'Proforma Invoice'
company_name =  'Orion Corp'
company_address = 'P-2/03, Tower 3B, Purvanchal Silver City -2\nSector Pi-2, Greater Noida\nIndia, 201308, Phone:0120-4543418'
warehouse_address = 'abcd'
bank_details = 'efgh'

company_details = company_name+'\n'+company_address+'\n'+warehouse_address+'\n'+bank_details

#retieving data from excel filename
#fake data



xl_file = pd.read_excel('/Users/Sanjay/Desktop/OrionCorp/bill_new1.xlsx',0)
bill_data = xl_file.values.tolist()
bill_data_head_row = ['SI','Description of Goods','HSN/SAC','Quantity','Rate','per','GST%','Amount']
bill_data.insert(0,bill_data_head_row)
calculate_table(bill_data)
bill_data.append(['','','','','','','',str(amount)])
bill_data.append(['','IGST 5%','','','','%','',str(tax_5)])
bill_data.append(['','IGST 18%','','','','%','',str(tax_18)])
bill_data.append(['','Total','',str(total_pieces),'','%','',str(total_amount)])


#dynamic data
invoice_number = 13
bill_date = date.today()
delivery_note =''
mode_of_payment = ''
supplier_ref = ''
other_ref = ''
buyer_order_number = 0
buyer_order_date = date.today()
despatch_document_number = 0
delivery_note_date = date.today()
despatched_through = ''
destination = ''

terms_of_delivery = ['Delivery schedule will be alloted after receipt of payment',
                    'Order will be on first come first receipt basis',
                    'Logistics arrangements to be made by the client',]
for term in terms_of_delivery:
    styles = getSampleStyleSheet()
    term = Paragraph(term,styles['Normal'])


buyer_name = 'OBAT MEDICARE'
buyer_address = 'ROY SADAN PLOT NO 42F, ROAD NO 10B,\nRAJENDERA NAGAR\nP.O RAJENDRA NAGAR, P.S KADAM KUAN, PATNA'
buyer_phone_number = 8210064245
buyer_gst_number = '10AAECB2157Q1ZP'

#styling tables and layout
#section 1
fixed_data = [
    [company_details,'Invoice No.\n'+str(invoice_number),'Dated\n'+str(bill_date)],
    ['','Delivery Note\n'+delivery_note,'Mode/Terms of Payment\n'+mode_of_payment],
    ['','Supplier\'s Ref.\n'+supplier_ref,'Other Reference(s)\n'+other_ref],
    ['','Buyer\'s Order No.\n'+str(buyer_order_number),'Dated\n'+str(buyer_order_date)],
    ['','Despatch Document No.\n'+str(despatch_document_number),'Delivery Note Date\n'+str(delivery_note_date)],
    ['','Despatch through\n'+despatched_through,'Destination\n'+destination],
]

style = TableStyle([
    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
    ('SPAN',(0,0),(0,-1)),
    ('VALIGN',(0,0),(-1,-1),'TOP'),


])

fixed_details_table =Table(fixed_data)
fixed_details_table.setStyle(style)
fixed_details_table._argW[0]=280
fixed_details_table._argW[1]=130
fixed_details_table._argW[2]=130

#section 2
buyer_details = 'BUYER\n'+buyer_name+'\n'+buyer_address+'\nMobile: '+str(buyer_phone_number)+'\nGST NO: '+buyer_gst_number
terms_block = 'Terms of Delivery\n*'+terms_of_delivery[0]+'\n*'+terms_of_delivery[1]+'\n*'+terms_of_delivery[2]
styles = getSampleStyleSheet()
#buyer_details = Paragraph(buyer_details,styles['Normal'])
#terms_block = Paragraph(terms_block,styles['Normal'])
#print(buyer_details)
terms_and_buyer_data = [
    [buyer_details,terms_block],
]

terms_and_buyer_table = Table(terms_and_buyer_data)
terms_and_buyer_style = TableStyle([
    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
    ('VALIGN',(0,0),(-1,-1),'TOP'),

])
terms_and_buyer_table.setStyle(terms_and_buyer_style)
terms_and_buyer_table._argW[0]=280
terms_and_buyer_table._argW[1]=260

#bill table
for row in bill_data:
    styles = getSampleStyleSheet()
    row[1] = Paragraph(row[1],styles['Normal'])

bill_table = Table(bill_data)
bill_table_style = TableStyle([
    ('GRID',(0,0),(-1,-5),0.5,colors.grey),
    ('GRID',(0,-1),(-1,-1),0.5,colors.grey),
    ('BOX',(0,-4),(-1,-1),0.5,colors.grey),
    ('LINEBEFORE',(-1,-4),(-1,-1),1,colors.grey),
    ('VALIGN',(0,0),(-1,-1),'TOP'),
])
bill_table.setStyle(bill_table_style)
bill_table._argW[0]=20
bill_table._argW[1]=210
bill_table._argW[2]=60
bill_table._argW[3]=60
bill_table._argW[4]=50
bill_table._argW[5]=30
bill_table._argW[6]=40
bill_table._argW[7]=70

#amount chargeable table
amount_chargeable = [
    ['Amount chargeable (inwords)', 'E. &O.E'],
    [amount_in_words(total_amount),'']
]
amount_chargeable_table = Table(amount_chargeable)
amount_chargeable_table_style = TableStyle([
    ('SPAN',(0,1),(1,1)),
    ('BOX',(0,0),(-1,-1),0.5,colors.grey),
    ('VALIGN',(0,0),(1,0),'TOP'),
    ('ALIGN',(1,0),(1,0),'RIGHT'),

])
amount_chargeable_table.setStyle(amount_chargeable_table_style)
amount_chargeable_table._argW[0]=470
amount_chargeable_table._argW[1]=70

#tax HSN/SAC Table
tax_data = [
    ['HSN/SAC','Taxable Value', 'Integrated Tax','','Total Tax Amount'],
    ['','', 'Rate','Amount',''],
    ['6307',tax_5/0.05,5,tax_5,tax_5],
    ['3921',tax_18/0.18,18,tax_18,tax_18],
    ['Total',amount,'',tax_18+tax_5,tax_5+tax_18]
]

tax_table = Table(tax_data)
tax_table_style = TableStyle([
    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
    ('VALIGN',(0,0),(-1,-1),'TOP'),
    ('SPAN',(0,0),(0,1)),
    ('SPAN',(1,0),(1,1)),
    ('SPAN',(2,0),(3,0)),
])
tax_table.setStyle(tax_table_style)
tax_table._argW[0]=240
tax_table._argW[1]=80
tax_table._argW[2]=40
tax_table._argW[3]=80
tax_table._argW[4]=100

#Amount and declaration table
amount_declaration_data =[
    ['Tax Amount (in words): INR '+ amount_in_words(tax_18+tax_5) +' only',''],
    ['Declaration','for Orion Corp'],
    ['We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct', 'Authorised Signatory']
]
styles = getSampleStyleSheet()
amount_declaration_data[2][0] = Paragraph(amount_declaration_data[2][0],styles['Normal'])

amount_declaration_table = Table(amount_declaration_data)
amount_declaration_style = TableStyle([
    ('BOX',(0,0),(-1,-1),0.5,colors.grey),
    ('BOX',(1,1),(1,-1),0.5,colors.grey),
    ('BOX',(0,1),(-1,-1),0.5,colors.grey),
    ('VALIGN',(0,0),(1,1),'TOP'),
    ('VALIGN',(0,2),(0,2),'TOP'),

    ('ALIGN',(1,1),(1,-1),'RIGHT'),
    ('SPAN',(0,0),(1,0)),
])
amount_declaration_table.setStyle(amount_declaration_style)
amount_declaration_table._argW[0]=280
amount_declaration_table._argW[1]=260
amount_declaration_table._argH[2]=50
amount_declaration_table._argH[0]=20


#structure layout pdf
styles.add(ParagraphStyle(name='my_center', alignment=TA_CENTER))

layout=[]
layout.append(Spacer(1, 20))
layout.append(Paragraph('<font size="16"><b>Proforma Invoice</b></font>', styles["my_center"]))
layout.append(Spacer(1, 20))
layout.append(fixed_details_table)
layout.append(terms_and_buyer_table)
layout.append(bill_table)
layout.append(amount_chargeable_table)
layout.append(tax_table)
layout.append(amount_declaration_table)
layout.append(Spacer(1, 6))
layout.append(Paragraph('<font size="10">This is a Computer Generated Invoice</font>', styles["my_center"]))
#build pdf
file_name = invoice_type+'_no.'+str(invoice_number)+'.pdf'
invoice = SimpleDocTemplate(file_name,rightMargin=0,leftMargin=0,
                        topMargin=0,bottomMargin=0)
invoice.build(layout)
