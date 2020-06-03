amount=0
tax_5 = 0
tax_12 = 0
tax_18 = 0
tax_5_set= set()
tax_12_set= set()
tax_18_set= set()

total_pieces = 0
total_amount = 0

def generate(bill_data):
    global amount
    global tax_5
    global tax_18
    global tax_12
    global total_pieces
    global total_amount
    global tax_5_set
    global tax_12_set
    global tax_18_set


    amount=0
    tax_5 = 0
    tax_12 = 0
    tax_18 = 0
    total_pieces = 0
    total_amount = 0
    tax_5_set= set()
    tax_12_set= set()
    tax_18_set= set()

    #imports
    import time
    import os
    from datetime import date
    from reportlab.platypus import tables, SimpleDocTemplate,Table,Paragraph,Spacer,TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from num2words import num2words
    import pandas as pd


    #calculate tax_table and total amount
    #amount=0
    #tax_5 = 0
    #tax_18 = 0
    #total_pieces = 0
    #total_amount = 0
    def calculate_table(bill_data):

        #tax_percentage = {6307:0.05,3921:0.18,40159100}
        global amount
        global tax_5
        global tax_18
        global tax_12
        global total_pieces
        global total_amount
        global tax_5_set
        global tax_12_set
        global tax_18_set

        for i, row in enumerate(bill_data):
            if i!= 0:
                row[1]= str(row[1])
                styles = getSampleStyleSheet()
                styles.add(ParagraphStyle(name='custom', alignment=TA_CENTER, fontSize= 6))
                row[1] = Paragraph(row[1],styles['custom'])
                amount = amount +row[7]
                total_pieces = total_pieces + row[3]
                if row[6]==5 :
                    tax_5 = tax_5 + 0.05*row[7]
                    tax_5_set.add(row[2])
                elif row[6]==18 :
                    tax_18 = tax_18 + 0.18*row[7]
                    tax_18_set.add(row[2])
                elif row[6]==12 :
                    tax_12 = tax_12 + 0.12*row[7]
                    tax_12_set.add(row[2])

        total_amount = amount + tax_18 + tax_5+ tax_12

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
    #invoice_type = 'Proforma_Invoice'
    styles = getSampleStyleSheet()
    company_name =  'Orion Corp'
    company_address = 'P-2/03, Tower 3B, Purvanchal Silver City -2\nSector Pi-2, Greater Noida\nIndia, 201308, Phone:0120-4543418'
    warehouse_address = 'WAREHOUSE- PLOT NO-239,\nGiani Compound, Giani Boarder, Opposite Metro Pillar No.160, Behind Giani Gill Transport, Post Ckikamberpur, GSTIN/UIN: 09AKGPG4906P1ZQ State Name: Uttar Pradesh, Code: 09'
    bank_details = 'Orion Corp\nPunjab & Sind Bank\nA/C no.: 09701100000176\nIFSC Code: PSIB0020970\nMICR-110023123\nBranch Address\nKasons Tower, Alpha-1 Commercial Belt, First floor,\nGreater Noida-201306\nPhone: 0120-4199100\nFax: 236098'
    company_details = company_name+'\n'+company_address+'\n'+warehouse_address+'\n'+bank_details
    company_details =   Paragraph('''<font size="8">
    <b>Orion Corp</b> <br/>
    P-2/03, Tower 3B, Purvanchal Silver City -2, Sector Pi-2, Greater Noida, India, 201308, Phone:0120-4543418 <br/>
    WAREHOUSE- PLOT NO-239, Giani Compound, Giani Boarder, Opposite Metro Pillar No.160, Behind Giani Gill Transport, Post Ckikamberpur <br/>
    GSTIN/UIN: 09AKGPG4906P1ZQ <br/>
    State Name: Uttar Pradesh, Code: 09 <br/>
    <b>Bank Details</b><br/>
    Orion Corp<br/>
    ICICI BANK LTD A/C no.: 003105037390 <br/>
    IFSC Code: ICIC0000031 <br/>
    Branch Address <br/>
    K-1,SENIOR MALL, SECTOR 18, NOIDA, UTTAR PRADESH, PIN CODE : 201301 </font>''', styles["Normal"])


    #retieving data from excel filename
    #fake data



    #xl_file = pd.read_excel('/Users/Sanjay/Desktop/OrionCorp/bill_new1.xlsx',0)
    #bill_data = xl_file.values.tolist()
    #bill_data_head_row = ['SI','Description of Goods','HSN/SAC','Quantity','Rate','per','GST%','Amount']
    #bill_data.insert(0,bill_data_head_row)
    styles = getSampleStyleSheet()

    invoice_type = bill_data[0][0]
    invoice_number = bill_data[0][1]

    buyer_name = bill_data[0][2]
    buyer_name = 'Buyer Name: '+buyer_name

    buyer_address = bill_data[0][3]
    buyer_address = 'Address: '+ buyer_address
    styles.add(ParagraphStyle(name='buyer', alignment=TA_JUSTIFY, fontSize= 8))
    buyer_address = Paragraph(buyer_address,styles['buyer'],)

    buyer_phone_number = bill_data[0][4]
    buyer_phone_number = 'Phone no.: '+str(buyer_phone_number)
    buyer_phone_number = Paragraph(buyer_phone_number,styles['buyer'],)

    buyer_gst_number = bill_data[0][5]
    buyer_gst_number = 'GST no.: '+buyer_gst_number
    buyer_gst_number = Paragraph(buyer_gst_number,styles['buyer'],)

    bill_date = bill_data[0][6]
    delivery_note = bill_data[0][7]
    mode_of_payment = bill_data[2][0]
    supplier_ref = bill_data[2][1]
    other_ref = bill_data[2][2]
    buyer_order_number = bill_data[2][3]
    buyer_order_date = bill_data[2][4]
    despatch_document_number = bill_data[2][5]
    delivery_note_date = bill_data[2][6]
    despatched_through = bill_data[2][7]
    destination = bill_data[4][0]
    destination = 'SHIP TO: '+str(destination)
    destination = Paragraph(destination,styles['buyer'],)

    term1 = bill_data[4][1]
    term1 = '*'+ term1
    term1 = Paragraph(term1,styles['buyer'])

    term2 = bill_data[4][2]
    term2 = '*'+ term2
    term2 = Paragraph(term2,styles['buyer'])

    term3 = bill_data[4][3]
    term3 = '*'+ term3
    term3 = Paragraph(term3,styles['buyer'])


    print(bill_data)

    bill_data.pop(0)
    bill_data.pop(0)
    bill_data.pop(0)
    bill_data.pop(0)
    bill_data.pop(0)
    bill_data.pop(0)
    bill_data_head_row = ['SI','Description of Goods','HSN/SAC','Quantity','Rate','per','GST%','Amount']
    bill_data.insert(0,bill_data_head_row)
    calculate_table(bill_data)
    bill_data.append(['','','','','','','',str(amount)])
    flag_5=0
    flag_18=0
    flag_12=0

    if(tax_5!=0):
        flag_5 = 1
        bill_data.append(['','IGST 5%','','','','%','',str(tax_5)])
    if(tax_18!=0):
        flag_18 = 1
        bill_data.append(['','IGST 18%','','','','%','',str(tax_18)])
    if(tax_12!=0):
        flag_12 = 1
        bill_data.append(['','IGST 12%','','','','%','',str(tax_12)])
    bill_data.append(['','Total','',str(total_pieces),'','%','',str(total_amount)])
    flag = flag_5+flag_18+flag_12

    #dynamic data
    #invoice_number = 13
    #bill_date = date.today()
    #delivery_note =''
    #mode_of_payment = ''
    #supplier_ref = ''
    #other_ref = ''
    #buyer_order_number = 0
    #buyer_order_date = date.today()
    #despatch_document_number = 0
    #delivery_note_date = date.today()
    #despatched_through = ''
    #destination = ''




    #buyer_name = 'OBAT MEDICARE'
    #buyer_address = 'ROY SADAN PLOT NO 42F, ROAD NO 10B,\nRAJENDERA NAGAR\nP.O RAJENDRA NAGAR, P.S KADAM KUAN, PATNA'
    #buyer_phone_number = 8210064245
    #buyer_gst_number = '10AAECB2157Q1ZP'

    #styling tables and layout
    #section 1
    fixed_data = [
        [company_details,'Inv No.\n'+str(invoice_number),'Dated\n'+str(bill_date)],
        ['','Delivery Note\n'+str(delivery_note),'Mode/Terms of Payment\n'+str(mode_of_payment)],
        ['','Other Reference(s)\n'+str(other_ref),'Buyer\'s Order No.\n'+str(buyer_order_number)],
        ['','Dated\n'+str(buyer_order_date),'Despatch Document No.\n'+str(despatch_document_number)],
        ['','Delivery Note Date\n'+str(delivery_note_date),'Despatch through\n'+despatched_through],
        ['',destination,''],
    ]

    style = TableStyle([
        ('GRID',(0,0),(-1,-1),0.5,colors.grey),
        ('SPAN',(0,0),(0,-1)),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('SPAN',(-2,-1),(-1,-1)),


    ])

    fixed_details_table =Table(fixed_data)
    fixed_details_table.setStyle(style)
    fixed_details_table._argW[0]=280
    fixed_details_table._argW[1]=130
    fixed_details_table._argW[2]=130

    #section 2
    #buyer_details = 'BUYER\n'+str(buyer_name)+'\n'+buyer_address+'\nMobile: '+str(buyer_phone_number)+'\nGST NO: '+buyer_gst_number
    #terms_block = 'Terms of Delivery\n*'+terms_of_delivery[0]+'\n*'+terms_of_delivery[1]+'\n*'+terms_of_delivery[2]
    #styles = getSampleStyleSheet()
    #buyer_details = Paragraph(buyer_details,styles['Normal'])
    #terms_block = Paragraph(terms_block,styles['Normal'])
    #print(buyer_details)
    terms_and_buyer_data = [
        [buyer_name,'Terms of Delivery'],
        [buyer_address,term1],
        [buyer_phone_number,term2],
        [buyer_gst_number,term3],
    ]

    terms_and_buyer_table = Table(terms_and_buyer_data)
    terms_and_buyer_style = TableStyle([
        ('GRID',(0,0),(1,0),0.5,colors.grey),
        ('BOX',(0,0),(-1,-1),0.5,colors.grey),
        ('BOX',(0,1),(0,-1),0.5,colors.grey),
        ('VALIGN',(0,0),(-1,-1),'TOP'),

    ])
    terms_and_buyer_table.setStyle(terms_and_buyer_style)
    terms_and_buyer_table._argW[0]=280
    terms_and_buyer_table._argW[1]=260

    #bill table
    #for row in bill_data:
    #    styles = getSampleStyleSheet()
    #    row[1] = Paragraph(row[1],styles['Normal'])

    bill_table = Table(bill_data)
    bill_table_style = TableStyle([
        ('GRID',(0,0),(-1,-3-flag),0.5,colors.grey),
        ('GRID',(0,-1),(-1,-1),0.5,colors.grey),
        ('BOX',(0,-2-flag),(-1,-1),0.5,colors.grey),
        ('LINEBEFORE',(-1,-2-flag),(-1,-1),1,colors.grey),
        ('LINEBEFORE',(-2,-2-flag),(-2,-1),1,colors.grey),
        ('LINEBEFORE',(-3,-2-flag),(-3,-1),1,colors.grey),
        ('LINEBEFORE',(-4,-2-flag),(-4,-1),1,colors.grey),
        ('LINEBEFORE',(-5,-2-flag),(-5,-1),1,colors.grey),
        ('LINEBEFORE',(-6,-2-flag),(-6,-1),1,colors.grey),
        ('LINEBEFORE',(-7,-2-flag),(-7,-1),1,colors.grey),

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
    ]
    if(flag_5==1):
        tax_data.append([tax_5_set,tax_5/0.05,5,tax_5,tax_5])
    if(flag_18==1):
        tax_data.append([tax_18_set,tax_18/0.18,18,tax_18,tax_18],)
    if(flag_12==1):
        tax_data.append([tax_12_set,tax_12/0.12,12,tax_12,tax_12])
    tax_data.append(['Total',amount,'',tax_18+tax_5+tax_12,tax_5+tax_18+tax_12])

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
        ['Tax Amount (in words):'+ amount_in_words(tax_18+tax_5+tax_12) +' only',''],
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
    styles.add(ParagraphStyle(name='heading', alignment=TA_CENTER, fontSize=12))

    layout=[]
    layout.append(Spacer(1, 20))
    layout.append(Paragraph(invoice_type, styles["heading"]))
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

    file_name = 'download.pdf'
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    STATIC_DIR = os.path.join(BASE_DIR,'static')
    file_name = os.path.join(STATIC_DIR,file_name)
    invoice = SimpleDocTemplate(file_name,rightMargin=0,leftMargin=0,
                                topMargin=0,bottomMargin=0,fontSize=6)
    invoice.build(layout)
