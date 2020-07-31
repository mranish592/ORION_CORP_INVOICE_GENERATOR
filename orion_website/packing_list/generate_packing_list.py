#imports
import time
import os
from datetime import date
from reportlab.platypus import tables, SimpleDocTemplate,Table,Paragraph,Spacer,TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT
from reportlab.lib import colors
from reportlab.lib.units import inch
from num2words import num2words
import pandas as pd

def generate(data):
    layout=[]
    invoice_number = data[0][0]
    invoice_number = 'Invoice no: ' +str(invoice_number)
    shipping_address = data[0][1]
    shipping_address = "Address: "+ str(shipping_address)
    address_line = int(len(shipping_address)/50)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='address', alignment=TA_LEFT))
    shipping_address = Paragraph(shipping_address,styles['address'],)

    buyer_name = data[2][0]
    buyer_name = 'Buyer: '+buyer_name
    date = data[2][1]
    date = 'Date: '+str(date)
    contact_no = data[4][0]
    contact_no = 'Contact No: '+str(contact_no)

    data.pop(0)
    data.pop(0)
    data.pop(0)
    data.pop(0)
    data.pop(0)
    heading = data[0];
    data.pop(0)
    spaceleft=42
    layout.append(Spacer(1, 40))

    for i, row in enumerate(data):
        packing_list=[]
        x=0
        for j, ele in enumerate(row):
            if(j==0):
                packing_list.append(['Box No: '+str(ele)])
                packing_list.append([invoice_number,date])
                packing_list.append([buyer_name,contact_no])
                packing_list.append([shipping_address])
                packing_list.append([''])
                packing_list.append(['Product','Quantity'])

            elif(ele!=0):
                x=x+1
                packing_list.append([heading[j],ele])

        fixed_details_table =Table(packing_list)
        style = TableStyle([
            ('GRID',(0,5),(-1,-1),0.5,colors.grey),
            #('SPAN',(0,3),(1,3)),
            ('ALIGN',(1,1),(1,2),'RIGHT'),
        ])


        fixed_details_table.setStyle(style)
        fixed_details_table._argW[0]=300
        fixed_details_table._argW[1]=240
        #fixed_details_table._argW[2]=130

        spaceleft = spaceleft - 10 - address_line - x
        if(spaceleft<2):
            layout.append(PageBreak())
            spaceleft=42
            layout.append(Spacer(1, 40))


        layout.append(fixed_details_table)
        layout.append(Spacer(1, 40))


    file_name = 'packing_download.pdf'
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    STATIC_DIR = os.path.join(BASE_DIR,'static')
    file_name = os.path.join(STATIC_DIR,file_name)
    invoice = SimpleDocTemplate(file_name,rightMargin=0,leftMargin=0,
                                topMargin=0,bottomMargin=0,fontSize=6)
    invoice.build(layout)
