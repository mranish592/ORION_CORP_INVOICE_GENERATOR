from django.shortcuts import render
import pandas as pd
import numpy as np
from invoice_generator import generate_pdf


from django.conf import settings
from django.templatetags.static import static
# Create your views here.
def index(request):
    return render(request, 'invoice_generator/index.html')

def invoice_upload(request):
    path = settings.MEDIA_ROOT

    if request.method =='POST':
        df = request.FILES['bill_data']
        xl_file0 = pd.read_excel(df,0)
        xl_file = xl_file0.replace(np.nan, '', regex=True)
        bill_data = xl_file.values.tolist()

        generate_pdf.generate(bill_data)
        return render(request,'invoice_generator/invoice_upload.html',context={})
    else:
        return render(request,'invoice_generator/invoice_upload.html',context={})
