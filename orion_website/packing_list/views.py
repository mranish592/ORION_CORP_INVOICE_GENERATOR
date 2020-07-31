from django.shortcuts import render
import pandas as pd
import numpy as np
from invoice_generator import generate_pdf
from packing_list import generate_packing_list

from django.conf import settings
from django.templatetags.static import static
# Create your views here.
def index(request):
    return render(request, 'invoice_generator/index.html')

def packing_upload(request):
    path = settings.MEDIA_ROOT

    if request.method =='POST':
        df = request.FILES['packing_list_data']
        xl_file0 = pd.read_excel(df,0)
        xl_file = xl_file0.replace(np.nan, 0, regex=True)
        packing_list_data = xl_file.values.tolist()

        generate_packing_list.generate(packing_list_data)
        return render(request,'packing_list/packing_upload.html',context={})
    else:
        return render(request,'packing_list/packing_upload.html',context={})
