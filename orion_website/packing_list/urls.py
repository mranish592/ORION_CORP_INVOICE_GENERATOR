from django.urls import path
from packing_list import views
from django.conf.urls.static import static
from django.conf import settings

app_name = 'packing_list'

urlpatterns = [
    path('',views.index,name = 'index'),
    path('packing_upload',views.packing_upload,name = 'packing_upload'),

]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
