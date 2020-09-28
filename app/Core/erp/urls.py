from django.urls import path
from Core.erp.views.dashboard.views import *
from Core.erp.views.huawei.views import *
from Core.erp.views.mantenimiento.sitio.views import *
from Core.erp.views.mantenimiento.aplicacion.views import *
from Core.erp.views.mantenimiento.tipoProceso.views import *
from Core.erp.views.mantenimiento.tipoArchivoCarga.views import *
from Core.erp.views.wireless.views import *



app_name = 'erp'

urlpatterns = [
    #mantenimiento
    path('mantenimient/sitio/lista', SitioListView.as_view(), name='lista_sitio'),
    path('mantenimient/aplicacion/lista', AplicacionListView.as_view(), name='lista_aplicacion'),
    path('mantenimient/tipoProceso/lista', TipoProcesoListView.as_view(), name='lista_tipoProceso'),
    path('mantenimient/tipoArchivoCarga/lista', TipoArchivoCargaListView.as_view(), name='lista_tipoArchivo'),
    #huawei
    path('huawei/lista/', HuaweiArchivoListView.as_view(), name='lista_archivoHuawei'),

    #wireless
    path('wireless/lista/', WirelessArchivoListView.as_view(), name='lista_archivoWireless'),

    #home
    path('dashboard/', DashboardView.as_view(), name='dashboard'),

]
