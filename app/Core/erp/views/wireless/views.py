from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.generic import ListView
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import ArchivoCargaDetail, ArchivoCarga, tipoArchivoCarga, TipoProceso
from Core.erp.forms import CargaArchivoForm
from django.db import transaction, IntegrityError
from django.conf import settings
from django.core import serializers
import xlrd as excel

class WirelessArchivoListView(ListView):
    model = ArchivoCargaDetail
    template_name = 'wireless/lista.html'

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        data = {}
        sid = transaction.savepoint()
        try:
            action = request.POST['action']
            if action == 'searchdata':
                data = []
                for i in ArchivoCarga.objects.all().filter(AplicacionId_id=2):
                    data.append(i.toJSON())
            elif action == 'getTipo':
                data = []
                for i in tipoArchivoCarga.objects.all().filter(tipoProcesoId_id=request.POST['id']):
                    data.append(i.toJSON())
            elif action == 'getProceso':
                data = []
                for i in TipoProceso.objects.all().filter(AplicacionId_id=request.POST['id']):
                    data.append(i.toJSON())
            elif action == 'add':
                try:
                    #cabezara de archivo
                    archivoCarga = ArchivoCarga()
                    archivoCarga.UserId_id = request.user.id
                    archivoCarga.IdSitio_id = request.POST['sitios']
                    archivoCarga.AplicacionId_id = request.POST['aplicaciones']
                    archivoCarga.save()

                    #Guardar detalle
                    for filename in request.FILES:
                        archivoCargaDetail = ArchivoCargaDetail()
                        archivoCargaDetail.ArchivoCargaId_id = archivoCarga.pk
                        tipo = tipoArchivoCarga.objects.get(abreviatura=filename, tipoProcesoId=request.POST['tipoProcesos'])
                        if(tipo.limiteArchivo > 1):
                            for filename2 in request.FILES.getlist(filename):
                                archivoCargaDetail = ArchivoCargaDetail()
                                archivoCargaDetail.ArchivoCargaId_id = archivoCarga.pk
                                archivoCargaDetail.tipoArchivoCargaId_id = tipo.id
                                archivoCargaDetail.ArchivoCheck = filename2
                                archivoCargaDetail.ExtencionArchivo = filename2.content_type
                                archivoCargaDetail.NombreArchivo = filename2.name
                                archivoCargaDetail.save()
                                transaction.savepoint_commit(sid)
                        else:
                            archivoCargaDetail.tipoArchivoCargaId_id = tipo.id
                            archivoCargaDetail.ArchivoCheck = request.FILES[filename]
                            archivoCargaDetail.ExtencionArchivo = request.FILES[filename].content_type
                            archivoCargaDetail.NombreArchivo = request.FILES[filename].name
                            archivoCargaDetail.save()
                            transaction.savepoint_commit(sid)
                except IntegrityError as ex:
                    transaction.savepoint_rollback(sid)
                    data['error'] = str(ex)
            elif action == 'getDetalle':
                data = []
                for i in ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id']):
                    data.append(i.toJSON())
            elif action == 'delete':
                for detalle in ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id']):
                    detalle.delete()
                archivoCarga = ArchivoCarga.objects.get(pk=request.POST['id'])
                archivoCarga.delete()
            elif action == 'getParametro':
                data = {}
                try:
                    data = []
                    #Obtener el archivo a evaluar
                    archBase = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'], tipoArchivoCargaId__Validar=True)
                    #Ruta del archivo base
                    rutaBase = '{0}{1}'.format(settings.MEDIA_ROOT, archBase[0].ArchivoCheck)
                    archsComp = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'], tipoArchivoCargaId__Validar=False)

                    #Realizar a formar un solo archivo
                    for file in archsComp:
                        rutaComp = '{0}{1}'.format(settings.MEDIA_ROOT, file.ArchivoCheck)
                        wb = excel.open_workbook(rutaComp)

                        if '1.PW_GL_BaseLine_HNG' in wb.sheet_names():
                            sheet = wb.sheet_by_name('1.PW_GL_BaseLine_HNG')
                            for row in range(sheet.nrows):
                                if sheet.cell_value(row, 2) != 'MO' and sheet.cell_value(row, 6) != '':
                                    data.append(
                                        {'parametro': sheet.cell_value(row, 3), 'valor': sheet.cell_value(row, 6),
                                         'mo': sheet.cell_value(row, 2)})
                        elif 'DF4G_PW' in wb.sheet_names():
                            sheet = wb.sheet_by_name('DF4G_PW')
                            for column in range(34, 43):
                                data.append({'parametro': sheet.cell_value(0, column), 'valor': sheet.cell_value(418, column), 'mo': ''})
                        else:
                            data['error'] = 'No se encontro hojas en el archivo'

                except FileNotFoundError:
                    data['error'] = 'Archivo no encontrado'
            elif action == 'validarArchivo':

                print(request.POST['dataParametro'])
            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            transaction.savepoint_rollback(sid)
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || WIRELESS'
        context['project'] = 'WIRELESS'
        context['create_url'] = '' #reverse_lazy('erp:client_create')
        context['list_url'] = '' #reverse_lazy('erp:client_list')
        context['entity'] = 'Carga de archivos'
        context['form'] = CargaArchivoForm()
        return context


