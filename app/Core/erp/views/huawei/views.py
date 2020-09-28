import uuid

from django.http import JsonResponse, HttpResponse
from django.utils.decorators import method_decorator
from django.views.generic import ListView
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import ArchivoCargaDetail, ArchivoCarga, tipoArchivoCarga, TipoProceso
from Core.erp.forms import CargaArchivoForm
from django.db import transaction, IntegrityError
from django.conf import settings
import xlrd as excel
from django.utils.encoding import smart_str

from Core.erp.views.huawei.process import Validate_Huawei


class HuaweiArchivoListView(ListView):
    model = ArchivoCargaDetail
    template_name = 'huawei/lista.html'

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
                for i in ArchivoCarga.objects.all().filter(AplicacionId_id=1):
                    data.append(i.toJSON())
                print(data)
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
                    # cabezara de archivo
                    archivoCarga = ArchivoCarga()
                    archivoCarga.UserId_id = request.user.id
                    archivoCarga.IdSitio_id = request.POST['sitios']
                    archivoCarga.AplicacionId_id = request.POST['aplicaciones']
                    archivoCarga.save()

                    # Guardar detalle
                    for filename in request.FILES:
                        archivoCargaDetail = ArchivoCargaDetail()
                        archivoCargaDetail.ArchivoCargaId_id = archivoCarga.pk
                        tipo = tipoArchivoCarga.objects.get(abreviatura=filename,
                                                            tipoProcesoId=request.POST['tipoProcesos'])
                        if (tipo.limiteArchivo > 1):
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
                    archLTE = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                      tipoArchivoCargaId__abreviatura="LTE")
                    rutaLTE = '{0}{1}'.format(settings.MEDIA_ROOT, archLTE[0].ArchivoCheck)
                    wb = excel.open_workbook(rutaLTE)
                    if 'Baseline_Unique' in wb.sheet_names():
                        sheet = wb.sheet_by_name('Baseline_Unique')
                        for row in range(sheet.nrows):
                            if sheet.cell_value(row, 1) != 'MO' and sheet.cell_value(row, 5) != '':
                                data.append({'parametro': sheet.cell_value(row, 2), 'valor': sheet.cell_value(row, 5),
                                             'mo': sheet.cell_value(row, 1)})
                    else:
                        data = ""

                except FileNotFoundError:
                    data['error'] = 'Archivo no encontrado'
            elif action == 'descargar':
                if action == 'des':
                    response = HttpResponse(mimetype='application/force-download')
                    response['Content-Disposition'] = 'attachment; filename=%s' % smart_str(request.POST['fileName'])
                    response['X-Sendfile'] = smart_str(request.POST['ruta'])
                    return response
            elif action == 'procesar':
                # le puedes poner la ruta aqui :v
                dest_filename = 'Result_book_Huawei' + str(uuid.uuid4().fields[-1])[:6] + '.xlsx'
                rutaResul = '{0}Result/{1}'.format(settings.MEDIA_ROOT, dest_filename)

                archLTE = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                  tipoArchivoCargaId__abreviatura="LTE")
                archSite = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                   tipoArchivoCargaId__abreviatura="Site")
                archDF = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                 tipoArchivoCargaId__abreviatura="DF")

                arch3gCELL = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                     tipoArchivoCargaId__abreviatura="3gCELL")
                Site_EUTRA = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                     tipoArchivoCargaId__abreviatura="Site_EUTRA")
                s3gRadNetw = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                     tipoArchivoCargaId__abreviatura="3gRadNetw")

                Site_EUTRA_count_data = ArchivoCargaDetail.objects.all().filter(ArchivoCargaId_id=request.POST['id'],
                                                                                tipoArchivoCargaId__abreviatura="Site_EUTRA").count()
                site1 = ""
                site2 = ""
                site3 = ""
                for i in range(0, 3):
                    param = str(archSite[i].ArchivoCheck)
                    if param.find('ConfigurationData_Cell') > 0:
                        site1 = param
                    if param.find('RnpData_BTS3900') > 0:
                        site2 = param
                    if param.find('ConfigurationData_eNodeB') > 0:
                        site3 = param
                if len(site1) > 0 and len(site2) > 0 and len(site3) > 0:

                    Validate_ = Validate_Huawei(
                        url_lte=settings.MEDIA_ROOT + str(archLTE[0].ArchivoCheck),
                        url_site1=settings.MEDIA_ROOT + str(site1),
                        url_site2=settings.MEDIA_ROOT + str(site2),
                        url_site3=settings.MEDIA_ROOT + str(site3),
                        url_df=settings.MEDIA_ROOT + str(archDF[0].ArchivoCheck),
                        url_3g_cell=settings.MEDIA_ROOT + str(arch3gCELL[0].ArchivoCheck),
                        url_3g_rad_net=settings.MEDIA_ROOT + str(s3gRadNetw[0].ArchivoCheck),
                        name_file_lte='NTA_LTE_Ran_Sharing_IPT_Minimacro_V4_20200805.xlsx')

                    list_output = Validate_.validate_general_huawei(rutaResul)

                    index_result_1 = list_output["index_result_1"]
                    index_result_2 = list_output["index_result_2"]

                    if Site_EUTRA_count_data > 0:
                        for i in range(0, Site_EUTRA_count_data):
                            nombre_archivo = str(Site_EUTRA[i].NombreArchivo).replace('.xlsx', '')
                            Url_archivo = settings.MEDIA_ROOT + str(Site_EUTRA[i].ArchivoCheck)
                            print(nombre_archivo, '|||||' + Url_archivo)
                            list_output = Validate_.validate_external_Huawei_final(nombre_archivo, Url_archivo,
                                                                                   index_result_1,
                                                                                   index_result_2)
                            index_result_1 = list_output["index_result_1"]
                            index_result_2 = list_output["index_result_2"]
                    Validate_.save_workbook(rutaResul)
                    # Site_EUTRA

                    data = '/media/Result/' + dest_filename
                else:
                    data['error'] = 'Ha ocurrido un error,los nombres de los archivos son incorrectos.'

            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            transaction.savepoint_rollback(sid)
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || HUAWEI'
        context['project'] = 'HUAWEI'
        context['create_url'] = ''  # reverse_lazy('erp:client_create')
        context['list_url'] = ''  # reverse_lazy('erp:client_list')
        context['entity'] = 'Carga de archivos'
        context['form'] = CargaArchivoForm()
        return context
