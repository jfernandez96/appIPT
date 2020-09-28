from django.views.generic import TemplateView
from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import tipoArchivoCarga, TipoProceso
from Core.erp.forms import TipoArchivoCargaForm,cboSelect

class TipoArchivoCargaListView(TemplateView):
    model = tipoArchivoCarga
    template_name = 'mantenimiento/tipoArchivoCarga/lista.html'

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        data = {}
        try:
            action = request.POST['action']
            if action == 'searchdata':
                data = []
                for i in tipoArchivoCarga.objects.all():
                    data.append(i.toJSON())
            elif action == 'getTipoProceso':
                data = []
                for i in TipoProceso.objects.all().filter(AplicacionId_id=request.POST['id']):
                    data.append(i.toJSON())
            elif action == 'add':
                tipo = tipoArchivoCarga()
                tipo.tipoProcesoId_id = request.POST['tipoProcesoId']
                tipo.NombreTipoCarga = request.POST['NombreTipoCarga']
                tipo.limiteArchivo = request.POST['limiteArchivo']
                tipo.TipoArchivopermitido = request.POST['TipoArchivopermitido']
                tipo.abreviatura = request.POST['abreviatura']

                if request.POST['Validar'] == 'on':
                    tipo.Validar = True
                else:
                    tipo.Validar = False

                tipo.save()
            elif action == 'edit':

                tipo = tipoArchivoCarga.objects.get(pk=request.POST['id'])
                tipo.tipoProcesoId_id = request.POST['tipoProcesoId']
                tipo.NombreTipoCarga = request.POST['NombreTipoCarga']
                tipo.limiteArchivo = request.POST['limiteArchivo']
                tipo.TipoArchivopermitido = request.POST['TipoArchivopermitido']
                tipo.abreviatura = request.POST['abreviatura']

                if request.POST['Validar'] == 'on':
                    tipo.Validar = True
                else:
                    tipo.Validar = False

                tipo.save()
            elif action == 'delete':
                tipo = tipoArchivoCarga.objects.get(pk=request.POST['id'])
                tipo.delete()
            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || Archivo de carga'
        context['project'] = 'Mantenimiento'
        context['entity'] = 'Tipo archivo de carga'
        context['cboAplicacion'] = cboSelect()
        context['form'] = TipoArchivoCargaForm()
        return context