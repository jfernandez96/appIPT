from django.views.generic import TemplateView
from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import TipoProceso
from Core.erp.forms import TipoProcesoForm

class TipoProcesoListView(TemplateView):
    model = TipoProceso
    template_name = 'mantenimiento/tipoProceso/lista.html'

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        data = {}
        try:
            action = request.POST['action']
            if action == 'searchdata':
                data = []
                for i in TipoProceso.objects.all():
                    data.append(i.toJSON())
            elif action == 'add':
                tipo = TipoProceso()
                tipo.nombretipoproceso = request.POST['nombretipoproceso']
                tipo.AplicacionId_id = request.POST['AplicacionId']
                tipo.save()
            elif action == 'edit':
                tipo = TipoProceso.objects.get(pk=request.POST['id'])
                tipo.nombretipoproceso = request.POST['nombretipoproceso']
                tipo.AplicacionId_id = request.POST['AplicacionId']
                tipo.save()
            elif action == 'delete':
                tipo = TipoProceso.objects.get(pk=request.POST['id'])
                tipo.delete()
            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || Tipo de proceso'
        context['project'] = 'Mantenimiento'
        context['create_url'] = ''  # reverse_lazy('erp:client_create')
        context['list_url'] = ''  # reverse_lazy('erp:client_list')
        context['entity'] = 'Tipo proceso'
        context['form'] = TipoProcesoForm()
        return context