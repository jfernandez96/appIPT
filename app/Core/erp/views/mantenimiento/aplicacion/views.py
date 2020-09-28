from django.views.generic import TemplateView
from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import Aplicacion
from Core.erp.forms import AplicacionForm

class AplicacionListView(TemplateView):
    model = Aplicacion
    template_name = 'mantenimiento/aplicacion/lista.html'

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        data = {}
        try:
            action = request.POST['action']
            if action == 'searchdata':
                data = []
                for i in Aplicacion.objects.all():
                    data.append(i.toJSON())
            elif action == 'add':
                aplicacion = Aplicacion()
                aplicacion.nombreAplicacion = request.POST['nombreAplicacion']
                aplicacion.save()
            elif action == 'edit':
                aplicacion = Aplicacion.objects.get(pk=request.POST['id'])
                aplicacion.nombreAplicacion = request.POST['nombreAplicacion']
                aplicacion.save()
            elif action == 'delete':
                aplicacion = Aplicacion.objects.get(pk=request.POST['id'])
                aplicacion.delete()
            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || Aplicaciones'
        context['project'] = 'Mantenimiento'
        context['create_url'] = ''  # reverse_lazy('erp:client_create')
        context['list_url'] = ''  # reverse_lazy('erp:client_list')
        context['entity'] = 'Aplicaciones'
        context['form'] = AplicacionForm()
        return context