from django.contrib.auth.mixins import LoginRequiredMixin
from django.views.generic import ListView, CreateView, UpdateView, DeleteView
from django.views.generic import TemplateView
from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
from Core.erp.models import Sitios
from Core.erp.forms import SitioForm

class SitioListView(TemplateView):
    model = Sitios
    template_name = 'mantenimiento/sitio/lista.html'

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        data = {}
        try:
            action = request.POST['action']
            if action == 'searchdata':
                data = []
                for i in Sitios.objects.all():
                    data.append(i.toJSON())
            elif action == 'add':
                sitio = Sitios()
                sitio.nombresitio = request.POST['nombresitio']
                sitio.save()
            elif action == 'edit':
                sitio = Sitios.objects.get(pk=request.POST['id'])
                sitio.nombresitio = request.POST['nombresitio']
                sitio.save()
            elif action == 'delete':
                sitio = Sitios.objects.get(pk=request.POST['id'])
                sitio.delete()
            else:
                data['error'] = 'Ha ocurrido un error'
        except Exception as e:
            data['error'] = str(e)
        return JsonResponse(data, safe=False)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'IPT || Sitios'
        context['project'] = 'Mantenimiento'
        context['create_url'] = ''  # reverse_lazy('erp:client_create')
        context['list_url'] = ''  # reverse_lazy('erp:client_list')
        context['entity'] = 'Sitio'
        context['form'] = SitioForm()
        return context




