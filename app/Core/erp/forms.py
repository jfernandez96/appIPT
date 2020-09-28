from django.forms import *
from Core.erp.models import *


class SitioForm(ModelForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['nombresitio'].widget.attrs['autofocus'] = True

    class Meta:
        model = Sitios
        fields = '__all__'
        widgets = {
            'nombresitio': TextInput(
                attrs={
                    'placeholder': 'Ingrese sitio',
                    'class': 'form-control input-sm'
                }

            )
        }
        exclude = ['user_updated', 'user_creation']

    def save(self, commit=True):
        data = {}
        form = super()
        try:
            if form.is_valid():
                form.save()
            else:
                data['error'] = form.errors
        except Exception as e:
            data['error'] = str(e)
        return data


class AplicacionForm(ModelForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['nombreAplicacion'].widget.attrs['autofocus'] = True

    class Meta:
        model = Aplicacion
        fields = '__all__'
        widgets = {
            'nombreAplicacion': TextInput(
                attrs={
                    'placeholder': 'Ingrese Aplicacion',
                    'class': 'form-control input-sm'
                }

            )
        }
        exclude = ['user_updated', 'user_creation']

    def save(self, commit=True):
        data = {}
        form = super()
        try:
            if form.is_valid():
                form.save()
            else:
                data['error'] = form.errors
        except Exception as e:
            data['error'] = str(e)
        return data


class TipoProcesoForm(ModelForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['nombretipoproceso'].widget.attrs['autofocus'] = True

    class Meta:
        model = TipoProceso
        fields = '__all__'
        exclude = ['user_updated', 'user_creation']
        widgets = {
            'nombretipoproceso': TextInput(
                attrs={
                    'placeholder': 'Ingrese nombre',
                    'class': 'form-control input-sm'
                }

            ),
            'AplicacionId': Select(
                attrs={
                    'class': 'form-control input-sm'
                }
            )
        }

    def save(self, commit=True):
        data = {}
        form = super()
        try:
            if form.is_valid():
                form.save()
            else:
                data['error'] = form.errors
        except Exception as e:
            data['error'] = str(e)
        return data


class TipoArchivoCargaForm(ModelForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['NombreTipoCarga'].widget.attrs['autofocus'] = True

    class Meta:
        model = tipoArchivoCarga
        fields = '__all__'
        exclude = ['user_updated', 'user_creation', 'tipoProcesoId','Validar']

        widgets = {

            'NombreTipoCarga': TextInput(
                attrs={
                    'placeholder': 'Ingrese nombre tipo carga',
                    'class': 'form-control input-sm'
                }
            ),
            'limiteArchivo': NumberInput(
                attrs={
                    'class': 'form-control input-sm'
                }
            ),
            'TipoArchivopermitido': TextInput(
                attrs={
                    'class': 'form-control input-sm'
                }
            ),
            'abreviatura': TextInput(
                attrs={
                    'class': 'form-control input-sm'
                }
            )
        }

    def save(self, commit=True):
        data = {}
        form = super()
        try:
            if form.is_valid():
                form.save()
            else:
                data['error'] = form.errors
        except Exception as e:
            data['error'] = str(e)
        return data


class cboSelect(Form):
    aplicaciones = ModelChoiceField(queryset=Aplicacion.objects.all(), widget=Select(attrs={
        'class': 'form-control input-sm',
        'style': 'width: 100%'
    }))

    sitios = ModelChoiceField(queryset=Sitios.objects.all(), to_field_name='id', widget=Select(attrs={
        'class': 'form-control',
        'style': 'width: 100%'
    }))


class CargaArchivoForm(ModelForm):
    sitios = ModelChoiceField(queryset=Sitios.objects.all(), widget=Select(attrs={
        'class': 'form-control input-sm select2',
        'style': 'width: 100%'
    }))
    aplicaciones = ModelChoiceField(queryset=Aplicacion.objects.all(), widget=Select(attrs={
        'class': 'form-control input-sm',
        'style': 'width: 100%'
    }))

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    class Meta:
        model = ArchivoCargaDetail
        fields = ['sitios', 'aplicaciones']

    def save(self, commit=True):
        data = {}
        form = super()
        try:
            if form.is_valid():
                form.save()
            else:
                data['error'] = form.errors
        except Exception as e:
            data['error'] = str(e)
        return data

