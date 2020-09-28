
from django.db import models
from django.forms import model_to_dict
from datetime import datetime
from app.settings import *
from Core.user.models import User
from django.conf import settings
# Create your models here.

def user_directory_path(request, filename):
    # file will be uploaded to MEDIA_ROOT / user_<id>/<filename>
    return '{0}/{1}'.format(request.user.id, filename)

class Sitios(models.Model):
    nombresitio = models.CharField(max_length=100, unique=True, verbose_name='Nombre')

    def __str__(self):
        return self.nombresitio

    def toJSON(self):
        item = model_to_dict(self)
        return item

    class Meta:
        verbose_name = 'Sitio'
        verbose_name_plural = 'Sitios'
        ordering = ['id']


class Aplicacion(models.Model):
    nombreAplicacion = models.CharField(max_length=30, unique=True, verbose_name='Nombre')

    def __str__(self):
        return self.nombreAplicacion

    def toJSON(self):
        item = model_to_dict(self)
        return item

    class Meta:
        verbose_name = 'Aplicacion'
        verbose_name_plural = 'Aplicaciones'
        ordering = ['id']


class TipoProceso(models.Model):
    AplicacionId = models.ForeignKey(Aplicacion, on_delete=models.CASCADE, verbose_name='Aplicacion')
    nombretipoproceso = models.CharField(max_length=50, unique=False, verbose_name='Nombre')

    def __str__(self):
        return self.nombretipoproceso

    def toJSON(self):
        item = model_to_dict(self)
        item['AplicacionId'] = self.AplicacionId.toJSON()
        return item

    class Meta:
        verbose_name = 'Tipo Proceso'
        verbose_name_plural = 'Tipos Procesos'
        ordering = ['id']


class tipoArchivoCarga(models.Model):
    tipoProcesoId = models.ForeignKey(TipoProceso, on_delete=models.CASCADE, verbose_name='Tipo Proceso')
    NombreTipoCarga = models.CharField(max_length=100, verbose_name='Nombre')
    limiteArchivo = models.IntegerField(default=0, verbose_name='Limite')
    TipoArchivopermitido = models.CharField(max_length=20,  default='.xlsx|.xls', verbose_name='Tipo archivo permitido')
    abreviatura = models.CharField(max_length=100, blank=True, verbose_name='abreviatura')
    Validar = models.BooleanField(default=False, blank=True, verbose_name="a validar")

    def __str__(self):
        return self.NombreTipoCarga

    def toJSON(self):
        item = model_to_dict(self)
        item['tipoProcesoId'] = self.tipoProcesoId.toJSON()
        return item

    class Meta:
        verbose_name = 'Tipo archivo carga'
        verbose_name_plural = 'Tipos de archivos de carga'
        ordering = ['id']


class ArchivoCarga(models.Model):
    UserId = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, verbose_name='id de usuario login')
    IdSitio = models.ForeignKey(Sitios, on_delete=models.CASCADE, verbose_name='Sitios')
    AplicacionId = models.ForeignKey(Aplicacion, on_delete=models.CASCADE, verbose_name='Aplicacion')
    FechaSubida = models.DateTimeField(default=datetime.now, verbose_name='Fecha de carga')
    Estado = models.BooleanField(default=True, verbose_name='Estado')

    def __str__(self):
        return self.FechaSubida

    def toJSON(self):
        item = model_to_dict(self)
        item['FechaSubida'] = self.FechaSubida.strftime('%d-%m-%Y %H:%M')
        item['IdSitio'] = self.IdSitio.toJSON()
        item['AplicacionId'] = self.AplicacionId.toJSON()
        return item

    class Meta:
        verbose_name = 'Archivo carga'
        verbose_name_plural = 'Archivos de carga'
        ordering = ['id']


class ArchivoCargaDetail(models.Model):

    ArchivoCargaId = models.ForeignKey(ArchivoCarga, on_delete=models.CASCADE, verbose_name='Archivo de carga')
    tipoArchivoCargaId = models.ForeignKey(tipoArchivoCarga, on_delete=models.CASCADE, verbose_name='tipo de archivo')
    NombreArchivo = models.CharField(max_length=100, blank=True, verbose_name='Nombre Archivo')
    ExtencionArchivo = models.CharField(max_length=100, blank=True, verbose_name='Extencion de archivo')
    ArchivoCheck = models.FileField(upload_to='%d_%m_%Y', blank=True, max_length=300, verbose_name='Archivos')

    def __str__(self):
        return self.NombreArchivo

    def get_file(self):
        if self.ArchivoCheck:
            return '{}{}'.format(MEDIA_URL, self.ArchivoCheck)
        return '{}{}'.format(STATIC_URL, '')

    def toJSON(self):
        item = model_to_dict(self, exclude=['tipoArchivoCargaId'])
        item['ArchivoCheck'] = '{}{}'.format(MEDIA_URL, self.ArchivoCheck)
        #item['ArchivoCargaId'] = self.ArchivoCargaId.toJSON()
        item['tipoArchivoCargaId'] = self.tipoArchivoCargaId.toJSON()
        return item

    class Meta:
        verbose_name = 'Archivo de carga con detalle'
        verbose_name_plural = 'Archivos de cargas con detalle'
        ordering = ['id']


