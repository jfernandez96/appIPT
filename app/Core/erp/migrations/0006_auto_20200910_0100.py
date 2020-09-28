# Generated by Django 3.1 on 2020-09-10 06:00

from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('erp', '0005_auto_20200827_2256'),
    ]

    operations = [
        migrations.AlterField(
            model_name='archivocarga',
            name='UserId',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to=settings.AUTH_USER_MODEL, verbose_name='id de usuario login'),
        ),
        migrations.AlterField(
            model_name='tipoarchivocarga',
            name='Validar',
            field=models.BooleanField(blank=True, default=False, verbose_name='a validar'),
        ),
    ]