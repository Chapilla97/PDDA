from django.contrib import admin
from .models import Proyecto, RegistroActividad

# Registramos el Proyecto
@admin.register(Proyecto)
class ProyectoAdmin(admin.ModelAdmin):
    list_display = ('titulo', 'nombre_analista', 'estado', 'fecha_creacion')
    list_filter = ('estado',)
    search_fields = ('titulo', 'nombre_analista')

# Registramos el Audit Trail (Solo lectura para seguridad)
@admin.register(RegistroActividad)
class RegistroActividadAdmin(admin.ModelAdmin):
    list_display = ('fecha', 'usuario', 'proyecto', 'accion')
    list_filter = ('usuario', 'fecha')
    search_fields = ('accion', 'proyecto__titulo')
    
    # Esto hace que nadie pueda modificar el historial desde el admin (Seguridad)
    def has_add_permission(self, request):
        return False
    def has_change_permission(self, request, obj=None):
        return False
    def has_delete_permission(self, request, obj=None):
        return False