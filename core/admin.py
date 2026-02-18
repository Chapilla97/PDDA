from django.contrib import admin
from .models import Proyecto, RegistroActividad, Equipo, JustificacionSelectividad

@admin.register(Proyecto)
class ProyectoAdmin(admin.ModelAdmin):
    list_display = ('titulo', 'nombre_analista', 'estado', 'fecha_creacion')

@admin.register(RegistroActividad)
class RegistroActividadAdmin(admin.ModelAdmin):
    list_display = ('fecha', 'usuario', 'proyecto', 'accion')
    def has_add_permission(self, request): return False
    def has_change_permission(self, request, obj=None): return False
    def has_delete_permission(self, request, obj=None): return False

# --- NUEVOS CATÁLOGOS ---
@admin.register(Equipo)
class EquipoAdmin(admin.ModelAdmin):
    list_display = ('nombre', 'marca', 'activo')

@admin.register(JustificacionSelectividad)
class JustificacionSelectividadAdmin(admin.ModelAdmin):
    list_display = ('titulo', 'activo')