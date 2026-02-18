from django.urls import path
from . import views

urlpatterns = [
    # Rutas existentes
    path('', views.lista_proyectos, name='lista_proyectos'),
    path('crear-proyecto/', views.crear_proyecto, name='crear_proyecto'),
    path('proyecto/<int:proyecto_id>/', views.detalle_proyecto, name='detalle_proyecto'),
    path('eliminar/<int:proyecto_id>/', views.eliminar_proyecto, name='eliminar_proyecto'),
    path('generar-validacion/<int:proyecto_id>/', views.generar_informe_validacion, name='generar_informe_validacion'),
    
    # NUEVAS RUTAS DE USUARIOS
    path('proyecto/<int:proyecto_id>/audit-pdf/', views.exportar_audit_trail_pdf, name='exportar_audit_trail_pdf'),
    path('usuarios/', views.administrar_usuarios, name='administrar_usuarios'),
    path('usuarios/nuevo/', views.crear_usuario_nuevo, name='crear_usuario_nuevo'),
]