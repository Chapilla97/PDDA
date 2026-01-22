from django.urls import path
from . import views

urlpatterns = [
    # Ruta principal: El Dashboard
    path('', views.lista_proyectos, name='lista_proyectos'),
    
    # Ruta para crear nuevos proyectos (Solo Calidad)
    path('crear-proyecto/', views.crear_proyecto, name='crear_proyecto'),
    
    # Ruta para ver el detalle y subir archivos (La nueva "inicio")
    path('proyecto/<int:proyecto_id>/', views.detalle_proyecto, name='detalle_proyecto'),
]