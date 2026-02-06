from django.contrib import admin
from django.urls import path, include # <--- AGREGAR 'include'
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    
    # RUTAS DE AUTENTICACIÓN (Login, Logout, Password Reset)
    # Esto habilita /accounts/login/ y /accounts/logout/
    path('accounts/', include('django.contrib.auth.urls')), 
    
    path('', include('core.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)