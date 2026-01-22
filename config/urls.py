from django.contrib import admin
from django.urls import path, include
from django.conf import settings             # <--- IMPORTANTE
from django.conf.urls.static import static   # <--- IMPORTANTE (Faltaba este)

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('core.urls')),
] 

# Agregamos esto para servir archivos en modo DEBUG
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)