from django.db import models
from django.contrib.auth.models import User

class Proyecto(models.Model):
    titulo = models.CharField(max_length=200)
    nombre_analista = models.CharField(max_length=150, blank=True, null=True, verbose_name="Analista Asignado")
    creado_por = models.ForeignKey(User, on_delete=models.CASCADE, related_name='proyectos_creados')
    fecha_creacion = models.DateTimeField(auto_now_add=True)
    
    # Datos JSON
    datos_validacion = models.JSONField(default=dict, blank=True)
    datos_estudio = models.JSONField(default=dict, blank=True)

    # Archivo Final (PDF Firmado)
    informe_final_firmado = models.FileField(upload_to='informes_firmados/', null=True, blank=True)
    fecha_subida_firmado = models.DateTimeField(null=True, blank=True)

    # Estado
    estado = models.CharField(max_length=20, default='en_proceso', choices=[
        ('en_proceso', 'En Proceso'),
        ('terminado', 'Terminado (Histórico)')
    ])

    def __str__(self):
        return self.titulo

class RegistroActividad(models.Model):
    proyecto = models.ForeignKey(Proyecto, on_delete=models.CASCADE, related_name='actividades')
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    accion = models.CharField(max_length=255) 
    # Nuevo campo para comentarios largos
    comentario = models.TextField(blank=True, null=True) 
    fecha = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.usuario} - {self.accion}"
# --- CATÁLOGOS DINÁMICOS ---
class Equipo(models.Model):
    nombre = models.CharField(max_length=100) # Ej. Agilent 1260
    marca = models.CharField(max_length=100, blank=True, null=True)
    activo = models.BooleanField(default=True)

    def __str__(self):
        return f"{self.nombre} ({self.marca})" if self.marca else self.nombre

class JustificacionSelectividad(models.Model):
    titulo = models.CharField(max_length=100) # Ej. "Opción 1: Sin interferencias"
    descripcion = models.TextField() # El texto largo que se pone en el informe
    activo = models.BooleanField(default=True)

    def __str__(self):
        return self.titulo