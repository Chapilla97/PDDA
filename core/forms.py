from django import forms
from .models import Proyecto
from django.contrib.auth.models import User

class ProyectoForm(forms.ModelForm):
    class Meta:
        model = Proyecto
        fields = ['titulo', 'nombre_analista'] # Agregamos el analista
        
        widgets = {
            'titulo': forms.TextInput(attrs={
                'class': 'w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500', 
                'placeholder': 'Ej. Validación Lote X-2026'
            }),
            'nombre_analista': forms.TextInput(attrs={
                'class': 'w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500', 
                'placeholder': 'Nombre del Analista responsable'
            }),
        }
class CrearUsuarioForm(forms.Form):
    username = forms.CharField(label="Usuario (Login)", max_length=150, widget=forms.TextInput(attrs={'class': 'w-full border rounded px-3 py-2'}))
    first_name = forms.CharField(label="Nombre(s)", max_length=150, widget=forms.TextInput(attrs={'class': 'w-full border rounded px-3 py-2'}))
    last_name = forms.CharField(label="Apellidos", max_length=150, widget=forms.TextInput(attrs={'class': 'w-full border rounded px-3 py-2'}))
    
    ROL_CHOICES = [
        ('analista', 'Analista (Sin acceso al sistema, solo para asignar)'),
        ('calidad', 'Calidad (Con acceso completo al sistema)'),
    ]
    rol = forms.ChoiceField(choices=ROL_CHOICES, widget=forms.Select(attrs={'class': 'w-full border rounded px-3 py-2'}))
    
    password = forms.CharField(label="Contraseña (Solo para Calidad)", required=False, widget=forms.PasswordInput(attrs={'class': 'w-full border rounded px-3 py-2'}))