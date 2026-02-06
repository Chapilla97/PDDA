from django.contrib import messages
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import user_passes_test
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.conf import settings
from docxtpl import DocxTemplate
from docx import Document
import openpyxl
import os
from .models import Proyecto, RegistroActividad 
from .forms import ProyectoForm
import locale
from datetime import datetime
from django.contrib.auth.models import User
from jinja2 import Environment, Undefined # <--- IMPORTANTE PARA LIMPIAR VARIABLES VACÍAS

# --- CONFIGURACIÓN DE LOCALE (FECHAS EN ESPAÑOL) ---
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') 
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain') # Intento para Windows
    except:
        pass

# --- MOTOR DE EXTRACCIÓN (Igual que antes) ---
def extraer_tags_de_archivo(archivo_en_memoria):
    nombre_archivo = archivo_en_memoria.name.lower()
    datos_encontrados = {}

    print(f"\n--- 🕵️ INICIANDO ANÁLISIS DE: {nombre_archivo} ---")

    # --- CASO 1: WORD (.docx) ---
    if nombre_archivo.endswith('.docx'):
        try:
            document = Document(archivo_en_memoria)
            namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            for element in document.element.body.iter():
                if element.tag.endswith('sdt'):
                    sdt_pr = element.find(namespace + 'sdtPr')
                    if sdt_pr is not None:
                        tag_element = sdt_pr.find(namespace + 'tag')
                        if tag_element is not None:
                            tag_name = tag_element.get(namespace + 'val')
                            sdt_content = element.find(namespace + 'sdtContent')
                            texto = "".join([t.text for t in sdt_content.iter(namespace + 't') if t.text]) if sdt_content else ""
                            datos_encontrados[tag_name] = texto
        except Exception as e:
            print(f"Error DOCX: {e}")

    # --- CASO 2: EXCEL (.xlsx) ---
    elif nombre_archivo.endswith('.xlsx'):
        try:
            wb = openpyxl.load_workbook(archivo_en_memoria, data_only=True)
            hojas_reales = wb.sheetnames
            
            for nombre_rango, objeto_definicion in wb.defined_names.items():
                if nombre_rango.startswith('_xlnm') or nombre_rango.startswith('Print_Area'):
                    continue

                try:
                    try:
                        destinations = list(objeto_definicion.destinations)
                    except:
                        continue

                    for sheet_title, coord in destinations:
                        sheet_title_clean = sheet_title.strip("'")
                        
                        if sheet_title_clean not in hojas_reales:
                            if len(hojas_reales) == 1:
                                ws = wb[hojas_reales[0]]
                            else:
                                continue
                        else:
                            ws = wb[sheet_title_clean]
                        
                        celda_limpia = coord.replace('$', '')
                        if ':' in celda_limpia:
                            celda_limpia = celda_limpia.split(':')[0]
                            
                        try:
                            cell = ws[celda_limpia]
                            valor = cell.value
                            
                            fmt = cell.number_format
                            if isinstance(valor, (int, float)):
                                if '%' in fmt: valor = f"{valor * 100:.2f}%"
                                elif '0.0000' in fmt: valor = f"{valor:.4f}"
                                elif '0.000' in fmt: valor = f"{valor:.3f}"
                                elif '0.00' in fmt: valor = f"{valor:.2f}"
                                elif '0.0' in fmt: valor = f"{valor:.1f}"
                                elif fmt == '0': valor = f"{valor:.0f}"
                                else: valor = round(valor, 4)
                            
                            val_str = str(valor) if valor is not None else ""
                            datos_encontrados[nombre_rango] = val_str
                        except Exception:
                            pass 
                except Exception:
                    pass
        except Exception as e:
            print(f"Error General XLSX: {e}")

    return datos_encontrados

# --- VISTAS ---
@user_passes_test(lambda u: u.is_superuser)
def eliminar_proyecto(request, proyecto_id):
    proyecto = get_object_or_404(Proyecto, pk=proyecto_id)
    if request.method == 'POST':
        titulo = proyecto.titulo
        proyecto.delete()
        messages.error(request, f"🗑️ Proyecto '{titulo}' eliminado correctamente.")
        return redirect('lista_proyectos')
    return redirect('lista_proyectos')

@login_required
def lista_proyectos(request):
    proyectos = Proyecto.objects.all().order_by('-fecha_creacion')
    return render(request, 'core/dashboard.html', {'proyectos': proyectos})

# --- HELPER: REGISTRO DE AUDITORÍA ---
def registrar_log(proyecto, usuario, accion):
    RegistroActividad.objects.create(
        proyecto=proyecto,
        usuario=usuario,
        accion=accion
    )

@login_required
def crear_proyecto(request):
    # Paso 1: Obtener analistas para el selector
    analistas = User.objects.filter(is_active=True).order_by('username')
    
    if request.method == 'POST':
        titulo = request.POST.get('titulo')
        analista_id = request.POST.get('analista_id')
        
        if titulo:
            # Lógica para guardar el nombre bonito del analista
            analista_user = User.objects.get(pk=analista_id) if analista_id else None
            nombre_analista = f"{analista_user.first_name} {analista_user.last_name}" if analista_user else "Sin asignar"
            if analista_user and not analista_user.first_name: 
                nombre_analista = analista_user.username 
            
            proyecto = Proyecto.objects.create(
                titulo=titulo,
                nombre_analista=nombre_analista,
                creado_por=request.user
            )
            registrar_log(proyecto, request.user, "Creó el proyecto")
            messages.success(request, "✨ Proyecto creado exitosamente.")
            return redirect('detalle_proyecto', proyecto_id=proyecto.id)
            
    return render(request, 'core/crear_proyecto.html', {'analistas': analistas})

@login_required
def detalle_proyecto(request, proyecto_id):
    proyecto = get_object_or_404(Proyecto, pk=proyecto_id)
    
    docs_validacion = [('01_LS', 'Hoja: 01 LS'), ('02_IF', 'Hoja: 02 IF'), ('03_Estabilidad', 'Hoja: 03 Estabilidad'), ('04_LM', 'Hoja: 04 LM'), ('05_R', 'Hoja: 05 R'), ('06_S', 'Hoja: 06 S'), ('Protocolo_Val', 'Protocolo Validación'), ('Informe_Val', 'Informe Validación')]
    docs_estudio = [('Factor_Similitud', 'Hoja: Factor Similitud'), ('Porcentaje_Disuelto', 'Hoja: % Disuelto'), ('Protocolo_Perfiles', 'Protocolo Perfiles'), ('Informe_Perfiles', 'Informe Perfiles')]

    if request.method == 'POST':
        
        # A. SUBIDA DE ARCHIVOS DE DATOS
        if 'subir_archivo' in request.POST and proyecto.estado != 'terminado':
            tipo_doc = request.POST.get('tipo_documento')
            parte = request.POST.get('parte')
            archivo = request.FILES.get('archivo')
            
            if archivo and tipo_doc:
                datos = extraer_tags_de_archivo(archivo)
                
                # Caso Especial: Estabilidad (Datos Manuales)
                if tipo_doc == '03_Estabilidad':
                    campos_manuales = ['estabilidad_eau_horas', 'estabilidad_eau_p_inicial', 'estabilidad_eau_p_final', 'estabilidad_eau_diferencia']
                    for campo in campos_manuales:
                        valor = request.POST.get(campo)
                        if valor:
                            datos[campo] = valor
                
                if parte == 'validacion':
                    proyecto.datos_validacion[tipo_doc] = datos
                elif parte == 'estudio':
                    proyecto.datos_estudio[tipo_doc] = datos
                proyecto.save()
                
                registrar_log(proyecto, request.user, f"Subió/Actualizó documento: {tipo_doc}")
                messages.success(request, f"✅ Documento {tipo_doc} cargado y datos extraídos.")
                return redirect('detalle_proyecto', proyecto_id=proyecto.id)

        # B. NUEVO: GUARDAR DATOS GENERALES (FECHAS Y TÉCNICA) -> PASOS 2 Y 3
        if 'guardar_datos_generales' in request.POST:
            datos = {
                'fecha_inicio': request.POST.get('fecha_inicio'),
                'fecha_fin': request.POST.get('fecha_fin'),
                'fecha_emision': request.POST.get('fecha_emision'),
                'tecnica': request.POST.get('tecnica'), # 'croma' o 'espectro'
                'equipo_modelo': request.POST.get('equipo_modelo'),
                'selectividad_opcion': request.POST.get('selectividad_opcion'),
                'selectividad_texto': request.POST.get('selectividad_texto'),
            }

            # Lógica de Fechas
            def formatear_fecha(str_fecha, tipo):
                if not str_fecha: return ""
                try:
                    dt = datetime.strptime(str_fecha, '%Y-%m-%d')
                    mes = dt.strftime('%B').capitalize()
                    if tipo == 'periodo': return f"{dt.day}"
                    elif tipo == 'completa': return f"{dt.day} de {mes} de {dt.year}"
                except:
                    return str_fecha

            # Periodo: "Del X al Y de Mes de Año"
            f_ini = datos['fecha_inicio']
            f_fin = datos['fecha_fin']
            try:
                dt_ini = datetime.strptime(f_ini, '%Y-%m-%d')
                dt_fin = datetime.strptime(f_fin, '%Y-%m-%d')
                mes_fin = dt_fin.strftime('%B').capitalize()
                datos['periodo_validacion_txt'] = f"Del {dt_ini.day} al {dt_fin.day} de {mes_fin} de {dt_fin.year}"
            except:
                datos['periodo_validacion_txt'] = ""

            # Emisión
            datos['fecha_emision_txt'] = formatear_fecha(datos['fecha_emision'], 'completa')

            # Lógica de Técnica (Cromatografía vs Espectro)
            if datos['tecnica'] == 'croma':
                atos['var_tecnica_nombre1'] = "Cromatografo"
                datos['var_tecnica_nombre'] = "Cromatografía"
                datos['var_tecnica_adj_pl'] = "Cromatográficas"
                datos['var_tecnica_adj_sg'] = "Cromatográfico"
                datos['var_unidades'] = "unidades de área"
                datos['bloque_estabilidad_automuestreador'] = "Estabilidad en automuestreador"
                
            elif datos['tecnica'] == 'espectro':
                datos['var_tecnica_nombre1'] = "Espectrofotometro"
                datos['var_tecnica_nombre'] = "Espectrofotometría"
                datos['var_tecnica_adj_pl'] = "Espectrofotométricas"
                datos['var_tecnica_adj_sg'] = "Espectrofotométrico"
                datos['var_unidades'] = "unidades de absorbancia"
                datos['bloque_estabilidad_automuestreador'] = "" # Se borra en Espectro
            
            else:
                datos['var_tecnica_nombre'] = ""
                datos['var_tecnica_adj_pl'] = ""
                datos['var_tecnica_adj_sg'] = ""
                datos['var_unidades'] = ""
                datos['bloque_estabilidad_automuestreador'] = ""

            # Guardar en el JSON del proyecto
            if not isinstance(proyecto.datos_validacion, dict):
                proyecto.datos_validacion = {}
            
            proyecto.datos_validacion['datos_generales'] = datos
            proyecto.save()
            
            registrar_log(proyecto, request.user, "Actualizó Datos Generales del Informe")
            messages.success(request, "💾 Datos generales guardados y procesados.")
            return redirect('detalle_proyecto', proyecto_id=proyecto.id)

        # C. GENERAR INFORME WORD (Estudio - Botón Derecho)
        if 'generar_informe' in request.POST:
            registrar_log(proyecto, request.user, "Generó el Informe Final (Word)")
            datos_finales = {**proyecto.datos_validacion, **proyecto.datos_estudio}
            datos_planos = {}
            for doc in datos_finales.values():
                if isinstance(doc, dict): datos_planos.update(doc)
            return generar_documento_descarga(datos_planos)

        # D. SUBIR PDF FINAL FIRMADO
        if 'subir_pdf_firmado' in request.POST:
            archivo_pdf = request.FILES.get('archivo_pdf')
            if archivo_pdf:
                proyecto.informe_final_firmado = archivo_pdf
                proyecto.fecha_subida_firmado = timezone.now()
                proyecto.save()
                registrar_log(proyecto, request.user, "Subió el PDF Final Firmado")
                messages.success(request, "🚀 ¡Felicidades! Informe firmado cargado. Proyecto Completado.")
                return redirect('detalle_proyecto', proyecto_id=proyecto.id)

        # E. MOVER A TERMINADOS
        if 'finalizar_proyecto' in request.POST:
            proyecto.estado = 'terminado'
            proyecto.save()
            registrar_log(proyecto, request.user, "Finalizó el proyecto (Movido a Histórico)")
            messages.info(request, "📦 Proyecto movido al histórico.")
            return redirect('lista_proyectos')
            
        # F. REACTIVAR
        if 'reactivar_proyecto' in request.POST:
            proyecto.estado = 'en_proceso'
            proyecto.save()
            registrar_log(proyecto, request.user, "Reactivó el proyecto")
            messages.warning(request, "⚠️ Proyecto reactivado para edición.")
            return redirect('detalle_proyecto', proyecto_id=proyecto.id)
        
    historial = proyecto.actividades.all().order_by('-fecha')

    return render(request, 'core/detalle_proyecto.html', {
        'proyecto': proyecto,
        'docs_validacion': docs_validacion,
        'docs_estudio': docs_estudio,
        'historial': historial
    })

def generar_documento_descarga(contexto_datos):
    ruta_plantilla = os.path.join(settings.BASE_DIR, 'Plantilla_Informe.docx')
    if not os.path.exists(ruta_plantilla):
        return HttpResponse("Error: Falta la Plantilla_Informe.docx", status=404)

    doc = DocxTemplate(ruta_plantilla)
    contexto_seguro = {k: str(v) for k, v in contexto_datos.items()}
    doc.render(contexto_seguro)
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="Informe_Final_Completo.docx"'
    doc.save(response)
    return response

# --- CLASE HELPER PARA VARIABLES VACÍAS ---
class SilentUndefined(Undefined):
    def __str__(self):
        return ""

@login_required
def generar_informe_validacion(request, proyecto_id):
    proyecto = get_object_or_404(Proyecto, pk=proyecto_id)
    ruta_plantilla = os.path.join(settings.BASE_DIR, 'Plantilla_Validacion.docx')
    
    if not os.path.exists(ruta_plantilla):
        return HttpResponse("Error: No encuentro 'Plantilla_Validacion.docx' en la carpeta raíz.", status=404)

    # 1. Recopilar datos de Validación (Excels)
    datos_finales = {}
    if proyecto.datos_validacion and isinstance(proyecto.datos_validacion, dict):
        for doc in proyecto.datos_validacion.values():
            if isinstance(doc, dict):
                datos_finales.update(doc)
    
    # 2. Asegurar que los datos generales estén incluidos
    if proyecto.datos_validacion and 'datos_generales' in proyecto.datos_validacion:
        datos_finales.update(proyecto.datos_validacion['datos_generales'])

    # 3. Sanitizar
    contexto = {k: str(v) for k, v in datos_finales.items()}

    try:
        doc = DocxTemplate(ruta_plantilla)
        
        # 4. CONFIGURAR JINJA2 PARA LIMPIAR VARIABLES VACÍAS (Paso 4)
        jinja_env = Environment(undefined=SilentUndefined)
        
        doc.render(contexto, jinja_env)
        
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        nombre_archivo = f"Borrador_Validacion_{proyecto.titulo}.docx"
        response['Content-Disposition'] = f'attachment; filename="{nombre_archivo}"'
        doc.save(response)
        
        registrar_log(proyecto, request.user, "Generó Borrador de Informe Validación")
        return response

    except Exception as e:
        return HttpResponse(f"Error generando Word: {e}", status=500)