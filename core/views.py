from django.shortcuts import render, redirect, get_object_or_404
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


# --- MOTOR DE EXTRACCIÓN (Igual que antes) ---
import openpyxl
from docx import Document

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
                            # print(f"   [DOCX] ✅ Leído: {tag_name}")
        except Exception as e:
            print(f"Error DOCX: {e}")

    # --- CASO 2: EXCEL (.xlsx) ---
    elif nombre_archivo.endswith('.xlsx'):
        try:
            wb = openpyxl.load_workbook(archivo_en_memoria, data_only=True)
            hojas_reales = wb.sheetnames
            print(f"   [XLSX] Hojas reales en el archivo: {hojas_reales}")
            
            for nombre_rango, objeto_definicion in wb.defined_names.items():
                
                if nombre_rango.startswith('_xlnm') or nombre_rango.startswith('Print_Area'):
                    continue

                try:
                    # openpyxl devuelve un generador
                    try:
                        destinations = list(objeto_definicion.destinations)
                    except:
                        continue

                    for sheet_title, coord in destinations:
                        sheet_title_clean = sheet_title.strip("'")
                        
                        # --- LÓGICA DE AUTO-CORRECCIÓN ---
                        # Si la hoja que busca el tag NO existe...
                        if sheet_title_clean not in hojas_reales:
                            # ...pero el archivo solo tiene 1 hoja, asumimos que es esa.
                            if len(hojas_reales) == 1:
                                # print(f"   [XLSX] 🔧 Redireccionando '{nombre_rango}' de '{sheet_title_clean}' a '{hojas_reales[0]}'")
                                ws = wb[hojas_reales[0]]
                            else:
                                # Si hay muchas hojas y no coincide el nombre, no podemos adivinar.
                                print(f"   [XLSX] ⚠️ SALTADO '{nombre_rango}': No encuentro la hoja '{sheet_title_clean}'")
                                continue
                        else:
                            # Si coincide, todo normal
                            ws = wb[sheet_title_clean]
                        
                        # Limpieza de coordenadas
                        celda_limpia = coord.replace('$', '')
                        
                        # Manejo de Celdas Combinadas (Rango A1:B2 -> Tomar A1)
                        if ':' in celda_limpia:
                            celda_limpia = celda_limpia.split(':')[0]
                            
                        # Lectura del Valor
                        try:
                            cell = ws[celda_limpia]
                            valor = cell.value
                            
                            # Formato WYSIWYG
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
                            pass # Si falla leer la celda, ignoramos silenciosamente

                except Exception:
                    pass

        except Exception as e:
            print(f"Error General XLSX: {e}")

    print(f"--- FIN: {len(datos_encontrados)} variables extraídas ---\n")
    return datos_encontrados
# --- VISTAS ---

@login_required
def lista_proyectos(request):
    # Todos son calidad, ven todo
    proyectos = Proyecto.objects.all().order_by('-fecha_creacion')
    return render(request, 'core/dashboard.html', {'proyectos': proyectos})
# --- HELPER: REGISTRO DE AUDITORÍA ---
def registrar_log(proyecto, usuario, accion):
    # Esta función ahora sí encontrará el modelo porque ya lo importamos arriba
    RegistroActividad.objects.create(
        proyecto=proyecto,
        usuario=usuario,
        accion=accion
    )

@login_required
def crear_proyecto(request):
    if request.method == 'POST':
        form = ProyectoForm(request.POST)
        if form.is_valid():
            proyecto = form.save(commit=False)
            proyecto.creado_por = request.user
            proyecto.save()
            # LOG
            registrar_log(proyecto, request.user, "Creó el proyecto")
            return redirect('detalle_proyecto', proyecto_id=proyecto.id)
    else:
        form = ProyectoForm()
    return render(request, 'core/crear_proyecto.html', {'form': form})

@login_required
def detalle_proyecto(request, proyecto_id):
    proyecto = get_object_or_404(Proyecto, pk=proyecto_id)
    
    # Listas de documentos (Igual que antes)
    docs_validacion = [('01_LS', 'Hoja: 01 LS'), ('02_IF', 'Hoja: 02 IF'), ('03_Estabilidad', 'Hoja: 03 Estabilidad'), ('04_LM', 'Hoja: 04 LM'), ('05_R', 'Hoja: 05 R'), ('06_S', 'Hoja: 06 S'), ('Protocolo_Val', 'Protocolo Validación'), ('Informe_Val', 'Informe Validación')]
    docs_estudio = [('Factor_Similitud', 'Hoja: Factor Similitud'), ('Porcentaje_Disuelto', 'Hoja: % Disuelto'), ('Protocolo_Perfiles', 'Protocolo Perfiles'), ('Informe_Perfiles', 'Informe Perfiles')]

    if request.method == 'POST':
        
        # A. SUBIDA DE ARCHIVOS DE DATOS (Solo si NO está terminado)
        if 'subir_archivo' in request.POST and proyecto.estado != 'terminado':
            tipo_doc = request.POST.get('tipo_documento')
            parte = request.POST.get('parte')
            archivo = request.FILES.get('archivo')
            
            if archivo and tipo_doc:
                datos = extraer_tags_de_archivo(archivo)
                if parte == 'validacion':
                    proyecto.datos_validacion[tipo_doc] = datos
                elif parte == 'estudio':
                    proyecto.datos_estudio[tipo_doc] = datos
                proyecto.save()
                
                # LOG
                registrar_log(proyecto, request.user, f"Subió/Actualizó documento: {tipo_doc}")
                return redirect('detalle_proyecto', proyecto_id=proyecto.id)

        # B. GENERAR INFORME (Siempre permitido, crea log)
        if 'generar_informe' in request.POST:
            registrar_log(proyecto, request.user, "Generó el Informe Final (Word)")
            datos_finales = {**proyecto.datos_validacion, **proyecto.datos_estudio} # Unir todo
            # Aplanar diccionarios anidados
            datos_planos = {}
            for doc in datos_finales.values():
                datos_planos.update(doc)
            return generar_documento_descarga(datos_planos)

        # C. SUBIR PDF FINAL FIRMADO (Acción Final)
        if 'subir_pdf_firmado' in request.POST:
            archivo_pdf = request.FILES.get('archivo_pdf')
            if archivo_pdf:
                proyecto.informe_final_firmado = archivo_pdf
                proyecto.fecha_subida_firmado = timezone.now()
                proyecto.save()
                registrar_log(proyecto, request.user, "Subió el PDF Final Firmado")
                return redirect('detalle_proyecto', proyecto_id=proyecto.id)

        # D. MOVER A TERMINADOS / FINALIZAR
        if 'finalizar_proyecto' in request.POST:
            proyecto.estado = 'terminado'
            proyecto.save()
            registrar_log(proyecto, request.user, "Finalizó el proyecto (Movido a Histórico)")
            return redirect('lista_proyectos')
            
        # E. REACTIVAR PROYECTO (Por si hubo error)
        if 'reactivar_proyecto' in request.POST:
            proyecto.estado = 'en_proceso'
            proyecto.save()
            registrar_log(proyecto, request.user, "Reactivó el proyecto")
            return redirect('detalle_proyecto', proyecto_id=proyecto.id)

    # Obtenemos el historial para mostrarlo
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
    # Convertimos todo a string por si acaso
    contexto_seguro = {k: str(v) for k, v in contexto_datos.items()}
    
    doc.render(contexto_seguro)
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="Informe_Final_Completo.docx"'
    doc.save(response)
    return response