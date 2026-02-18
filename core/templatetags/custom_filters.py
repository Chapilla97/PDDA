from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    if not isinstance(dictionary, dict):
        return None
    return dictionary.get(key)

@register.filter
def calcular_progreso(proyecto):
    # 1. Si ya está terminado o tiene PDF final firmado -> 100%
    if proyecto.estado == 'terminado' or proyecto.informe_final_firmado:
        return 100

    # 2. Definir Listas de Requisitos
    docs_validacion_req = ['01_LS', '02_IF', '03_Estabilidad', '04_LM', '05_R', '06_S', 'Protocolo_Val']
    docs_estudio_req = ['Factor_Similitud', 'Porcentaje_Disuelto', 'Protocolo_Perfiles']
    
    # 3. Contar cumplimiento
    total_items = len(docs_validacion_req) + len(docs_estudio_req) + 1 # +1 por la Configuración
    encontrados = 0
    
    # Verificar Validación
    datos_val = proyecto.datos_validacion
    if datos_val:
        for doc in docs_validacion_req:
            if doc in datos_val: encontrados += 1
        # Verificar Configuración (Datos Generales)
        if 'datos_generales' in datos_val: encontrados += 1

    # Verificar Estudio
    datos_est = proyecto.datos_estudio
    if datos_est:
        for doc in docs_estudio_req:
            if doc in datos_est: encontrados += 1
    
    # 4. Cálculo final (Tope 95% hasta que se firme)
    if encontrados == 0: return 5
    
    porcentaje = int((encontrados / total_items) * 95)
    return min(porcentaje, 95)