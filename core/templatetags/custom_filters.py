from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """
    Retorna el valor de un diccionario usando una clave dinámica.
    Devuelve el objeto real (dict) para permitir acceso a atributos anidados.
    """
    if not isinstance(dictionary, dict):
        return None
    return dictionary.get(key)

@register.filter
def calcular_progreso(proyecto):
    """
    Calcula el porcentaje de avance del proyecto basado en documentos cargados.
    """
    # 1. Si ya está terminado o tiene PDF final firmado -> 100%
    if proyecto.estado == 'terminado' or proyecto.informe_final_firmado:
        return 100

    # 2. Lista de documentos clave que suman puntos
    # (Basado en los 6 Excels principales + Protocolo + Informe)
    documentos_clave = [
        '01_LS', '02_IF', '03_Estabilidad', 
        '04_LM', '05_R', '06_S', 
        'Protocolo_Val', 'Informe_Val'
    ]
    
    total_docs = len(documentos_clave)
    encontrados = 0
    
    # Verificamos cuántos de estos existen en los datos guardados
    datos = proyecto.datos_validacion
    if datos and isinstance(datos, dict):
        for doc in documentos_clave:
            if doc in datos:
                encontrados += 1
    
    # 3. Calculamos porcentaje (Tope 90% si no está firmado)
    # Si tiene todo cargado pero no firmado, se queda en 90%
    if encontrados == 0:
        return 5 # Un 5% de "cortesía" por haber creado el proyecto
    
    porcentaje = int((encontrados / total_docs) * 90)
    
    return porcentaje