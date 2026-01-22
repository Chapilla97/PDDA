from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    # Verificamos que sea un diccionario de verdad
    if not isinstance(dictionary, dict):
        return None
    
    val = dictionary.get(key)
    # Solo devolvemos .items() si encontramos un valor (que debe ser otro dict)
    if val:
        return val.items()
    return None