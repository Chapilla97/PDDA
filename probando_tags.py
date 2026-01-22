from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    """
    Función auxiliar para recorrer todos los elementos del documento (párrafos y tablas)
    en orden, buscando dentro del XML.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Error al leer el documento")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extraer_tags(docx_path):
    document = Document(docx_path)
    datos_encontrados = {}

    # El namespace de Word para los tags
    namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    # Recorremos todo el XML del cuerpo del documento buscando las etiquetas 'sdt' (Structured Document Tag)
    for element in document.element.body.iter():
        if element.tag.endswith('sdt'):
            # Buscar el elemento 'tag' dentro de las propiedades del sdt
            sdt_pr = element.find(namespace + 'sdtPr')
            if sdt_pr is not None:
                tag_element = sdt_pr.find(namespace + 'tag')
                if tag_element is not None:
                    # Obtenemos el valor del Tag (la etiqueta que tú pusiste en Word)
                    tag_name = tag_element.get(namespace + 'val')
                    
                    # Ahora buscamos el contenido de texto dentro de ese control
                    sdt_content = element.find(namespace + 'sdtContent')
                    texto = ""
                    if sdt_content is not None:
                        # Extraemos todo el texto que haya dentro
                        texto = "".join([t.text for t in sdt_content.iter(namespace + 't') if t.text])
                    
                    datos_encontrados[tag_name] = texto
                    print(f"✅ ENCONTRADO -> Tag: [{tag_name}] | Valor actual: '{texto}'")

    return datos_encontrados

if __name__ == "__main__":
    # --- CONFIGURACIÓN ---
    nombre_archivo = "Protocolo.docx"  # <--- CAMBIA ESTO POR EL NOMBRE DE TU ARCHIVO REAL
    # ---------------------

    try:
        print(f"--- Analizando {nombre_archivo} ---")
        datos = extraer_tags(nombre_archivo)
        
        if not datos:
            print("⚠️ No encontré ningún Tag. Asegúrate de haber usado 'Propiedades' en el modo Programador.")
        else:
            print("\n--- RESUMEN DE EXTRACCIÓN ---")
            print(datos)
            
    except Exception as e:
        print(f"❌ Error: {e}")
        print("Asegúrate de que el archivo .docx está en la misma carpeta que este script.")