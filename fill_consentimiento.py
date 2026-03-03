#!/usr/bin/env python3
"""
Genera el documento de Consentimiento Informado (.docx)
llenando los datos del paciente en la plantilla original.

Placeholders en el XML:
  Cláusula 5:
    - "…………………………………………," → Nombre del paciente
    - "…………………………," → Cédula + texto
    - "Empresa/Institución ……………………………………" → Proveedor (VitalWork S.A.S.)
  Cláusula 6:
    - "empresa ……………..." → Nombre de la empresa
  Fecha final:
    - "Guayaquil, " → Ciudad
    - "XXX de XXX de 2025." → Fecha
  Firma:
    - "……………………………." → Nombre del paciente
"""

import os, shutil, zipfile, tempfile, re

PLANTILLA = os.path.join(os.path.dirname(__file__), 'consentimiento_plantilla.docx')

MESES = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
}

def fill_consentimiento(data, output_path):
    """
    data = {
        'nombre': 'GARCÍA LÓPEZ MARÍA ISABEL',
        'cedula': '0701234567',
        'empresa': 'CONSTRUCTORA MACHALA S.A.',
        'proveedor': 'VitalWork S.A.S.',
        'ciudad': 'Machala',
        'fecha': '2026-03-03',  # AAAA-MM-DD
    }
    """
    nombre = (data.get('nombre', '') or '').upper()
    cedula = data.get('cedula', '') or ''
    empresa = (data.get('empresa', '') or '').upper()
    proveedor = data.get('proveedor', 'VitalWork S.A.S.')
    ciudad = data.get('ciudad', 'Machala')
    fecha_str = data.get('fecha', '')
    
    # Formatear fecha: "3 de marzo de 2026"
    if fecha_str:
        parts = fecha_str.split('-')
        if len(parts) == 3:
            anio = parts[0]
            mes = int(parts[1])
            dia = str(int(parts[2]))  # quitar cero al inicio
            fecha_legible = f'{dia} de {MESES.get(mes, "")} de {anio}'
        else:
            fecha_legible = fecha_str
    else:
        from datetime import date
        hoy = date.today()
        fecha_legible = f'{hoy.day} de {MESES.get(hoy.month, "")} de {hoy.year}'
    
    # Descomprimir docx
    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(PLANTILLA, 'r') as z:
        z.extractall(tmpdir)
    
    # Leer document.xml
    doc_xml_path = os.path.join(tmpdir, 'word', 'document.xml')
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        xml = f.read()
    
    # === CLÁUSULA 5 ===
    # Reemplazar nombre: "…………………………………………," → "NOMBRE,"
    xml = xml.replace('…………………………………………,', f'{nombre},')
    
    # Reemplazar cédula + proveedor en el párrafo siguiente
    # "…………………………, otorgo mi consentimiento... Empresa/Institución …………………………………… y al"
    xml = xml.replace(
        '…………………………, otorgo mi consentimiento libre, previo, informado, específico e inequívoco a la Empresa/Institución …………………………………… y al profesional médico correspondiente',
        f'{cedula}, otorgo mi consentimiento libre, previo, informado, específico e inequívoco a la Empresa/Institución {proveedor} y al profesional médico correspondiente'
    )
    
    # === CLÁUSULA 6 ===
    # "empresa ……………..." → "empresa NOMBRE_EMPRESA"
    xml = xml.replace('……………...</w:t>', f'{empresa}</w:t>')
    
    # === FECHA FINAL ===
    # Reemplazar "Guayaquil, " → "Ciudad, "
    xml = xml.replace('Guayaquil, </w:t>', f'{ciudad}, </w:t>')
    
    # Reemplazar bloque "al XXX de XXX de 2025." con regex
    # El patrón en XML compactado es: <w:t>XXX</w:t>...<w:t>de</w:t>...<w:t>XXX</w:t>...<w:t>de</w:t>...<w:t>2025.</w:t>
    # Reemplazamos desde "al</w:t>" hasta "2025.</w:t>" 
    import re
    pattern = r'(<w:t[^>]*>)al(</w:t>)(.*?)<w:t[^>]*>XXX</w:t>.*?<w:t[^>]*>2025\.</w:t>'
    replacement = rf'\1al {fecha_legible}.\2'
    xml = re.sub(pattern, replacement, xml, flags=re.DOTALL)
    
    # === FIRMA ===
    # "……………………………. " → "NOMBRE "
    xml = xml.replace(
        '<w:t xml:space="preserve">……………………………. </w:t>',
        f'<w:t xml:space="preserve">{nombre} </w:t>'
    )
    
    # Guardar XML modificado
    with open(doc_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    # Reempaquetar como docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmpdir):
            for file in files:
                filepath = os.path.join(root, file)
                arcname = os.path.relpath(filepath, tmpdir)
                zout.write(filepath, arcname)
    
    # Limpiar
    shutil.rmtree(tmpdir)
    
    return output_path


# Test
if __name__ == '__main__':
    result = fill_consentimiento({
        'nombre': 'GARCÍA LÓPEZ MARÍA ISABEL',
        'cedula': '0701234567',
        'empresa': 'CONSTRUCTORA MACHALA S.A.',
        'proveedor': 'VitalWork S.A.S.',
        'ciudad': 'Machala',
        'fecha': '2026-03-03',
    }, '/home/claude/test_consentimiento.docx')
    print(f'✅ Generado: {result}')
