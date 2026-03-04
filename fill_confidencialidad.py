#!/usr/bin/env python3
"""
Genera el Acuerdo de Confidencialidad (.docx)
llenando datos del médico y empresa en la plantilla.

Placeholders:
  Comparecientes:
    - "NOMBRE FUNCIONARIO/A, SERVIDOR/A EL/LA TRABAJADOR/A con cargo" → Nombre médico + cargo
    - "cédula ... No. ………" → Cédula médico
  Cláusula 2:
    - "(incorporar nombre centro de trabajo) …" → Empresa
  Cláusula 8 (fecha):
    - "Quito Distrito Metropolitano al XXX de XXX de 2025" → Ciudad + fecha
"""

import os, shutil, zipfile, tempfile, re

PLANTILLA = os.path.join(os.path.dirname(__file__), 'confidencialidad_plantilla.docx')

MESES = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
}

def fill_confidencialidad(data, output_path):
    """
    data = {
        'medico': 'DR. LEONARDO ARIAS ESPINOZA',
        'cedula_medico': '0701234567',
        'cargo': 'Médico Ocupacional',
        'empresa': 'CONSTRUCTORA MACHALA S.A.',
        'ciudad': 'Machala',
        'fecha': '2026-03-03',
    }
    """
    medico = (data.get('medico', '') or '').upper()
    cedula_med = data.get('cedula_medico', '') or ''
    cargo = data.get('cargo', 'Médico Ocupacional')
    empresa = (data.get('empresa', '') or '').upper()
    ciudad = data.get('ciudad', 'Machala')
    fecha_str = data.get('fecha', '')
    
    # Formatear fecha
    if fecha_str:
        parts = fecha_str.split('-')
        if len(parts) == 3:
            anio = parts[0]
            mes = int(parts[1])
            dia = str(int(parts[2]))
            fecha_legible = f'{dia} de {MESES.get(mes, "")} de {anio}'
        else:
            fecha_legible = fecha_str
    else:
        from datetime import date
        hoy = date.today()
        fecha_legible = f'{hoy.day} de {MESES.get(hoy.month, "")} de {hoy.year}'
    
    # Descomprimir
    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(PLANTILLA, 'r') as z:
        z.extractall(tmpdir)
    
    doc_xml_path = os.path.join(tmpdir, 'word', 'document.xml')
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        xml = f.read()
    
    # === COMPARECIENTES: Nombre del médico ===
    # "NOMBRE FUNCIONARIO/A, SERVIDOR/A EL/LA TRABAJADOR/A con cargo"
    xml = xml.replace(
        'NOMBRE FUNCIONARIO/A, SERVIDOR/A EL/LA TRABAJADOR/A con cargo',
        f'{medico}, con cargo'
    )
    
    # === Cargo: "(médico, enfermera, psicólogo, odontólogo, trabajadora social)" → cargo real ===
    # Solo en comparecientes (primera ocurrencia con "y con cédula")
    xml = xml.replace(
        '(médico, enfermera, psicólogo, odontólogo, trabajadora social) y con cédula',
        f'{cargo} y con cédula'
    )
    
    # === Cédula: puntos después de "No. " ===
    # El patrón tiene proofErr: "No. …</w:t>...<w:t>………</w:t>...<w:t>…….</w:t>"
    xml = re.sub(
        r'No\. …</w:t></w:r>.*?……\.</w:t></w:r>',
        f'No. </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial MT" w:hAnsi="Arial MT"/><w:b/><w:bCs/><w:sz w:val="20"/></w:rPr><w:t>{cedula_med}</w:t></w:r>',
        xml,
        flags=re.DOTALL,
        count=1
    )
    
    # === Médico Ocupacional en comparecientes ===
    xml = xml.replace(
        'se denominará Médico Ocupacional',
        f'se denominará {cargo}'
    )
    
    # === CLÁUSULA 2: Centro de trabajo ===
    xml = xml.replace(
        '(incorporar nombre centro de trabajo) …',
        f'{empresa}'
    )
    
    # === CLÁUSULA 2: cargo en "EL/LA FUNCIONARIO/A..." ===
    xml = xml.replace(
        'EL/LA FUNCIONARIO/A, SERVIDOR/A EL/LA TRABAJADOR/A con cargo de Médico de Salud Ocupacional',
        f'{medico} con cargo de {cargo}'
    )
    
    # === Todas las ocurrencias de "(médico, enfermera, psicólogo, odontólogo, trabajadora social)" ===
    xml = xml.replace(
        '(médico, enfermera, psicólogo, odontólogo, trabajadora social)',
        cargo
    )
    # También la versión con "El" al inicio
    xml = xml.replace(
        'El (médico, enfermera, psicólogo, odontólogo, trabajadora social)',
        f'El {cargo}'
    )
    
    # === FECHA ===
    # El bloque es: "de Quito Distrito Metropolitano al XXX de XXX de </w:t>...</w:t>2025.</w:t>"
    # Reemplazar todo de una vez
    xml = re.sub(
        r'de Quito Distrito Metropolitano al XXX de XXX de </w:t></w:r><w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="20"/></w:rPr><w:t>2025\.</w:t>',
        f'de {ciudad}, al {fecha_legible}.</w:t>',
        xml,
        count=1
    )
    
    # Guardar
    with open(doc_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    # Reempaquetar
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmpdir):
            for file in files:
                filepath = os.path.join(root, file)
                arcname = os.path.relpath(filepath, tmpdir)
                zout.write(filepath, arcname)
    
    shutil.rmtree(tmpdir)
    return output_path


if __name__ == '__main__':
    result = fill_confidencialidad({
        'medico': 'DR. LEONARDO ARIAS ESPINOZA',
        'cedula_medico': '0701234567',
        'cargo': 'Médico Ocupacional',
        'empresa': 'CONSTRUCTORA MACHALA S.A.',
        'ciudad': 'Machala',
        'fecha': '2026-03-03',
    }, '/home/claude/test_confidencialidad.docx')
    print(f'✅ Generado: {result}')
