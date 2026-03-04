#!/usr/bin/env python3
"""
server.py - Microservicio para generar PDFs del formulario MSP
POST /generar-pdf → recibe JSON, devuelve PDF
"""
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fill_formulario import fill_formulario
from fill_consentimiento import fill_consentimiento
from fill_confidencialidad import fill_confidencialidad
import tempfile, os, uuid

app = Flask(__name__)
CORS(app)  # Permite llamadas desde tu app

TEMPLATE = os.path.join(os.path.dirname(__file__), 'plantilla.xlsx')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE)})

@app.route('/generar-pdf', methods=['POST'])
def generar_pdf():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No se recibieron datos JSON'}), 400
        
        # Modo: 'certificado' (hoja 1), 'formulario' (hojas 2,3,4), 'todo'
        modo = data.pop('_modo', 'todo')
        
        tmp_dir = tempfile.mkdtemp()
        pdf_path = os.path.join(tmp_dir, f'formulario_{uuid.uuid4().hex[:8]}.pdf')
        
        result = fill_formulario(data, TEMPLATE, pdf_path, modo=modo)
        
        if not os.path.exists(pdf_path):
            return jsonify({'error': 'Error generando PDF'}), 500
        
        a = data.get('a', {})
        nombre_archivo = f"{modo}_{a.get('ap1','')}_{a.get('n1','')}.pdf"
        nombre_archivo = nombre_archivo.replace(' ', '_').upper()
        
        return send_file(
            pdf_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=nombre_archivo
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generar-xlsx', methods=['POST'])
def generar_xlsx():
    """Alternativa: devuelve el Excel llenado en vez de PDF"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No se recibieron datos JSON'}), 400
        
        tmp_dir = tempfile.mkdtemp()
        xlsx_path = os.path.join(tmp_dir, f'formulario_{uuid.uuid4().hex[:8]}.xlsx')
        
        # Llenar Excel (sin convertir a PDF)
        fill_formulario(data, TEMPLATE, xlsx_path)
        
        pac = data.get('paciente', {})
        nombre = f"formulario_{pac.get('apellido1','')}_{pac.get('nombre1','')}.xlsx"
        
        return send_file(
            xlsx_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nombre.replace(' ', '_').upper()
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generar-confidencialidad', methods=['POST'])
def generar_confidencialidad():
    """Genera el Acuerdo de Confidencialidad (.docx)"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No se recibieron datos JSON'}), 400
        
        tmp_dir = tempfile.mkdtemp()
        docx_path = os.path.join(tmp_dir, f'confidencialidad_{uuid.uuid4().hex[:8]}.docx')
        
        fill_confidencialidad(data, docx_path)
        
        if not os.path.exists(docx_path):
            return jsonify({'error': 'Error generando confidencialidad'}), 500
        
        medico = data.get('medico', 'MEDICO').replace(' ', '_')
        nombre_archivo = f'CONFIDENCIALIDAD_{medico}.docx'.upper()
        
        return send_file(
            docx_path,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generar-consentimiento', methods=['POST'])
def generar_consentimiento():
    """Genera el documento de Consentimiento Informado (.docx)"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No se recibieron datos JSON'}), 400
        
        tmp_dir = tempfile.mkdtemp()
        docx_path = os.path.join(tmp_dir, f'consentimiento_{uuid.uuid4().hex[:8]}.docx')
        
        fill_consentimiento(data, docx_path)
        
        if not os.path.exists(docx_path):
            return jsonify({'error': 'Error generando consentimiento'}), 500
        
        nombre = data.get('nombre', 'PACIENTE').replace(' ', '_')
        nombre_archivo = f'CONSENTIMIENTO_{nombre}.docx'.upper()
        
        return send_file(
            docx_path,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
