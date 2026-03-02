#!/usr/bin/env python3
"""
server.py - Microservicio para generar PDFs del formulario MSP
POST /generar-pdf → recibe JSON, devuelve PDF
"""
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from fill_formulario import fill_formulario
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
        
        # Crear archivo temporal para el PDF
        tmp_dir = tempfile.mkdtemp()
        pdf_path = os.path.join(tmp_dir, f'formulario_{uuid.uuid4().hex[:8]}.pdf')
        
        # Llenar y generar PDF
        result = fill_formulario(data, TEMPLATE, pdf_path)
        
        if not os.path.exists(pdf_path):
            return jsonify({'error': 'Error generando PDF'}), 500
        
        # Nombre del archivo: apellido_nombre_fecha
        pac = data.get('paciente', {})
        nombre_archivo = f"formulario_{pac.get('apellido1','')}_{pac.get('nombre1','')}.pdf"
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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
