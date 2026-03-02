# PDF Formulario MSP - Microservicio

Genera PDFs del formulario oficial del MSP (Ministerio de Salud Pública) para evaluaciones médicas ocupacionales.

## Endpoints

### `POST /generar-pdf`
Recibe JSON con datos del paciente, devuelve PDF llenado.

### `POST /generar-xlsx`
Recibe JSON con datos del paciente, devuelve Excel llenado.

### `GET /health`
Health check.

## Deploy en Render.com (GRATIS)

### Opción 1: Desde GitHub
1. Sube esta carpeta a un repo de GitHub
2. Ve a [render.com](https://render.com) → New → Web Service
3. Conecta tu repo
4. Render detecta el Dockerfile automáticamente
5. Plan: Free
6. Deploy

### Opción 2: Manual
1. Ve a [render.com](https://render.com) → New → Web Service
2. Selecciona "Docker"
3. Sube los archivos
4. Environment: Docker
5. Port: 10000

## Ejemplo de uso desde JavaScript

```javascript
const response = await fetch('https://tu-servicio.onrender.com/generar-pdf', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    paciente: {
      apellido1: 'ARIAS',
      apellido2: 'ESPINOZA',
      nombre1: 'JORGE',
      nombre2: 'LEONARDO',
      cedula: '0705191229',
      sexo: 'Masculino',
      fecha_nacimiento: '1990-05-15',
      edad: '35'
    },
    empresa: {
      nombre: 'VitalWork',
      ruc: '0791234567001',
      ciiu: 'Q8690',
      establecimiento: 'Clínica Central'
    },
    profesional: {
      nombre: 'Dr. Juan Pérez',
      codigo: 'MSP-12345'
    },
    ficha: {
      b: { tipo: 'INGRESO', fecha: '2026-03-01', puesto_ciuo: 'Operador' },
      c: { clinicos: 'Sin antecedentes', familiares: 'Padre HTA' },
      d: { descripcion: 'Evaluación preocupacional' },
      e: { temperatura: '36.5', presion: '120/80', fc: '72', fr: '18', sat_o2: '98', peso: '75', talla: '170', imc: '25.9', perimetro: '85' },
      k: { diagnosticos: [{ codigo: 'Z00.0', descripcion: 'Examen médico general', tipo: 'DEF' }] },
      l: { aptitud: 'APTO', observaciones: '' },
      m: { descripcion: 'Control en 12 meses' }
    }
  })
});

const blob = await response.blob();
const url = URL.createObjectURL(blob);
window.open(url); // Abre el PDF
```

## Archivos
- `server.py` - Servidor Flask
- `fill_formulario.py` - Lógica de llenado del Excel
- `plantilla.xlsx` - Formulario MSP oficial (sin diagonales, con fit-to-page)
- `Dockerfile` - Imagen Docker con LibreOffice
- `render.yaml` - Configuración de deploy en Render
