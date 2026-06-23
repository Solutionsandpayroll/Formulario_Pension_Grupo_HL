# Formulario de Pensión - Grupo HL

Aplicación web para la generación automática de formularios PDF de pensión para el Grupo HL, utilizando datos desde archivos Excel.

## Características Principales

- **Carga de archivos Excel** (.xlsx/.xls) con datos de empleados mediante drag & drop o selección manual
- **Mapeo automático** de columnas Excel a campos del formulario PDF (nombres, apellidos, cédula, ciudad, cargo, etc.)
- **Formato de fechas colombiano** (día-mes-año) con parseo automático y dibujado dígito por dígito en las casillas del PDF
- **Campos manuales editables** para datos que no provienen del Excel (NIT, razón social, género, fecha de expedición, etc.)
- **Generación individual** del PDF por empleado con vista previa integrada
- **Generación masiva** de todos los formularios en un solo archivo PDF
- **Descarga** de PDF individual o combinado con nombres personalizados
- **Mapeo de columnas editable** para ajustar la correspondencia Excel → PDF
- **Valores personalizados opcionales** por campo para sobreescribir datos del Excel
- **Interfaz responsive** con diseño corporativo de Solutions & Payroll

## Estructura del PDF (3 páginas)

| Página | Contenido |
|--------|-----------|
| 1 | Autorización: Nombre, N° Documento |
| 2 | Formulario principal: datos personales, residencia, información bancaria |
| 3 | Firmas y constancia: Nombre, N° Documento |

## Campos del Formulario

### Página 1 - Autorización
- Nombres y Apellidos
- Número de Documento

### Página 2 - Datos Principales
- N° de NIT
- Razón Social Entidad Patrocinadora
- N° de Documento de Identidad
- Fecha de Expedición (dígitos individuales)
- Lugar de Expedición
- Primer Apellido / Otros Apellidos
- Primer Nombre / Otros Nombres
- Género (M/F)
- Lugar de Nacimiento
- Fecha de Nacimiento (dígitos individuales)
- Cargo / Ocupación u Oficio
- Profesión
- Dirección de Residencia
- Ciudad / Municipio
- Departamento
- Teléfono Celular
- Correo Electrónico
- Número de Cuenta
- Nombre del Banco

### Página 3 - Firmas
- Nombres y Apellidos
- N° Documento de Identidad

## Tecnologías

- **React 18** + Vite
- **pdf-lib** para manipulación y generación de PDFs
- **xlsx** para lectura de archivos Excel
- **PDF.js** para renderizado de vista previa

## Requisitos

- Node.js 16+
- npm 7+

## Instalación y Uso

```bash
# Instalar dependencias
npm install

# Desarrollo
npm run dev

# Build de producción
npm run build

# Preview del build
npm run preview
```

## Archivos en `/public`

| Archivo | Propósito |
|---------|-----------|
| `Archivo final.pdf` | Plantilla PDF base del formulario |
| `Excel inicial.xlsx` | Plantilla Excel de ejemplo con datos de empleados |
| `Logo syp.png` | Logo corporativo |

## Flujo de Trabajo

1. Cargar el archivo Excel con los datos de empleados
2. Revisar y ajustar los campos manuales (NIT, razón social, género, etc.)
3. Verificar el mapeo de columnas Excel → PDF en la tabla de mapeo editable
4. Seleccionar un empleado de la tabla
5. Generar PDF individual o generar todos los formularios
6. Descargar el PDF resultante

---

© 2026 Solutions & Payroll. Todos los derechos reservados.
