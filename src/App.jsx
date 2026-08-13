import { useState, useRef, useCallback, useEffect, useMemo } from 'react'
import * as XLSX from 'xlsx'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import './App.css'

const MESES_ESP = {
  'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5,
  'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11
}

function excelSerialToDate(serial) {
  const ms = (serial - 25569) * 86400 * 1000
  const d = new Date(ms)
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())
}

function formatDateYMD(d) {
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`
}

function parseFechaColombia(str) {
  if (str === null || str === undefined || str === '') return ''

  if (str instanceof Date) {
    return isNaN(str.getTime()) ? '' : formatDateYMD(str)
  }

  if (typeof str === 'number') {
    if (str > 1000 && str < 80000) return formatDateYMD(excelSerialToDate(str))
    return ''
  }

  const s = String(str).trim()
  if (!s) return ''

  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s)
    if (serial > 1000 && serial < 80000) return formatDateYMD(excelSerialToDate(serial))
    return ''
  }

  const normalized = s
    .replace(/\s+de\s+/gi, ' ')
    .replace(/[\/\-\.]/g, '-')
    .replace(/\s+/g, '-')

  const parts = normalized.split('-').filter(Boolean)
  if (parts.length === 3) {
    let dia, mes, anio
    const p0 = parseInt(parts[0], 10)
    const p1 = parseInt(parts[1], 10)
    const p2 = parseInt(parts[2], 10)
    const mes1 = MESES_ESP[parts[1].toLowerCase()]

    if (!isNaN(p0) && mes1 !== undefined && !isNaN(p2)) {
      dia = p0
      mes = mes1
      anio = p2
    } else if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
      if (p0 > 1900 && p1 >= 1 && p1 <= 12 && p2 >= 1 && p2 <= 31) {
        anio = p0; mes = p1 - 1; dia = p2
      } else if (p0 >= 1 && p0 <= 31 && p1 >= 1 && p1 <= 12 && p2 > 1900) {
        dia = p0; mes = p1 - 1; anio = p2
      } else if (p0 >= 1 && p0 <= 12 && p1 >= 1 && p1 <= 31 && p2 > 1900) {
        mes = p0 - 1; dia = p1; anio = p2
      }
    }

    if (dia !== undefined && mes !== undefined && anio !== undefined && !isNaN(dia) && !isNaN(mes) && !isNaN(anio)) {
      const fullYear = anio < 100 ? (anio > 50 ? 1900 + anio : 2000 + anio) : anio
      if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && fullYear > 1900 && fullYear < 2100) {
        return `${fullYear}/${String(mes + 1).padStart(2, '0')}/${String(dia).padStart(2, '0')}`
      }
    }
  }

  const textParts = s.split(/\s+/).filter(Boolean)
  if (textParts.length >= 3) {
    const dia = parseInt(textParts[0], 10)
    let mes = -1
    let anio = NaN
    for (const p of textParts) {
      const lower = p.toLowerCase().replace(/[,\.]/g, '')
      if (MESES_ESP[lower] !== undefined) mes = MESES_ESP[lower]
      else if (/^\d{2,4}$/.test(p)) anio = parseInt(p, 10)
    }
    if (!isNaN(dia) && mes >= 0 && !isNaN(anio)) {
      const fullYear = anio < 100 ? (anio > 50 ? 1900 + anio : 2000 + anio) : anio
      if (fullYear > 1900 && fullYear < 2100) {
        return `${fullYear}/${String(mes + 1).padStart(2, '0')}/${String(dia).padStart(2, '0')}`
      }
    }
  }

  const d = new Date(s)
  if (!isNaN(d.getTime())) return formatDateYMD(d)

  return s
}

function splitName(str) {
  if (!str) return { first: '', rest: '' }
  const parts = String(str).trim().split(/\s+/)
  return { first: parts[0] || '', rest: parts.slice(1).join(' ') }
}

function formatExcelValue(val) {
  if (val === null || val === undefined) return ''
  if (typeof val === 'number') {
    const str = String(val)
    if (str.includes('e+') || str.includes('e-') || str.includes('E+') || str.includes('E-')) {
      return val.toLocaleString('es-CO', { maximumFractionDigits: 0 })
    }
    return str
  }
  return String(val).trim()
}

function getFullName(row, headers) {
  const nomIdx = headers.findIndex(h => h && String(h).toUpperCase().trim() === 'NOMBRES')
  const apeIdx = headers.findIndex(h => h && String(h).toUpperCase().trim() === 'APELLIDOS')
  let name = ''
  if (nomIdx >= 0 && row[nomIdx]) name = formatExcelValue(row[nomIdx])
  if (apeIdx >= 0 && row[apeIdx]) {
    name = (name ? name + ' ' : '') + formatExcelValue(row[apeIdx])
  }
  return name
}

function getExcelValue(row, headers, colName) {
  const idx = headers.findIndex(h => h && String(h).trim().toUpperCase() === colName.toUpperCase())
  if (idx >= 0 && idx < row.length) return formatExcelValue(row[idx])
  return ''
}

const PDF_FIELDS = [
  // ===== PAGE 0 (Autorización) =====
  { key: 'nombreDoc_p1', label: 'Nombres y Apellidos', page: 0, x: 790, y: 123, size: 16, w: 300, editable: false },
  { key: 'numDoc_p1', label: 'Número de Documento', page: 0, x: 790, y: 100, size: 16, w: 200, editable: false },

  // ===== PAGE 1 (Formulario principal) =====
  { key: 'nit', label: 'N° de NIT', page: 1, x: 30, y: 880, size: 12, w: 120, editable: true, default: 'N/A' },
  { key: 'razonSocial', label: 'Razón Social Entidad Patrocinadora', page: 1, x: 212, y: 880, size: 11, w: 210, editable: true, default: 'N/A' },
  { key: 'numDocumento', label: 'N° de Documento de Identidad', page: 1, x: 230, y: 829, size: 13, w: 160, editable: false, col: 'CEDULA' },
  { key: 'fechaExpedicion', label: 'Fecha de Expedición', page: 1, type: 'dateDigits', y: 831, size: 11, editable: false, col: 'FECHA DE EXPEDICION_date',
    digitPositions: [
      { x: 392 },  // Año d1
      { x: 405 },  // Año d2
      { x: 418 },  // Año d3
      { x: 431 },  // Año d4
      { x: 442 },  // Mes d1
      { x: 455 },  // Mes d2
      { x: 466 },  // Día d1
      { x: 479 },  // Día d2
    ],
  },
  { key: 'lugarExpedicion', label: 'Lugar de Expedición', page: 1, x: 507, y: 830, size: 11, w: 150, editable: false, col: 'CIUDAD EXPEDICION' },
  { key: 'primerApellido', label: 'Primer Apellido', page: 1, x: 27, y: 802, size: 10, w: 100, editable: false, col: 'APELLIDOS_first' },
  { key: 'otrosApellidos', label: 'Otros Apellidos', page: 1, x: 172, y: 802, size: 10, w: 100, editable: false, col: 'APELLIDOS_rest' },
  { key: 'primerNombre', label: 'Primer Nombre', page: 1, x: 317, y: 802, size: 10, w: 100, editable: false, col: 'NOMBRES_first' },
  { key: 'otrosNombres', label: 'Otros Nombres', page: 1, x: 462, y: 802, size: 10, w: 100, editable: false, col: 'NOMBRES_rest' },
  { key: 'genero', label: 'Género', page: 1, type: 'gender', editable: false, col: 'GENERO' },
  { key: 'lugarNacimiento', label: 'Lugar de Nacimiento', page: 1, x: 317, y: 774, size: 10, w: 160, editable: false, col: 'CIUDAD NACIMIENTO' },
  { key: 'fechaNacimiento', label: 'Fecha de Nacimiento', page: 1, type: 'dateDigits', y: 776, size: 10, editable: false, col: 'FECHA NACIMIENTO_date',
    digitPositions: [
      { x: 479 },  // Año d1
      { x: 492 },  // Año d2
      { x: 505 },  // Año d3
      { x: 517 },  // Año d4
      { x: 529 },  // Mes d1
      { x: 542 },  // Mes d2
      { x: 555 },  // Día d1
      { x: 568 },  // Día d2
    ],
  },
  { key: 'cargo', label: 'Cargo / Ocupación u Oficio', page: 1, x: 37, y: 738, size: 12, w: 180, editable: false, col: 'OFICIO' },
  { key: 'profesion', label: 'Profesión', page: 1, x: 323, y: 739, size: 11, w: 130, editable: false, col: 'PROFESION' },
  { key: 'direccion', label: 'Dirección de Residencia', page: 1, x: 42, y: 654, size: 11, w: 260, editable: false, col: 'DIRECCION' },
  { key: 'ciudad', label: 'Ciudad / Municipio', page: 1, x: 380, y: 654, size: 12, w: 140, editable: false, col: 'CIUDAD' },
  { key: 'departamento', label: 'Departamento', page: 1, x: 40, y: 625, size: 12, w: 120, editable: false, col: 'DEPARTAMENTO' },
  { key: 'telefono', label: 'Teléfono Celular', page: 1, x: 440, y: 625, size: 12, w: 130, editable: false, col: 'TELEFONO CELULAR' },
  { key: 'email', label: 'Correo Electrónico', page: 1, x: 42, y: 597, size: 14, w: 280, editable: false, col: 'CORREO ELECTRONICO' },
  { key: 'numCuenta', label: 'Número de Cuenta', page: 1, x: 30, y: 319, size: 12, w: 160, editable: false, col: 'CTA. No.' },
  { key: 'nombreBanco', label: 'Nombre del Banco', page: 1, x: 210, y: 319, size: 12, w: 200, editable: false, col: 'BANCO' },

  // ===== PAGE 2 (Firmas y constancia) =====
  { key: 'nombreFirma_p3', label: 'Nombres y Apellidos (14. Firma)', page: 2, x: 641, y: 561, size: 8, w: 300, editable: false },
  { key: 'docFirma_p3', label: 'N° Documento Identidad (14. Firma)', page: 2, x: 641, y: 533, size: 11, w: 200, editable: false, col: 'CEDULA' },
]

const STATIC_FIELDS = PDF_FIELDS.filter(f => f.editable)
const EXCEL_FIELDS = PDF_FIELDS.filter(f => !f.editable)

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / 1048576).toFixed(1) + ' MB'
}

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [dragActive, setDragActive] = useState(false)
  const [excelFile, setExcelFile] = useState(null)
  const [excelData, setExcelData] = useState(null)
  const [headers, setHeaders] = useState([])
  const [rows, setRows] = useState([])
  const [selectedRow, setSelectedRow] = useState(null)
  const [staticValues, setStaticValues] = useState({})
  const [fieldMapping, setFieldMapping] = useState({})
  const [fieldOverrides, setFieldOverrides] = useState({})
  const [loading, setLoading] = useState(false)
  const [pdfReady, setPdfReady] = useState(false)
  const [pdfUrl, setPdfUrl] = useState(null)
  const [generatingAll, setGeneratingAll] = useState(false)
  const [allPdfs, setAllPdfs] = useState([])
  const [razonesSociales, setRazonesSociales] = useState([])
  const [selectedRazon, setSelectedRazon] = useState('TODAS')
  const fileInputRef = useRef(null)

  const processExcel = useCallback((file) => {
    setExcelFile(file)
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const wb = XLSX.read(data, { type: 'array' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 })
        if (jsonData.length < 2) {
          alert('El archivo Excel no tiene datos suficientes.')
          return
        }
        const hdrs = jsonData[0] || []
        const dataRows = jsonData.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''))
        setHeaders(hdrs)
        setRows(dataRows)
        setExcelData(jsonData)

        const defaults = {}
        STATIC_FIELDS.forEach(f => { defaults[f.key] = f.default || '' })
        setStaticValues(defaults)

        const initialMapping = {}
        const initialOverrides = {}
        EXCEL_FIELDS.forEach(f => {
          if (f.col) {
            const baseCol = f.col.replace(/_first$/, '').replace(/_rest$/, '').replace(/_date$/, '')
            const idx = hdrs.findIndex(h => h && String(h).trim().toUpperCase() === baseCol.toUpperCase())
            initialMapping[f.key] = idx >= 0 ? idx : null
          } else {
            initialMapping[f.key] = null
          }
          initialOverrides[f.key] = ''
        })
        setFieldMapping(initialMapping)
        setFieldOverrides(initialOverrides)

        const razonIdx = hdrs.findIndex(h => h && String(h).trim().toUpperCase() === 'RAZON SOCIAL')
        const nitIdx = hdrs.findIndex(h => h && String(h).trim().toUpperCase() === 'NIT')
        const pairsMap = new Map()
        dataRows.forEach(row => {
          const razon = razonIdx >= 0 ? formatExcelValue(row[razonIdx]) : ''
          const nit = nitIdx >= 0 ? formatExcelValue(row[nitIdx]) : ''
          if (razon && !pairsMap.has(razon)) pairsMap.set(razon, nit)
        })
        const uniques = Array.from(pairsMap, ([razon, nit]) => ({ razon, nit }))
        setRazonesSociales(uniques)
        setSelectedRazon('TODAS')

        setSelectedRow(0)
      } catch (err) {
        console.error(err)
        alert('Error al leer el archivo Excel. Verifica que sea un archivo .xlsx válido.')
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleDrag = useCallback((e) => { e.preventDefault(); e.stopPropagation() }, [])

  useEffect(() => {
    if (selectedRazon === 'TODAS') return
    const rs = razonesSociales.find(r => r.razon === selectedRazon)
    if (rs) {
      setStaticValues(prev => ({
        ...prev,
        nit: rs.nit || prev.nit || 'N/A',
        razonSocial: rs.razon
      }))
    }
  }, [selectedRazon, razonesSociales])

  const filteredRows = useMemo(() => {
    const list = rows.map((row, idx) => ({ row, idx }))
    if (selectedRazon === 'TODAS') return list
    const razonIdx = headers.findIndex(h => h && String(h).trim().toUpperCase() === 'RAZON SOCIAL')
    if (razonIdx < 0) return list
    return list.filter(({ row }) => formatExcelValue(row[razonIdx]) === selectedRazon)
  }, [rows, selectedRazon, headers])

  const handleDragIn = useCallback((e) => { e.preventDefault(); e.stopPropagation(); setDragActive(true) }, [])
  const handleDragOut = useCallback((e) => { e.preventDefault(); e.stopPropagation(); setDragActive(false) }, [])
  const handleDrop = useCallback((e) => {
    e.preventDefault(); e.stopPropagation()
    setDragActive(false)
    const file = e.dataTransfer.files?.[0]
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) processExcel(file)
    else alert('Solo se aceptan archivos Excel (.xlsx, .xls)')
  }, [processExcel])

  const handleFileChange = useCallback((e) => {
    const file = e.target.files?.[0]
    if (file) processExcel(file)
  }, [processExcel])

  const getFieldValue = useCallback((field, row) => {
    if (field.editable) return staticValues[field.key] || field.default || ''

    const override = fieldOverrides[field.key]
    if (override && override.trim()) return override.trim()

    const colIdx = fieldMapping[field.key]
    if (colIdx !== null && colIdx !== undefined && colIdx >= 0 && colIdx < row.length) {
      const rawVal = formatExcelValue(row[colIdx])
      const col = field.col
      if (col) {
        if (col.endsWith('_first')) return splitName(rawVal).first
        if (col.endsWith('_rest')) return splitName(rawVal).rest
        if (col.endsWith('_date')) return parseFechaColombia(rawVal)
      }
      return rawVal
    }

    if (field.key === 'nombreDoc_p1' || field.key === 'nombreFirma_p3') return getFullName(row, headers)
    if (field.key === 'numDoc_p1' || field.key === 'docFirma_p3') return getExcelValue(row, headers, 'CEDULA')
    return ''
  }, [headers, staticValues, fieldMapping, fieldOverrides])

  const fillPdfPages = useCallback((pages, font, row) => {
    for (const field of PDF_FIELDS) {
      if (field.type === 'gender') {
        const raw = getFieldValue(field, row)
        const val = String(raw || '').toUpperCase().trim()
        if (!val) continue
        const page = pages[field.page]
        if (!page) continue
        const isMale = val === 'M' || val === 'MASCULINO' || val === 'HOMBRE' || val === 'H'
        const x = isMale ? 64 : 36
        const y = 776
        page.drawText('X', { x, y, size: 10, font, color: rgb(0, 0, 0) })
        continue
      }
      if (field.type === 'dateDigits') {
        const val = getFieldValue(field, row)
        if (!val) continue
        const page = pages[field.page]
        if (!page) continue
        const digits = val.replace(/\D/g, '')
        for (let i = 0; i < field.digitPositions.length && i < digits.length; i++) {
          const pos = field.digitPositions[i]
          page.drawText(digits[i], {
            x: pos.x,
            y: field.y,
            size: field.size,
            font,
            color: rgb(0, 0, 0),
          })
        }
        continue
      }
      const val = getFieldValue(field, row)
      if (!val) continue
      const page = pages[field.page]
      if (!page) continue
      page.drawText(val, {
        x: field.x,
        y: field.y,
        size: field.size,
        font,
        color: rgb(0, 0, 0),
        maxWidth: field.w,
      })
    }
  }, [getFieldValue, staticValues])

  const generatePdf = useCallback(async () => {
    if (!excelData || selectedRow === null) return
    setLoading(true)
    try {
      const pdfBytes = await fetch('/Archivo final.pdf').then(r => r.arrayBuffer())
      const pdfDoc = await PDFDocument.load(pdfBytes)
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
      const pages = pdfDoc.getPages()

      const row = rows[selectedRow]
      if (!row) { setLoading(false); return }

      fillPdfPages(pages, font, row)

      const modifiedPdfBytes = await pdfDoc.save()
      const blob = new Blob([modifiedPdfBytes], { type: 'application/pdf' })
      const url = URL.createObjectURL(blob)
      if (pdfUrl) URL.revokeObjectURL(pdfUrl)
      setPdfUrl(url)
      setPdfReady(true)
    } catch (err) {
      console.error(err)
      alert('Error al generar el PDF: ' + err.message)
    }
    setLoading(false)
  }, [excelData, selectedRow, rows, fillPdfPages, pdfUrl])

  const generateAllPdfs = useCallback(async () => {
    if (!excelData || rows.length === 0) return
    setGeneratingAll(true)
    try {
      const pdfBytes = await fetch('/Archivo final.pdf').then(r => r.arrayBuffer())

      allPdfs.forEach(p => URL.revokeObjectURL(p.url))

      const targets = filteredRows.map(f => f.row)
      const results = []
      for (let i = 0; i < targets.length; i++) {
        const row = targets[i]
        const pdfDoc = await PDFDocument.load(pdfBytes)
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
        const pages = pdfDoc.getPages()

        fillPdfPages(pages, font, row)

        const singlePdfBytes = await pdfDoc.save()
        const blob = new Blob([singlePdfBytes], { type: 'application/pdf' })
        const url = URL.createObjectURL(blob)
        const fullName = getFullName(row, headers) || `Empleado_${i + 1}`
        const safeName = fullName.replace(/[^\w\s-]/g, '').replace(/\s+/g, '_') || `Empleado_${i + 1}`
        const prefijo = selectedRazon !== 'TODAS'
          ? `${safeName}_${razonesSociales.find(r => r.razon === selectedRazon)?.nit || 'PDF'}`
          : safeName
        results.push({ name: `${prefijo}_formulario_pension.pdf`, url, fullName })
      }

      setAllPdfs(results)
    } catch (err) {
      console.error(err)
      alert('Error al generar PDFs: ' + err.message)
    }
    setGeneratingAll(false)
  }, [excelData, rows, fillPdfPages, headers, allPdfs, filteredRows, selectedRazon, razonesSociales])

  const downloadPdf = () => {
    if (pdfUrl) {
      const a = document.createElement('a')
      a.href = pdfUrl
      const name = rows[selectedRow] ? getFullName(rows[selectedRow], headers).replace(/\s+/g, '_') || 'empleado' : 'empleado'
      a.download = `${name}_formulario_pension.pdf`
      a.click()
    }
  }

  const downloadAllPdf = () => {
    if (allPdfs.length === 0) return
    allPdfs.forEach((pdf, i) => {
      setTimeout(() => {
        const a = document.createElement('a')
        a.href = pdf.url
        a.download = pdf.name
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
      }, i * 200)
    })
  }

  return (
    <div className="app">
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img src="/Logo syp.png" alt="Solutions & Payroll Logo" width="60" height="60" />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Formulario de Pensión - Grupo HL</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      <main className="main-content">
        <div className="container">

          <div className="help-section">
            <button
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Subí el Excel</strong>
                    <p>Arrastrá o seleccioná el archivo Excel con los datos de empleados.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Completá los campos manuales</strong>
                    <p>Los campos marcados como "N/A" por defecto se pueden editar abajo. Modificalos si es necesario.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Generá el PDF</strong>
                    <p>Seleccioná un empleado de la tabla y generá su formulario, o generá todos juntos.</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <h2>Formulario de Pensión - Porvenir</h2>
              <p className="description">
                Cargá el archivo Excel con los datos de empleados, ajustá los campos manuales y generá los formularios PDF.
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">

                {/* Upload */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                      <line x1="16" y1="13" x2="8" y2="13"/>
                      <line x1="16" y1="17" x2="8" y2="17"/>
                    </svg>
                    Archivo Excel de Empleados
                  </label>
                  <div
                    className={`drop-zone ${dragActive ? 'drag-active' : ''} ${excelFile ? 'has-file' : ''}`}
                    onDragEnter={handleDragIn}
                    onDragLeave={handleDragOut}
                    onDragOver={handleDrag}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                  >
                    {excelFile ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <div className="file-name">{excelFile.name}</div>
                          <div className="file-size">{formatFileSize(excelFile.size)} &middot; {rows.length} empleados</div>
                        </div>
                        <button className="btn-remove" onClick={(e) => { e.stopPropagation(); setExcelFile(null); setExcelData(null); setRows([]); setHeaders([]) }}>
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="7 10 12 15 17 10"/>
                          <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastrá tu archivo Excel aquí</span>
                          <span className="drop-zone-subtitle">o hacé clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    )}
                  </div>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleFileChange}
                    className="file-input"
                  />
                </div>

                {/* Static / Manual Fields */}
                {rows.length > 0 && STATIC_FIELDS.length > 0 && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                      </svg>
                      Campos Manuales (datos que no vienen del Excel)
                    </label>
                    <div className="manual-fields-grid">
                      {STATIC_FIELDS.map(f => (
                        <div key={f.key} className="manual-field-item">
                          <label className="manual-field-label">
                            {f.label} <span className="page-badge">Pág. {f.page + 1}</span>
                          </label>
                          {f.type === 'gender' ? (
                            <select
                              className="select-input"
                              value={staticValues[f.key] || ''}
                              onChange={(e) => setStaticValues(prev => ({ ...prev, [f.key]: e.target.value }))}
                            >
                              <option value="">-- Seleccionar --</option>
                              <option value="M">M</option>
                              <option value="F">F</option>
                            </select>
                          ) : (
                            <input
                              type="text"
                              className="select-input"
                              value={staticValues[f.key] || ''}
                              onChange={(e) => setStaticValues(prev => ({ ...prev, [f.key]: e.target.value }))}
                              placeholder={f.default || ''}
                            />
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Data Preview Table */}
                {rows.length > 0 && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>
                      </svg>
                      Datos Cargados ({filteredRows.length} de {rows.length} empleados) — Seleccioná uno para generar su PDF
                    </label>

                    {razonesSociales.length > 1 && (
                      <div className="manual-field-item" style={{ maxWidth: '420px' }}>
                        <label className="manual-field-label">
                          Filtrar por Razón Social
                          {selectedRazon !== 'TODAS' && razonesSociales.find(r => r.razon === selectedRazon) && (
                            <span style={{ marginLeft: '0.5rem', fontWeight: 400, color: 'var(--text-secondary)' }}>
                              (NIT: {razonesSociales.find(r => r.razon === selectedRazon).nit || 'N/A'})
                            </span>
                          )}
                        </label>
                        <select
                          className="select-input"
                          value={selectedRazon}
                          onChange={(e) => setSelectedRazon(e.target.value)}
                        >
                          <option value="TODAS">Todas ({rows.length})</option>
                          {razonesSociales.map((rs, i) => (
                            <option key={i} value={rs.razon}>
                              {rs.razon}{rs.nit ? ` — NIT: ${rs.nit}` : ''}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}

                    <div className="table-wrapper">
                      <table className="data-table">
                        <thead>
                          <tr>
                            <th>#</th>
                            <th>Nombre Completo</th>
                            <th>Cédula</th>
                            <th>Cargo</th>
                            <th>Ciudad</th>
                            <th>Razón Social</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredRows.map(({ row, idx: i }) => (
                            <tr
                              key={i}
                              className={selectedRow === i ? 'selected-row' : ''}
                              onClick={() => setSelectedRow(i)}
                            >
                              <td>{i + 1}</td>
                              <td>{getFullName(row, headers) || '-'}</td>
                              <td>{getExcelValue(row, headers, 'CEDULA') || '-'}</td>
                              <td>{getExcelValue(row, headers, 'OFICIO') || '-'}</td>
                              <td>{getExcelValue(row, headers, 'CIUDAD') || '-'}</td>
                              <td>{getExcelValue(row, headers, 'RAZON SOCIAL') || '-'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* Mapeo editable */}
                {rows.length > 0 && EXCEL_FIELDS.length > 0 && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                      </svg>
                      Mapeo Excel → PDF (editable)
                    </label>
                    <div className="table-wrapper">
                      <table className="data-table">
                        <thead>
                          <tr>
                            <th>Campo PDF</th>
                            <th>Pág.</th>
                            <th>Columna Excel</th>
                            <th>Valor personalizado (opcional)</th>
                            <th>Ejemplo</th>
                          </tr>
                        </thead>
                        <tbody>
                          {EXCEL_FIELDS.map(f => {
                            const sampleVal = selectedRow !== null && rows[selectedRow]
                              ? getFieldValue(f, rows[selectedRow])
                              : ''
                            return (
                              <tr key={f.key}>
                                <td><strong>{f.label}</strong></td>
                                <td>{f.page + 1}</td>
                                <td>
                                  <select
                                    className="select-input"
                                    style={{ padding: '0.3rem 0.4rem', fontSize: '0.75rem', width: '100%' }}
                                    value={fieldMapping[f.key] !== null && fieldMapping[f.key] !== undefined ? fieldMapping[f.key] : -1}
                                    onChange={(e) => {
                                      const idx = parseInt(e.target.value)
                                      setFieldMapping(prev => ({ ...prev, [f.key]: idx }))
                                    }}
                                  >
                                    <option value={-1}>-- Sin mapeo --</option>
                                    {headers.map((h, i) => (
                                      <option key={i} value={i}>{String(h || `Col. ${i + 1}`).trim()}</option>
                                    ))}
                                  </select>
                                </td>
                                <td>
                                  <input
                                    type="text"
                                    className="select-input"
                                    style={{ padding: '0.3rem 0.4rem', fontSize: '0.75rem', width: '100%' }}
                                    value={fieldOverrides[f.key] || ''}
                                    onChange={(e) => setFieldOverrides(prev => ({ ...prev, [f.key]: e.target.value }))}
                                    placeholder="Dejar vacío para usar columna"
                                  />
                                </td>
                                <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {sampleVal || '-'}
                                </td>
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* Acciones */}
                {rows.length > 0 && (
                  <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap', alignItems: 'center' }}>
                    {selectedRow !== null && (
                      <button className="btn-primary" onClick={generatePdf} disabled={loading}>
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                          <polyline points="14 2 14 8 20 8"/>
                        </svg>
                        {loading ? 'Generando...' : `Generar PDF: ${rows[selectedRow] ? getFullName(rows[selectedRow], headers) || `Empleado ${selectedRow + 1}` : ''}`}
                      </button>
                    )}
                    <button className="btn-secondary" onClick={generateAllPdfs} disabled={generatingAll}>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                      </svg>
                      {generatingAll ? 'Generando todos...' : `Generar Todos (${filteredRows.length})`}
                    </button>
                    {pdfReady && pdfUrl && (
                      <button className="btn-success" onClick={downloadPdf}>
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="7 10 12 15 17 10"/>
                          <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Descargar PDF Individual
                      </button>
                    )}
                    {allPdfs.length > 0 && (
                      <button className="btn-success" onClick={downloadAllPdf}>
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="7 10 12 15 17 10"/>
                          <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Descargar Todos los PDF ({allPdfs.length})
                      </button>
                    )}
                  </div>
                )}

                {/* PDF Preview */}
                {pdfReady && pdfUrl && (
                  <div className="form-group" style={{ marginTop: '1.5rem' }}>
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                        <circle cx="12" cy="12" r="3"/>
                      </svg>
                      Vista Previa del PDF Generado
                    </label>
                    <iframe
                      src={pdfUrl}
                      style={{ width: '100%', height: '600px', border: '1px solid var(--border)', borderRadius: 'var(--radius)' }}
                      title="PDF Preview"
                    />
                  </div>
                )}

              </div>
            </div>
          </div>
        </div>
      </main>

      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App
