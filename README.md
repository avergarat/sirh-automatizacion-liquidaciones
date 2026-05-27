<div align="center">

# ⚙️ Automatización de Descarga de Liquidaciones desde SIRH
### Caso de Estudio: Eliminación del Trabajo Manual Repetitivo en Gestión de RRHH del Sector Público

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![pywinauto](https://img.shields.io/badge/pywinauto-Win32_UI-F7931E?style=for-the-badge)](https://pywinauto.readthedocs.io)
[![Tkinter](https://img.shields.io/badge/Tkinter-GUI-4EABE6?style=for-the-badge)](https://docs.python.org/3/library/tkinter.html)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346?style=for-the-badge)](https://openpyxl.readthedocs.io)
[![Platform](https://img.shields.io/badge/Plataforma-Windows-0078D4?style=for-the-badge&logo=windows)](.)
[![Status](https://img.shields.io/badge/Estado-Completado-27AE60?style=for-the-badge)](.)

**Alexis Vergara Torres — Analista de Datos & Automatización de Procesos**

[⚙️ Ver Pipeline](#-pipeline-de-automatización) · [🖥️ Interfaz](#-interfaz-gráfica) · [🚀 Cómo Usar](#-cómo-usar) · [💼 Contratar](#-sobre-el-autor)

</div>

---

## ⚡ Impacto en Números

<div align="center">

| Métrica | Antes (manual) | Después (automatizado) |
|:--------|:--------------:|:----------------------:|
| ⏱️ Tiempo por liquidación | ~4–6 minutos | **< 15 segundos** |
| 🖱️ Clics manuales por folio | ~10 acciones | **0 intervención** |
| 📄 Nomenclatura de archivos | Variable / errores | **Estandarizada automáticamente** |
| 🔁 Reanudación si se interrumpe | Desde el inicio | **Desde cualquier folio** |
| ❌ Riesgo de omisión de folios | Alto (fatiga) | **Eliminado** |
| 📁 Formato de salida | Manual, inconsistente | **`AAAAmm_RUT_CORRPAGO_A.pdf`** |

</div>

---

## 🎯 ¿De Qué Trata Este Proyecto?

> **El problema:** El sistema SIRH (Sistema de Información de Recursos Humanos) gestiona liquidaciones accesorias de planillas para cientos de funcionarios públicos. Descargar manualmente cada liquidación —ingresar el folio, esperar la carga, hacer clic en "Copia Liqu.", accionar el preview, imprimir, nombrar el PDF y guardarlo— multiplica ese ciclo de ~10 acciones por cada funcionario del lote. Con cientos de folios por período, esto significa horas de trabajo administrativo repetitivo, con alto riesgo de errores de nomenclatura, omisiones y fatiga operacional.

> **La solución:** Desarrollé una aplicación de escritorio con interfaz gráfica profesional (Tkinter) que automatiza completamente el ciclo de descarga. La herramienta carga un Excel con los folios a procesar, se conecta a la ventana de SIRH ya abierta mediante automatización de interfaz Win32 (pywinauto), y ejecuta cada paso del proceso de forma autónoma: ingresa el folio, activa "Copia Liqu.", lanza el preview, hace clic en imprimir y OK, intercepta el diálogo "Guardar como", aplica la nomenclatura estandarizada y guarda el PDF — todo sin intervención humana, en un hilo separado para que la GUI permanezca responsiva.

> **El resultado:** Lo que antes tomaba horas de trabajo manual y concentración constante se reduce a cargar el Excel, definir la carpeta de salida, seleccionar el período y presionar **"Comenzar"**. El operador puede pausar, reanudar desde cualquier folio, y monitorear el avance en tiempo real desde el log de la aplicación.

---

## 🔍 Qué Hace Exactamente — Paso a Paso

### 🔗 1. Se Conecta a SIRH sin Modificarlo

La herramienta no modifica ni interviene el sistema SIRH. En cambio, se conecta a la **ventana ya abierta** de "Consulta de liquidación por folio" usando la API Win32, la toma como objeto controlable y opera sobre ella como si fuera un operador humano — pero sin errores y sin fatiga.

```
SIRH abierto por el operador
         │
         ▼
  pywinauto detecta la ventana
  title_re: ".*Consulta de liquidación por folio.*"
         │
         ▼
  Control total: click, set_focus, send_keys, child_window
```

### 📋 2. Lee el Excel de Folios — Sin Supuestos de Posición

El Excel de entrada tiene una estructura con columnas `FOLIO`, `NUMERO`, `RUT` y `CORRELATIVO DE PAGO`. El parser lee el **encabezado dinámicamente** — si las columnas están en distinto orden, la herramienta las encuentra igual. No asume posiciones fijas.

### 🔄 3. Ciclo Completo por Folio

```
Para cada FOLIO en el lote:
  ├── 1. Ingresar folio en el campo (por posición relativa en ventana)
  ├── 2. Enviar ENTER → SIRH carga la liquidación
  ├── 3. Clic en "Copia Liqu." (control_id=137)
  ├── 4. Esperar apertura del preview (3s estabilidad)
  ├── 5. Clic en ícono de impresión (coords absolutas)
  ├── 6. Clic en OK del cuadro de impresión
  ├── 7. Interceptar diálogo "Guardar como" (polling ~40s timeout)
  ├── 8. Escribir ruta PDF estandarizada: AAAAmm_RUT_CORR_A.pdf
  ├── 9. Clic en "Guardar"
  ├── 10. Cerrar preview y limpiar formulario
  └── 11. Log de resultado → siguiente folio
```

### 📁 4. Nomenclatura Estandarizada con Manejo de Duplicados

```python
# Formato: AAAAmm_RUT_CORRELATIVO_DE_PAGO_A.pdf
# Si ya existe: AAAAmm_RUT_CORRELATIVO_DE_PAGO_A_1.pdf
# Si sigue: AAAAmm_RUT_CORRELATIVO_DE_PAGO_A_2.pdf ...
```

Esto garantiza que ningún PDF se sobreescribe accidentalmente, incluso si un funcionario tiene múltiples liquidaciones en el mismo período.

### ⏸️ 5. Reanudable desde Cualquier Punto

El combo de inicio permite seleccionar desde qué folio comenzar. Si el proceso se interrumpe en el folio 47 de 200, se puede reanudar desde el 47 sin reprocesar los anteriores.

---

## 🖥️ Interfaz Gráfica

```
┌─────────────────────────────────────────────────────────────┐
│  EVIDANT · Descarga automática de liquidaciones             │
├─────────────────────────────────────────────────────────────┤
│  Archivo Excel de folios:  [____________________] [Buscar]  │
│  Carpeta salida PDF:       [____________________] [Selec.]  │
│  Periodo (AAAAmm):         [202505            ]             │
│  Comenzar desde:           [0001 - FOLIO 12345 · RUT...  ] │
├─────────────────────────────────────────────────────────────┤
│  [ Comenzar ]  [ Detener ]                                  │
├─────────────────────────────────────────────────────────────┤
│  ✔ Cargados 148 registros desde el Excel.                   │
│  ➡ Fila 1: FOLIO 12345 | RUT 12.345.678-9 | CORR 001       │
│  ✅ Terminado FOLIO 12345                                    │
│  ➡ Fila 2: FOLIO 12346 | RUT 11.222.333-4 | CORR 001       │
│  ✅ Terminado FOLIO 12346                                    │
│  ...                                                        │
└─────────────────────────────────────────────────────────────┘
```

- **Fondo oscuro** profesional (`#0B1F33` + turquesa `#00A6A6`)
- **Log en tiempo real** — el proceso corre en hilo separado, la GUI nunca se bloquea
- **Selector de carpeta** via PowerShell FolderBrowserDialog (compatible con modo Administrador)
- **Período con valor por defecto** — mes y año actual al abrir la aplicación

---

## 💼 ¿Qué Puede Hacer Este Tipo de Automatización por Su Organización?

<table>
<tr>
<th>¿Su equipo hace esto manualmente?</th>
<th>Lo que la automatización puede hacer</th>
</tr>
<tr>
<td>📋 <strong>Descargar documentos uno a uno desde un sistema legacy</strong></td>
<td>Pipeline automatizado que procesa cientos de registros sin intervención</td>
</tr>
<tr>
<td>🖱️ <strong>Repetir la misma secuencia de clics decenas de veces</strong></td>
<td>Automatización Win32 que replica exactamente el flujo humano — sin errores</td>
</tr>
<tr>
<td>📁 <strong>Nombrar archivos manualmente con riesgo de inconsistencia</strong></td>
<td>Nomenclatura estandarizada automática con manejo de duplicados</td>
</tr>
<tr>
<td>⏱️ <strong>Destinar horas de personal calificado a trabajo operativo</strong></td>
<td>Reducción de horas-persona a minutos de supervisión</td>
</tr>
<tr>
<td>🔁 <strong>Reiniciar desde cero si el proceso se interrumpe</strong></td>
<td>Reanudación exacta desde el último punto procesado</td>
</tr>
<tr>
<td>🏛️ <strong>Sistemas institucionales que no tienen API ni exportación directa</strong></td>
<td>Automatización de UI — funciona sobre cualquier sistema de escritorio Windows</td>
</tr>
</table>

### Sectores donde este tipo de automatización genera mayor impacto:

```
🏛️ Sector Público (SIRH, SII, SIGFE, CMF)     🏥 Salud (SIGGES, RIS, HIS legacy)
🏦 Banca & Seguros (sistemas core)              🏗️ Municipios y servicios descentralizados
📋 RRHH corporativo (ADP, SAP legacy)           ⚖️ Tribunales y organismos judiciales
```

---

## 🧠 Pipeline de Automatización

```
Excel de folios (FOLIO, NUMERO, RUT, CORRELATIVO DE PAGO)
         │
         ▼
   Validación de columnas (detección dinámica de encabezado)
         │
         ▼
   Conexión a ventana SIRH activa (pywinauto / Win32 backend)
         │
         ▼
   Para cada folio:
   ├── Ingreso de folio por coordenadas relativas
   ├── Activación de "Copia Liqu." por control_id
   ├── Clic en impresión (coords absolutas calibradas)
   ├── Intercepción de "Guardar como" (polling con timeout)
   ├── Escritura de ruta PDF estandarizada
   └── Cierre de preview + limpieza de formulario
         │
         ▼
   Log en tiempo real en GUI (hilo separado / thread-safe)
         │
         ▼
   PDFs organizados: AAAAmm_RUT_CORRPAGO_A.pdf
```

### Stack Técnico

```python
# Automatización de interfaz
pywinauto       # Conexión y control de ventanas Win32
send_keys       # Escritura de texto y atajos de teclado
mouse           # Clics en coordenadas absolutas

# Interfaz gráfica
tkinter         # GUI principal (stdlib)
ttk             # Widgets modernos (Combobox, estilo)
threading       # Proceso en hilo separado — GUI no se bloquea

# Datos
openpyxl        # Lectura de Excel con detección dinámica de columnas
pathlib         # Manejo de rutas y creación de carpetas

# Sistema
subprocess      # PowerShell FolderBrowserDialog (compatible modo Admin)
```

---

## 📁 Estructura del Repositorio

```
📁 sirh-automatizacion-liquidaciones/
├── 🐍 descarga_liquidaciones_sirh.py    # Aplicación completa
└── 📦 requirements.txt                  # Dependencias
```

> ⚠️ **Nota de privacidad:** Los archivos Excel con RUTs y datos de funcionarios no están incluidos. Son información sensible de carácter institucional y no deben versionarse.

---

## 🚀 Cómo Usar

### Requisitos

- **Windows** (la automatización Win32 es exclusiva de Windows)
- Python 3.10+
- SIRH abierto con la ventana **"Consulta de liquidación por folio"** visible

### Instalación

```bash
git clone https://github.com/avergarat/sirh-automatizacion-liquidaciones.git
cd sirh-automatizacion-liquidaciones
pip install -r requirements.txt
```

### Ejecución

```bash
python descarga_liquidaciones_sirh.py
```

### Preparar el Excel

El archivo Excel debe tener una hoja llamada `FOLIOS` con las columnas (en cualquier orden):

| FOLIO | NUMERO | RUT | CORRELATIVO DE PAGO |
|-------|--------|-----|---------------------|
| 12345 | 001    | 12.345.678-9 | 001 |
| 12346 | 002    | 11.222.333-4 | 001 |

### Calibración de coordenadas

Si la resolución o la posición de la pantalla cambia, ajusta estas constantes al inicio del script:

```python
OFFSET_FOLIO_X = 286    # posición relativa del campo FOLIO en la ventana SIRH
OFFSET_FOLIO_Y = 204
PRINT_ABS_X = 249       # botón imprimir en el preview
PRINT_ABS_Y = 38
OK_ABS_X = 1083         # botón OK del cuadro de impresión
OK_ABS_Y = 467
CLOSE_ABS_X = 1894      # botón cerrar del preview
CLOSE_ABS_Y = 7
```

---

## 👤 Sobre el Autor

<div align="center">

### Alexis Vergara Torres
**Analista de Datos Senior · Automatización de Procesos · Machine Learning**

</div>

Especializado en eliminar trabajo manual repetitivo en organizaciones públicas y privadas. Este proyecto nació de una necesidad real del sector público: sistemas institucionales legacy que no ofrecen exportación masiva y obligan a los equipos a operar manualmente durante horas. La automatización de UI es la solución cuando no hay API — y funciona sobre cualquier sistema de escritorio Windows sin modificarlo.

| Área | Capacidades |
|------|-------------|
| Automatización de UI (RPA ligero) | ✅ pywinauto, Win32, coordenadas, control_id |
| Interfaces gráficas de escritorio | ✅ Tkinter, ttk, threading, diseño profesional |
| Procesamiento de datos institucionales | ✅ openpyxl, pandas, validación de columnas |
| Integración con sistemas legacy | ✅ SIRH, SII, SIGFE, cualquier app Windows |
| Nomenclatura y organización de archivos | ✅ Estandarización automática, manejo de duplicados |

### 📬 Contacto

> ¿Su equipo destina horas a tareas repetitivas que un script podría hacer en minutos?  
> Automatizo procesos sobre cualquier sistema Windows — sin necesidad de acceso a la base de datos ni modificar el sistema existente.

📧 **Email:** alexis.vergara.torres@gmail.com  
💼 **GitHub:** [github.com/avergarat](https://github.com/avergarat)  
📍 **Ubicación:** Santiago, Chile — disponible para proyectos remotos y presenciales

---

## 📌 Notas Técnicas

- **Coordinadas absolutas:** Los valores `PRINT_ABS_X/Y`, `OK_ABS_X/Y` y `CLOSE_ABS_X/Y` dependen de la resolución y disposición de pantalla. Deben calibrarse una vez por equipo usando una herramienta de captura de coordenadas.
- **Timeout "Guardar como":** El polling espera hasta 40 segundos (`80 × 0.5s`) para que aparezca el diálogo — suficiente para conexiones lentas o sistemas con alta carga.
- **Thread safety:** Todas las actualizaciones de UI desde el hilo de procesamiento usan `root.after(0, ...)` para encolar en el hilo principal — sin condiciones de carrera.
- **Compatibilidad:** Probado sobre Windows 10/11 con SIRH institucional del Ministerio de Salud. Requiere que la ventana objetivo esté visible y al frente al iniciar.
- **Privacidad:** No almacena, transmite ni registra RUTs ni datos de funcionarios. Opera solo en memoria RAM durante la ejecución.

---

<div align="center">

⭐ Si este proyecto le resultó útil, deje una estrella en el repositorio  
🔔 Siga el perfil para ver más proyectos de automatización y Business Analytics

</div>
