# Sistema de Asignación de Monitores

Aplicación de **escritorio** (Windows/Linux) que automatiza la asignación de monitores a espacios/salas de estudio a partir de su disponibilidad horaria, aplicando reglas configurables de prioridad, balanceo de carga y restricciones horarias, y exportando el resultado a un archivo Excel con múltiples formatos de horario.

Construida en **Python** con **PySide6** (Qt) para la interfaz gráfica y **pandas / openpyxl** para el procesamiento y la exportación de datos.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![PySide6](https://img.shields.io/badge/GUI-PySide6%20(Qt)-41CD52?logo=qt&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-data%20processing-150458?logo=pandas&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux-lightgrey)

---

## Tabla de contenido

- [¿Qué resuelve?](#qué-resuelve)
- [Características](#características)
- [Arquitectura](#arquitectura)
- [Algoritmo de asignación](#algoritmo-de-asignación)
- [Formato de los archivos de entrada](#formato-de-los-archivos-de-entrada)
- [Configuración](#configuración)
- [Puesta en marcha](#puesta-en-marcha)
- [Empaquetado / distribución](#empaquetado--distribución)
- [Estructura del proyecto](#estructura-del-proyecto)

---

## ¿Qué resuelve?

Asignar manualmente monitores a decenas de bloques horarios (salas, días, cursos) respetando disponibilidad, mínimos y máximos de horas, y evitando choques de horario, es un proceso tedioso y propenso a errores. Esta herramienta carga la disponibilidad de los monitores y el listado de espacios a cubrir desde archivos Excel/ODS, ejecuta un algoritmo de asignación configurable y genera un Excel final listo para compartir, con hojas separadas por sala y por monitor.

## Características

- 📥 **Carga de datos desde Excel/ODS**: tanto la disponibilidad de monitores como el listado de espacios a cubrir se importan desde hojas de cálculo (`.xlsx`, `.xls`, `.ods`).
- 🧠 **Motor de asignación en dos fases**:
  1. Prioriza monitores que aún no alcanzan su mínimo de horas.
  2. Asigna el resto de espacios balanceando carga o respetando prioridades, según configuración.
- ⚖️ **Reglas configurables**: uso de sistema de prioridades (1–5), balanceo de carga, priorización de mínimos, máximo de horas continuas permitidas y horas mínimas/máximas por monitor.
- 🚫 **Prevención de conflictos**: valida que un monitor no quede asignado a dos espacios que se solapen en el mismo horario, y respeta el límite de horas seguidas configurado.
- 🧾 **Normalización robusta de datos**: interpreta celdas de disponibilidad en distintos formatos ("Libre", "9am-5pm", "No disponible", etc.), normaliza días de la semana (con o sin acentos, abreviados) y valida prioridades.
- 📊 **Reporte de resultados** con porcentaje de espacios asignados y espacios sin monitor disponible.
- 📤 **Exportación a Excel enriquecida**: genera hojas de horario consolidado, horario por sala (todas las salas lado a lado) y horario individual por monitor, con colores distintivos por monitor y estilos aplicados con `openpyxl`.
- 🖥️ **Interfaz gráfica moderna** construida con PySide6: tarjetas de estadísticas, roadmap visual del proceso, barra de progreso, tabla de resultados y consola de reporte, todo ejecutado en un hilo (`QThread`) separado de la UI.
- ⚙️ **Panel de configuración** para ajustar horas por defecto y las reglas del algoritmo, persistido en un archivo JSON local.
- 📦 **Distribución multiplataforma**: empaquetado como ejecutable de Windows (instalador con Inno Setup) y como AppImage para Linux.

## Arquitectura

```
GUI (PySide6)                    Núcleo (core)                  Exportación (export)
─────────────────                ───────────────                ─────────────────────
MainWindow  ──carga──►  loader.py (lee Excel/ODS)
            ──config──► config.py (config.json persistente)
            ──ejecuta──► AsignacionThread (QThread)
                              │
                              ▼
                        assignment.py (algoritmo de asignación)
                              │
                              ▼
                        excel_writer.py + schedule_formatter.py
                              │
                              ▼
                     Excel final (horarios por sala / por monitor)
```

- **`core/`**: lógica de negocio pura, sin dependencias de la interfaz — parsing de Excel, normalización de datos y algoritmo de asignación.
- **`gui/`**: toda la interfaz PySide6 (ventana principal, diálogos de configuración y descarga, widgets reutilizables como tarjetas y pasos del roadmap).
- **`export/`**: construcción del archivo Excel de salida, con formato y estilos.
- **`utils/threads.py`**: ejecuta la asignación en un hilo (`QThread`) para no bloquear la interfaz mientras se procesa.

## Algoritmo de asignación

Implementado en `src/core/assignment.py`, se ejecuta en dos fases sobre la lista de monitores y el DataFrame de espacios:

1. **Fase 1 — Cumplimiento de mínimos** (si `usar_prioridad` y `priorizar_minimo` están activos): para cada espacio, se buscan monitores que aún no alcanzan su mínimo de horas, estén disponibles, no violen restricciones de horas seguidas y no tengan conflicto de horario. Se elige el de mayor prioridad y menor holgura respecto a su mínimo.
2. **Fase 2 — Asignación del resto**: para los espacios no cubiertos en la fase 1, se buscan candidatos disponibles que no superen su máximo de horas. Se ordenan por prioridad (si está activada) o por menor carga horaria acumulada (si se prefiere balanceo), y se asigna el mejor candidato.

Un espacio queda marcado como `SIN MONITOR` si ningún candidato cumple todas las restricciones (disponibilidad, horas máximas, horas seguidas o conflicto de horario).

## Formato de los archivos de entrada

**Archivo de monitores** (fila de encabezado configurable, por defecto fila 5 en adelante):
- Columna con el nombre completo del monitor.
- Columna opcional de prioridad (entero 1–5).
- Columnas de disponibilidad por día y jornada (`lunes_mañana`, `lunes_tarde`, etc.), aceptando valores como `"Libre"`, `"9am-1pm"`, `"No disponible"`.

**Archivo de espacios**, con columnas (nombres configurables):
- `SALA`, `DIA`, `HORA_INICIO`, `HORA_FIN`, `CURSO`.

Ambos formatos (columnas, filas de encabezado, nombres) son configurables desde `core/config.py` / el panel de configuración de la app.

## Configuración

La configuración se persiste en `~/.asignacion_monitores/config.json` y cubre tres bloques (ver `DEFAULT_CONFIG` en `src/core/config.py`):

| Bloque        | Parámetros                                                                 |
|---------------|-------------------------------------------------------------------------------|
| `monitores`   | Fila de encabezado, fila de inicio de datos, nombres de columnas, horas mínimas/máximas por defecto |
| `espacios`    | Nombres de columnas esperadas (sala, día, hora inicio/fin, curso)                  |
| `asignacion`  | Balancear carga, priorizar mínimos, usar prioridad, máximo de horas seguidas, descanso mínimo, permitir sobrepasar el máximo |

Desde la interfaz, el diálogo **⚙️ Configuración del Sistema** permite ajustar horas mínimas/máximas por defecto y activar/desactivar el sistema de prioridades, el balanceo de carga y la priorización de mínimos, guardando los cambios automáticamente.

## Puesta en marcha

### Requisitos previos

- Python 3.10+
- Dependencias: `PySide6`, `pandas`, `openpyxl`, `xlsxwriter`, `numpy`, `odfpy` (para leer archivos `.ods`)

### Instalación

```bash
git clone https://github.com/N4him/Automatizador-de-horario.git
cd Automatizador-de-horario

python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate

pip install PySide6 pandas openpyxl xlsxwriter numpy odfpy
```

> El repositorio no incluye un `requirements.txt`; las dependencias en tiempo de ejecución están declaradas explícitamente como `hiddenimports` en `Asignacion_Monitores.spec` (usado por PyInstaller).

### Ejecución

```bash
python run.py
```

## Empaquetado / distribución

El proyecto incluye configuración para generar ejecutables independientes con **PyInstaller**:

- **`Asignacion_Monitores.spec`**: define el empaquetado con PyInstaller (sin consola, ícono incluido, con `hiddenimports` de PySide6, openpyxl, pandas, etc.).
- **`setup.iss`**: script de **Inno Setup** para generar un instalador de Windows (`Setup_Sistema_Asignacion_Monitores_v1.0.exe`, incluido en `installer_output/`).
- **`build_appimage.sh`**: script para empaquetar la app como **AppImage** en Linux, usando la carpeta `AppDir/`.

Flujo típico de build:

```bash
pyinstaller Asignacion_Monitores.spec       # genera dist/Asignacion_Monitores(.exe)

# Windows: compilar setup.iss con Inno Setup Compiler
# Linux:
./build_appimage.sh
```

> El repositorio contiene ramas dedicadas al empaquetado por plataforma: `Setup-windows` y `Setup-linux`, además de `main`.

## Estructura del proyecto

```
Automatizador-de-horario/
├── run.py                        # Punto de entrada (usado por el instalador)
├── src/
│   ├── main.py                     # Punto de entrada de la app (QApplication)
│   ├── core/
│   │   ├── config.py                 # Configuración persistente (JSON)
│   │   ├── loader.py                   # Carga de Excel/ODS → monitores y espacios
│   │   ├── parser.py                     # Normalización de horas, días y prioridades
│   │   ├── models.py                       # Dataclasses Monitor y Espacio
│   │   └── assignment.py                     # Algoritmo de asignación
│   ├── gui/
│   │   ├── widgets/main_window.py              # Ventana principal
│   │   ├── dialogs/config_dialog.py               # Diálogo de configuración
│   │   ├── dialogs/download_dialog.py               # Diálogo de exportación
│   │   ├── models/pandas_model.py                     # Modelo Qt para mostrar DataFrames
│   │   └── styles/app_styles.py                         # Hoja de estilos de la app
│   ├── export/
│   │   ├── excel_writer.py                                # Orquesta la exportación a Excel
│   │   ├── schedule_formatter.py                            # Construye horarios consolidados/por monitor
│   │   └── styles.py                                          # Estilos de las celdas de Excel
│   └── utils/threads.py                                         # QThread para ejecutar la asignación
├── Asignacion_Monitores.spec       # Configuración de PyInstaller
├── setup.iss                        # Instalador Windows (Inno Setup)
├── build_appimage.sh                  # Empaquetado AppImage (Linux)
├── AppDir/                              # Recursos para el AppImage
└── installer_output/                      # Instalador de Windows generado
```

---

Herramienta desarrollada para automatizar la asignación de monitores/horarios en espacios académicos.
