# MetaCleaner — Limpiador de Metadatos

> Aplicación de escritorio para Windows que **detecta y elimina metadatos sensibles** de documentos, hojas de cálculo, PDFs e imágenes, sin alterar el contenido.

---

## Tabla de contenidos

- [Descripción](#descripción)
- [Funcionalidades](#funcionalidades)
- [Requisitos del sistema](#requisitos-del-sistema)
- [Instalación](#instalación)
- [Guía de uso](#guía-de-uso)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Configuración avanzada](#configuración-avanzada)
- [Generar el ejecutable](#generar-el-ejecutable)
- [Solución de errores comunes](#solución-de-errores-comunes)
- [Historial de versiones](#historial-de-versiones)

---

## Descripción

Cuando compartes un archivo de Word, un PDF o una foto, ese archivo lleva consigo **metadatos ocultos**: tu nombre, empresa, fechas de edición, coordenadas GPS de la cámara, software utilizado y mucho más.

**MetaCleaner** te permite:

1. Ver exactamente qué información personal contiene cada archivo antes de enviarlo.
2. Eliminar esos metadatos con un solo clic.
3. Guardar una copia limpia sin modificar el archivo original.

La interfaz es completamente en español, con tema oscuro y funciona sin conexión a internet.

---

## Funcionalidades

| Función | Detalle |
|---|---|
| **Carga por arrastre** | Arrastra el archivo directamente sobre la ventana |
| **Carga por diálogo** | Botón de selección con filtro de formatos compatibles |
| **Visor de metadatos** | Tabla con todos los campos detectados antes de limpiar |
| **Limpieza con un clic** | Botón "Limpiar Metadatos" elimina toda la información sensible |
| **Guardado inteligente** | Sugiere automáticamente el nombre `archivo_cleaned.ext` |
| **Verificación post-limpieza** | Muestra los metadatos del archivo limpio para confirmar el resultado |
| **Barra de progreso** | Indicador visual durante la extracción y limpieza |
| **Mensajes de estado** | Colores diferenciados: verde (éxito), rojo (error), blanco (info) |
| **Procesamiento en segundo plano** | La UI nunca se congela mientras trabaja |

### Formatos soportados

| Formato | Metadatos que se eliminan |
|---|---|
| `.docx` | Autor, último editor, título, asunto, palabras clave, empresa, fechas de creación y modificación, número de revisión |
| `.xlsx` | Creador, último editor, título, asunto, descripción, empresa, palabras clave, fechas |
| `.pdf` | Autor, título, creador, productor (software), asunto, palabras clave, fechas de creación y modificación |
| `.jpg` / `.jpeg` | **Todos los datos EXIF**: modelo de cámara, coordenadas GPS, fecha y hora de disparo, ajustes de exposición, software de edición, datos del fabricante |
| `.png` | Todos los metadatos EXIF y de texto incrustados |
| `.json` | Re-serialización limpia que elimina comentarios no estándar y formato irregular |

---

## Requisitos del sistema

### Para ejecutar el instalable (`.exe`)

- **Sistema operativo**: Windows 10 / 11 (64-bit)
- **RAM**: mínimo 256 MB libres
- **Disco**: ~120 MB para la carpeta de la aplicación
- Python **no** es necesario

### Para ejecutar desde código fuente

- **Python** 3.11 o superior (probado con 3.13)
- **Sistema operativo**: Windows, macOS o Linux
- Conexión a internet para la primera instalación de dependencias

---

## Instalación

### Opción A — Ejecutable portátil (recomendado)

1. Descarga o copia la carpeta `dist/MetaCleaner/`.
2. Abre la carpeta y haz doble clic en `MetaCleaner.exe`.
3. No requiere instalación ni Python.

> La carpeta completa es portátil: puedes copiarla a un USB o a cualquier equipo Windows y funciona igual.

### Opción B — Desde código fuente

```bash
# 1. Clonar / descargar el proyecto
cd c:\Proyectos\appLimpiarMetadatos

# 2. (Recomendado) Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la aplicación
python app.py
```

---

## Guía de uso

### Paso 1 — Cargar un archivo

Tienes dos opciones:

- **Arrastrar y soltar**: arrastra el archivo desde el Explorador directamente sobre la zona de carga (recuadro con el icono de flecha).
- **Clic para seleccionar**: haz clic en la zona de carga y elige el archivo en el diálogo estándar de Windows.

La aplicación validará que el formato sea compatible. Si no lo es, verás un mensaje de error en la barra inferior.

### Paso 2 — Revisar metadatos

Una vez cargado el archivo, el panel derecho se poblará automáticamente con todos los metadatos detectados. Cada fila muestra:

- **Campo**: nombre del metadato (p. ej. "Autor", "GPS / GPSLatitude").
- **Valor**: dato almacenado en el archivo.

Los campos vacíos aparecen con un guión (`—`) en gris.

### Paso 3 — Limpiar

Haz clic en el botón morado **"Limpiar Metadatos"**. Aparecerá un diálogo para elegir dónde guardar el archivo limpio. Por defecto sugiere el mismo directorio con el sufijo `_cleaned`:

```
informe.docx  →  informe_cleaned.docx
foto.jpg      →  foto_cleaned.jpg
```

> El archivo **original no se modifica** en ningún momento.

### Paso 4 — Verificar

Tras guardar, la aplicación carga automáticamente los metadatos del archivo limpio en la tabla para que puedas confirmar que los campos sensibles han sido eliminados.

### Cargar otro archivo

Usa el botón **"Cargar otro archivo"** para resetear la aplicación y procesar un nuevo documento.

---

## Estructura del proyecto

```
appLimpiarMetadatos/
│
├── app.py                  # Aplicación completa (backend + UI)
├── requirements.txt        # Dependencias Python
├── metacleaner.spec        # Configuración de PyInstaller
├── build.bat               # Script de compilación automática
├── runtime_hook.py         # Hook de PyInstaller para rutas en modo frozen
│
├── assets/
│   ├── create_icon.py      # Generador del icono (ejecutar una vez)
│   └── icon.ico            # Icono de la aplicación
│
├── hooks/                  # Hooks personalizados de PyInstaller (opcional)
│
├── build/                  # Archivos temporales de compilación (ignorar)
│
└── dist/
    └── MetaCleaner/        # Carpeta portátil lista para distribuir
        ├── MetaCleaner.exe
        ├── customtkinter/
        ├── tkinterdnd2/
        └── (resto de dependencias)
```

### Arquitectura del código (`app.py`)

El archivo está organizado en capas bien separadas:

| Sección | Responsabilidad |
|---|---|
| `MetadataExtractor` | Lee archivos y devuelve `{campo: valor}` con todos los metadatos. Sin dependencia de la UI. |
| `MetadataCleaner` | Recibe rutas de entrada/salida, procesa el archivo y lo guarda limpio. Sin dependencia de la UI. |
| `TablaMetadatos` | Widget scrollable que renderiza el diccionario de metadatos como filas. |
| `ZonaCarga` | Frame con drag & drop y botón de selección. |
| `MetaCleanerApp` | Orquestador principal. Une UI y backend, gestiona hilos y estado. |

---

## Configuración avanzada

### Cambiar el tema de color

Los colores están definidos en el diccionario `C` al inicio de `app.py`. Puedes modificarlos libremente:

```python
C = {
    "accent":     "#7C6AF7",   # Color principal (morado por defecto)
    "success":    "#22C55E",   # Color de éxito (verde)
    "bg_main":    "#0D0D0D",   # Fondo principal
    ...
}
```

### Añadir soporte para un nuevo formato

1. En `MetadataExtractor`: añade el método `_nuevo_formato(ruta)` y regístralo en el diccionario `mapa` de `extraer()`.
2. En `MetadataCleaner`: añade el método `_nuevo_formato(entrada, salida)` y regístralo en `limpiar()`.
3. Añade la extensión al conjunto `MetadataExtractor.FORMATOS`.
4. Actualiza el filtro de archivos en `ZonaCarga._al_clic()`.

---

## Generar el ejecutable

### Compilación automática (recomendado)

```bat
# Doble clic en build.bat
# O desde CMD:
cd c:\Proyectos\appLimpiarMetadatos
build.bat
```

El script hace automáticamente:

1. Verifica que Python esté instalado.
2. Instala/actualiza las dependencias de `requirements.txt`.
3. Instala PyInstaller si no está presente.
4. Genera el icono si no existe.
5. Limpia la compilación anterior.
6. Ejecuta PyInstaller con el archivo `metacleaner.spec`.
7. Abre la carpeta `dist/MetaCleaner/` en el Explorador al terminar.

### Compilación manual

```bat
pip install pyinstaller
pyinstaller metacleaner.spec --clean --noconfirm
```

### Qué incluye el ejecutable

PyInstaller empaqueta en `dist/MetaCleaner/`:

- `MetaCleaner.exe` — punto de entrada
- `customtkinter/` — temas e imágenes de la UI
- `tkinterdnd2/` — DLLs de drag & drop nativo
- `_internal/` — todas las dependencias Python compiladas
- Runtime de Python embebido

> **Importante**: debes distribuir **toda la carpeta** `MetaCleaner/`, no solo el `.exe`. El ejecutable necesita los archivos junto a él.

### Notas para futuras versiones

Al actualizar el código y necesitar recompilar:

```bat
# Opción 1 — Rápida (sin limpiar caché)
pyinstaller metacleaner.spec --noconfirm

# Opción 2 — Limpia (recomendada si cambian dependencias)
pyinstaller metacleaner.spec --clean --noconfirm

# Opción 3 — Automática
build.bat
```

Si añades nuevas dependencias, actualiza también:

1. `requirements.txt` — añade el paquete.
2. `metacleaner.spec` — añade `collect_data_files('nuevo_paquete')` en la sección `datas` y los imports necesarios en `hiddenimports`.

---

## Solución de errores comunes

### La aplicación no arranca / pantalla negra

**Causa**: falta alguna DLL de Visual C++ Redistributable.  
**Solución**: instala [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe).

---

### Error: `No module named 'tkinterdnd2'`

**Causa**: el paquete de drag & drop no está instalado.  
**Solución**:
```bash
pip install tkinterdnd2
```
La app funciona igualmente sin drag & drop (solo con el botón de selección).

---

### Error al procesar PDF: `No hay librería PDF disponible`

**Causa**: ni `pypdf` ni `PyPDF2` están instalados.  
**Solución**:
```bash
pip install pypdf
```

---

### El ejecutable muestra consola negra al inicio

**Causa**: el antivirus intercepta el inicio o hay un error en el bootloader.  
**Solución**: añade la carpeta `dist/MetaCleaner/` a las exclusiones del antivirus.

---

### Error de permisos al guardar el archivo limpio

**Causa**: el archivo de destino está abierto en otro programa (p. ej. Word o Acrobat).  
**Solución**: cierra el programa que tiene el archivo abierto antes de guardar.

---

### Las imágenes JPEG pierden calidad

**Comportamiento esperado**: la limpieza de EXIF reconstruye la imagen desde los píxeles crudos y re-codifica el JPEG con calidad 95. Esto implica una recompresión mínima inevitable, ya que el formato JPEG es siempre con pérdida. Para imágenes críticas, usa el formato PNG como destino.

---

### PyInstaller falla con `RecursionError`

**Causa**: Python 3.12+ puede requerir aumentar el límite de recursión durante el análisis.  
**Solución**: añade al inicio de `metacleaner.spec`:
```python
import sys
sys.setrecursionlimit(5000)
```

---

### `UPX` no encontrado (advertencia durante la compilación)

**Causa**: UPX (compresor de ejecutables) no está en el PATH.  
**Efecto**: el ejecutable no se comprime, pero funciona igual.  
**Solución (opcional)**: descarga [UPX](https://upx.github.io/) y cópialo a `C:\Windows\System32\` o a la carpeta del proyecto.

---

## Historial de versiones

### v1.0.0 — 2025-05-25

- Lanzamiento inicial.
- Soporte para DOCX, XLSX, PDF, PNG, JPG y JSON.
- Interfaz en tema oscuro con CustomTkinter.
- Drag & drop mediante tkinterdnd2.
- Extracción de metadatos con tabla visual detallada.
- Limpieza no destructiva (el original no se modifica).
- Verificación automática post-limpieza.
- Barra de progreso animada.
- Procesamiento en hilo secundario (UI siempre responsiva).
- Ejecutable portátil para Windows generado con PyInstaller.

---

## Licencia

Uso personal y educativo. Las librerías de terceros (CustomTkinter, pypdf, Pillow, etc.) mantienen sus propias licencias.
