# Generador de Registro Fotográfico por Punto de Control

Script de Python que automatiza la creación de un documento Word con las fotografías de campo organizadas por punto de control kilométrico, a partir de archivos KMZ y un reporte geológico en Word.

> Ingeniería de prompt: **Ing. Vanessa Olayo Peñaloza**

---

## ¿Qué hace?

1. **Lee el reporte geológico** (`.docx`) y extrae automáticamente todos los cadenamienos del tramo (formato `KM XXX+XXX`) usando expresiones regulares.
2. **Obtiene las coordenadas GPS** de cada cadenamiento interpolando sobre la polilínea del eje vial contenida en el KMZ de diseño geométrico (TRAZO). Si existe un placemark con nombre coincidente, usa su coordenada exacta.
3. **Extrae las fotografías** y coordenadas GPS de cada placemark del KMZ de campo (Timemark).
4. **Asigna cada fotografía** al punto de control más cercano, calculando la distancia Haversine entre la coordenada de la foto y la de cada punto de control.
5. **Genera un documento Word** (`.docx`) con una sección por punto de control, donde las fotos se presentan en una tabla de 2 columnas con pie de foto (nombre del placemark, latitud y longitud).

---

## Archivos de entrada (Input)

Los tres archivos deben estar en la misma carpeta que el script:

| Archivo | Tipo | Descripción |
|---|---|---|
| `Timemark_XXXXXXXXXXXXXX.kmz` | KMZ de campo | Generado con la app Timemark. Contiene los placemarks con fotografías y coordenadas GPS capturadas en campo. |
| `TRAZO KM XXX-XXX.kmz` | KMZ de diseño | KMZ del proyecto geométrico del tramo. Contiene la polilínea del eje vial y los placemarks de diseño usados para georreferenciar los cadenamienos. |
| `REPORTE GEOLOGICO.docx` | Word | Reporte geológico de campo. El script extrae de este documento todos los cadenamienos con formato `KM XXX+XXX` para definir los puntos de control. |

> Las rutas de los tres archivos se configuran al inicio del script en las variables `KMZ_TIMEMARK`, `KMZ_TRAZO` y `DOCX_GEO`.

---

## Archivo de salida (Output)

| Archivo | Descripción |
|---|---|
| `Fotografias_por_Punto_Control.docx` | Documento Word generado automáticamente en la misma carpeta del script. |

### Estructura del Word generado

- Portada con nombre del tramo y fuente del KMZ.
- Una sección por cada punto de control encontrado en el reporte geológico.
- Cada sección incluye:
  - Encabezado: `Punto de control KM XXX+XXX`
  - Coordenadas de referencia del punto de control
  - Número total de fotografías asignadas
  - Tabla de 2 columnas con las fotografías y su pie de foto (nombre del placemark, latitud y longitud)

---

## Requisitos

### Python
Python 3.7 o superior.

### Dependencias
Solo se requiere instalar una librería externa:

```bash
pip install python-docx
```

Las demás librerías usadas (`zipfile`, `xml.etree.ElementTree`, `math`, `io`, `os`, `re`) son parte de la biblioteca estándar de Python y no requieren instalación.

---

## Instalación y uso

```bash
# 1. Instalar dependencia
pip install python-docx

# 2. Colocar los tres archivos de entrada en la misma carpeta que el script

# 3. Ejecutar
python generar_fotos_word.py
```

El script imprime en consola el progreso de cada etapa y al finalizar indica la ruta del archivo generado.

---

## Configuración

Si los nombres de los archivos de entrada cambian, actualizar las siguientes variables al inicio del script:

```python
KMZ_TIMEMARK = os.path.join(SCRIPT_DIR, "Timemark_202603222011.kmz")
KMZ_TRAZO    = os.path.join(SCRIPT_DIR, "TRAZO KM 100-110.kmz")
DOCX_GEO     = os.path.join(SCRIPT_DIR, "ZACUALTIPAN MOLANGO - KM 99+960-105+000.docx")
OUTPUT_DOCX  = os.path.join(SCRIPT_DIR, "Fotografias_por_Punto_Control.docx")
```

También se puede ajustar el cadenamiento de inicio del eje vial (en metros) si el tramo cambia:

```python
KM_INICIO = 99_960   # cadenamiento del primer punto del eje (metros)
```
