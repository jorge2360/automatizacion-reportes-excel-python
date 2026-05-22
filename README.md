# Automatización de Reportes Excel con Python

Proyecto desarrollado en Python para automatizar la generación de reportes empresariales en Excel a partir de archivos CSV.

---

## Descripción

El sistema procesa información de ventas e inventario utilizando Python y genera automáticamente:

- reportes Excel
- resúmenes ejecutivos
- análisis de ventas
- inventario valorizado
- gráficos automáticos
- hojas organizadas por categorías

El objetivo es demostrar automatización empresarial orientada a análisis de datos y generación de reportes.

---

## Tecnologías utilizadas

- Python
- pandas
- openpyxl
- matplotlib
- Excel Automation

---

## Estructura del proyecto

```text
automatizacion-reportes-excel-python/
├── entrada/
│   ├── ventas.csv
│   └── inventario.csv
├── salida/
│   └── reporte_empresarial.xlsx
├── src/
│   └── generar_reporte.py
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Funcionalidades implementadas

### Procesamiento de datos

- Lectura de archivos CSV
- Limpieza básica de datos
- Cálculo automático de ventas
- Cálculo de inventario valorizado

### Reportes Excel

- Generación automática de Excel
- Múltiples hojas organizadas
- Resumen ejecutivo
- Reportes empresariales

### Formato profesional

- Encabezados estilizados
- Bordes automáticos
- Ajuste de columnas
- Formato de moneda
- Congelación de filas

### Gráficos automáticos

- Ventas por categoría
- Ventas por producto

---

## Instalación

### Clonar repositorio

```bash
git clone https://github.com/jorge2360/automatizacion-reportes-excel-python.git
```

### Ingresar al proyecto

```bash
cd automatizacion-reportes-excel-python
```

### Crear entorno virtual

```bash
python -m venv .venv
```

### Activar entorno virtual

PowerShell:

```bash
.venv\Scripts\Activate
```

### Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## Ejecución

```bash
python src/generar_reporte.py
```

El sistema generará:

```text
salida/reporte_empresarial.xlsx
```

---

## Archivos de entrada

### ventas.csv

Contiene:
- fecha
- cliente
- producto
- categoría
- cantidad
- precio_unitario

### inventario.csv

Contiene:
- producto
- categoría
- stock
- precio_compra
- precio_venta

---

## Resultados generados

El Excel incluye:

- Resumen ejecutivo
- Ventas
- Inventario
- Ventas por categoría
- Ventas por producto
- Inventario valorizado
- Gráficos automáticos

---

## Buenas prácticas implementadas

- Modularización de funciones
- Automatización de procesos
- Uso de entorno virtual
- Organización de archivos
- Generación automática de reportes
- Formato profesional de Excel
- Visualización de datos

---

## Autor

Jorge García