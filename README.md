# App de Facturación Streamlit

Una aplicación Streamlit que genera facturas desde un Excel de productos y permite agregar nuevos productos al mismo archivo.

## Requisitos

- Python 3.11+ recomendado
- `streamlit`
- `pandas`
- `openpyxl`
- `XlsxWriter`

## Instalación

1. Crea y activa un entorno virtual:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instala dependencias:

```powershell
pip install -r requirements.txt
```

## Uso

```powershell
streamlit run app.py
```

## Archivos importantes

- `app.py`: aplicación principal
- `Productos.xlsx`: archivo de datos de productos
- `requirements.txt`: dependencias de Python

## Notas

- El archivo `Productos.xlsx` debe contener la hoja `Productos`.
- La app busca columnas de nombre de producto en el siguiente orden: `Producto_Web`, `Nombre_Drive`, `Nombre_Genesys`, `Producto`.
- La columna `Precio` es obligatoria para poder generar facturas correctamente.
