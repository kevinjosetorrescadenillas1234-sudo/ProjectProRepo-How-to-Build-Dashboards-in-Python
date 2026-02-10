# ProjectProRepo-How-to-Build-Dashboards-in-Python

Dashboard interactivo con **Dash + Plotly** que lee un archivo Excel y genera gráficos básicos de:

- Participación
- FOB total
- KG NT totales

## Requisitos

```bash
pip install -r requirements.txt
```

## Ejecutar

```bash
python app.py
```

Luego abre: `http://localhost:8050`

## Formato esperado del Excel

El dashboard intenta detectar nombres de columnas comunes (sin importar mayúsculas/minúsculas):

- Participación: se calcula automáticamente como `(suma FOB por país / FOB total) * 100`
- FOB: `fob total`, `fob`, `valor fob`, `total fob`
- KG NT: `kg nt totales`, `kg nt total`, `kg`, `kilogramos`

Además, para los gráficos de FOB y KG prioriza una columna de país (`pais`, `país`, `paises`, `países`, `country`, `countries`) y, si no existe, usa la primera columna no numérica.

Si no encuentra alguna columna, el dashboard muestra un mensaje y mantiene el resto de gráficos disponibles.
