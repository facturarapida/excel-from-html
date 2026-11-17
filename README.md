# excel-from-html

**Extraccion de tablas HTML a Excel** — convierte cualquier tabla HTML (de una pagina web, un informe, o un email) en un `xlsx` limpio y analizable, con los precios y porcentajes ya convertidos a numeros.

## El problema que resuelve

Un cliente recibe datos en una pagina web (o un email con una tabla HTML) y los necesita en Excel para analizarlos. Hacerlo a mano: copiar-pegar, limpiar cabeceras, convertir "12.450,00" a numero, interpretar el 51,8%... 15-30 min por tabla, con riesgo de errores.

## La solucion (este repo)

`html2excel.py` hace en ~1 segundo:

1. Lee el HTML (fichero local `--html` o pagina `--url`).
2. Extrae la tabla (o la N-esima si hay varias, `--table N`).
3. **Convierte a numero** las celdas que lo son: `12.450,00` → `12450.0`, `51,8%` → `0.518` (con formato de porcentaje, estandar de Excel), `45.581,25` → `45581.25`.
4. Escribe `<nombre>.xlsx` con cabecera en negrita, filtros automaticos y anchos de columna. Con `--csv` genera tambien un CSV UTF-8-BOM.

## Ejecucion

```bash
pip install -r requirements.txt
python html2excel.py --html datos.html --sheet "Ventas Q4" --csv
python html2excel.py --url https://sitio.com/informe --table 1
```

## Resultado real (ejecucion del 2026-11-17)

Entrada: `datos.html` — informe de ventas Q4 con 5 regiones + fila de totales (formato `12.450,00` y `51,8%`).

```
OK  7 filas x 6 columnas
XLSX output/datos.xlsx
CSV  output/datos.csv
```

Verificacion del `datos.xlsx` (sheet "Ventas Q4", 7 filas x 6 columnas):

| Celda | Valor en HTML | Valor en Excel | Tipo |
|-------|---------------|----------------|------|
| B2 | `12.450,00` | `12450.0` | numero |
| F2 | `51,8%` | `0.518` | numero (formato %) |
| D7 | `50.142,25` | `50142.25` | numero |
| F7 | `37,8%` | `0.378` | numero (formato %) |

Con los datos como numero, en Excel puedes hacer SUMA, PIVOT y graficos directamente. Como texto, no.

El fichero `muestra-resultado.xlsx` de este repo es la salida REAL de esa ejecucion: abrela en Excel y veras los valores numericos, los filtros activos y los porcentajes como numero con formato.

## Que se automatiza (vs. manual)

| Tarea | Manual | Con este repo |
|-------|--------|---------------|
| Copiar-pegar la tabla | ~10 min | 0 |
| Limpiar "12.450,00" a numero | ~10 min | automatico |
| Interpretar porcentajes | a ojo | 0.518 con formato % |
| Filtros en Excel | a mano | ya incluidos |
| Riesgo de error de tecleo | alto | 0 |

## Adaptar a tu sitio

* Si tu tabla esta dentro de un `<div>` con varias tablas, usa `--table N`.
* Si hay celdas con formato exotico (fechas, monedas raras), ajusta `to_number()`.
* Para paginas con JavaScript, combina con el repo hermano `rpa-ecommerce-bot` (Selenium renderiza la pagina y guarda el HTML; `html2excel.py` lo convierte).

## Dependencias

`requests`, `beautifulsoup4`, `lxml`, `openpyxl` (ver `requirements.txt`). Python >= 3.10.

## Licencia

MIT — uso libre como punto de partida para extracciones propias.
