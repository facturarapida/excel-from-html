# -*- coding: utf-8 -*-
# Extraccion de tablas HTML a Excel (CSV opcional).
#
# Problema: un cliente recibe datos en una pagina web (o un email con tabla HTML)
# y necesita los numeros en Excel para analizarlos. Hacerlo a mano:
# copiar-pegar, limpiar cabeceras, convertir precios... 15 min por tabla.
#
# Esta herramienta lo hace en 1 segundo:
#   python html2excel.py --html datos.html              # local
#   python html2excel.py --url https://ejemplo.com/t    # web
#   python html2excel.py --html datos.html --sheet "Q4" --csv
#
# Salida:
#   - <nombre>.xlsx : Excel con cabecera en negrita, filtros, anchos de columna
#   - <nombre>.csv  : solo con --csv (UTF-8-BOM)
import argparse
import csv
import html
import pathlib
import re
import sys

import requests
from bs4 import BeautifulSoup


def load_source(args):
    """Devuelve (html, nombre_base_para_salidas)."""
    if args.url:
        r = requests.get(args.url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
        r.raise_for_status()
        slug = re.sub(r"\W+", "-", args.url.split("/")[2]).strip("-")
        return r.text, slug or "web"
    p = pathlib.Path(args.html)
    if not p.exists():
        raise FileNotFoundError(args.html)
    return p.read_text(encoding="utf-8-sig"), p.stem


def extract_tables(soup, table_index):
    tables = soup.find_all("table")
    if not tables:
        raise ValueError("no hay <table> en el HTML")
    if table_index >= len(tables):
        raise ValueError(f"solo hay {len(tables)} tablas, pediste la {table_index}")
    rows = []
    for tr in tables[table_index].find_all("tr"):
        cells = tr.find_all(["th", "td"])
        row = [html.unescape(c.get_text(" ", strip=True)) for c in cells]
        if any(row):
            rows.append(row)
    return rows


def to_number(v):
    """'1.234,56' o '1234.56' o '50%' -> numero, si no, texto."""
    s = v.replace("\u00a0", "").strip()
    if s.endswith("%"):
        inner = s[:-1].strip()
        inner = inner.replace(".", "").replace(",", ".") if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", inner) else inner.replace(",", ".")
        m = re.fullmatch(r"-?\d+\.?\d*", inner)
        return float(m.group(0)) / 100 if m else s
    # formato es "12.450,00" (punto separa miles) o "1234.56" (punto decimal)
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d+(,\d+)?", s):
        s = s.replace(",", ".")
    m = re.fullmatch(r"-?\d+\.?\d*", s)
    return float(m.group(0)) if m else v


def write_xlsx(rows, path, sheet):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = sheet[:31]
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row, 1):
            if i > 1:
                val = to_number(val)
            ws.cell(row=i, column=j, value=val)
    for c in ws[1]:
        c.font = Font(bold=True)
    ncol = len(rows[0]) if rows else 1
    for j in range(1, ncol + 1):
        width = max(len(str(ws.cell(row=i, column=j).value)) for i in range(1, len(rows) + 1)) + 2
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(10, width))
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Extraccion de tablas HTML a Excel")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--html", help="fichero .html local")
    src.add_argument("--url", help="URL de una pagina HTML con tabla")
    ap.add_argument("--table", type=int, default=0, help="indice de tabla si hay varias (defecto 0)")
    ap.add_argument("--sheet", default="Datos", help="nombre de la hoja (defecto 'Datos')")
    ap.add_argument("--csv", action="store_true", help="generar tambien CSV")
    args = ap.parse_args()

    try:
        text, base = load_source(args)
        soup = BeautifulSoup(text, "lxml")
        rows = extract_tables(soup, args.table)
    except Exception as exc:
        print(f"FALLO: {exc}", file=sys.stderr)
        return 1
    if not rows:
        print("FALLO: tabla vacia", file=sys.stderr)
        return 1

    outdir = pathlib.Path(__file__).parent / "output"
    outdir.mkdir(exist_ok=True)
    xlsx = outdir / f"{base}.xlsx"
    write_xlsx(rows, xlsx, args.sheet)
    csv_path = None
    if args.csv:
        csv_path = outdir / f"{base}.csv"
        with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerows(rows)

    print(f"OK  {len(rows)} filas x {len(rows[0])} columnas")
    print(f"XLSX {xlsx}")
    if csv_path:
        print(f"CSV  {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
