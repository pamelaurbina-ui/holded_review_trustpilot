#!/usr/bin/env python3
"""
Actualiza un Excel con las reviews de Trustpilot (Holded) publicadas en un feed RSS de rss.app.

Uso:
    python update_reviews.py

Comportamiento:
- Descarga el XML del feed.
- Extrae: fecha, título, texto de la review, rating, nombre del reviewer y país.
- Si el Excel de salida ya existe, añade solo las reviews nuevas (dedupe por GUID).
- Si no existe, lo crea con cabeceras y formato.
"""

import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

FEED_URL = "https://rss.app/feeds/cgBhx0wJysNg5qoP.xml"
OUTPUT_XLSX = "holded_reviews.xlsx"
SHEET_NAME = "Reviews"

HEADERS = ["Fecha", "GUID", "Titulo", "Review", "Rating", "Reviewer", "Pais", "Link"]

# La descripcion viene como HTML tipo:
# <div>Texto de la review<br><br>Rating: 5/5 (Excellent)<br><br>Reviewer: Nombre, ES</div>
RATING_RE = re.compile(r"Rating:\s*(\d+)/5\s*\(([^)]+)\)")
REVIEWER_RE = re.compile(r"Reviewer:\s*(.*?),\s*([A-Z]{2})\s*$")


def fetch_feed(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def strip_html(raw: str) -> str:
    # Quita tags HTML y separa lineas donde habia <br>
    text = re.sub(r"<br\s*/?>", "\n", raw)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()


def parse_description(raw_html: str):
    text = strip_html(raw_html)

    rating_match = RATING_RE.search(text)
    rating = f"{rating_match.group(1)}/5 ({rating_match.group(2)})" if rating_match else ""

    reviewer_match = REVIEWER_RE.search(text.split("\n")[-1] if text else "")
    # Buscar la linea "Reviewer: ..." explicitamente por si no es la ultima
    reviewer_name, country = "", ""
    for line in text.split("\n"):
        line = line.strip()
        m = REVIEWER_RE.search(line)
        if m:
            reviewer_name, country = m.group(1), m.group(2)
            break

    # El cuerpo de la review es todo antes de "Rating:"
    body = text.split("Rating:")[0].strip()

    return body, rating, reviewer_name, country


def parse_feed(xml_bytes: bytes):
    root = ET.fromstring(xml_bytes)
    items = []
    for item in root.findall(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        guid = (item.findtext("guid") or "").strip()
        pub_date_raw = (item.findtext("pubDate") or "").strip()
        description_raw = item.findtext("description") or ""

        try:
            pub_date = datetime.strptime(pub_date_raw, "%a, %d %b %Y %H:%M:%S %Z")
        except ValueError:
            pub_date = None

        body, rating, reviewer_name, country = parse_description(description_raw)

        items.append({
            "fecha": pub_date.strftime("%Y-%m-%d") if pub_date else pub_date_raw,
            "guid": guid,
            "titulo": title,
            "review": body,
            "rating": rating,
            "reviewer": reviewer_name,
            "pais": country,
            "link": link,
        })
    return items


def load_or_create_workbook(path: Path):
    if path.exists():
        wb = load_workbook(path)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.create_sheet(SHEET_NAME)
        return wb, ws
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Arial", bold=True)
        cell.alignment = Alignment(horizontal="center")
    return wb, ws


def existing_guids(ws) -> set:
    guids = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and len(row) > 1 and row[1]:
            guids.add(row[1])
    return guids


def autosize_columns(ws):
    widths = {1: 12, 2: 34, 3: 30, 4: 60, 5: 16, 6: 22, 7: 8, 8: 45}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def main():
    output_path = Path(OUTPUT_XLSX)

    print(f"Descargando feed: {FEED_URL}")
    xml_bytes = fetch_feed(FEED_URL)

    print("Parseando reviews...")
    reviews = parse_feed(xml_bytes)
    print(f"Reviews encontradas en el feed: {len(reviews)}")

    wb, ws = load_or_create_workbook(output_path)
    known_guids = existing_guids(ws)

    new_count = 0
    for r in reviews:
        if r["guid"] in known_guids:
            continue
        ws.append([
            r["fecha"], r["guid"], r["titulo"], r["review"],
            r["rating"], r["reviewer"], r["pais"], r["link"],
        ])
        for col_idx in range(1, len(HEADERS) + 1):
            ws.cell(row=ws.max_row, column=col_idx).font = Font(name="Arial")
            ws.cell(row=ws.max_row, column=col_idx).alignment = Alignment(
                vertical="top", wrap_text=(col_idx == 4)
            )
        new_count += 1

    autosize_columns(ws)
    wb.save(output_path)

    print(f"Reviews nuevas anadidas: {new_count}")
    print(f"Total de reviews en el Excel: {ws.max_row - 1}")
    print(f"Guardado en: {output_path.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
