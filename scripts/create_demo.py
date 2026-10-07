"""Genera un documento ficticio para probar el portal sin datos de clientes."""

import argparse
from pathlib import Path

import pymupdf

parser = argparse.ArgumentParser(description="Crear un certificado PDF ficticio sin validez.")
parser.add_argument("--output", type=Path, default=Path("/tmp/qcp-demo.pdf"))
args = parser.parse_args()
if args.output.exists():
    parser.error("El archivo de destino ya existe. Elige otra ruta para no sobrescribirlo.")
with pymupdf.open() as doc:
    for number in range(1, 3):
        page = doc.new_page()
        page.insert_text((50, 65), "CERTIFICADO DE PRUEBA - SIN VALIDEZ", fontsize=18)
        page.insert_text((50, 105), "QCP-2026-0001", fontsize=12)
        page.insert_text((50, 145), f"Pagina {number} de 2. Documento ficticio para probar la correccion y el QR.", fontsize=11)
        page.insert_text((50, 175), "Resultado de ejemplo: conforme. No usar como certificado real.", fontsize=11)
    doc.save(args.output)
print(f"PDF de prueba creado: {args.output}")
