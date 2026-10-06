"""
Detalle de cupones de J.H. (pedido del usuario, 2026-10-06, chat 21): el PDF
de detalle de tarjetas que se imprime desde el portal de J.H., de un cupón
("Credit Card Batch Detail", trae el DDC) o de varios a la vez ("Credit Card
Detail", el DDC sale "N/A"). Cada grupo termina en su fila Totals y es una
fila del reporte mensual de cupones (Credit Card Daily Summary): un depósito
de uno o varios DDC. Trae cada batch con su fecha real de venta, número,
Gross, Fees y Net.

Lo validado con agosto-septiembre 2026 (67 grupos, 1,131 batches): la fecha
del batch es el día de Store Info; los batches del POS (número de 4 dígitos:
0xxx, 82xx, 92xx) suman lo vendido con tarjeta de ese día, salvo un batch que
cruza la medianoche y se compensa con el día de al lado; los demás ("slri…",
de 7 dígitos) no están en Store Info. Lectura sola; se guarda en eft_db
(cupon_detail_groups/cupon_detail_batches) y el control día por día está en
control_tarjetas.py.
"""

import re
from datetime import datetime

import pdfplumber

_NUM = r"(-?[\d,]+\.\d{2})"
_ROW_RE = re.compile(r"^(\d{1,2}/\d{1,2}/\d{4})\s+(\S+)\s+(\S+)\s+" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM + r"$")
_TOTALS_RE = re.compile(r"^Totals\s+" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM + r"$")
_TITLE_RE = re.compile(r"Credit Card (Batch )?Detail")
_COUPON_RE = re.compile(r"^DDC-\d+$")


def _num(text):
    return round(float(text.replace(",", "")), 2)


def extract_detail_groups(pdf_path):
    """
    [{"coupons": [DDC] o None, "batches": [{"batch_date" (ISO), "batch",
    "gross", "fees", "net"}], "totals": {"gross", "fees", "net"}}], un
    grupo por fila Totals (un grupo que sigue en la página siguiente se
    junta solo). ValueError si no es un detalle de tarjetas de J.H., si no
    tiene grupos o si un grupo no suma su fila Totals (no se guarda nada).
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if not _TITLE_RE.search(text):
        raise ValueError("Uno de los archivos no es un detalle de cupones de J.H. (Credit Card Detail).")
    groups, batches = [], []
    for raw in text.splitlines():
        line = raw.strip()
        row = _ROW_RE.match(line)
        if row:
            day, batch, description, gross, fees, net = row.groups()
            batches.append({
                "batch_date": datetime.strptime(day, "%m/%d/%Y").date().isoformat(),
                "batch": batch, "description": description,
                "gross": _num(gross), "fees": _num(fees), "net": _num(net),
            })
            continue
        totals = _TOTALS_RE.match(line)
        if totals:
            printed = [_num(v) for v in totals.groups()]
            read = [round(sum(b[k] for b in batches), 2) for k in ("gross", "fees", "net")]
            if not batches or any(abs(a - b) >= 0.005 for a, b in zip(read, printed)):
                raise ValueError(
                    f"Un grupo del detalle de cupones (Totals ${printed[0]:,.2f}) no suma sus batches: no se guardó el archivo."
                )
            coupons = sorted({b["description"] for b in batches if _COUPON_RE.match(b["description"])})
            groups.append({
                "coupons": coupons or None,
                "batches": [{k: v for k, v in b.items() if k != "description"} for b in batches],
                "totals": dict(zip(("gross", "fees", "net"), printed)),
            })
            batches = []
    if batches:
        raise ValueError("El detalle de cupones termina sin su fila Totals: no se guardó el archivo.")
    if not groups:
        raise ValueError("Un detalle de cupones no tiene ningún batch.")
    return groups
