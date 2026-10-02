"""
Productos de las facturas de proveedores -- pedido explícito del usuario
(2026-09-28): "crear un sistema organizado en el que extraigamos los
productos que compramos con su precio, para así utilizar esa info para ver
cómo van cambiando los costos de los mismos productos en cada compra y
también poder compararlo con el CMV... a qué proveedor se le compra ese
producto por su UPC". Lectura/extracción + cálculo, sin UI -- el guardado
vive en proveedores_db (tabla supplier_invoice_lines) y las pantallas en la
sección Proveedores de webapp.py.

Por ahora solo H.T. Hackney: es el único proveedor cuyas facturas traen
texto real (el resto son escaneos -- ver el relevamiento del 2026-09-28 en
HISTORIAL.md). Cada proveedor nuevo se suma en LINE_EXTRACTORS.

Regla de oro: el detalle de una factura se guarda solo si cierra al
centavo -- cantidad x neto = total de cada renglón, la suma de renglones =
INVOICE SUBTOTAL impreso, y subtotal + cargos = Total impreso. Si algo no
cierra, ValueError y no se guarda ningún renglón (nunca un detalle a medias).
"""

import re

import pdfplumber

# Columnas de la grilla de H.T., en el orden de los 16 bloques de guiones
# que la factura imprime debajo del encabezado ("- ------- -------------- ...").
_HT_COLUMNS = [
    "flag", "qty", "upc", "pack", "size", "description", "item_no", "srp_pct",
    "srp", "srp_ext", "price", "allowance", "tax", "net", "ext", "tx",
]
# Un cartón de cigarrillos (size "CTN") son 10 paquetes, y el POS tiene el
# costo por paquete: Marlboro Gold $102.95 el cartón = $10.30 en el POS.
_UNITS_PER_CARTON = 10


def normalize_upc(value):
    """Solo dígitos, sin ceros adelante -- así cruza con el UPC del POS (CMV)."""
    digits = re.sub(r"\D", "", str(value or ""))
    return digits.lstrip("0")


def product_key(supplier_key, upc, item_no):
    """Clave de producto: el UPC; sin UPC, el Item # propio del proveedor."""
    return upc if upc else f"item-{supplier_key}-{item_no}"


def _num(text):
    text = (text or "").replace(",", "").replace("$", "").strip()
    if not text:
        return None
    negative = text.endswith("-") or text.startswith("-")
    text = text.strip("-").strip()
    try:
        value = float(text)
    except ValueError:
        return None
    return -value if negative else value


def _page_rows(page):
    """Renglones de la página: palabras agrupadas por altura, de izquierda a derecha."""
    rows = []
    last_top = None
    for word in sorted(page.extract_words(), key=lambda w: (round(w["top"]), w["x0"])):
        top = round(word["top"])
        if last_top is not None and top - last_top <= 2:
            rows[-1].append(word)
        else:
            rows.append([word])
            last_top = top
    for row in rows:
        row.sort(key=lambda w: w["x0"])
    return rows


def _column_index(dash_ranges, word):
    center = (word["x0"] + word["x1"]) / 2

    def distance(index):
        start, end = dash_ranges[index]
        if start - 1 <= center <= end + 1:
            return 0
        return min(abs(center - start), abs(center - end))

    return min(range(len(dash_ranges)), key=distance)


def extract_ht_hackney_lines(pdf_path):
    """
    Renglones de producto de una factura de H.T. Hackney. Las columnas se
    leen por posición (los bloques de guiones del encabezado marcan dónde
    empieza y termina cada una), porque varias vienen vacías según el
    producto (SRP%, Alw, Tax, y a veces el UPC) y leer el texto corrido las
    correría de lugar.
    """
    invoice_no = None
    lines = []
    charges = {}
    subtotal = None
    total = None
    category = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            dash_ranges = None
            for row in _page_rows(page):
                text = " ".join(w["text"] for w in row)
                if invoice_no is None:
                    match = re.search(r"Invoice #:\s*(\d+)", text)
                    if match:
                        invoice_no = match.group(1)
                # El encabezado repite el total en cada página ("**For
                # Reference Purposes Only** Total: ..."): no es el renglón final.
                if "For Reference Purposes Only" in text:
                    continue
                # Cargos, subtotal y total se buscan en toda la página, no solo
                # debajo de la grilla: a veces el Total cae solo en una última
                # página que ya no repite la grilla de columnas.
                match = re.match(r"^(FUEL SURCHARGE|DELIVERY CHARGE)\s+([\d,.]+-?)$", text)
                if match:
                    charges[match.group(1)] = _num(match.group(2))
                    continue
                # Al final del renglón y no al principio: a veces viene pegado
                # a un aviso de retiro de producto de la FDA ("...RECALLS Total: 4995.51").
                # El monto puede venir sin cero adelante ("Tax: .66").
                match = re.search(r"(?:Tax:\s*([\d,]*\.\d{2}-?)\s+)?Total:\s*([\d,]*\.\d{2}-?)$", text)
                if match:
                    if match.group(1):
                        charges["TAX"] = _num(match.group(1))
                    total = _num(match.group(2))
                    continue
                match = re.search(r"\*\*\* INVOICE SUBTOTAL \*\*\*.*?([\d,]+\.\d{2}-?)$", text)
                if match:
                    subtotal = _num(match.group(1))
                    continue
                if dash_ranges is None:
                    if len(row) >= 14 and all(set(w["text"]) == {"-"} for w in row):
                        if len(row) != len(_HT_COLUMNS):
                            raise ValueError(
                                "la grilla de productos de H.T. cambió de formato "
                                f"({len(row)} columnas en vez de {len(_HT_COLUMNS)})."
                            )
                        dash_ranges = [(w["x0"], w["x1"]) for w in row]
                    continue

                cells = {name: [] for name in _HT_COLUMNS}
                for word in row:
                    cells[_HT_COLUMNS[_column_index(dash_ranges, word)]].append(word["text"])
                cells = {name: " ".join(parts) for name, parts in cells.items()}
                if (re.fullmatch(r"\d+", cells["qty"]) and re.fullmatch(r"\d{5,8}", cells["item_no"])
                        and _num(cells["ext"]) is not None and _num(cells["net"]) is not None):
                    lines.append(_ht_line(cells, category, len(lines) + 1))
                    continue
                # Encabezado de sección ("CANDY", "CIGS, FULL PRICE"): en
                # mayúsculas y sin números. Las notas ("ABOVE IS ...", "THE
                # ABOVE ITEM WAS ...") no son sección.
                if text.isupper() and not re.search(r"\d", text) and "ABOVE" not in text and "***" not in text:
                    category = text

    if not lines:
        raise ValueError("no se encontró ningún renglón de producto en la factura de H.T.")
    for line in lines:
        if abs(line["qty"] * line["net"] - line["ext"]) > 0.011:
            raise ValueError(
                f"el renglón {line['line_no']} ({line['description']}) no cierra: "
                f"{line['qty']:g} x {line['net']:.2f} ≠ {line['ext']:.2f}."
            )
    lines_total = round(sum(line["ext"] for line in lines), 2)
    if subtotal is None or abs(lines_total - subtotal) >= 0.01:
        raise ValueError(
            f"los renglones suman ${lines_total:,.2f} y el INVOICE SUBTOTAL de la factura es "
            f"{'$' + format(subtotal, ',.2f') if subtotal is not None else 'ilegible'}."
        )
    if total is None or abs(subtotal + sum(charges.values()) - total) >= 0.01:
        raise ValueError("el subtotal más los cargos (fuel, delivery, tax) no da el Total de la factura.")
    if invoice_no is None:
        raise ValueError("no se encontró el N° de invoice en la factura de H.T.")
    return {"invoice_no": invoice_no, "lines": lines, "subtotal": subtotal, "total": total}


def _ht_line(cells, category, line_no):
    pack = int(_num(cells["pack"].lstrip("*")) or 1)
    size = cells["size"]
    units = pack * (_UNITS_PER_CARTON if size.upper() == "CTN" else 1)
    net = _num(cells["net"])
    return {
        "line_no": line_no,
        "upc": normalize_upc(cells["upc"]),
        "item_no": cells["item_no"],
        "description": cells["description"],
        "category": category,
        "qty": _num(cells["qty"]),
        "pack": pack,
        "size": size,
        "units": units,
        "price": _num(cells["price"]),
        "allowance": _num(cells["allowance"]),
        "tax": _num(cells["tax"]),
        "net": net,
        "ext": _num(cells["ext"]),
        "unit_cost": round(net / units, 4),
        "srp": _num(cells["srp"]),
    }


# supplier_key (el de SUPPLIER_REGISTRY en proveedores.py) -> extractor de renglones.
LINE_EXTRACTORS = {
    "ht_hackney": extract_ht_hackney_lines,
}


def extract_lines(supplier_key, pdf_path):
    """Detalle de productos de la factura, o None si el proveedor todavía no tiene extractor."""
    extractor = LINE_EXTRACTORS.get(supplier_key)
    return extractor(pdf_path) if extractor else None


# ---------------------------------------------------------------------------
# Resumen por producto y cruce con el POS (CMV)
# ---------------------------------------------------------------------------

def _pct(part, whole):
    return round(part / whole * 100, 1) if whole else None


def _pos_diff(pos_cost, unit_cost):
    if pos_cost is None:
        return None
    diff = pos_cost - unit_cost
    return 0.0 if abs(diff) <= 0.0051 else round(diff, 2)


def _group_lines(lines):
    """{clave de producto: [renglones del más viejo al más nuevo]}."""
    by_product = {}
    for line in lines:
        key = product_key(line["supplier_key"], line["upc"], line["item_no"])
        by_product.setdefault(key, []).append(line)
    return by_product


def _pos_index(pos_costs):
    return {normalize_upc(row.get("upc")): row for row in pos_costs if normalize_upc(row.get("upc"))}


def _summary(key, product_lines, supplier_labels, pos_by_upc):
    last = product_lines[-1]
    # Compra anterior = el último renglón de OTRA factura (el mismo producto
    # puede venir dos veces en una misma factura).
    previous = next(
        (line for line in reversed(product_lines[:-1])
         if (line["supplier_key"], line["invoice_no"]) != (last["supplier_key"], last["invoice_no"])),
        None,
    )
    change = round(last["unit_cost"] - previous["unit_cost"], 4) if previous else None
    pos = pos_by_upc.get(last["upc"]) if last["upc"] else None
    pos_cost = pos.get("cost") if pos else None
    pos_price = pos.get("price") if pos else None
    costs = [line["unit_cost"] for line in product_lines]
    return {
        "key": key,
        "upc": last["upc"],
        "item_no": last["item_no"],
        "description": last["description"],
        "category": last["category"],
        "size": last["size"],
        "units_per_case": last["units"],
        "suppliers": sorted({supplier_labels.get(line["supplier_key"], line["supplier_key"]) for line in product_lines}),
        "purchases": len({(line["supplier_key"], line["invoice_no"]) for line in product_lines}),
        "units_bought": sum((line["qty"] or 0) * (line["units"] or 0) for line in product_lines),
        "first_date": product_lines[0]["invoice_date"],
        "last_date": last["invoice_date"],
        "last_cost": last["unit_cost"],
        "previous_cost": previous["unit_cost"] if previous else None,
        "previous_date": previous["invoice_date"] if previous else None,
        "change": change,
        "change_pct": _pct(change, previous["unit_cost"]) if previous else None,
        "min_cost": min(costs),
        "max_cost": max(costs),
        "srp": last["srp"],
        "pos_name": pos.get("name") if pos else None,
        "pos_dept": pos.get("dept_name") if pos else None,
        "pos_cost": pos_cost,
        "pos_price": pos_price,
        # Diferencia entre el costo que tiene cargado el POS y el costo real
        # de la última compra -- si no es 0, el CMV está calculando el margen
        # con un costo viejo. Medio centavo o menos es redondeo del POS
        # (ej. $12.35 la caja de 10 = $1.235 la unidad, en el POS $1.24).
        "pos_cost_diff": _pos_diff(pos_cost, last["unit_cost"]),
        "margin_pct": _pct(pos_price - last["unit_cost"], pos_price) if pos_price else None,
    }


def build_product_list(lines, supplier_labels, pos_costs):
    """
    Un renglón por producto comprado (clave = UPC, o Item # si la factura no
    trae UPC), del comprado más recientemente al más viejo.
    """
    pos_by_upc = _pos_index(pos_costs)
    products = [
        _summary(key, product_lines, supplier_labels, pos_by_upc)
        for key, product_lines in _group_lines(lines).items()
    ]
    products.sort(key=lambda p: (p["last_date"], p["description"] or ""), reverse=True)
    return products


# ---------------------------------------------------------------------------
# Carpetas: proveedor -> facturas por fecha -> productos (2026-10-02)
# ---------------------------------------------------------------------------
# Pedido del usuario: "distintos módulos con fechas dentro de ellos como si
# fueran carpetas", solo de los proveedores que se leen bien (los de
# LINE_EXTRACTORS -- los únicos que guardan renglones), y cada factura
# compara sus precios contra la factura ANTERIOR POR FECHA, no por orden de
# carga: si se sube una del medio después, la más nueva pasa a compararse
# contra esa. Por eso todo se recalcula al mostrar, ordenando por fecha.

def _compare_supplier_invoices(supplier_lines):
    """
    Facturas de un proveedor del más vieja a la más nueva (fecha y N°), cada
    renglón con el costo de la última compra ANTERIOR de ese producto (de
    una factura con fecha previa; dos renglones de una misma factura nunca
    se comparan entre sí). Devuelve [(fecha, N°, [renglones])].
    """
    by_invoice = {}
    for line in sorted(supplier_lines, key=lambda l: (l["invoice_date"], str(l["invoice_no"]), l["line_no"] or 0)):
        by_invoice.setdefault((line["invoice_date"], line["invoice_no"]), []).append(line)
    last_by_product = {}
    result = []
    for (invoice_date, invoice_no), invoice_lines in sorted(by_invoice.items(), key=lambda item: (item[0][0], str(item[0][1]))):
        rows = []
        seen = {}
        for line in invoice_lines:
            key = product_key(line["supplier_key"], line["upc"], line["item_no"])
            previous = last_by_product.get(key)
            change = round(line["unit_cost"] - previous["unit_cost"], 4) if previous else None
            if change is None:
                state = "Nuevo"
            elif change > 0.0001:
                state = "Subió"
            elif change < -0.0001:
                state = "Bajó"
            else:
                state = "Igual"
            rows.append({
                **line,
                "product_key": key,
                "previous_cost": previous["unit_cost"] if previous else None,
                "previous_date": previous["invoice_date"] if previous else None,
                "previous_invoice_no": previous["invoice_no"] if previous else None,
                "change": change,
                "change_pct": _pct(change, previous["unit_cost"]) if previous else None,
                "state": state,
            })
            seen[key] = line
        last_by_product.update(seen)
        result.append((invoice_date, invoice_no, rows))
    return result


def _invoice_summary(invoice_date, invoice_no, rows, previous):
    count = lambda state: sum(1 for row in rows if row["state"] == state)
    return {
        "invoice_date": invoice_date,
        "invoice_no": invoice_no,
        "lines": len(rows),
        "total": round(sum(row["ext"] or 0.0 for row in rows), 2),
        "up": count("Subió"),
        "down": count("Bajó"),
        "same": count("Igual"),
        "new": count("Nuevo"),
        "previous_date": previous[0] if previous else None,
        "previous_invoice_no": previous[1] if previous else None,
    }


def build_supplier_invoices(supplier_key, lines):
    """Facturas del proveedor con su resumen contra la anterior, de la más nueva a la más vieja."""
    compared = _compare_supplier_invoices([line for line in lines if line["supplier_key"] == supplier_key])
    summaries = []
    for index, (invoice_date, invoice_no, rows) in enumerate(compared):
        previous = compared[index - 1][:2] if index > 0 else None
        summaries.append(_invoice_summary(invoice_date, invoice_no, rows, previous))
    summaries.reverse()
    return summaries


def build_invoice_products(supplier_key, invoice_date, invoice_no, lines):
    """(resumen, renglones) de una factura con cada producto contra su compra anterior; None si no existe."""
    compared = _compare_supplier_invoices([line for line in lines if line["supplier_key"] == supplier_key])
    for index, (date_value, number, rows) in enumerate(compared):
        if date_value == invoice_date and str(number) == str(invoice_no):
            previous = compared[index - 1][:2] if index > 0 else None
            return _invoice_summary(date_value, number, rows, previous), rows
    return None


def build_supplier_folders(lines, supplier_labels):
    """Una carpeta por proveedor con renglones guardados: cuántas facturas, de qué fechas y cómo vino la última."""
    keys = sorted({line["supplier_key"] for line in lines}, key=lambda k: supplier_labels.get(k, k).lower())
    folders = []
    for key in keys:
        invoices = build_supplier_invoices(key, lines)
        folders.append({
            "key": key,
            "label": supplier_labels.get(key, key),
            "invoices": len(invoices),
            "first_date": invoices[-1]["invoice_date"] if invoices else None,
            "last": invoices[0] if invoices else None,
            "years": sorted({inv["invoice_date"][:4] for inv in invoices}, reverse=True),
        })
    return folders


def build_product_detail(key, lines, supplier_labels, pos_costs, monthly_sales):
    """
    Resumen + historial de compras de un producto, y sus ventas por mes del
    POS (CMV) con el costo real de ese momento -- para ver a cuánto se vende
    contra lo que costó. None si ese producto nunca se compró.
    """
    product_lines = _group_lines(lines).get(key)
    if not product_lines:
        return None
    summary = _summary(key, product_lines, supplier_labels, _pos_index(pos_costs))

    purchases = []
    previous_cost = None
    for line in product_lines:
        change = round(line["unit_cost"] - previous_cost, 4) if previous_cost is not None else None
        purchases.append({
            **line,
            "supplier_label": supplier_labels.get(line["supplier_key"], line["supplier_key"]),
            "change": change,
            "change_pct": _pct(change, previous_cost) if change is not None else None,
        })
        previous_cost = line["unit_cost"]
    purchases.reverse()

    sales = []
    if summary["upc"]:
        for row in monthly_sales:
            if normalize_upc(row.get("upc")) != summary["upc"] or not row.get("count"):
                continue
            month_end = f"{row['year']:04d}-{row['month']:02d}-31"
            # Costo vigente en ese mes = el de la última compra hasta fin de mes.
            cost_lines = [line for line in product_lines if line["invoice_date"] <= month_end]
            cost = cost_lines[-1]["unit_cost"] if cost_lines else None
            avg_price = row["amount"] / row["count"]
            sales.append({
                "year": row["year"],
                "month": row["month"],
                "count": row["count"],
                "amount": row["amount"],
                "avg_price": round(avg_price, 2),
                "cost": cost,
                "margin_pct": _pct(avg_price - cost, avg_price) if cost is not None and avg_price else None,
            })
        sales.reverse()

    return {"product": summary, "purchases": purchases, "sales": sales}


# ---------------------------------------------------------------------------
# Reporte para el manager (pedido del usuario, 2026-10-02): "enviar un
# reporte en Excel o PDF directamente al manager de la estación de servicio
# con los productos que cambiaron de precio". Solo los que subieron o
# bajaron contra su compra anterior; con el precio del POS y el margen que
# deja el costo nuevo, para que sepa qué precio revisar.
# ---------------------------------------------------------------------------

def price_change_rows(rows, pos_costs):
    """Renglones de una factura (build_invoice_products) que cambiaron de costo: primero los que subieron."""
    pos_by_upc = _pos_index(pos_costs)
    changed = []
    for row in rows:
        if row["state"] not in ("Subió", "Bajó"):
            continue
        pos = pos_by_upc.get(row["upc"]) if row["upc"] else None
        pos_price = pos.get("price") if pos else None
        margin = _pct(pos_price - row["unit_cost"], pos_price) if pos_price else None
        changed.append({**row, "pos_price": pos_price, "margin_pct": margin})
    changed.sort(key=lambda r: (0 if r["state"] == "Subió" else 1, r.get("category") or "", r.get("description") or ""))
    return changed


def _ddmmyyyy(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[0:4]}" if iso else ""


def _pack_label(row):
    return " ".join(str(part) for part in (row.get("pack"), row.get("size")) if part not in (None, ""))


def build_price_change_pdf(supplier_label, invoice, changed, dest_path):
    from pdf_export import build_simple_table_pdf

    def money(value):
        # Hasta 4 decimales (costos por unidad de centavos), nunca menos de 2.
        if value is None:
            return ""
        text = f"{abs(value):,.4f}".rstrip("0")
        if len(text.split(".")[1]) < 2:
            text = f"{abs(value):,.2f}"
        return f"{'-' if value < 0 else ''}${text}"

    table_rows = []
    for row in changed:
        sign = "+" if row["change"] > 0 else ""
        table_rows.append([
            row["upc"] or "",
            row.get("description") or "",
            _pack_label(row),
            f"{money(row['previous_cost'])} ({_ddmmyyyy(row['previous_date'])})",
            money(row["unit_cost"]),
            f"{sign}{money(row['change'])} ({sign}{row['change_pct']}%)",
            money(row.get("srp")),
            money(row.get("pos_price")),
            "" if row.get("margin_pct") is None else f"{row['margin_pct']}%",
        ])
    up = sum(1 for row in changed if row["state"] == "Subió")
    build_simple_table_pdf(
        dest_path,
        f"{supplier_label} — Cambios de precio — Factura N° {invoice['invoice_no']}",
        ["UPC", "Producto", "Pack", "Costo anterior (fecha)", "Costo nuevo", "Cambio", "SRP factura", "Precio POS", "Margen POS"],
        table_rows,
        col_widths_mm=[30, 82, 22, 40, 24, 36, 22, 22, 22],
        company_header=True,
        period_label=(
            f"Factura del {_ddmmyyyy(invoice['invoice_date'])} — {up} subieron, {len(changed) - up} bajaron"
            + (" — comparado con la compra anterior de cada producto" if changed else " — ningún producto cambió de precio")
        ),
        footer_note="Costo por unidad. Margen POS = (precio del POS − costo nuevo) / precio del POS.",
    )
    return dest_path


def build_price_change_workbook(supplier_label, invoice, changed, dest_path):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Cambios de precio"
    sheet.append([f"{supplier_label} — Factura N° {invoice['invoice_no']} del {_ddmmyyyy(invoice['invoice_date'])}"])
    sheet["A1"].font = Font(bold=True, size=12)
    sheet.append([])
    headers = ["Cambio", "UPC", "Producto", "Categoría", "Pack", "Costo anterior", "Fecha anterior",
               "Costo nuevo", "Diferencia", "Diferencia %", "SRP factura", "Precio POS", "Margen POS %"]
    sheet.append(headers)
    for cell in sheet[3]:
        cell.font = Font(bold=True)
    up_fill = PatternFill("solid", fgColor="FEE2E2")
    down_fill = PatternFill("solid", fgColor="DCFCE7")
    money_format = '"$"#,##0.00##'
    for row in changed:
        sheet.append([
            row["state"], row["upc"], row.get("description"), row.get("category"), _pack_label(row),
            row["previous_cost"], _ddmmyyyy(row["previous_date"]), row["unit_cost"], row["change"],
            row["change_pct"], row.get("srp"), row.get("pos_price"), row.get("margin_pct"),
        ])
        new_row = sheet[sheet.max_row]
        new_row[0].fill = up_fill if row["state"] == "Subió" else down_fill
        for index in (5, 7, 8, 10, 11):
            new_row[index].number_format = money_format
        for cell in new_row[5:]:
            cell.alignment = Alignment(horizontal="center")
    for col_letter, width in zip("ABCDEFGHIJKLM", (9, 15, 38, 16, 10, 14, 14, 12, 12, 12, 12, 12, 13)):
        sheet.column_dimensions[col_letter].width = width
    if changed:
        sheet.auto_filter.ref = f"A3:M{sheet.max_row}"
    sheet.freeze_panes = "A4"
    workbook.save(dest_path)
    return dest_path
