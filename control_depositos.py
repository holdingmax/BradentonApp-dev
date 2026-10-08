"""
Control Depósitos (pedido del usuario, 2026-10-07): todos los depósitos del
mes contra Chase, de una vez. Cálculo puro, sin UI ni base.

- Recibos de depósito (depositos_db, los PDF "Transaccion #…" del cajero):
  normales, Ice Machine, Food Truck, Vaccumms (y cualquier otra aclaración
  del nombre del archivo, como "Fruit Stand"). Cada recibo es un depósito de
  Chase con la misma fecha e importe (validado contra septiembre 2026).
- Payment Summary de Cantaloupe (ice_machine_db): la venta con tarjeta de la
  máquina de hielo, que Chase acredita como "Cantaloupe ... IND ID:<Reference>"
  el día del "To" del resumen (ver ice_machine.py).

Al revés, todo pago de Cantaloupe de Chase del mes tiene que tener su
Payment Summary; un depósito de Chase sin recibo solo se avisa ("notices"). Lo posterior al último día cargado de Chase queda "todavía
no se puede controlar".
"""

from ice_machine import chase_cantaloupe

TOLERANCE = 0.005

NORMAL = "Depósito"
ICE_MACHINE = "Ice Machine"
FOOD_TRUCK = "Food Truck"
VACCUMMS = "Vaccumms"
KNOWN_KINDS = (NORMAL, ICE_MACHINE, FOOD_TRUCK, VACCUMMS)
GROUP_LABELS = {NORMAL: "Depósitos normales", ICE_MACHINE: "Ice Machine — efectivo",
                FOOD_TRUCK: FOOD_TRUCK, VACCUMMS: VACCUMMS}
# Cómo queda categorizado en Chase cada tipo (detalle).
CHASE_DETALLE = {NORMAL: "DEPOSITO", ICE_MACHINE: "DEPOSITO VENTA ICE", FOOD_TRUCK: "FOOD TRUCK",
                 VACCUMMS: "DEPOSITO VACCUMMS"}


def receipt_kind(kind):
    """El grupo de un recibo según su aclaración (None = depósito normal)."""
    import re
    lowered = (kind or "").lower()
    if not lowered.strip():
        return NORMAL
    if re.search(r"\b(ice|hielo)\b", lowered):
        return ICE_MACHINE
    if re.search(r"\b(food|truck)\b", lowered):
        return FOOD_TRUCK
    if re.search(r"\bvac", lowered):
        return VACCUMMS
    return kind.strip()


def chase_kind(detalle):
    """El grupo de un depósito de Chase según su categoría (sin categoría = normal)."""
    upper = (detalle or "").strip().upper()
    if upper == CHASE_DETALLE[ICE_MACHINE]:
        return ICE_MACHINE
    if upper == CHASE_DETALLE[FOOD_TRUCK]:
        return FOOD_TRUCK
    if upper == CHASE_DETALLE[VACCUMMS] or upper.startswith("VAC"):
        return VACCUMMS
    if upper in ("", CHASE_DETALLE[NORMAL]):
        return NORMAL
    return upper.title()  # "FRUIT STAND" -> "Fruit Stand"


def chase_deposits(rows):
    """Depósitos de Chase (descripción DEPOSIT, importe positivo)."""
    return [{"date": r["posting_date"], "amount": r["amount"], "detalle": r.get("detalle"),
             "description": r.get("description")}
            for r in rows if (r.get("amount") or 0) > 0 and (r.get("description") or "").upper().startswith("DEPOSIT")]


def _dm(iso):
    return f"{iso[8:10]}/{iso[5:7]}"


def build_month_control(year, month, summaries, deposits, chase_rows, chase_last_date, known_references=()):
    """
    summaries: los Payment Summary pagados en el mes (fecha "To"); deposits:
    todos los recibos de Depósitos del mes; chase_rows: los movimientos de
    Chase del mes (y los primeros días del siguiente, para los pagos de fin de
    mes); chase_last_date: último día cargado de Chase (ISO) o None;
    known_references: los Reference de los resúmenes del mes anterior (un
    pago que entra los primeros días del mes ya tiene su resumen allá).
    """
    prefix = f"{year:04d}-{month:02d}-"

    def covered(iso):
        return chase_last_date is not None and iso <= chase_last_date

    issues = []

    # Payment Summary <-> pago de Cantaloupe, por Reference (= IND ID) e importe.
    payments = chase_cantaloupe(chase_rows)
    used = set()
    summary_rows = []
    for s in sorted(summaries, key=lambda s: s["to_date"]):
        match = next((i for i, p in enumerate(payments) if i not in used and p["reference"] == s["reference"]), None)
        if match is None:  # pagos viejos sin IND ID: misma fecha e importe
            match = next((i for i, p in enumerate(payments) if i not in used and p["reference"] is None
                          and p["date"] == s["to_date"] and abs(p["amount"] - s["net"]) <= TOLERANCE), None)
        row = dict(s, chase=None, status=None)
        if match is not None:
            used.add(match)
            row["chase"] = payments[match]
            row["status"] = "ok" if abs(payments[match]["amount"] - s["net"]) <= TOLERANCE else "diff"
            if row["status"] == "diff":
                issues.append(f"Payment Summary #{s['summary_no']}: Chase acreditó ${payments[match]['amount']:,.2f} "
                              f"y el resumen dice ${s['net']:,.2f}.")
        elif covered(s["to_date"]):
            row["status"] = "missing"
            issues.append(f"Payment Summary #{s['summary_no']} (${s['net']:,.2f} del {_dm(s['to_date'])}) no está en Chase.")
        else:
            row["status"] = "after"
        summary_rows.append(row)
    known = set(known_references)
    orphan_payments = [p for i, p in enumerate(payments)
                       if i not in used and p["date"].startswith(prefix) and p["reference"] not in known]
    for p in orphan_payments:
        issues.append(f"Chase: pago de Cantaloupe de ${p['amount']:,.2f} del {_dm(p['date'])} sin su Payment Summary.")

    # Recibo <-> depósito de Chase, misma fecha e importe; si hay varios
    # iguales, primero el que ya está categorizado como el tipo del recibo.
    chase_deps = chase_deposits(chase_rows)
    taken = set()
    deposit_rows = []
    for d in sorted(deposits, key=lambda d: (d["deposit_date"] or "9999", d.get("tx_number") or 0)):
        kind = receipt_kind(d.get("kind"))
        row = dict(d, group=kind, chase=None, status=None, uncategorized=False)
        if d["deposit_date"] is None or d["amount"] is None:
            row["status"] = "incomplete"
            issues.append(f"Recibo #{d.get('tx_number') or '?'}: falta la fecha o el importe (corregilo en Carga de Datos → Depósitos).")
            deposit_rows.append(row)
            continue
        candidates = [i for i, c in enumerate(chase_deps) if i not in taken and c["date"] == d["deposit_date"]
                      and abs(c["amount"] - d["amount"]) <= TOLERANCE]
        candidates.sort(key=lambda i: chase_kind(chase_deps[i]["detalle"]) != kind)
        if candidates:
            taken.add(candidates[0])
            row["chase"] = chase_deps[candidates[0]]
            row["status"] = "ok"
            row["uncategorized"] = chase_kind(row["chase"]["detalle"]) != kind
        elif covered(d["deposit_date"]):
            row["status"] = "missing"
            label = "Depósito" if kind == NORMAL else kind
            issues.append(f"{label} de ${d['amount']:,.2f} del {_dm(d['deposit_date'])} "
                          f"(recibo #{d.get('tx_number') or '?'}) no está en Chase.")
        else:
            row["status"] = "after"
        deposit_rows.append(row)
    orphan_deposits = [dict(c, group=chase_kind(c["detalle"])) for i, c in enumerate(chase_deps)
                       if i not in taken and c["date"].startswith(prefix)]

    # Un depósito de Chase sin recibo es un aviso, no algo mal (pedido del
    # usuario, 2026-10-08: suele ser un recibo que el manager no mandó).
    kinds = list(KNOWN_KINDS) + sorted({r["group"] for r in deposit_rows + orphan_deposits} - set(KNOWN_KINDS))
    notices = []
    for kind in kinds:
        for o in (o for o in orphan_deposits if o["group"] == kind):
            label = "Depósito" if kind == NORMAL else GROUP_LABELS.get(kind, kind)
            notices.append(f"{label} de ${o['amount']:,.2f} del {_dm(o['date'])}")

    def total(items, key):
        return round(sum(i[key] or 0 for i in items), 2)

    groups = {kind: [d for d in deposit_rows if d["group"] == kind] for kind in kinds}
    ice = {
        "transactions": sum(s["transactions"] for s in summary_rows),
        "gross": total(summary_rows, "gross"),
        "process_fees": total(summary_rows, "process_fees"),
        "service_fees": total(summary_rows, "service_fees"),
        "net": total(summary_rows, "net"),
        "cash": total(groups[ICE_MACHINE], "amount"),
    }
    ice["total"] = round(ice["net"] + ice["cash"], 2)

    # Todo junto: por tipo, lo que dicen los papeles contra lo que entró a
    # Chase por esos papeles; lo de Chase sin papel va aparte ("Sin papel"):
    # en depósitos es un aviso, en Cantaloupe (sin Payment Summary) un error.
    # Lo posterior al último día de Chase queda afuera.
    def overview_row(key, label, papers, amount_key, orphans, orphans_are_errors):
        checked = [p for p in papers if p["status"] != "after"]
        paper_total = total(checked, amount_key)
        chase_total = round(sum(p["chase"]["amount"] for p in checked if p["chase"]), 2)
        diff = round(chase_total - paper_total, 2)
        return {"key": key, "label": label, "count": len(checked), "papers": paper_total, "chase": chase_total,
                "diff": diff, "orphans": len(orphans), "orphans_total": total(orphans, "amount"),
                "ok": abs(diff) <= TOLERANCE and not (orphans and orphans_are_errors)
                      and all(p["status"] == "ok" for p in checked)}

    overview = [overview_row(kind, GROUP_LABELS.get(kind, kind), groups[kind], "amount",
                             [o for o in orphan_deposits if o["group"] == kind], False) for kind in kinds]
    overview.insert(1, overview_row("cantaloupe", "Ice Machine — tarjeta (Cantaloupe)", summary_rows, "net",
                                    orphan_payments, True))
    overview = [r for r in overview if r["count"] or r["chase"] or r["orphans"]
                or r["key"] in KNOWN_KINDS or r["key"] == "cantaloupe"]
    overview_total = {
        "label": "Total", "count": sum(r["count"] for r in overview),
        "papers": round(sum(r["papers"] for r in overview), 2), "chase": round(sum(r["chase"] for r in overview), 2),
        "orphans": sum(r["orphans"] for r in overview),
        "orphans_total": round(sum(r["orphans_total"] for r in overview), 2),
    }
    overview_total["diff"] = round(overview_total["chase"] - overview_total["papers"], 2)
    overview_total["ok"] = all(r["ok"] for r in overview)
    return {
        "overview": overview,
        "overview_total": overview_total,
        "summaries": summary_rows,
        "kinds": kinds,
        "group_labels": {k: GROUP_LABELS.get(k, k) for k in kinds},
        "groups": {kind: {"rows": rows, "total": total(rows, "amount"),
                          "orphans": [o for o in orphan_deposits if o["group"] == kind]}
                   for kind, rows in groups.items()},
        "ice": ice,
        "orphan_payments": orphan_payments,
        "uncategorized": [d for d in deposit_rows if d["uncategorized"]],
        "issues": issues,
        "notices": notices,
        "chase_last_date": chase_last_date,
        "pending_after": [r for r in summary_rows + deposit_rows if r["status"] == "after"],
        "ok": not issues,
    }
