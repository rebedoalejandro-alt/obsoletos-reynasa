#!/usr/bin/env python3
"""Actualiza el panel de obsoletos de Reynasa a partir de dos exportaciones de Sage X3.

Uso:
    python actualizar.py --stock ZCON13.xlsx --ventas ZDL25.xlsx --fecha 2026-11-25 --datos datos

- --stock : peticionario ZCON13 "Stock sin estado Q ni cascos" (todas las plantas), sacado
            al cierre del dia --fecha.
- --ventas: peticionario ZDL25 "Ventas por referencia" del 01.08.2026 al dia --fecha
            (cantidad e importe netos; las devoluciones de cliente restan).
- --datos : carpeta con detalle.json y serie.json actuales. Se reescriben en el sitio.

Criterio (el mismo del Excel de evolucion de obsoletos):
- Perimetro congelado a 31.12.2025: categoria, unidades obsoletas, umbral y PMP por referencia.
- Obsoleto a una fecha = MIN( MAX(stock de la fecha - umbral; 0) ; unidades obsoletas de diciembre ).
- Tramo 1 (dic -> 31.07.2026) queda fijo. Tramo 2 (31.07.2026 -> fecha) se recalcula cada dia:
    salida = obsoleto 31.07 - obsoleto fecha
    vendido = MIN( MAX(salida;0) ; MAX(ventas netas 01.08 -> fecha;0) )
    otras salidas = MAX(salida;0) - vendido
    reentrada = MAX(-salida;0)
- Todo valorado a PMP unitario del obsoleto de diciembre 2025, congelado.
"""
import argparse
import calendar
import datetime as dt
import json
import math
import os
import re
import sys

CORTES_FIJOS = {"2025-12-31", "2026-07-20", "2026-07-31", "2026-10-31", "2026-11-10"}
CORTE_TRAMO = "2026-07-31"
SITUACIONES = [
    "Igual que en diciembre",
    "Parcial",
    "Liquidada por venta",
    "Liquidada (venta + otras salidas)",
    "Liquidada sin venta",
]
EPS = 1e-9


def normalizar(ref):
    """Referencia normalizada: mayusculas, sin guiones, puntos, barras ni espacios."""
    if ref is None:
        return ""
    if isinstance(ref, float) and ref.is_integer():
        ref = int(ref)
    return re.sub(r"[-./\\\s]", "", str(ref)).upper()


def r(x, n=4):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return 0
    v = round(float(x), n)
    return int(v) if v == int(v) else v


def situacion(ud_dic, ud_hoy, vendido, otras):
    if ud_hoy >= ud_dic - EPS:
        return 0
    if ud_hoy <= EPS:
        if otras <= EPS:
            return 2
        if vendido <= EPS:
            return 4
        return 3
    return 1


def es_fin_de_mes(fecha):
    d = dt.date.fromisoformat(fecha)
    return d.day == calendar.monthrange(d.year, d.month)[1]


# ---------------------------------------------------------------- calculo

def calcular(detalle, serie, fecha, stock_ud, ventas_ud, ventas_imp, stock_eur=None,
             plantas=None, fuente=""):
    """Recalcula el tramo 2 y el corte de `fecha` para todas las referencias del perimetro.

    stock_ud / ventas_ud / ventas_imp / stock_eur: dict referencia normalizada -> valor.
    plantas: dict referencia normalizada -> {planta: ud} (opcional).
    """
    campos = detalle["campos"]
    ix = {c: i for i, c in enumerate(campos)}
    cortes = list(detalle["cortes"])
    i_tramo = cortes.index(CORTE_TRAMO)

    # Columna de corte: se conserva el ultimo corte si es fijo o fin de mes; si no, se sustituye.
    sustituir = cortes[-1] == fecha or (cortes[-1] not in CORTES_FIJOS and not es_fin_de_mes(cortes[-1]))
    if sustituir:
        cortes[-1] = fecha
    else:
        cortes.append(fecha)

    lista_plantas = list(detalle.get("plantas") or [])
    if plantas:
        for p in sorted({p for d in plantas.values() for p in d}):
            if p not in lista_plantas:
                lista_plantas.append(p)
    pidx = {p: i for i, p in enumerate(lista_plantas)}

    tot = dict(ud=0.0, eur=0.0, v_ud=0.0, v_eur=0.0, o_ud=0.0, o_eur=0.0, r_ud=0.0, r_eur=0.0)
    for fila in detalle["filas"]:
        ref = normalizar(fila[ix["ref"]])
        pmp = fila[ix["pmp"]]
        umbral = fila[ix["umbral"]]
        u = fila[ix["u"]]
        ud_dic = u[0]
        ud_3107 = u[i_tramo]
        st = float(stock_ud.get(ref, 0.0))
        ud_hoy = min(max(st - umbral, 0.0), ud_dic)
        salida = ud_3107 - ud_hoy
        vs2 = float(ventas_ud.get(ref, 0.0))
        iv2 = float(ventas_imp.get(ref, 0.0))
        v2 = min(max(salida, 0.0), max(vs2, 0.0))
        o2 = max(salida, 0.0) - v2
        r2 = max(-salida, 0.0)
        v1 = fila[ix["v1"]]
        o1 = fila[ix["o1"]]
        if sustituir:
            u[-1] = r(ud_hoy, 3)
        else:
            u.append(r(ud_hoy, 3))
        fila[ix["v2"]] = r(v2, 3)
        fila[ix["o2"]] = r(o2, 3)
        fila[ix["r2"]] = r(r2, 3)
        fila[ix["vs2"]] = r(vs2, 3)
        fila[ix["iv2"]] = r(iv2, 2)
        fila[ix["st"]] = r(st, 3)
        if stock_eur is not None:
            fila[ix["ste"]] = r(stock_eur.get(ref, 0.0), 2)
        fila[ix["sit"]] = situacion(ud_dic, ud_hoy, v1 + v2, o1 + o2)
        if plantas is not None:
            fila[ix["pl"]] = [[pidx[p], r(q, 3)] for p, q in sorted(plantas.get(ref, {}).items()) if abs(q) > EPS]

        tot["ud"] += ud_hoy
        tot["eur"] += ud_hoy * pmp
        tot["v_ud"] += v1 + v2
        tot["v_eur"] += (v1 + v2) * pmp
        tot["o_ud"] += o1 + o2
        tot["o_eur"] += (o1 + o2) * pmp
        tot["r_ud"] += r2
        tot["r_eur"] += r2 * pmp

    detalle["cortes"] = cortes
    detalle["plantas"] = lista_plantas
    detalle["ultimo"] = fecha
    detalle["fuenteUltimo"] = fuente
    try:
        from zoneinfo import ZoneInfo
        ahora = dt.datetime.now(ZoneInfo("Europe/Madrid"))
    except Exception:
        ahora = dt.datetime.now()
    detalle["generado"] = ahora.replace(microsecond=0, tzinfo=None).isoformat()

    punto = {
        "fecha": fecha,
        "ud": r(tot["ud"], 2), "eur": r(tot["eur"], 2),
        "vendidoUd": r(tot["v_ud"], 2), "vendidoEur": r(tot["v_eur"], 2),
        "otrasUd": r(tot["o_ud"], 2), "otrasEur": r(tot["o_eur"], 2),
        "reentUd": r(tot["r_ud"], 2), "reentEur": r(tot["r_eur"], 2),
        "fuente": fuente,
    }
    serie = [p for p in serie if p["fecha"] != fecha] + [punto]
    serie.sort(key=lambda p: p["fecha"])
    return detalle, serie, punto


# ---------------------------------------------------------------- lectura de Sage

def _filas_excel(ruta):
    """Devuelve las filas de la primera hoja (xlsx o csv) como listas de valores."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".csv", ".txt"):
        import csv
        with open(ruta, "r", encoding="utf-8-sig", errors="replace") as f:
            muestra = f.read(4096)
            f.seek(0)
            sep = ";" if muestra.count(";") > muestra.count(",") else ","
            if muestra.count("\t") > max(muestra.count(";"), muestra.count(",")):
                sep = "\t"
            return [row for row in csv.reader(f, delimiter=sep)]
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("€", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _sin_tildes(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").lower().strip()


def _buscar_cabecera(filas, columnas):
    """Localiza la fila de cabecera y el indice de cada columna pedida.

    columnas: dict nombre_logico -> lista de textos posibles (sin tildes, minusculas).
    Devuelve (indice_fila_cabecera, {nombre_logico: indice_columna}).
    """
    for i, fila in enumerate(filas[:60]):
        textos = [_sin_tildes(c) if c is not None else "" for c in fila]
        encontrado = {}
        for logico, opciones in columnas.items():
            for j, t in enumerate(textos):
                if t and any(t == o or t.startswith(o) for o in opciones):
                    encontrado[logico] = j
                    break
        obligatorias = [k for k in columnas if not k.endswith("?")]
        if all(k in encontrado for k in obligatorias):
            return i, encontrado
    raise SystemExit("No encuentro la cabecera esperada en el fichero. Columnas buscadas: %s" % columnas)


def leer_stock(ruta, mapa=None):
    """ZCON13: suma por referencia normalizada (todas las plantas y ubicaciones)."""
    filas = _filas_excel(ruta)
    columnas = mapa or {
        "ref": ["articulo", "referencia", "codigo articulo", "cod. articulo"],
        "ud": ["stock", "cantidad", "unidades", "ud", "existencias"],
        "eur?": ["valor", "importe", "valoracion", "total pmp", "valor pmp"],
        "planta?": ["planta", "almacen"],
    }
    h, col = _buscar_cabecera(filas, columnas)
    stock, eur, plantas = {}, {}, {}
    lineas = 0
    for fila in filas[h + 1:]:
        if fila is None or len(fila) <= col["ref"]:
            continue
        ref = normalizar(fila[col["ref"]])
        if not ref:
            continue
        lineas += 1
        q = _num(fila[col["ud"]])
        stock[ref] = stock.get(ref, 0.0) + q
        if "eur?" in col:
            eur[ref] = eur.get(ref, 0.0) + _num(fila[col["eur?"]])
        if "planta?" in col:
            p = str(fila[col["planta?"]] or "").strip()
            d = plantas.setdefault(ref, {})
            d[p] = d.get(p, 0.0) + q
    control = {
        "lineas": lineas,
        "ud": round(sum(stock.values()), 2),
        "eur": round(sum(eur.values()), 2) if eur else None,
        "plantas": len({p for d in plantas.values() for p in d}) if plantas else None,
        "columnas": {k: (filas[h][v] if v < len(filas[h]) else None) for k, v in col.items()},
    }
    return stock, (eur or None), (plantas or None), control


def leer_ventas(ruta, mapa=None):
    """ZDL25: cantidad e importe netos por referencia normalizada."""
    filas = _filas_excel(ruta)
    columnas = mapa or {
        "ref": ["articulo", "referencia", "codigo articulo", "cod. articulo"],
        "ud": ["cantidad", "unidades", "ud", "cant."],
        "imp": ["importe", "base imponible", "venta", "total"],
    }
    h, col = _buscar_cabecera(filas, columnas)
    ud, imp = {}, {}
    lineas = 0
    for fila in filas[h + 1:]:
        if fila is None or len(fila) <= col["ref"]:
            continue
        ref = normalizar(fila[col["ref"]])
        if not ref or _sin_tildes(ref).startswith("total"):
            continue
        lineas += 1
        ud[ref] = ud.get(ref, 0.0) + _num(fila[col["ud"]])
        imp[ref] = imp.get(ref, 0.0) + _num(fila[col["imp"]])
    control = {
        "lineas": lineas,
        "ud": round(sum(ud.values()), 2),
        "imp": round(sum(imp.values()), 2),
        "columnas": {k: (filas[h][v] if v < len(filas[h]) else None) for k, v in col.items()},
    }
    return ud, imp, control


# ---------------------------------------------------------------- E/S

def cargar(carpeta):
    with open(os.path.join(carpeta, "detalle.json"), encoding="utf-8") as f:
        detalle = json.load(f)
    with open(os.path.join(carpeta, "serie.json"), encoding="utf-8") as f:
        serie = json.load(f)
    return detalle, serie


def guardar(carpeta, detalle, serie):
    with open(os.path.join(carpeta, "detalle.json"), "w", encoding="utf-8") as f:
        json.dump(detalle, f, ensure_ascii=False, separators=(",", ":"))
    with open(os.path.join(carpeta, "serie.json"), "w", encoding="utf-8") as f:
        json.dump(serie, f, ensure_ascii=False, indent=1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--stock", required=True)
    ap.add_argument("--ventas", required=True)
    ap.add_argument("--fecha", required=True, help="Fecha de corte AAAA-MM-DD (cierre del dia)")
    ap.add_argument("--datos", default="datos")
    ap.add_argument("--simular", action="store_true", help="Calcula y muestra el resultado sin escribir")
    a = ap.parse_args()

    dt.date.fromisoformat(a.fecha)
    detalle, serie = cargar(a.datos)
    if a.fecha < CORTE_TRAMO:
        raise SystemExit("La fecha debe ser posterior al 31.07.2026")
    stock, stock_eur, plantas, c_stock = leer_stock(a.stock)
    ventas, importes, c_ventas = leer_ventas(a.ventas)
    print("ZCON13:", json.dumps(c_stock, ensure_ascii=False, default=str))
    print("ZDL25:", json.dumps(c_ventas, ensure_ascii=False, default=str))
    if c_stock["lineas"] < 100000:
        raise SystemExit("ZCON13 con muy pocas lineas (%d): exportacion incompleta, no publico." % c_stock["lineas"])

    d = dt.date.fromisoformat(a.fecha)
    fuente = "Sage X3 · ZCON13 al cierre del %s y ZDL25 del 01.08.2026 al %s" % (
        d.strftime("%d.%m.%Y"), d.strftime("%d.%m.%Y"))
    detalle, serie, punto = calcular(detalle, serie, a.fecha, stock, ventas, importes,
                                     stock_eur=stock_eur, plantas=plantas, fuente=fuente)
    print("Corte", a.fecha, json.dumps(punto, ensure_ascii=False))
    prev = [p for p in serie if p["fecha"] < a.fecha]
    if prev and abs(punto["eur"] - prev[-1]["eur"]) > 0.15 * prev[-1]["eur"]:
        print("AVISO: el obsoleto cambia mas de un 15 %% frente al corte anterior (%s: %.0f -> %.0f). Revisar antes de publicar."
              % (prev[-1]["fecha"], prev[-1]["eur"], punto["eur"]))
        sys.exit(2)
    if not a.simular:
        guardar(a.datos, detalle, serie)
        print("Escrito en", a.datos)


if __name__ == "__main__":
    main()
