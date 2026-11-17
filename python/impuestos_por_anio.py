"""
Impuestos de los despachos de la DNIT, por año, en Excel (una fila por tipo de impuesto).

Para cada año baja los archivos mensuales "Nivel Item", aplica los mismos filtros que el resto del
proyecto (solo regímenes definitivos; exportaciones de origen paraguayo) y suma por:
    AÑO, OPERACION, NCM, DESCRIPCION, PAIS, TIPO_ACUERDO, ACUERDO, TIPO_IMPUESTO
los montos en guaraníes (como los publica la DNIT) y en dólares (divididos por la cotización de cada
despacho). Una segunda hoja trae el valor, la base imponible y el peso, sin repetirlos por impuesto.

Uso: python python/impuestos_por_anio.py --anios 2011-2026
Salida: salidas/impuestos/DNIT_impuestos_<año>.xlsx (con tabla dinámica, Calibri 10)
"""
import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import procesar_despachos as pdesp  # noqa: E402
from excel_descargas import escribir, etiquetas_en_espanol  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "salidas" / "impuestos"
NOMBRES = {"DERECHO": "Derecho aduanero (arancel)", "IVA": "IVA", "ISC": "Impuesto selectivo al consumo (ISC)",
           "SERVICIO": "Servicio de valoración", "RENTA": "Anticipo de renta", "OTROS": "Otros tributos"}
CLAVES = ["AÑO_INT", "OPERACION", "POSICION_NCM", "DESCRIPCION_PRODUCTO", "PAIS_PROC_DEST",
          "TIPO_ACUERDO", "ACUERDO_CONSOLIDADO"]


def anio(a, regimenes_si, acuerdos):
    temporal = RAIZ / "salidas" / "tmp_dnit" / str(a)
    subprocess.run([sys.executable, str(RAIZ / "python" / "descargar_dnit.py"), "--anios", str(a),
                    "--destino", str(temporal), "--fuente", "github"], check=True)
    partes_imp, partes_base, meses = [], [], 0
    for ruta in sorted(p for p in temporal.rglob("*") if p.suffix.lower() in (".csv", ".json", ".txt")):
        df = pdesp.estandarizar_columnas(pdesp.leer_archivo(ruta))
        for c in pdesp.TRIBUTOS_GS:
            df[c + "_GS"] = df[c]
        df = pdesp.transformar(df, regimenes_si, acuerdos)
        df = df[df["MOTIVO_EXCLUSION"] == ""]
        meses += 1
        for c in CLAVES:
            df[c] = df[c].fillna("")
        g = df.groupby(CLAVES, dropna=False)
        partes_imp.append(g[[c for t in pdesp.TRIBUTOS_GS for c in (t, t + "_GS")]].sum())
        partes_base.append(g[["VALOR_DOLAR", "IMPONIBLE_DOLAR", "PESO_NETO_KG"]].sum())
    shutil.rmtree(temporal, ignore_errors=True)
    if not partes_imp:
        return None
    imp = pd.concat(partes_imp).groupby(level=list(range(len(CLAVES)))).sum().reset_index()
    base = pd.concat(partes_base).groupby(level=list(range(len(CLAVES)))).sum().reset_index()
    largo = []
    for t in pdesp.TRIBUTOS_GS:
        x = imp[CLAVES + [t + "_GS", t]].rename(columns={t + "_GS": "MONTO_GS", t: "MONTO_USD"})
        x = x[(x["MONTO_GS"] != 0) | (x["MONTO_USD"] != 0)]
        x.insert(len(CLAVES), "TIPO_IMPUESTO", NOMBRES[t])
        largo.append(x)
    largo = pd.concat(largo, ignore_index=True)
    ren = {"AÑO_INT": "AÑO", "POSICION_NCM": "NCM", "DESCRIPCION_PRODUCTO": "DESCRIPCION",
           "PAIS_PROC_DEST": "PAIS", "ACUERDO_CONSOLIDADO": "ACUERDO"}
    largo, base = largo.rename(columns=ren), base.rename(columns=ren)
    largo[["MONTO_GS", "MONTO_USD"]] = largo[["MONTO_GS", "MONTO_USD"]].round(2)
    # control: el total por tipo de impuesto se conserva al pasar a formato largo
    for t in pdesp.TRIBUTOS_GS:
        assert abs(largo.loc[largo["TIPO_IMPUESTO"] == NOMBRES[t], "MONTO_GS"].sum() - imp[t + "_GS"].sum()) < 1
    return largo, base, meses


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anios", required=True)
    args = ap.parse_args()
    anios = pdesp.leer_rango_anios(args.anios)
    cfg = pdesp.cargar_config()
    regimenes_si, acuerdos = cfg[0], cfg[2]
    SALIDA.mkdir(parents=True, exist_ok=True)
    lista = []
    for a in anios:
        r = anio(a, regimenes_si, acuerdos)
        if r is None:
            print(a, "sin archivos", flush=True)
            continue
        largo, base, meses = r
        destino = SALIDA / f"DNIT_impuestos_{a}.xlsx"
        notas = [f"Impuestos de los despachos de la DNIT, {a}, por tipo de impuesto",
                 f"Meses incluidos: {meses}." + (" Año incompleto." if meses < 12 else ""),
                 "Fuente: Dirección Nacional de Ingresos Tributarios (DNIT), datos abiertos de aduana (nivel ítem). Elaboración propia.",
                 "Solo regímenes definitivos; exportaciones solo de origen paraguayo (mismos criterios que el resto del proyecto).",
                 "Una fila por año, operación, NCM (8 dígitos), país, acuerdo y TIPO_IMPUESTO. Filas en cero omitidas.",
                 "MONTO_GS: monto en guaraníes tal como lo publica la DNIT. MONTO_USD: dividido por la cotización de cada despacho.",
                 "PAIS: procedencia (importación) o destino (exportación).",
                 "Hoja 'Base': valor (exportación FOB, importación CIF), base imponible y peso, sin repetir por impuesto."]
        escribir(destino, largo, notas)
        with pd.ExcelWriter(destino, engine="openpyxl", mode="a") as w:
            base.to_excel(w, sheet_name="Base", index=False)
        lista.append({"archivo": str(destino), "hoja_datos": "Datos", "filas": ["TIPO_IMPUESTO"],
                      "columnas": ["OPERACION"], "filtros": ["PAIS", "ACUERDO"],
                      "valores": [["MONTO_GS", "suma", "Monto (Gs.)"]]})
        print(f"{destino.name}: {len(largo):,} filas; {meses} meses", flush=True)
    cfgj = SALIDA / "_tablas.json"
    cfgj.write_text(json.dumps(lista, ensure_ascii=False), encoding="utf-8")
    subprocess.run(["/usr/bin/python3", str(RAIZ / "python" / "tabla_dinamica.py"), str(cfgj)], check=True)
    cfgj.unlink()
    for t in lista:
        etiquetas_en_espanol(t["archivo"])


if __name__ == "__main__":
    main()
