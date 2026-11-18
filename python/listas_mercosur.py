"""
Arma config/listas_py.csv con las listas de Paraguay notificadas al MERCOSUR (POLCOM):
Bienes de Capital (BK), Informática y Telecomunicaciones (BIT), Lista Nacional de Excepciones (LNE)
y Régimen de Materias Primas (RMP). Una fila por lista, fecha de notificación y código NCM (8 dígitos).
Uso: python python/listas_mercosur.py <carpeta con los xlsx descomprimidos>
"""
import re
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
PATRON = re.compile(r"^\s*(\d{4})\.?(\d{2})\.?(\d{2})")


def lista_de(hoja, titulo):
    t = (hoja + " " + titulo).upper()
    if "EXCEPCION" in t or re.search(r"\bLNE\b", t): return "LNE"
    if "INFORM" in t or re.search(r"\bBIT\b", t): return "BIT"
    if "CAPITAL" in t or re.search(r"\bBK\b", t): return "BK"
    if "MATERIA" in t or "RMP" in t: return "RMP"
    return None


def fecha_de(texto, nombre):
    m = re.search(r"(\d{1,2})[/.\-_](\d{1,2})[/.\-_](\d{2,4})", texto) or re.search(r"(\d{1,2})[/.\-_](\d{1,2})[/.\-_](\d{2,4})", nombre)
    if m:
        d, mo, a = m.groups(); a = int(a) + (2000 if len(a) == 2 else 0)
        return f"{a:04d}-{int(mo):02d}-{int(d):02d}"
    m = re.search(r"(20\d{2})", texto + " " + nombre)
    return f"{m.group(1)}-01-01" if m else None


filas = []
for f in sorted(Path(sys.argv[1]).rglob("*.xls*")):
    if "uruguay" in f.name.lower():
        continue
    try:
        hojas = pd.read_excel(f, sheet_name=None, header=None, dtype=str)
    except Exception as e:  # noqa: BLE001
        print("no se pudo leer", f.name, e); continue
    for hoja, df in hojas.items():
        cab = " ".join(df.head(6).fillna("").astype(str).values.ravel())
        if "URUGUAY" in cab.upper():
            continue
        lista = lista_de(hoja, cab) or lista_de("", f.name)
        fecha = fecha_de(cab, f.name)
        if not lista or not fecha:
            continue
        ext = df[0].dropna().astype(str).str.extract(PATRON).dropna()
        cods = (ext[0] + ext[1] + ext[2]).unique() if len(ext) else []
        for c in cods:
            filas.append((lista, fecha, c, f.name.encode("utf-8", "surrogateescape").decode("cp850", "replace")))
out = pd.DataFrame(filas, columns=["LISTA", "FECHA_NOTIFICACION", "NCM", "ARCHIVO"]).drop_duplicates(["LISTA", "FECHA_NOTIFICACION", "NCM"])
out.sort_values(["LISTA", "FECHA_NOTIFICACION", "NCM"]).to_csv(RAIZ / "config" / "listas_py.csv", index=False)
print(out.groupby(["LISTA", "FECHA_NOTIFICACION"]).size().to_string())
