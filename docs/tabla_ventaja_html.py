"""Escribe en docs/aduana.html la tabla de importaciones con acuerdo y con ventaja arancelaria potencial.

Uso (desde la raíz del repositorio), después de python/importaciones_ventaja.py:
    python docs/tabla_ventaja_html.py
"""
import re
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
HTML = RAIZ / "docs" / "aduana.html"


def coma(v, d=1):
    return f"{v:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def main():
    d = pd.read_csv(RAIZ / "docs" / "datos" / "importaciones_acuerdo_ventaja.csv", encoding="utf-8-sig")
    tope = 40.0
    filas = []
    for _, r in d.iterrows():
        anio = f"{int(r['AÑO'])}*" if r["MESES"] < 12 else str(int(r["AÑO"]))
        a, v = r["PCT_CON_ACUERDO"], r["PCT_CON_VENTAJA_POTENCIAL"]
        sin = (r["CON_ACUERDO_USD"] - r["CON_ACUERDO_CON_VENTAJA_POTENCIAL_USD"]) / 1e6
        filas.append(
            f'<tr><td>{anio}</td><td class="anid"><span class="nest">'
            f'<i class="n1" style="width:{a / tope * 100:.1f}%"></i><i class="n3" style="width:{v / tope * 100:.1f}%"></i>'
            f'</span></td><td>{coma(a)}%</td><td>{coma(v)}%</td><td>{coma(sin, 0)}</td></tr>')
    tabla = ('<thead><tr><th>Año</th><th></th><th>Usó un acuerdo</th><th>Con ventaja potencial</th>'
             '<th>Sin ventaja (mill. USD)</th></tr></thead>\n        <tbody>' + "".join(filas) + "</tbody>")
    s = HTML.read_text(encoding="utf-8")
    s, n = re.subn(r'<thead><tr><th>Año</th><th>(?:Importaciones con acuerdo|</th><th>Usó un acuerdo).*?</tbody>', tabla, s, count=1, flags=re.S)
    assert n == 1
    HTML.write_text(s, encoding="utf-8")


if __name__ == "__main__":
    main()
