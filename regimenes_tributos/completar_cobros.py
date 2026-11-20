"""Completa la lista de cobros y regímenes especiales con lo que se ve en los datos de la DNIT.

Uso (desde la raíz del repositorio):
    python regimenes_tributos/completar_cobros.py

Lee regimenes_tributos/cobros_original.xlsx (la lista armada a partir de las normas, con la tasa y el artículo de
cada cobro) y salidas/dnit_anual/DNIT_2025_importacion.xlsx. Escribe regimenes_tributos/Cobros_y_regimenes_especiales.xlsx:
- celdas en verde: tasa leída en la norma (como en el original);
- celdas en celeste: tasa observada en lo que cobró la DNIT en 2025, sin norma leída todavía;
- hoja Control_DNIT_2025: la tasa efectiva que se ve en los datos para cada cobro.
"""
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

RAIZ = Path(__file__).resolve().parent.parent
CARPETA = RAIZ / "regimenes_tributos"
CELESTE = PatternFill("solid", fgColor="DCEBFA")
VERDE = PatternFill("solid", fgColor="B9EEE1")


def control(anio=2025):
    d = pd.read_excel(RAIZ / "salidas" / "dnit_anual" / f"DNIT_{anio}_importacion.xlsx", sheet_name="Datos", engine="calamine",
                      usecols=["SE_CUENTA", "IMPONIBLE GS", "IMPONIBLE DOLAR", "IVA", "RENTA", "SERVICIO", "OTROS",
                               "REGIMEN_TRIBUTARIO"])
    d = d[(d["SE_CUENTA"] == "SI") & (d["IMPONIBLE GS"] > 0)]
    w = d["IMPONIBLE DOLAR"]
    rtc = d["REGIMEN_TRIBUTARIO"].astype(str).str.startswith("Turismo de Compras")

    def pct(col, m=None):
        m = (d[col] > 0) if m is None else (m & (d[col] > 0))
        r = d.loc[m, col] / d.loc[m, "IMPONIBLE GS"] * 100
        return round(r.median(), 2), round(100 * w[m].sum() / w.sum(), 1)

    filas = []
    for nombre, col, m, norma in [
        ("Servicio de valoración", "SERVICIO", None, "Norma de creación no encontrada todavía"),
        ("Otros (en la importación)", "OTROS", None, "Puede incluir el canon informático y el arancel consular; sin desglose en los datos"),
        ("IVA, tasa general y reducida", "IVA", ~rtc & (d["IVA"] / d["IMPONIBLE GS"] > 0.03), "Ley 6380/2019, art. 90 (10 % y 5 %)"),
        ("IVA de Turismo de Compras", "IVA", rtc, "Decreto 2063/2024, art. 7 (10 % sobre el 12,5 % de la base = 1,25 %)"),
        ("Anticipo del IRE, fuera de Turismo de Compras", "RENTA", ~rtc, "Decreto 3182/2019, anexo, art. 87 (0,4 %)"),
        ("Anticipo del IRE, Turismo de Compras", "RENTA", rtc, "Decreto 2063/2024, art. 7 (1 %)"),
    ]:
        med, part = pct(col, m)
        filas.append([nombre, col, med, part, norma])
    return filas


def main():
    wb = openpyxl.load_workbook(CARPETA / "cobros_original.xlsx")
    ws = wb["Lista"]
    fila = {ws.cell(r, 3).value: r for r in range(2, ws.max_row + 1)}

    def poner(nombre, tasa=None, base=None, articulo=None, nota=None, relleno=CELESTE):
        r = fila[nombre]
        for col, v in zip((12, 13, 14, 15), (tasa, base, articulo, nota)):
            if v is not None:
                ws.cell(r, col).value = v
        if tasa is not None:
            ws.cell(r, 12).fill = relleno

    poner("Tasa por servicio de valoración aduanera", "0,5 %", "Valor en aduana (base gravable de las mercaderías importadas)",
          "Ley 2422/2004 (Código Aduanero), arts. 263 y 393 a)",
          "El servicio lo creó el Decreto-Ley 5/1974, aprobado por la Ley 489/1974. En 2025 el 99 % del valor importado "
          "pagó 0,5 % (hoja Control_DNIT_2025).", relleno=VERDE)
    ws.cell(fila["Tasa por servicio de valoración aduanera"], 7).value = "Ley 2422/2004, Código Aduanero"
    poner("Importación de vehículos eléctricos (2012)", "0 % de arancel y 0 % de IVA", "Vehículos eléctricos e híbridos importados",
          "Ley 4601/2012, art. 1", "Modificada por la Ley 5183/2014 (carpeta normas). Texto de la BACN.", relleno=VERDE)
    poner("Donaciones al Estado y a otras instituciones", "Exentas de todo tributo", "Bienes donados",
          "Ley 302/1993, art. 2", "No se pueden vender ni transferir por 5 años (art. 5). Texto de la BACN (escaneado, leído con OCR).",
          relleno=VERDE)
    poner("Canon informático", nota="La DNIT no lo muestra aparte: va dentro de OTROS, que en 2025 fue cerca de 0,4 % "
          "del valor en aduana (hoja Control_DNIT_2025). El monto lo fija la Resolución MH 1144/1992, que falta.")
    poner("Impuesto selectivo al consumo (ISC) en la importación",
          nota="OJO: son las tasas MÁXIMAS de la ley. La tasa real de cada producto la fija el Decreto 3109/2019 "
               "y sus cambios (copia escaneada, falta leerla). Detalle: hoja Tasas_ISC.")
    poner("Maquila: nuevo régimen",
          nota="El tributo único se paga a la DNIT como impuesto interno (sobre el valor agregado en el país), "
               "no en el despacho aduanero: no aparece en los datos de comercio.")
    poner("Turismo de compras (RTC)", nota="Confirmado en los datos 2025: IVA de cerca de 1,3 % y anticipo del IRE "
          "de cerca de 1 % de la base en guaraníes (hoja Control_DNIT_2025).")
    poner("Anticipo del IRE en la importación", nota="En los datos se ve cerca de 0,4 % a 0,5 % (la base de la DNIT "
          "puede incluir otros tributos). Solo lo pagan los contribuyentes del IRE general.")

    poner("Maquila: Ley 1064/1997 (régimen anterior)", "Tributo único de 1 %", "Valor agregado en territorio nacional",
          "Ley 1064/1997, art. 29", "Igual tasa que la Ley 7547/2025. Se paga como impuesto interno, no en el despacho.",
          relleno=VERDE)
    poner("Franquicias diplomáticas y consulares", "Exentos del gravamen aduanero y de los tributos internos a la importación",
          "Bienes importados por misiones, organismos internacionales y funcionarios que regresan", "Ley 110/1992, art. 3",
          "Bajo condición de reciprocidad (art. 2). No se eximen las tasas por servicios prestados.", relleno=VERDE)
    poner("Zonas francas", "Impuesto de Zona Franca: 0,5 %", "Ingresos brutos por ventas a terceros países",
          "Ley 523/1995, arts. 13 y 14",
          "Exentos de todo otro tributo (art. 13). Lo que sale de la zona franca al resto del país paga todos los tributos de "
          "importación (art. 20), salvo productos que cumplen el régimen de origen. En los datos de la DNIT el 0,5 % figura "
          "en OTROS en las exportaciones desde zona franca (ZF2E).", relleno=VERDE)
    for nombre, tasa, base, art, nota in [
        ("Política automotriz nacional", "0 % de arancel; IVA sobre el 20 % del valor aduanero (equivale a 2 %); motos (87.11): IVA pleno",
         "Bienes de capital, materias primas, partes e insumos para producir o ensamblar vehículos (cap. 87)",
         "Ley 4838/2012, arts. 6 y 7 (art. 7 en la redacción de la Ley 5819)", "La Ley 5819 subió la base del IVA de las motos del 40 % al 100 % del valor aduanero y destina lo recaudado al Hospital del Trauma, el SEME y el MEC."),
        ("Transporte eléctrico", "0 % de impuesto aduanero y 0 % de IVA",
         "Vehículos eléctricos, sus repuestos y equipos para ensamblarlos", "Ley 6925/2022, arts. 10, 11 y 12",
         "Copia escaneada, leída con OCR. La exoneración de repuestos dura 10 años desde la publicación."),
        ("Bienes de alta tecnología", "0 % de arancel; IVA 10 % sobre el 15 % del valor aduanero (equivale a 1,5 %)",
         "Materias primas, partes y piezas para producir bienes de alta tecnología", "Ley 4427/2012, arts. 7 y 8",
         "Texto de la BACN. Vigencia a confirmar: no figura en el art. 153 de la Ley 6380/2019."),
        ("Garantías para las inversiones y el empleo", "Sin exoneración propia en el despacho",
         "Maquinaria y equipos que no se producen en el país", "Ley 5542/2015, art. 9",
         "Remite al régimen arancelario de la Ley 60/90. Lo propio de esta ley es la invariabilidad del impuesto a la renta."),
        ("Leasing", "IVA en suspenso en la importación", "Bienes importados por el dador para entregarlos en leasing",
         "Ley 1295/1998, arts. 71 y 75", "Con fianza o con el contrato inscripto. Si el tomador tiene una exoneración, "
         "el dador importa bajo ese régimen."),
        ("Minería", "Exentos de derechos de importación, IVA y todo otro impuesto",
         "Maquinaria, vehículos, insumos y materiales para prospección y exploración", "Ley 3180/2007, art. 49",
         "Solo mientras dure la prospección y exploración; no en la explotación (art. 48)."),
        ("Cooperativas", "Exentas de aranceles aduaneros, adicionales y recargos",
         "Bienes de capital para el objeto social", "Ley 438/1994, art. 113, en la redacción de la Ley 6178 (antes, Ley 5501/2015)",
         "No se pueden transferir libremente hasta 5 años después de su ingreso. La Ley 6178 mantuvo esta exención."),
        ("Negocios fiduciarios", "Sin exoneración en el despacho",
         "Negocios y servicios fiduciarios (operación interna)", "Ley 921/1996, arts. 38 y 39",
         "Exentos de IVA y del impuesto a los actos y documentos; no tocan los tributos a la importación."),
        ("Vacunas", "Exento de todo tributo a la importación",
         "Vacunas, jeringas y equipos de cadena de frío que importa el Ministerio de Salud", "Ley 7577/2025, art. 2",
         "Despacho directo a consumo con prioridad (art. 3)."),
        ("Eventos deportivos internacionales", "Exentos de todo impuesto, tasa, contribución o canon a la importación",
         "Equipos e implementos deportivos para eventos declarados por decreto", "Ley 7467/2025, arts. 6, 7 y 8",
         "Copia escaneada, leída con OCR. También admisión temporaria sin tasas ni garantías (art. 7)."),
    ]:
        poner(nombre, tasa, base, art, nota, relleno=VERDE)
    poner("Diabetes", "Sin exoneración en el despacho", "Medicamentos e insumos que entrega el Ministerio de Salud",
          "Ley 5372/2014, art. 16", "La ley obliga a entregar gratis insulina e insumos, pero no exonera tributos a la importación.",
          relleno=VERDE)
    poner("Ley 285/1993", "Sin relación con la importación", "Beneficios adicionales que la empresa da a sus trabajadores",
          "Ley 285/1993, arts. 2 y 7", "Reglamenta el art. 93 de la Constitución: esos beneficios son gasto deducible "
          "exento de tributos, hasta el 10 % de la renta neta.", relleno=VERDE)
    poner("Acuerdos, convenios y tratados internacionales", "Preferencia arancelaria de cada acuerdo",
          "Productos con certificado de origen", "Cada acuerdo y su ley aprobatoria",
          "Lo que se usa y lo que se ahorra se mide en la página principal (comercio por acuerdo).")
    r_rmp = fila["Régimen de materias primas (RMP)"]
    ws.cell(r_rmp, 7).value = "Decreto 11.771/2000 (modifica el art. 12 del Decreto 1053/1993); prorrogado por el Decreto 8664/2022"
    poner("Régimen de materias primas (RMP)", "0 % de arancel", "Materias primas e insumos para el proceso productivo propio",
          "Decreto 1053/1993, art. 12, en la redacción del Decreto 11.771/2000, art. 1",
          "Requisitos: inscripción anual, pedido de al menos USD 1.500 FOB y que no haya producción nacional. "
          "Texto tomado de una copia de relacionespublicas.com.py (no oficial). Prórrogas: Dec. 6957/2005 y 11.525/2007 "
          "(hasta el 31/12/2008, según la Dec. CMC 14/2007), luego otras hasta el Dec. 8664/2022 (2030). "
          "La Res. MH 1/2001 integra la Comisión Técnica que dictamina cada pedido.", relleno=VERDE)
    poner("Canon informático", nota="La DNIT no lo muestra aparte: va dentro de OTROS, que en 2025 fue cerca de 0,4 % "
          "del valor en aduana (hoja Control_DNIT_2025). La Res. DGA 25/1994 lo cobra por despacho de importación desde el "
          "14/02/1994 y dice que la exportación no paga, pero el cuadro de montos (Res. MH 1144/1992, art. 2) no está en la copia.")
    # Fila nueva: el cobro de la carne exportada (ya documentado en este repositorio).
    r = ws.max_row + 1
    ws.cell(r, 1).value = r - 1
    for col, v in {2: "2. Cobros y reglas en la exportación", 3: "Cobro sobre la carne bovina exportada",
                   4: "Porcentaje sobre la exportación de carne (NCM 0201, 0202 y 0206); figura en OTROS",
                   5: "Exportación", 6: "Directa", 7: "Ley 808/1996, art. 14 c)", 8: "Resolución CIEFA 3/2021 (texto en la carpeta normas)",
                   9: "Vigente", 10: "Sí (Ley 808/1996)", 11: "Texto de la BACN; datos de la DNIT (OTROS en EC01 y EC03)", 12: "0,25 % por cada tonelada de carne",
                   13: "Valor de la carne importada y exportada", 14: "Ley 808/1996, art. 14 c) (tope); Res. CIEFA 3/2021, art. 1 (valor desde el 01/01/2022)",
                   15: "La Res. CIEFA 3/2021 (17/12/2021, Acta CIEFA 385) fija el valor en 0,25 % desde el 01/01/2022. En los datos de la DNIT se ve cerca de 0,25 %."}.items():
        ws.cell(r, col).value = v
        ws.cell(r, col).font = Font(name="Arial", size=10)
        ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(r, 12).fill = VERDE
    # Las fórmulas del resumen cubrían hasta la fila 45.
    res = wb["Resumen"]
    for row in res.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = c.value.replace("$45", f"${r}")

    # Hoja de control.
    hc = wb.create_sheet("Control_DNIT_2025")
    hc.append(["Cobro", "Columna DNIT", "Tasa mediana (% de la base en guaraníes)", "% del valor importado que lo pagó", "Qué dice la norma"])
    for f in control():
        hc.append(f)
    hc.append([])
    hc.append(["Importaciones definitivas (SE_CUENTA = SI) de 2025. Tasa = cobro / IMPONIBLE GS, mediana de las filas que pagan. "
               "Programa: regimenes_tributos/completar_cobros.py."])
    for c in hc[1]:
        c.font = Font(name="Arial", size=10, bold=True)
    for w_, col in zip((42, 14, 22, 20, 70), "ABCDE"):
        hc.column_dimensions[col].width = w_

    # Leeme: colores, avisos y normas que faltan.
    le = wb["Leeme"]
    for t in ["", "Agregado en el repositorio (20/11/2026)",
              "Celdas en celeste: tasa observada en lo que cobró la DNIT en 2025, sin norma leída todavía.",
              "Leídas en el texto el 20/11/2026: Leyes 808/1996, 302/1993, 4601/2012, 5542/2015 (BACN), 1064/1997, 110/1992, 523/1995, Decretos 11.771/2000 y 11.525/2007 Res. CIEFA 3/2021 y Leyes 4838/2012, 6925/2022, 4427/2012, 1295/1998, 3180/2007, 5501/2015, 921/1996, 7577/2025, 7467/2025, 5372/2014, 285/1993, 5819 y 6178; tasa de valoración en el Código Aduanero.",
              "Hoja Control_DNIT_2025: la tasa que se ve en los datos para cada cobro, comparada con la norma.",
              "Las tasas del ISC son las máximas de la ley; la de cada producto la fija el Decreto 3109/2019.",
              "El tributo único de maquila (1 %) se paga como impuesto interno: no figura en los despachos.",
              "El Código Aduanero (Ley 2422/2004) ya está en la carpeta normas del repositorio."]:
        le.append([t])
    for c in le["A"]:
        if c.value in ("Agregado en el repositorio (20/11/2026)",):
            c.font = Font(name="Arial", bold=True)
    wb.save(CARPETA / "Cobros_y_regimenes_especiales.xlsx")
    print("Listo:", CARPETA / "Cobros_y_regimenes_especiales.xlsx")


if __name__ == "__main__":
    main()
