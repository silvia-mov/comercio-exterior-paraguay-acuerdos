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
    ws.cell(fila["Tasa por servicio de valoración aduanera"], 9).value = "Vigente (se cobró en 2025: hoja Control_DNIT_2025)"
    poner("Importación de vehículos eléctricos (2012)", "0 % de arancel y 0 % de IVA", "Vehículos eléctricos e híbridos importados",
          "Ley 4601/2012, art. 1", "Modificada por la Ley 5183/2014 (carpeta normas). Texto de la BACN.", relleno=VERDE)
    poner("Donaciones al Estado y a otras instituciones", "Exentas de todo tributo", "Bienes donados",
          "Ley 302/1993, art. 2", "No se pueden vender ni transferir por 5 años (art. 5). Texto de la BACN (escaneado, leído con OCR).",
          relleno=VERDE)
    poner("Canon informático",
          "Monto fijo por despacho, en jornales mínimos diarios: remesa expresa 0 % a 50 % de un jornal; fracciones complementarias "
          "25 %; exportación simplificada 25 %",
          "Cada despacho oficializado (no depende del valor)",
          "Res. DNA 198/2017, art. 2 (remesa expresa); Res. DNA 218/2020, art. 7 (IFC1); Res. DGA 25/1994, art. 1",
          "Remesa expresa: documentos 0 %, mercaderías de menos de USD 100 25 %, de USD 100 a 1.000 50 % (vigente desde el 01/04/2017). "
          "Para el despacho general, un resumen sin verificar habla de 2,5 jornales; el monto original está en la Res. MH 1144/1992, "
          "que falta. Lo percibe el Centro de Despachantes de Aduana (CDAP), no va al presupuesto de Aduanas. La exportación "
          "general no paga (Res. DGA 25/1994).", relleno=VERDE)
    poner("Impuesto selectivo al consumo (ISC) en la importación",
          nota="OJO: son las tasas MÁXIMAS de la ley. La tasa real de cada producto la fija el Decreto 3109/2019 "
               "y sus cambios (copia escaneada, falta leerla). Detalle: hoja Tasas_ISC.")
    poner("Maquila: nuevo régimen",
          nota="El tributo único se paga a la DNIT como impuesto interno (sobre el valor agregado en el país), "
               "no en el despacho aduanero: no aparece en los datos de comercio.")
    poner("Turismo de compras (RTC)", nota="Confirmado en los datos 2025: IVA de cerca de 1,3 % y anticipo del IRE "
          "de cerca de 1 % de la base en guaraníes (hoja Control_DNIT_2025).")
    poner("Anticipo del IRE en la importación", nota="Confirmado en el texto del Decreto 3182/2019 (BACN, extracto en normas/): 0,4 % "
          "sobre el valor aduanero más los tributos aduaneros; por eso en los datos se ve cerca de 0,46 %. Solo contribuyentes del IRE general.")

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
         "Leída en la imagen de la copia escaneada (SILpy). Las tres exoneraciones duran 10 años desde la publicación. "
         "Art. 10 incluye los equipos para centros de recarga. Art. 12 exige transferencia tecnológica con universidades "
         "públicas y colegios técnicos. Art. 8: fondo con el 10 % del ISC de naftas de 95 octanos o más, turbo fuel, nafta de "
         "aviación y gasoil premium."),
        ("Bienes de alta tecnología", "0 % de arancel; IVA 10 % sobre el 15 % del valor aduanero (equivale a 1,5 %)",
         "Materias primas, partes y piezas para producir bienes de alta tecnología", "Ley 4427/2012, arts. 7 y 8",
         "Texto de la BACN. Ya no rige (ver Vigencia); hoy aplica la Ley 7546/2025."),
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
         "Texto de la Gaceta Oficial N.º 88 del 24/04/2025 (normas/). También admisión temporaria sin tasas ni garantías (art. 7)."),
    ]:
        poner(nombre, tasa, base, art, nota, relleno=VERDE)
    poner("Emergencia por COVID-19", "Sin exoneración en el despacho",
          "Solicitudes de devolución de créditos fiscales (incluida la devolución del IVA al exportador)",
          "Ley 6524/2020, art. 15", "En 2020 suspendió las nuevas solicitudes de devolución de créditos fiscales y la resolución "
          "de las que estaban en trámite (salvo el régimen acelerado). El art. 14 permitió prorrogar vencimientos. No exoneró "
          "tributos a la importación. Texto del SILpy (normas/).", relleno=VERDE)
    poner("Diabetes", "Exoneración impositiva en la importación y en la venta",
          "Todos los medicamentos e insumos para el tratamiento y autocontrol de la diabetes",
          "Ley 5372/2014, art. 21", "La DNIT aplica la exoneración del IVA con listas por marca y principio activo (consultas "
          "vinculantes de 2016, normas/). La Ley 5372 está entre las excepciones del art. 153 de la Ley 6380/2019, así que sigue vigente.",
          relleno=VERDE)
    poner("Ley 285/1993", "Sin relación con la importación", "Beneficios adicionales que la empresa da a sus trabajadores",
          "Ley 285/1993, arts. 2 y 7", "Reglamenta el art. 93 de la Constitución: esos beneficios son gasto deducible "
          "exento de tributos, hasta el 10 % de la renta neta.", relleno=VERDE)
    def celda(nombre, col, valor):
        ws.cell(fila[nombre], col).value = valor
    poner("Tasas de legalización y aranceles consulares", "Sin tasa consular en la importación desde 2023",
          "Visado de facturas, conocimientos de embarque y demás documentos de importación",
          "Ley 4033/2010, art. 11 (tabla); Ley 7196/2023, arts. 2 a 4",
          "La Ley 7196/2023 derogó la 1.ª sección (operaciones de importación) del capítulo de comercio y los capítulos "
          "de transporte; siguen las tasas por otros trámites consulares.", relleno=VERDE)
    celda("Tasas de legalización y aranceles consulares", 7, "Ley 4033/2010, Del Arancel Consular")
    celda("Tasas de legalización y aranceles consulares", 9, "Vigente para otros trámites; derogada para la importación (Ley 7196/2023)")
    poner("Ajuste de precios en la exportación de granos", "No es una tasa: el precio de exportación se ajusta a la cotización de bolsa del día del embarque",
          "Bienes con cotización internacional (granos y derivados)", "Ley 6380/2019, art. 38, num. 7 (antes Ley 5061/2013, art. 4)",
          "La misma regla pasó de la Ley 5061/2013 a la Ley 6380/2019, para el Impuesto a la Renta Empresarial.", relleno=VERDE)
    celda("Ajuste de precios en la exportación de granos", 7, "Ley 6380/2019, art. 38, num. 7")
    celda("Ajuste de precios en la exportación de granos", 9, "Vigente (Ley 6380/2019)")
    celda("Bienes de alta tecnología", 9, "Derogada: Ley 6380/2019, art. 153, num. 5 (no está entre las excepciones) y "
          "expresamente por la Ley 7546/2025, art. 22")
    ws.cell(fila["Arancel aduanero (derecho de importación)"], 14).value = (
        "Ley 1095/1984, art. 2 (tope: hasta 30 % del valor imponible) y art. 10 a) (el Poder Ejecutivo fija y modifica los "
        "niveles); tasa de cada código: decretos de arancel y Arancel Nacional Vigente")
    poner("Sector azucarero (AZ)", "30 % en las 5 líneas de la partida 17.01 (AEC: 16 %), también para lo originario del MERCOSUR",
          "Azúcar (partida 17.01 y otras del anexo)",
          "Decreto 6898/2022, arts. 2 y 3 (Anexo III, columna A.I.); Decreto 7306/1995; Dec. CMC 19/94, arts. 3 y 4",
          "El Decreto 7306/1995 incorporó la Dec. CMC 19/94 (cada país aplica su propio arancel al azúcar, también dentro del "
          "MERCOSUR). El 30 % es el tope del arancel que permite la Ley 1095/1984, art. 2. El Decreto 6898/2022 fija hoy los aranceles intrazona del azúcar y del sector automotor; la copia "
          "Anexo III «Azúcar», columna A.I.: 30 % en 1701.12.00, 1701.13.00, 1701.14.00, 1701.91.00 y 1701.99.00 (Gaceta Oficial N.º 68 "
          "del 05/04/2022, normas/). Coincide con el Arancel Nacional Vigente de octubre 2025 (pág. 64).", relleno=VERDE)
    celda("Sector azucarero (AZ)", 7, "Decreto 6898/2022 (antes Decreto 7306/1995 y 6654/2016); Dec. CMC 19/94")
    celda("Régimen automotor (AUT)", 8, "Decreto 6898/2022, arts. 1, 3 y 4 (Anexos I y II): aranceles intrazona de vehículos y autopartes; "
          "Dec. CMC 29/94. Bienes usados del cap. 87: Decreto 7244/2017")
    v = ws.cell(fila["Diabetes"], 15).value
    celda("Diabetes", 15, v + " Además, la Ley 77/1992 (arts. 1 y 2) exonera de tributos aduaneros e internos la insulina, los "
          "insumos para diabetes y medicamentos para el cáncer y otras afecciones; no figura entre las excepciones del art. 153 de la "
          "Ley 6380/2019 y, según un resumen sin verificar, las leyes de presupuesto 2025 y 2026 no la exceptuaron: su alcance hoy es "
          "dudoso, pero la insulina sigue cubierta por la Ley 5372.")
    # Ley 6380/2019, art. 153, num. 5: siguen vigentes solo las exoneraciones de esta lista.
    for nombre in ["Donaciones al Estado y a otras instituciones", "Franquicias diplomáticas y consulares", "Ley 285/1993",
                   "Cooperativas", "Negocios fiduciarios", "Leasing", "Minería", "Zonas francas",
                   "Inversión: Ley 60/90 (régimen anterior)", "Maquila: Ley 1064/1997 (régimen anterior)",
                   "Política automotriz nacional", "Importación de vehículos eléctricos (2012)", "Diabetes",
                   "Garantías para las inversiones y el empleo"]:
        v = ws.cell(fila[nombre], 9).value or ""
        if "Ley 6380" not in v:
            celda(nombre, 9, (v + ". " if v and v != "Vigente" else "Vigente. ") + "Exceptuada de la derogación general de la Ley 6380/2019, art. 153, num. 5")
    poner("Inversión: Ley 60/90 (régimen anterior)",
          "0 % de gravámenes aduaneros (bienes de capital, materias primas e insumos del proyecto); IVA 0 % para bienes de capital",
          "Lo previsto en el proyecto de inversión aprobado", "Ley 60/1990, art. 5 c) (texto de 1990); IVA: Ley 6380/2019, art. 100, num. 4",
          "Derogada por la Ley 7548/2025 (art. 34); los proyectos ya aprobados siguen (art. 35).", relleno=VERDE)
    # Qué normas están ya en la carpeta normas/ del repositorio o en la release normas_aduanas.
    en_repo = {
        "Arancel aduanero (derecho de importación)": "Sí: Ley 1095/1984 (BACN, escaneada, leída con OCR) y Código Aduanero (normas/). El archivo F07 de la release normas_aduanas es otra ley (1095/1997)",
        "Tasa por servicio de valoración aduanera": "Sí: Código Aduanero (normas/)",
        "Canon informático": "Sí: Res. DGA 25/1994, Res. DNA 198/2017 y 218/2020 (normas/); falta la Res. MH 1144/1992",
        "Franquicias diplomáticas y consulares": "Sí (normas/)", "Donaciones al Estado y a otras instituciones": "Sí (normas/, BACN)",
        "Zonas francas": "Sí (normas/, BACN)", "Inversión: Ley 60/90 (régimen anterior)": "Sí (normas/)",
        "Inversión: nuevo régimen de incentivos fiscales": "Sí, con su reglamento (Decreto 5432/2026, release normas_aduanas)",
        "Maquila: Ley 1064/1997 (régimen anterior)": "Sí (normas/)",
        "Maquila: nuevo régimen": "Sí, con su reglamento (Decreto 5714/2026, release normas_aduanas)",
        "Política automotriz nacional": "Sí, con la Ley 5819 (normas/)", "Importación de vehículos eléctricos (2012)": "Sí (normas/, BACN)",
        "Bienes de alta tecnología": "Sí (normas/, BACN)", "Garantías para las inversiones y el empleo": "Sí (normas/, BACN)",
        "Leasing": "Sí (normas/, BACN)", "Minería": "Sí (normas/, BACN)", "Diabetes": "Sí (normas/)",
        "Cooperativas": "Sí: Ley 438/1994 y sus cambios (Leyes 5501/2015 y 6178) en normas/",
        "Régimen de materias primas (RMP)": "Sí: Decretos 11.771/2000, 11.525/2007 y 6539/2021 y Res. MH 1/2001 (normas/); el 8664/2022 está citado",
        "Negocios fiduciarios": "Sí (normas/, BACN)", "Ley 285/1993": "Sí (normas/)",
        "Vacunas": "Sí (normas/, BACN)", "Eventos deportivos internacionales": "Sí (normas/, Gaceta Oficial N.º 88, con texto)",
        "Tasas de legalización y aranceles consulares": "Sí: Leyes 4033/2010 y 7196/2023 (release normas_aduanas)",
        "Ajuste de precios en la exportación de granos": "Sí: Ley 5061/2013 (normas/) y Ley 6380/2019 (BACN, leída con OCR)",
        "Lista nacional de excepciones (LNE)": "En parte: Decreto 3339/2025 (normas/)",
        "Sector azucarero (AZ)": "Sí: Decreto 6898/2022 con anexos (Gaceta Oficial N.º 68) y Dec. CMC 19/94",
    }
    for nombre, art in [("Bienes de capital (BK)", "art. 5 (vigencia hasta el 31/12/2030, Dec. CMC 08/21)"),
                        ("Informática y telecomunicaciones (BIT)", "art. 6 (vigencia hasta el 31/12/2030, Dec. CMC 08/21)"),
                        ("Lista nacional de excepciones (LNE)", "art. 4 (vigencia hasta el 31/12/2030, Dec. CMC 11/21)"),
                        ("Régimen automotor (AUT)", "art. 7 (vigente hasta que el MERCOSUR incorpore el sector)"),
                        ("Sector azucarero (AZ)", "art. 7 (vigente hasta que el MERCOSUR incorpore el sector)")]:
        r_ = fila[nombre]
        prev = ws.cell(r_, 8).value or ""
        ws.cell(r_, 8).value = (prev + "; " if prev else "") + ("Decreto 6897/2022 (arancel consolidado para importaciones de fuera "
                                "del MERCOSUR; la columna «Lista» identifica cada lista), " + art +
                                ". El Decreto 6542/2021 incorporó las Dec. CMC 08/21, 11/21 y 12/21.")
        en_repo.setdefault(nombre, "Sí: Decreto 6897/2022 (MEF, normas/)")
    for nombre, v in en_repo.items():
        celda(nombre, 10, v)
    ws.cell(fila["Maquila: nuevo régimen"], 8).value = "Reglamento: Decreto 5714/2026 (release normas_aduanas)"
    ws.cell(fila["Inversión: nuevo régimen de incentivos fiscales"], 8).value = "Reglamento: Decreto 5432/2026 (release normas_aduanas)"
    poner("Impuesto selectivo al consumo (ISC) en la importación",
          "Tabaco 20 % (Decreto 6619/2022; antes 18 %); bebidas 5 % a 11 %; alto contenido calórico 2 %; combustibles 0 % a 38 %; otros bienes 1 % a 5 % "
          "(detalle: hoja ISC_Decreto_3109)", None, "Decreto 3109/2019, anexo, arts. 8 a 12 (tasas máximas: Ley 6380/2019, arts. 115 a 119)",
          "Tasas del texto de 2019 del Decreto 3109 (normas/). Puede haber cambios posteriores, por ejemplo el Decreto 7229/2026 "
          "sobre la base del gasoil. Control con la DNIT 2025 (ISC / (valor + arancel), mediana): celulares 1,0 %, cerveza 9,05 %, "
          "whisky 11,05 %, perfumes 5,02 % = decreto. Gasoil 14,3 %: el Decreto 7229/2026 fija una base especial de "
          "G 1.888,89 por litro hasta el 31/12/2026 (normas/, leído con OCR). Cigarrillos 20,1 %: el Decreto 6619/2022 subió la tasa de "
          "tabacos, cigarrillos y esencias de 18 % a 20 % (confirmado en el Informe del Primer Semestre 2022 del Ministerio de Hacienda, "
          "pág. 16, normas/). Según un resumen sin verificar, el Decreto 8878/2023 volvió a modificar ese artículo manteniendo el 20 %.",
          relleno=VERDE)
    import docx as _docx
    _d = _docx.Document(str(RAIZ / "normas" / "Decreto_3109_2019_Reglamento_ISC.docx"))
    hi = wb.create_sheet("ISC_Decreto_3109")
    hi.append(["Artículo del anexo", "Grupo", "Bienes gravados", "Tasa"])
    for tb, (art, grupo) in zip(_d.tables, [("8", "Tabaco"), ("9", "Bebidas"), ("10", "Alto contenido calórico"),
                                            ("11", "Combustibles derivados del petróleo"), ("12", "Otros bienes")]):
        for fila_t in tb.rows[1:]:
            bien = fila_t.cells[0].text.strip().replace("\n", " ")
            tasa = fila_t.cells[-1].text.strip()
            hi.append([art, grupo, bien, float(tasa.rstrip("%").replace(",", ".")) / 100 if tasa.endswith("%") else tasa])
    for c in hi[1]:
        c.font = Font(name="Arial", size=10, bold=True)
    for fila_t in hi.iter_rows(min_row=2):
        fila_t[3].number_format = "0%"
        fila_t[2].alignment = Alignment(wrap_text=True, vertical="top")
    for col, w_ in zip("ABCD", (12, 26, 90, 8)):
        hi.column_dimensions[col].width = w_
    hi.freeze_panes = "A2"
    poner("Acuerdos, convenios y tratados internacionales", "Preferencia arancelaria de cada acuerdo",
          "Productos con certificado de origen", "Cada acuerdo y su ley aprobatoria",
          "Lo que se usa y lo que se ahorra se mide en la página principal (comercio por acuerdo).",
          relleno=PatternFill(fill_type=None))
    r_rmp = fila["Régimen de materias primas (RMP)"]
    ws.cell(r_rmp, 7).value = "Decreto 11.771/2000 (modifica el art. 12 del Decreto 1053/1993); prorrogado por el Decreto 8664/2022"
    poner("Régimen de materias primas (RMP)", "0 % de arancel", "Materias primas e insumos para el proceso productivo propio",
          "Decreto 1053/1993, art. 12, en la redacción del Decreto 11.771/2000, art. 1",
          "Requisitos: inscripción anual, pedido de al menos USD 1.500 FOB y que no haya producción nacional. "
          "Texto tomado de una copia de relacionespublicas.com.py (no oficial). Prórrogas: Dec. 6957/2005 y 11.525/2007 "
          "(hasta el 31/12/2008, según la Dec. CMC 14/2007), luego otras hasta el Dec. 8664/2022 (2030). "
          "La Res. MH 1/2001 integra la Comisión Técnica que dictamina cada pedido.", relleno=VERDE)
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
              "Ley 6380/2019, art. 153, num. 5: deroga todas las exoneraciones de impuestos salvo las de su lista (leído con OCR). Leídas en el texto el 20/11/2026: Leyes 808/1996, 302/1993, 4601/2012, 5542/2015 (BACN), 1064/1997, 110/1992, 523/1995, Decretos 11.771/2000 y 11.525/2007 Res. CIEFA 3/2021 y Leyes 4838/2012, 6925/2022, 4427/2012, 1295/1998, 3180/2007, 5501/2015, 921/1996, 7577/2025, 7467/2025, 5372/2014, 285/1993, 5819, 6178 y 6524/2020; tasa de valoración en el Código Aduanero.",
              "Hoja Control_DNIT_2025: la tasa que se ve en los datos para cada cobro, comparada con la norma.",
              "Las tasas del ISC de cada producto están en la hoja ISC_Decreto_3109 (texto de 2019 del decreto).",
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
