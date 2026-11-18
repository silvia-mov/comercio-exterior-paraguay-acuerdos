# Normas a buscar y subir

Subirlas en PDF a la carpeta `normas/` de este repositorio (o pasarlas por el chat).

Ya están en `normas/` (bajadas del digesto aduanero de la DNIT): Ley 2422/04 (Código Aduanero),
Decretos 859, 1952/2024 y 3465 (modifican el Decreto 4672/05), Instructivo GGA 06/2024 y Ley 7196.
Dónde buscar: Biblioteca y Archivo Central del Congreso (BACN, www.bacn.gov.py), Gaceta Oficial
(www.gacetaoficial.gov.py), digesto aduanero de la DNIT (www.dnit.gov.py/web/portal-institucional/digesto-normativo,
normas de 2023 en adelante), histórico de la ex DNA (digesto.aduana.gov.py, caído en noviembre de 2026), MIC y SILpy.

## Prioridad 1: explican los números

| # | Norma | Para qué |
|---|---|---|
| 1 | ~~Ley 2422/2004, Código Aduanero~~ **Encontrada** en el digesto aduanero de la DNIT ([PDF](https://www.dnit.gov.py/documents/44828/0/1517412693.pdf/c817f24a-046b-8282-7b24-9ba74d935396)); falta una versión con las modificaciones posteriores | Base de todos los regímenes y de qué paga cada uno |
| 2 | Decreto 4672/2005 (reglamento del Código Aduanero), texto completo. Sus modificaciones (Dec. 859, 1952/2024 y 3465) ya están en el digesto de la DNIT | Detalle de cada régimen |
| 3 | Ley y decreto del **Régimen de Turismo** + resolución de la DNIT que dice cómo se declara | Pedido del jefe: identificarlo en los datos |
| 4 | Ley o decreto de incentivo a **vehículos eléctricos e híbridos** | Explica casi todo el arancel 0 "sin identificar" (autos 8703, motos 8711.60) |
| 5 | ~~Norma del Servicio de Valoración Aduanera (0,5 %)~~ **Resuelto:** Ley 2422/04 arts. 263 y 393 inc. a) | Columna SERVICIO |
| 6 | Qué cobros suma la DNIT en **OTROS** (≈ 0,4 % en importaciones). Revisar también la Ley 4033/2010 del Arancel Consular (en parte derogada por la Ley 7196) | Columna OTROS (conviene preguntar a la DNIT) |
| 7 | Decreto del **Arancel Nacional vigente** de cada año y Ley 1095/1984 (Arancel de Aduanas) | Arancel cobrado por NCM |

## Prioridad 2: exoneraciones y regímenes especiales

| # | Norma | Para qué |
|---|---|---|
| 8 | Ley 60/1990 de inversiones y su decreto | Bienes de capital con arancel 0 fuera de las listas |
| 9 | Leyes de exoneración de **medicamentos** y **maquinaria agrícola** que mencionó el jefe | Arancel 0 "sin acuerdo" |
| 10 | Ley 523/1995 Zonas Francas y Decreto 15.554/1996 | Impuesto de zona franca 0,5 %; ZF01, ZF2I, ZF2E |
| 11 | Ley 1295/1998 Leasing | Tributos suspendidos (IML1, IML2) |
| 12 | Ley 608/1995 Regularización | Régimen IC03 |
| 13 | Norma del cobro de ≈ 0,25 % sobre **carne bovina exportada** (posible tasa de SENACSA; su sitio bloquea accesos automáticos) | Aparece en OTROS de EC01 (no es el Decreto 11.915/08) |
| 14 | Resolución DNA 136/2018 Remesa expresa (courier). El Código Aduanero art. 223 ya fija la exención hasta USD 100 FOB | IRE1 a IRE4 |
| 15 | Resolución DNA 1252/2020 (OEA/OEC) | IC19, IC20, TDA5 |
| 16 | Convenciones de Viena y sus leyes aprobatorias (el Código Aduanero arts. 237-238 remite a ellas) | ID04, ID09 |

## Prioridad 3: detalles

| # | Norma | Para qué |
|---|---|---|
| 17 | Decreto 5714/2026 (reglamento de la Ley 7547/25 de maquila) | Regímenes de maquila |
| 18 | RG DNIT 14/2024 (aprovisionamiento a bordo). El Instructivo GGA 06/2024 ya está en el digesto de la DNIT | AB01, AFC1, DAAB, DAAR, ICAB |
| 19 | Resolución DNA 550/2022 (depósitos aduaneros) | IDA2, IDA3, IC09 |
| 20 | Resolución DNA 218/2020 (despacho fraccionado) | IFC1 |
| 21 | Resolución DNA 285/2017 (exportación menor simplificada) | DSEM |
| 22 | Normas de materiales aeronáuticos (asistencia y salvamento ya resuelto: Código Aduanero arts. 239-240) | IMAR, IDMA, ICMA, ERMA, SPA1 |
| 23 | ~~Qué son los códigos AFC1, DAAR y TZF1~~ **Resuelto:** AFC1 = aprovisionamiento fraccionado complementario (Instructivo GGA 06/2024); DAAR = aprovisionamiento de repuestos y rancho; TZF1 = traspaso dentro de zona franca | Columna SE_CUENTA |

Detalle con las destinaciones de cada norma: `regimenes_tributos/normas_faltantes.csv`.
