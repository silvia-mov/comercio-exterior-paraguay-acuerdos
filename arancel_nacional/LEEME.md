# Arancel Nacional Vigente de Paraguay frente al AEC

El **Arancel Nacional Vigente (ANV) de Paraguay, actualizado a octubre de 2025**, pasado de PDF a
Excel, con la diferencia frente al **Arancel Externo Común (AEC)** del MERCOSUR.

**Fuente principal:** "Arancel Nacional Vigente (ANV) actualizado a octubre 2025", Dirección General
de Política Comercial e Integración, Ministerio de Economía y Finanzas (NCM 2022, 455 páginas).
El propio documento aclara que es un documento de trabajo y no tiene fuerza legal. Este
repositorio no está asociado al Ministerio ni cuenta con su participación o apoyo.

## Qué hay en el Excel

`Arancel_Nacional_Vigente_oct2025_vs_AEC.xlsx`. La hoja **Leeme** del Excel explica cada columna.

| Hoja | Contenido |
|---|---|
| Resumen | Líneas con ANV menor, igual o mayor que el AEC, por lista y en el Régimen de Materias Primas |
| Listas | Qué es cada lista del PDF (BK, BIT, LNE, AUT, AZ*, Decisiones CMC 08/22, 37/07, 29/15 y 30/15) |
| Arancel | Una fila por línea arancelaria: AEC, ANV, diferencia ANV − AEC y control con la NCM de noviembre 2026 |
| Historia_BK, Historia_BIT, Historia_LNE | Códigos de cada lista según las notificaciones de Paraguay al MERCOSUR, 2018 a enero 2026 |
| Materias_primas | Códigos del Régimen de Materias Primas por año, 2018 a 2025 (arancel 0 % dentro del régimen) |
| Estructura | Todas las filas del PDF en orden (partidas, subpartidas, ítems, códigos y desgloses) |
| NCM_nov2026 | NCM de noviembre 2026 de SAREM (`../NCM_2026_11.xlsx`), solo para control |

Resultado: el 38 % de las líneas tiene un ANV menor que el AEC (8,2 puntos menos en promedio),
el 61 % igual y 38 líneas mayor.

## Controles

1. Con la NCM de noviembre 2026: 10.486 de 10.496 códigos comunes tienen el mismo AEC. Los 10
   distintos se revisaron en el PDF y son cambios reales.
2. Con la notificación de Paraguay al 31/07/2025: las listas BK (1.257), BIT (361) y LNE (715)
   coinciden código por código y arancel por arancel.
3. Con la Decisión CMC 08/22 (rebaja del 10 %): 1.089 de 1.095 líneas cumplen ANV = AEC × 0,9.
   Las 6 restantes están anotadas en la hoja Leeme del Excel.

## Programas

| Programa | Qué hace |
|---|---|
| `leer_pdf_arancel.py` | Lee el texto del PDF y arma una fila por código, con AEC, ANV, lista y desgloses nacionales |
| `leer_notificaciones.py` | Lee las notificaciones de BK, BIT y LNE (código y arancel) |
| `historia_listas.py` | Junta las notificaciones de 2018 a 2026 y las listas del Régimen de Materias Primas |
| `armar_excel.py` | Arma el Excel con fórmulas |

Para volver a generar el Excel (necesita Python con `openpyxl` y `pdftotext`, de poppler):

```bash
pdftotext -layout "Arancel Nacional Vigente_octubre 2025_1.pdf" arancel.txt
python leer_pdf_arancel.py arancel.txt arancel.json
python historia_listas.py carpeta_notificaciones historia.json
python armar_excel.py arancel.json ../NCM_2026_11.xlsx Arancel_Nacional_Vigente_oct2025_vs_AEC.xlsx historia.json
```

`carpeta_notificaciones` es la carpeta con los cinco zip de notificaciones descomprimidos (BK, BIT,
excepciones al AEC, materias primas y regímenes especiales). El PDF y los zip no se publican aquí.

El Excel se guarda con fórmulas sin calcular: Excel las calcula al abrirlo.
