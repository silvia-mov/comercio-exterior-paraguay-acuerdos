# Impuestos por destinación (régimen) de la DNIT

Para cada destinación de `config/regimenes.csv`: qué impuestos se pagan (arancel, IVA, ISC,
servicio de valoración, anticipo de renta, otros), sobre qué base y con qué norma, y un control con
lo que efectivamente cobró la DNIT en 2025.

| Archivo | Qué es |
|---|---|
| `Tributos_por_regimen.xlsx` | El resultado (Calibri 10). Hojas: *Por destinación*, *Control 2025 (detalle)*, *Códigos fuera de la lista*, *Normas faltantes* y *Notas* |
| `normas_por_regimen.csv` | Lo que dice la norma para cada destinación. Se edita a mano |
| `normas_faltantes.csv` | Normas que no están en la biblioteca y conviene buscar |
| `control_2025.csv` | Suma por destinación de la base imponible y de cada impuesto (datos DNIT 2025, todos los regímenes) |
| `tributos_por_regimen.py` | El programa |

```bash
python python/descargar_dnit.py --anios 2025                                # baja los 12 meses a datos_dnit/2025
python regimenes_tributos/tributos_por_regimen.py --control --anio 2025     # rehace control_2025.csv
python regimenes_tributos/tributos_por_regimen.py --excel --anio 2025       # rehace el Excel
```

**Resultado 2025:** de las 63 destinaciones con datos, en las 63 lo cobrado coincide con la norma.
Las 27 restantes no tuvieron despachos en 2025.
- Importación a consumo: se pagan arancel, IVA, ISC, 0,5 % de servicio de valoración, 0,4 % de
  anticipo de IRE (solo contribuyentes del régimen general) y cerca de 0,4 % de "Otros".
- Suspensivos (maquila, depósito, zona franca, leasing): no pagan tributos. La admisión temporaria
  (IT04, IT14) paga el servicio de valoración y, en algunos ítems, una parte del arancel.
- Exportación: no paga, salvo cerca de 0,25 % sobre la carne bovina y el 0,5 % del Impuesto de
  Zona Franca en ZF2E (ambos en "Otros").

**Normas:** IVA, ISC y anticipo de IRE se leyeron en la biblioteca: Ley 6380/19 y Decretos 3107,
3109 y 3182 de 2019. Maquila: Ley 7547/25. Inversión: Ley 7548/25. Una norma con asterisco (*) no
está en la biblioteca: falta, sobre todo, el Código Aduanero (Ley 2422/04), su reglamento y las
normas del servicio de valoración y de "Otros". La lista completa está en `normas_faltantes.csv`.
