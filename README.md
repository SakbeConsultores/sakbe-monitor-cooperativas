# Monitor de Cooperativas Financieras Ecuador — Segmento 1

Dashboard de indicadores financieros para las cooperativas del Segmento 1 supervisadas por la SEPS. Cubre 48 cooperativas con datos mensuales 2018–2025.

## Estructura

```
src/etl.py          Script ETL: lee Excel SEPS → genera cooperativas.json
docs/index.html     Dashboard web (GitHub Pages)
docs/data/
  cooperativas.json Datos procesados (generado por el ETL)
source/             Carpeta local para los Excel de la SEPS (no se sube al repo)
```

## Actualizar los datos

1. Descarga el nuevo archivo Excel de la SEPS y colócalo en `source/`
2. Corre el ETL desde la carpeta raíz del proyecto:

```bash
pip install openpyxl
python src/etl.py --source source/ --output docs/data/cooperativas.json
```

3. Commit y push del JSON actualizado:

```bash
git add docs/data/cooperativas.json
git commit -m "datos: actualizar a YYYY-MM"
git push
```

GitHub Pages sirve el dashboard automáticamente desde `docs/`.

## Tabs del dashboard

1. **Alertas del sistema** — cooperativas con señales de deterioro
2. **Vista ejecutiva** — KPIs principales del sistema o de una coop
3. **Cartera de crédito** — calidad de cartera, morosidad y cobertura
4. **Liquidez** — posición de liquidez histórica y comparativa
5. **Solvencia y patrimonio técnico** — cumplimiento regulatorio
6. **Resultados y rentabilidad** — P&L waterfall y eficiencia
7. **Benchmarking** — comparación entre cooperativas

## Fuente de datos

Superintendencia de Economía Popular y Solidaria (SEPS) — Ecuador.
Archivos: `Coop Ind Financieros {año} sin f.xlsx` (2018–2025).
