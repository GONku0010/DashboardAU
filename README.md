# Autlan Micro-Macro Executive Dashboard

Dashboard ejecutivo en Streamlit construido sobre los outputs auditables generados por `Micro-Macro.ipynb`.

## Estructura

```text
Dashboard/
  app.py
  requirements.txt
  README.md
  .streamlit/
    config.toml
Outputs análisis/
  *.csv
  Dashboards HTML/
Micro-Macro.ipynb
```

## Fuente de datos

La app no recalcula los modelos principales. Lee directamente los CSVs exportados por el notebook:

- `autlan_analysis_dataset.csv`
- `autlan_financials_clean.csv`
- `executive_sensitivity_summary.csv`
- `model_comparison_all.csv`
- `best_models_by_outcome.csv`
- `monthly_macro_financial_trends_long.csv`
- `annual_macro_financial_trends_long.csv`
- `monthly_macro_seasonality_summary.csv`
- `quarterly_business_seasonality_summary.csv`
- `rate_regime_summary.csv`
- `rolling_correlations_by_rate_regime.csv`
- `usd_mxn_manganese_*`

Si cambian los datos o el notebook, primero ejecuta `Micro-Macro.ipynb` para regenerar los outputs.

## Correr localmente

Desde la raiz del proyecto:

```bash
streamlit run Dashboard/app.py
```

Si quieres usar un puerto especifico:

```bash
streamlit run Dashboard/app.py --server.port 8502 --server.fileWatcherType none
```

## Tabs incluidas

- Resumen Ejecutivo
- Proyecciones y Escenarios
- Sensibilidades
- Margenes
- Top-Line / Ingresos
- Bottom-Line / Rentabilidad
- Ciclos y Estacionalidad
- Correlaciones y Regresiones
- Data Explorer

## Notas metodologicas

- Las sensibilidades se muestran con `R2`, signo economico y uso recomendado.
- Las proyecciones USD/MXN y manganeso usan los escenarios ya exportados: Bear, Base, Bull y Drift.
- Las graficas de forecast mantienen las correcciones visuales acordadas: lineas principales con contraste, fan chart tenue, labels terminales directos, separacion automatica de labels y periodo forecast sombreado.
- El dashboard esta preparado para deploy cloud siempre que se suban tambien los CSVs de `Outputs análisis`.
