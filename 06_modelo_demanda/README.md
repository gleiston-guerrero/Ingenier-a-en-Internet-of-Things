# Modelo de demanda y oferta de profesionales IoT en Los Ríos

`Modelo_demanda_oferta_IoT_LosRios.xlsx` cruza las empresas formales de Los Ríos (SRI e INEC) con supuestos de adopción de IoT y las compara con la oferta de graduados TIC. Tiene fórmulas vivas: al cambiar una celda amarilla de la hoja `Supuestos`, se recalculan `Demanda`, `Oferta`, `Brecha` y `Sensibilidad`.

**Advertencia.** Los porcentajes de adopción y de contratación son supuestos de trabajo, sin fuente oficial. Deben reemplazarse con la encuesta a empleadores.

## Contenido

| Ruta | Qué es |
|---|---|
| `Modelo_demanda_oferta_IoT_LosRios.xlsx` | El modelo |
| `datos/` | Conteos agregados del catastro RUC del SRI y los supuestos de adopción, en CSV |
| `src/preparar_datos_sri.py` | Agrega el catastro RUC de Los Ríos y escribe solo conteos |
| `src/generar_modelo.py` | Genera el libro de Excel a partir de `datos/` |

## Regenerar

```bash
pip install pandas openpyxl
python src/preparar_datos_sri.py RUTA/SRI_RUC_Los_Rios.csv
python src/generar_modelo.py
```

El catastro `SRI_RUC_Los_Rios.csv` se descarga de https://descargas.sri.gob.ec/download/datosAbiertos/SRI_RUC_Los_Rios.zip. No se guarda en este repositorio porque incluye nombres de personas naturales.
