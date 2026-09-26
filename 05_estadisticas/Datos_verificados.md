# Datos estadísticos usados en el estudio de pertinencia

Cada cifra del borrador (sección 4) se verificó contra el documento oficial archivado en esta carpeta, salvo donde se indica. Consulta: 26 de septiembre de 2026.

| Dato | Valor | Documento y ubicación |
|---|---|---|
| Carreras TIC ofertadas por universidades, 2024 | 197 | `../01_normativa/Otros/MINTEL_Politica_Transformacion_Digital_2025-2030.pdf`, Tabla 8, p. 55 (datos SENESCYT 2024) |
| Carreras de Desarrollo de software | 62 | Mismo documento, Tabla 7, p. 52 |
| Carreras "Internet de las cosas" y "Técnico superior en internet de las cosas" | 2 y 1 | Mismo documento, Tabla 7, p. 53 |
| Graduados TIC 2021 y 2020 | 7.476 y 5.079 | Mismo documento, texto de la p. 53. La Tabla 8 consigna 7.485 para 2021: discrepancia en la fuente |
| Habilidades avanzadas de software (2019, INEC) | 15,27 % | Mismo documento, Tabla 9, p. 55 |
| IoT como tendencia de transformación digital | Lista de 14 tendencias | Mismo documento, p. 46 |
| Producción de banano en Los Ríos, 2024 | 39,8 % de la nacional | `INEC_ESPAC_2024_resultados.pdf`, sección de cultivos permanentes |
| Producción de maíz duro seco en Los Ríos, 2024 | 43,8 % de la nacional | Mismo documento, cultivos transitorios |
| Producción de palma africana en Los Ríos, 2024 | 28,6 % de la nacional | Mismo documento, cultivos permanentes |
| Producción de arroz en Los Ríos, 2024 | 27,3 % de la nacional | Mismo documento. El texto dice que Guayas concentra 67,7 % y el gráfico muestra 66,7 %: discrepancia menor en la fuente |
| Superficie cultivada con riego, 2024 | 28,2 % (1,30 millones de ha) | `INEC_MIATA_ESPAC_2024_tecnificacion.pdf`, resumen de resultados |
| Cultivos permanentes bajo riego, 2024 | 40,9 % | Mismo documento |
| Unidades de producción que nunca hicieron análisis de suelo | 88,9 % | Mismo documento, tecnología agrícola |
| Unidades de producción con capacitación o asistencia técnica, 2024 | 12,5 % | Mismo documento, tecnología agrícola |
| Empleo adecuado en Los Ríos, 2025 y 2024 | 32,5 % y 36,3 % | `INEC_ENEMDU_anual_2025_boletin.pdf`, Tabla 5 |
| Subempleo en Los Ríos, 2025 y 2024 | 28,4 % y 24,0 % | Mismo documento, Tabla 6 |
| Desempleo en Los Ríos, 2025 | 1,7 % | Mismo documento, Tabla 3 |
| Empleo adecuado nacional, 2025 | 37,1 % | Mismo documento, sección 1.4.1 |
| Desempleo nacional, 2025 | 3,6 % | Mismo documento, sección 1.1 |
| Hogares con acceso a internet, 2025 | 71,3 % nacional; 76,6 % urbano; 58,7 % rural | `INEC_TIC_ENEMDU_2025-07.pdf`, sección Acceso a internet |
| Hogares rurales con internet, 2022 | 38,0 % | Mismo documento, serie histórica (lectura de un gráfico) |
| Penetración de internet fijo y móvil, cuarto trimestre de 2024 | 17,48 % y 65,6 % | `ARCOTEL_Boletin_cierre_2024.pdf`, introducción |
| Los Ríos entre las cinco provincias con más radiobases | Guayas, Pichincha, Manabí, Azuay, Los Ríos | Mismo documento, figura 13 y su comentario |
| Empresas del universo REEM en Los Ríos, 2025 | 32.993 (11.ª provincia) | `INEC_REEM_2025_tabulados.xlsx`, cuadro 1.1.8 |
| Empresas con ventas y empleo IESS en Los Ríos, 2025 | 1.968 (8.ª provincia) | Mismo documento, cuadro 2.1.7 |
| Empleo registrado equivalente en Los Ríos, 2025 | 66.742,5 (7.ª provincia) | Mismo documento, cuadro 1.3.18 |
| RUC registrados y activos en Los Ríos | 276.712 y 112.164 | Catastro RUC del SRI, `SRI_RUC_Los_Rios.csv`, actualizado el 1-IX-2026. Cálculo propio; el archivo no se guarda en el repositorio (contiene nombres de personas naturales) |
| Empresas formales (sociedades y personas naturales obligadas a contabilidad) | 6.957 (5.332 y 1.625) | Mismo catastro; agregados en `../06_modelo_demanda/datos/sri_grupos.csv` |
| Graduados de grado de la UTEQ en 2024 | 1.620 (669 y 951) | `UTEQ_Informe_rendicion_cuentas_2024.pdf`, sección Graduados |
| Tasa de titulación UTEQ, cohorte 2018-2019 | 60,96 % | Mismo documento, tabla 6 |
| Graduados de las carreras TIC de la UTEQ, cohorte 2018-2019 | Telemática 16 y 19; Software 10; una fila sin nombre 24 | Mismo documento, tabla 6 (la fila sin nombre aparece bajo la Facultad de Ciencias de la Ingeniería) |
| Graduados de la UTEQ empleados en su área | 67,32 % | Mismo documento, seguimiento a graduados |

## Límites

- Las cifras del MIATA son nacionales. No se encontró el desglose para Los Ríos.
- La serie de hogares con internet en 2022 sale de un gráfico del tabulado del INEC, no de una tabla numérica.
- No se extrajo el subempleo nacional de 2025.
- La matrícula y los graduados TIC por provincia de la SENESCYT no se descargaron; el portal de datos abiertos las publica.
- No se consultaron los mapas de cobertura por parroquia de la ARCOTEL ni indicadores del CACES sobre carreras TIC.
- La base de matrícula 2015-2023 de la SENESCYT (datosabiertos.gob.ec) no se pudo descargar: su servidor no respondió. No se encontró una serie de graduados por provincia.
- Las horas de matemática y física por tipo de bachillerato salen de un reportaje de Primicias sobre la malla 2024-2025 de la Sierra; no se encontró la malla oficial de la Costa ni la reforma 2026.
- Los graduados de Los Ríos por tipo de bachillerato (AMIE) no se obtuvieron.
