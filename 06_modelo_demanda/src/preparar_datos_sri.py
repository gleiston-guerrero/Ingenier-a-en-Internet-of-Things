"""Agrega el catastro RUC de Los Ríos (SRI, datos abiertos) y escribe solo conteos.

Uso: python preparar_datos_sri.py RUTA/SRI_RUC_Los_Rios.csv
El archivo de origen contiene nombres de personas naturales: no se copia ni se guarda.
Descarga: https://descargas.sri.gob.ec/download/datosAbiertos/SRI_RUC_Los_Rios.zip
"""
import os
import sys

import pandas as pd

origen = sys.argv[1]
salida = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'datos')
os.makedirs(salida, exist_ok=True)

cols = ['NUMERO_RUC', 'ESTADO_CONTRIBUYENTE', 'OBLIGADO', 'TIPO_CONTRIBUYENTE', 'NUMERO_ESTABLECIMIENTO',
        'DESCRIPCION_CANTON_EST', 'CODIGO_CIIU', 'FECHA_INICIO_ACTIVIDADES']
df = pd.read_csv(origen, sep='|', dtype=str, encoding='utf-8', usecols=cols)
df['NUM'] = df.NUMERO_ESTABLECIMIENTO.astype(int)
matriz = df.sort_values('NUM').drop_duplicates('NUMERO_RUC', keep='first')
a = matriz[matriz.ESTADO_CONTRIBUYENTE == 'ACTIVO'].copy()
a['SEC'] = a.CODIGO_CIIU.str[0]
a['DIV'] = a.CODIGO_CIIU.str[1:3]


def grupo(r):
    s, d = r.SEC, r.DIV
    if s == 'A' and d in ('01', '02'):
        return 'G1 Agricultura, ganadería y silvicultura'
    if s == 'A' and d == '03':
        return 'G2 Pesca y acuicultura'
    if s == 'C' and d in ('10', '11'):
        return 'G3 Alimentos y bebidas (agroindustria)'
    if s == 'C':
        return 'G4 Otras manufacturas'
    if s in ('D', 'E'):
        return 'G5 Energía, agua y saneamiento'
    if s == 'H':
        return 'G6 Transporte y almacenamiento'
    if s == 'G' and d == '46':
        return 'G7 Comercio al por mayor'
    if s == 'O':
        return 'G8 Administración pública (GAD y entidades)'
    if s == 'J' or (s == 'M' and d in ('71', '72')):
        return 'G9 Proveedores tecnológicos (TIC, ingeniería, I+D)'
    if s == 'F':
        return 'G10 Construcción'
    if s == 'Q':
        return 'G11 Salud'
    return 'Otros sectores'


a['grupo'] = a.apply(grupo, axis=1)
soc = a[a.TIPO_CONTRIBUYENTE == 'SOCIEDAD'].groupby('grupo').size()
pn = a[(a.TIPO_CONTRIBUYENTE != 'SOCIEDAD') & (a.OBLIGADO == 'S')].groupby('grupo').size()
act = a.groupby('grupo').size()
t = pd.DataFrame({'sociedades': soc, 'pn_obligadas': pn, 'activos_total': act}).fillna(0).astype(int)
t['formales'] = t.sociedades + t.pn_obligadas
orden = sorted(t.index, key=lambda g: (g == 'Otros sectores', int(g.split(' ')[0][1:]) if g.startswith('G') else 99))
t = t.loc[orden]
t[['sociedades', 'pn_obligadas', 'formales', 'activos_total']].to_csv(os.path.join(salida, 'sri_grupos.csv'), encoding='utf-8')

formales = a[(a.TIPO_CONTRIBUYENTE == 'SOCIEDAD') | (a.OBLIGADO == 'S')]
objetivo = formales[formales.grupo.str.match(r'G[1-9] ')]
objetivo.DESCRIPCION_CANTON_EST.value_counts().head(13).to_csv(os.path.join(salida, 'sri_cantones.csv'), encoding='utf-8', header=['empresas'], index_label='canton')
altas = a[a.TIPO_CONTRIBUYENTE == 'SOCIEDAD'].FECHA_INICIO_ACTIVIDADES.str[:4].value_counts().sort_index()
altas.loc['2012':'2026'].to_csv(os.path.join(salida, 'sri_altas.csv'), encoding='utf-8', header=['sociedades'], index_label='anio')
print('formales', len(formales), 'objetivo', len(objetivo), 'O,P,Q,U', formales.SEC.isin(['O', 'P', 'Q', 'U']).sum())
