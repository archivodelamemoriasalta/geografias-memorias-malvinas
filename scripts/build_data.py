#!/usr/bin/env python3
"""Convierte datos/veteranos.xlsx -> data/veteranos.json (lo que lee el sitio).

Uso local:   python scripts/build_data.py
En GitHub:   lo ejecuta .github/workflows/build-data.yml al subir el Excel.
Si hay errores de validación, termina con código 1 y NO escribe el JSON.
"""
import json, sys, datetime as dt
from pathlib import Path
from openpyxl import load_workbook

XLSX = Path('datos/veteranos.xlsx')
OUT  = Path('data/veteranos.json')

# ── Política de privacidad (ajustar según lo que decida el Archivo) ──────────
FECHA_NAC_VIVOS = 'anio'   # 'completa' | 'anio' | 'oculta'   (fallecidos: siempre completa)
LOCALIDAD_VIVOS = True     # False = para vivos solo se publica departamento / provincia

FUERZAS = {'Armada', 'Ejército', 'Fuerza Aérea', 'Otro'}
DEPTOS = {"Anta","Cachi","Cafayate","Capital","Cerrillos","Chicoana","General Güemes","Guachipas","Iruya","La Caldera",
          "La Candelaria","La Poma","La Viña","Los Andes","Metán","Molinos","Orán","Rivadavia","Rosario de la Frontera",
          "Rosario de Lerma","San Carlos","San Martín","Santa Victoria"}

def s(x): return ' '.join(str(x).split()) if x not in (None, '') else ''
def fecha(x):
    if isinstance(x, (dt.datetime, dt.date)): return x.date() if isinstance(x, dt.datetime) else x
    return None
def fmt(d): return d.strftime('%d/%m/%Y') if d else ''

wb = load_workbook(XLSX, data_only=True, read_only=True)
ws = wb['Veteranos']
it = ws.iter_rows(values_only=True)
head = [s(h) for h in next(it)]
col = {h: i for i, h in enumerate(head)}
need = ['nro','apellido','nombre','fuerza','unidad','estado','fecha_nacimiento','lugar_nacimiento','fecha_fallecimiento',
        'provincia_residencia','localidad_residencia','departamento','publicar']
falta = [c for c in need if c not in col]
if falta: sys.exit(f'ERROR: faltan columnas en el Excel: {falta}')

errores, avisos, out, vistos = [], [], [], set()
for n, row in enumerate(it, start=2):
    g = lambda c: row[col[c]] if col[c] < len(row) else None
    ap, no = s(g('apellido')), s(g('nombre'))
    if not ap: continue                       # fila vacía
    tag = f'fila {n} ({ap}, {no})'
    if s(g('publicar')).upper() == 'NO': continue
    nro = g('nro')
    if nro in (None, ''): avisos.append(f'{tag}: sin número')
    elif nro in vistos: errores.append(f'{tag}: número {nro} repetido')
    vistos.add(nro)
    fz = s(g('fuerza'))
    if fz not in FUERZAS: errores.append(f'{tag}: fuerza inválida "{fz}"')
    est = s(g('estado'))
    if est not in ('Vivo', 'Fallecido'): errores.append(f'{tag}: estado inválido "{est}"')
    fn, ff = fecha(g('fecha_nacimiento')), fecha(g('fecha_fallecimiento'))
    if g('fecha_nacimiento') and not fn: errores.append(f'{tag}: fecha_nacimiento no es una fecha de Excel')
    if g('fecha_fallecimiento') and not ff: errores.append(f'{tag}: fecha_fallecimiento no es una fecha de Excel')
    if fn and ff and ff < fn: errores.append(f'{tag}: fallece antes de nacer')
    if est == 'Fallecido' and not ff: avisos.append(f'{tag}: fallecido sin fecha')
    prov, loc, dep = s(g('provincia_residencia')), s(g('localidad_residencia')), s(g('departamento'))
    if dep and dep not in DEPTOS: errores.append(f'{tag}: departamento inválido "{dep}"')
    if prov == 'Salta' and not dep: avisos.append(f'{tag}: reside en Salta sin departamento (irá a la ventana aparte)')
    if dep and prov != 'Salta': errores.append(f'{tag}: tiene departamento pero la provincia no es Salta')
    vivo = est == 'Vivo'
    if vivo:
        fnac = {'completa': fmt(fn), 'anio': str(fn.year) if fn else '', 'oculta': ''}[FECHA_NAC_VIVOS]
    else:
        fnac = fmt(fn)
    res_partes = [] if (vivo and not LOCALIDAD_VIVOS) else ([loc] if loc else [])
    if prov and prov != 'Salta': res_partes.append(prov)
    if prov == 'Salta' and not res_partes: res_partes.append('Salta')
    out.append({
        'nro': str(nro) if nro not in (None, '') else '',
        'nombre': f'{ap}, {no}' if no else ap,
        'apellido': ap,
        'fuerza': fz, 'unidad': s(g('unidad')),
        'fallecido': not vivo,
        'fechaNac': fnac, 'birthplace': s(g('lugar_nacimiento')),
        'fechaFallec': fmt(ff),
        'residencia': ', '.join(res_partes),
        'depto': dep, 'enMapa': bool(prov == 'Salta' and dep),
    })

for a in avisos: print('AVISO ', a)
if errores:
    for e in errores: print('ERROR ', e)
    sys.exit(f'\n{len(errores)} error(es): no se generó {OUT}')
OUT.parent.mkdir(exist_ok=True)
OUT.write_text(json.dumps(out, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
print(f'OK: {len(out)} registros -> {OUT}  ({sum(v["enMapa"] for v in out)} en mapa, {sum(v["fallecido"] for v in out)} fallecidos, {len(avisos)} avisos)')
