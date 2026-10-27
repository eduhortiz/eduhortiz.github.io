#!/usr/bin/env python3
"""Convierte un HTML "bundle" (exportado con __bundler/manifest) en archivos normales.

Uso: python3 -I desempaquetar_bundle.py <bundle.html> <carpeta_salida>

Crea en la carpeta de salida:
  index.html  plantilla con rutas locales
  js/         runtime y librerías externas (React, etc.) que venían dentro del bundle
  fonts/      fuentes woff2
Así la app no depende de CDNs y puede publicarse en GitHub Pages.
"""
import base64, gzip, json, os, re, sys

src_path, out = sys.argv[1], sys.argv[2]
src = open(src_path, encoding='utf-8').read()

def bloque(tipo):
    m = re.search(r'<script type="__bundler/' + tipo + r'">(.*?)</script>', src, re.S)
    return m.group(1) if m else None

manifest = json.loads(bloque('manifest'))
template = json.loads(bloque('template'))
ext = json.loads(bloque('ext_resources') or '[]')  # [{id: url_cdn, uuid}]
ext_por_uuid = {e['uuid']: e['id'] for e in ext}

os.makedirs(os.path.join(out, 'js'), exist_ok=True)
os.makedirs(os.path.join(out, 'fonts'), exist_ok=True)
recursos = {}
for uuid, e in manifest.items():
    data = base64.b64decode(e['data'])
    if e.get('compressed'):
        data = gzip.decompress(data)
    mime = e.get('mime', '')
    if mime.startswith('font/') or data[:4] == b'wOF2':
        ruta = 'fonts/' + uuid + '.woff2'
    elif 'javascript' in mime:
        nombre = os.path.basename(ext_por_uuid[uuid]) if uuid in ext_por_uuid else uuid + '.js'
        ruta = 'js/' + nombre
    else:
        ruta = 'assets/' + uuid
        os.makedirs(os.path.join(out, 'assets'), exist_ok=True)
    open(os.path.join(out, ruta), 'wb').write(data)
    if uuid in ext_por_uuid:
        recursos[ext_por_uuid[uuid]] = ruta  # el runtime busca window.__resources[url_cdn]
    template = template.replace('"' + uuid + '"', '"' + ruta + '"')
    print(f'{ruta}  ({mime}, {len(data)} bytes)')

template = re.sub(r'\s+integrity="[^"]*"', '', template)
template = re.sub(r'\s+crossorigin="[^"]*"', '', template)
if recursos:
    tag = '<script>window.__resources=' + json.dumps(recursos) + ';</script>\n'
    m = re.search(r'<head[^>]*>', template, re.I)
    template = template[:m.end()] + '\n' + tag + template[m.end():]
open(os.path.join(out, 'index.html'), 'w', encoding='utf-8').write(template)
print('index.html listo. Revisa con grep si quedan URLs de CDN: grep -o "https://[^\\"\' ]*" index.html | sort -u')
