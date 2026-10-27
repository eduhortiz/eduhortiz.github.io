---
name: html-a-app-segura
description: Convierte un HTML (incluido un "bundle" exportado con __bundler/manifest) en una app instalable y segura. Pasos - revisión de seguridad, backend Apps Script con login por token, publicación en GitHub Pages como PWA, APK Android con PWABuilder y assetlinks.json, e instalación en iPhone. Usar cuando el usuario pida "pasar el HTML a una aplicación", "crear una app/APK desde este HTML", "instalar en el celular/iOS" o revisar la seguridad de una app HTML + Google Sheets.
---

# HTML → app segura (PWA + APK + iOS)

Responder en español, conciso. Usuario de GitHub: `eduhortiz`. Sitio: `https://eduhortiz.github.io/` (repo `eduhortiz/eduhortiz.github.io`, ya tiene `.nojekyll` y `.well-known/assetlinks.json` con varias apps).

## 0. Preguntas iniciales (una sola vez, con AskUserQuestion)
- Método: **GitHub Pages + PWABuilder** (lo que prefiere el usuario) o repo privado + Actions.
- ¿Endurecer el backend? Recomendar "Sí, todo".
- Proponer el Package ID `com.eduhortiz.<nombreapp>` y avisar que no se puede cambiar después.

## 1. Desempaquetar y auditar
- Si el HTML es un bundle (`<script type="__bundler/manifest">`), ejecuta `python3 -I scripts/desempaquetar_bundle.py <html> <salida>`. Deja React y las fuentes como archivos locales, sin CDN.
- Busca y reporta antes de tocar nada:
  - Datos personales incrustados en el código (DNI, nombres, listas de personal).
  - Claves por defecto, por ejemplo `123456`.
  - IDs de hojas y URLs `script.google.com/.../exec`.
  - Lecturas `gviz/tq` (indican que la hoja es pública).
  - `fetch(...accion=...)` sin token.
  - Login validado solo en el cliente.
  - Acciones de borrado o edición sin autenticación.
  - Un `doGet` que devuelva usuarios o claves.
- Comandos útiles: `grep -n "fetch(\|localStorage\|clave\|script.google\|gviz\|docs.google" index.html` y extraer `doGet`/`doPost` del backend.

## 2. Backend Apps Script endurecido
Base: `assets/seguridad_backend.gs`. Es el bloque probado de la app de Uniformes/EPPs; adapta `datos_`, `ejecutar_` y `puede_` a las acciones de la nueva app.
- `doPost` es la única entrada: `{accion, token, data}`. `doGet` no devuelve datos.
- `login_`:
  - Hash SHA-256 con sal (`h1$sal$hash`).
  - Migra al vuelo las claves en texto plano.
  - Bloquea 15 minutos tras 5 intentos (CacheService).
  - Avisa `claveDebil`.
- `sesion_(token)`:
  - Token de 64 caracteres hexadecimales en CacheService, válido 6 horas.
  - Relee la hoja Usuarios en cada llamada: bloquear o cambiar el rol aplica al instante.
- `puede_(u, accion, d)`: permisos por rol en el servidor. Borrar, editar usuarios y ajustar stock quedan solo para el administrador.
- `publico_()`: nunca devuelve `clave`. `guardarUsuarios_` recibe la clave solo al crearla o cambiarla, y conserva el hash si llega vacía.
- `limpiar_()`: neutraliza fórmulas (`=`, `+`, `@`) en los textos entrantes.
- `cifrarClavesExistentes()`: se ejecuta una vez desde el editor.
- Escrituras con `LockService`. Lo que audita (por ejemplo "Entregado por") se toma de la sesión, no del cliente.
- En el repo, `SHEET_ID = 'PEGA_AQUI_EL_ID_DE_TU_HOJA'`. Al usuario se le entrega aparte una copia con el ID real (`Codigo_AppsScript.txt` vía SendUserFile).
- Verifica con un test en Node que simule `Utilities`, `CacheService`, `LockService`, `ContentService` y `SpreadsheetApp` mediante `vm`. Casos mínimos:
  - Sin token → `auth:false`.
  - Clave mala → error. Clave buena → token.
  - Migración a hash.
  - El operador no puede borrar ni editar usuarios.
  - Bloqueo del usuario → pierde la sesión.
  - Fuerza bruta → bloqueo.
  - Logout invalida el token.

## 3. Cliente
- Agrega un método `api(accion, data)`:
  - Hace un POST `text/plain;charset=utf-8` a la URL `/exec` (valida que sea `https://script.google.com/macros/s/.../exec`).
  - Envía el token guardado en `localStorage['…_session']`.
  - Si la respuesta trae `auth:false`, llama a `endSession()`.
- Sustituye todos los `fetch` con `?accion=` y los `mode:'no-cors'` por `api()`, y comprueba `r.ok` (antes se daba por sincronizado sin confirmar).
- Login con correo + clave escritos. No muestres una lista de usuarios antes de iniciar sesión.
- El usuario actual sale de la sesión que devuelve el servidor, no de buscarlo en la lista local.
- No cargues datos hasta que haya sesión. Logout: llama a `api('logout')` y borra del `localStorage` los datos sensibles (personal, entregas, usuarios).
- Quita del código: datos de personal, catálogo con costos, clave por defecto, ID de la hoja y lecturas `gviz` (la hoja pasa a ser privada).
- Lista de usuarios para el admin:
  - Las claves no se ven (`• • • • • •`).
  - Para enviarla por WhatsApp hay que generar antes una "Nueva clave"; queda solo en memoria.
- Carga el código del backend en la pestaña de configuración con `fetch('backend/Code.gs')`, no incrustado.
- Prueba con Playwright interceptando `https://script.google.com/**`:
  - 0 llamadas externas antes del login.
  - Mensaje de error con clave mala.
  - Login → panel. La recarga mantiene la sesión.
  - Logout limpia el storage.
  - Sin violaciones de CSP.

## 4. PWA y endurecimiento web
- En `<head>`:
  - CSP: `default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval'` (más los CDN imprescindibles); `connect-src 'self' https://script.google.com https://script.googleusercontent.com`; `object-src 'none'; base-uri 'self'; form-action 'none'`.
  - `referrer no-referrer`.
  - `<title>`, `theme-color`, el manifest y el registro del service worker.
- Usa `assets/manifest.webmanifest` (ajusta `id`, `start_url` y `scope` a `/<carpeta>/`) y `assets/sw.js`: red primero, solo el shell de la app, nunca respuestas del API. Sube `CACHE` en cada cambio.
- Íconos generados con Playwright (`setContent` + `screenshot`):
  - `icon-192`, `icon-512`, `maskable-512`.
  - `apple-touch-icon` de 180×180, cuadrado y opaco.
- iOS:
  - `apple-mobile-web-app-capable`, `mobile-web-app-capable`.
  - `apple-mobile-web-app-status-bar-style` = `default`.
  - `apple-mobile-web-app-title`.
  - CSS `@media (max-width:760px){input,select,textarea{font-size:16px!important}}` para evitar el zoom al escribir.
- Si un CDN está bloqueado por el proxy (por ejemplo `cdn.sheetjs.com`), cárgalo en runtime, agrégalo a `script-src` y anótalo como limitación.

## 5. Publicar en GitHub
- Carpeta `/<nombreapp>/` dentro de `eduhortiz.github.io`; la URL queda `https://eduhortiz.github.io/<nombreapp>/`.
- Incluye `README.md` en español: pasos del Apps Script, cerrar la hoja, cambiar claves, APK y tabla "Antes/Ahora" de seguridad.
- Commit en una rama → push → fast-forward a `main` (`git push origin rama:main`).
- Si el push da 403: pide reconectar GitHub y entrega un zip mientras tanto.
- Desde el contenedor no se puede abrir `github.io`. Para verificar, usa `https://digitalassetlinks.googleapis.com/v1/statements:list?source.web.site=https://eduhortiz.github.io&relation=delegate_permission/common.handle_all_urls`.

## 6. APK Android
- El usuario genera el APK en pwabuilder.com: Package for stores → Android, con el Package ID acordado. Pídele que guarde `signing.keystore` y `signing-key-info.txt`.
- Cuando te pase el `assetlinks.json`, **agrega** un bloque nuevo en `.well-known/assetlinks.json` sin borrar los existentes. Valida con `python3 -m json.tool`, sube a `main` y verifica con la API de Digital Asset Links.
- Instalación: el `.apk`, no el `.aab`. Permitir orígenes desconocidos. Si ya estaba instalado y se ve la barra del navegador, reinstalar.

## 7. iPhone
- Safari → Compartir → "Agregar a pantalla de inicio".
- App Store: solo si lo pide. Requiere Mac + Xcode + cuenta de Apple Developer (99 USD/año) + revisión de Apple; no se recomienda para uso interno.

## Orden de pasos para el usuario (siempre recordarlo)
1. Pegar el nuevo Apps Script → ejecutar `cifrarClavesExistentes` → Implementar → Gestionar implementaciones → Editar → **Nueva versión**. La URL no cambia; el acceso sigue en "Cualquier usuario".
2. Hoja → Compartir → **Restringido**.
3. Cambiar todas las claves: las anteriores se consideran expuestas. No compartir el HTML antiguo.
4. APK (PWABuilder + assetlinks) e iPhone (Agregar a pantalla de inicio).
5. En equipos compartidos, pulsar siempre **Salir**.
