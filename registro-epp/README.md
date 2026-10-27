# Registro de Uniformes y EPPs

App web (PWA) para registrar ingresos, entregas y stock de uniformes y EPPs. Los datos viven en una hoja de Google Sheets y se leen y escriben a través de un Apps Script.

- App: `https://eduhortiz.github.io/registro-epp/`
- Servidor: `backend/Code.gs` (se pega en Apps Script)

## 1. Actualizar el Apps Script (obligatorio)

La app nueva **no funciona con el script anterior**, porque ahora todo pasa por un login validado en el servidor.

1. Abre tu proyecto de Apps Script y reemplaza todo el código por el de `backend/Code.gs`.
2. En la línea 2 cambia `PEGA_AQUI_EL_ID_DE_TU_HOJA` por el ID de tu hoja: es la parte de la URL que va entre `/d/` y `/edit`.
3. Guarda. En el menú de funciones elige **`cifrarClavesExistentes`** y pulsa **Ejecutar** una sola vez. Así las claves de la pestaña "Usuarios" quedan cifradas.
4. Ve a **Implementar → Gestionar implementaciones → ✏️ Editar → Versión: Nueva versión → Implementar**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**. Esto es necesario para que la app pueda conectarse, y ahora es seguro porque sin sesión el servidor no entrega datos.
   - Al editar la implementación existente, la URL `/exec` no cambia.

## 2. Cerrar la hoja de Google

Antes la app leía la hoja directamente, y por eso tenía que estar compartida con "cualquiera con el enlace". Ya no hace falta:

**Compartir → Acceso general → Restringido.** Deja solo a las personas que de verdad la editan.

## 3. Cambiar la clave del administrador

La clave `123456` estaba escrita en el código de la versión anterior, así que hay que considerarla pública. Entra a la app y, en **Usuarios**, pulsa **Nueva clave** para tu propio usuario y para cualquier otro que use una clave fácil.

## 4. APK para Android

1. Con la app publicada, entra a <https://www.pwabuilder.com>, pega `https://eduhortiz.github.io/registro-epp/` y elige **Package for stores → Android**.
2. Usa el Package ID `com.eduhortiz.registroepp` (no se puede cambiar después) y descarga el zip.
3. Guarda `signing.keystore` y `signing-key-info.txt` en un lugar seguro. Sin ellos no podrás publicar actualizaciones.
4. Para que la app abra sin la barra del navegador, agrega en `/.well-known/assetlinks.json`, en la raíz de este repositorio, un segundo bloque con el `package_name` y la huella `sha256_cert_fingerprints` que trae el zip (junto al bloque de carnet-sanidad).
5. Instala el `.apk` en el celular (hay que permitir "instalar apps de origen desconocido").

## Seguridad: qué cambió

| Antes | Ahora |
|---|---|
| `…/exec?accion=usuarios` devolvía todos los correos y claves a cualquiera | El servidor nunca devuelve claves. Se guardan cifradas (SHA-256 con sal) |
| El login se comprobaba solo en el celular | El login se valida en el servidor, que entrega un token de sesión de 6 horas |
| Cualquiera con la URL podía borrar entregas, cambiar stock o editar usuarios | Cada acción exige un token válido y el rol adecuado (por ejemplo, borrar es solo para el administrador) |
| El código traía 126 colaboradores con DNI, más el catálogo con costos | Esos datos solo llegan tras iniciar sesión y se borran del equipo al pulsar "Salir" |
| Clave de administrador `123456` escrita en el código | Eliminada. El servidor avisa si la clave es débil |
| Intentos de clave ilimitados | 5 intentos y luego 15 minutos de bloqueo por correo |
| La hoja tenía que ser pública (lectura `gviz`) | La hoja puede ser privada |
| React y las fuentes se cargaban desde CDNs | React y las fuentes vienen con la app. CSP restringe a qué servidores se conecta |
| Texto como `=HYPERLINK(…)` se guardaba como fórmula | El servidor neutraliza fórmulas en los textos recibidos |
| Bloquear a un usuario no cerraba su sesión | Bloquear o cambiar el rol aplica al instante en la siguiente acción |

### Limitaciones que siguen

- La clave sigue siendo de 6 números. El límite de intentos frena los ataques, pero una clave más larga sería más fuerte.
- La exportación a Excel carga SheetJS desde `cdn.sheetjs.com`.
- El token se guarda en el navegador del equipo. En un equipo compartido, pulsa siempre **Salir**.
