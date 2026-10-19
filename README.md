# Control del Jefe de Ronda – ESCONBOL

Sistema para que el administrador asigne cada semana a los dos responsables de ronda y cargue el santo y seña de cada día. Cada responsable llena desde su celular la Lista de Verificación del Jefe de Ronda, con 4 fotografías y su firma. El sistema genera el PDF con el formato oficial. Solo el administrador puede descargar el PDF. El Cmdte y el 2do Cmdte reciben únicamente la confirmación de que se pasó ronda.

## Partes

| Carpeta | Qué es | Dónde corre |
|---|---|---|
| `docs/` | Página web: `index.html` (formulario del Jefe de Ronda) y `admin.html` (panel del administrador) | GitHub Pages |
| `apps-script/` | Base de datos, fotos, PDF y envío diario de las 08:00 | Google Apps Script + Google Sheets + Google Drive |

El santo y seña, los celulares, las fotos y los PDF se guardan **solo en Google**, nunca en este repositorio.

## Instalación (una sola vez)

### 1. Google (base de datos)
1. Con la cuenta de Gmail del sistema, crea una hoja de cálculo nueva en Google Sheets llamada **Control Jefe de Ronda**.
2. Menú **Extensiones > Apps Script**.
3. Copia `apps-script/Code.gs` en el archivo `Código.gs`. Crea un archivo HTML llamado **Plantilla** y pega `apps-script/Plantilla.html`.
4. En **Configuración del proyecto** (engranaje), marca "Mostrar el archivo de manifiesto appsscript.json" y pega `apps-script/appsscript.json`.
5. En **Configuración del proyecto > Propiedades del script** agrega:
   - `ADMIN_PASSWORD`: la contraseña del administrador.
   - `WEB_URL`: la dirección de GitHub Pages (paso 2), por ejemplo `https://anave300989.github.io/jefe-de-ronda-nido/`.
6. En el editor elige la función **instalar** y presiona **Ejecutar**. Acepta los permisos. Así se crean las hojas, la carpeta en Drive y el envío diario de las 08:00.
7. **Implementar > Nueva implementación > Aplicación web**. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**. Copia la URL que termina en `/exec`.
8. (Opcional) Ejecuta **probarPdf** para ver un PDF de prueba en la carpeta de Drive.

### 2. GitHub Pages (página web)
1. Pega la URL `/exec` en `docs/config.js`.
2. En el repositorio: **Settings > Pages > Branch: main / carpeta /docs > Save**.
3. Panel del administrador: `https://<usuario>.github.io/jefe-de-ronda-nido/admin.html`.

## Uso semanal (administrador)
1. Entra a `admin.html` con la contraseña.
2. Carga los dos responsables (grado y nombre, celular y PIN de 4 dígitos) y el santo y seña de los 7 días. Presiona **Guardar semana**.
3. Carga una vez los celulares del Cmdte y del 2do Cmdte.

## Cada día
- 08:00: el sistema prepara un enlace personal para cada responsable con el santo y seña del día.
  - Con la API oficial de WhatsApp configurada (`WA_TOKEN` y `WA_PHONE_ID` en Propiedades del script), el mensaje se envía solo.
  - Sin la API, el administrador presiona **WhatsApp** junto a cada responsable en el panel y se abre WhatsApp con el mensaje listo.
- El Jefe de Ronda abre el enlace, ingresa su PIN, marca SI/NO en los 40 puntos, toma las 4 fotos, firma y envía.
- Se genera el PDF (lista + anexo fotográfico) en Google Drive. Solo el administrador lo descarga desde el panel.
- El Cmdte y el 2do Cmdte reciben la confirmación de que se pasó ronda (sin el PDF).
