# Status de campañas · Advisors

Tablero donde cada advisor (Mafer, Adriana, Cintia) lleva el status de sus campañas desde su propia laptop. Todo se guarda al momento en **una sola hoja de Google**, que es la base de datos. Andrea ve a todas y los lunes se presenta desde el tablero.

```
 Laptop de Mafer ─┐
 Laptop de Adriana ┼─▶  tablero (GitHub Pages)  ─▶  Apps Script  ─▶  Hoja de Google
 Laptop de Cintia ─┘          web/                  apps-script/       (base de datos)
 Andrea / Presentación ─┘
```

- **Cada advisor entra con su perfil y PIN.** Solo ve y edita sus campañas (su pestaña de la hoja).
- **Historial:** cada cambio queda en la pestaña `_Historial` con fecha, hora y quién lo hizo.
- **Puntaje:** A1 es la campaña más sencilla y C3 la más complicada.
- **En este repo no hay datos de clientes ni PIN.** El `.gitignore` bloquea los Excel y CSV.

## Estructura

| Ruta | Qué es |
|---|---|
| `web/index.html` | El tablero. Se publica solo en GitHub Pages. |
| `web/config.js` | La URL del servidor (`API_URL`) y de la hoja (`SHEET_URL`). |
| `apps-script/Code.gs` | El servidor. Vive dentro de la hoja de Google (Apps Script). |
| `apps-script/appsscript.json` | Zona horaria (Guatemala), permisos y modo de aplicación web. |
| `.github/workflows/pages.yml` | Publica `web/` en GitHub Pages en cada cambio. |
| `.github/workflows/apps-script.yml` | Opcional: publica `apps-script/` en la hoja en cada cambio. |

## Puesta en marcha (una sola vez)

### 1. Convertir el Excel en hoja de Google
1. Entra a Google Drive con la cuenta que va a ser dueña de la base de datos (lo ideal es la de Andrea o una cuenta de la empresa, no una personal).
2. **Nuevo → Subir archivo** y sube `Advisors - Status de campañas.xlsx`.
3. Ábrelo y elige **Archivo → Guardar como hoja de cálculo de Google**. Desde ahí se trabaja solo con esa versión; el Excel ya no se usa.
4. Copia la dirección de la hoja (termina en `/edit`). Es tu `SHEET_URL`.

> Las pestañas deben llamarse igual que en la columna **Hoja** de `_Usuarios` (`Mafer`, `Adriana`, `Cintia`). Las demás pestañas (`Cami`, `Manuel`, `Campañas firmados`) se quedan en la hoja pero el tablero no las toca.

### 2. Instalar el servidor
1. En la hoja: **Extensiones → Apps Script**.
2. Borra lo que haya, pega `apps-script/Code.gs` y guarda.
3. En **Configuración del proyecto (⚙)** marca *Mostrar el archivo de manifiesto "appsscript.json"* y pega el contenido de `apps-script/appsscript.json`.
4. Elige la función `configurar` y pulsa **Ejecutar**. Acepta los permisos. Esto:
   - crea la pestaña `_Usuarios` con un **PIN al azar** para cada perfil (míralos ahí y pásaselos a cada una);
   - crea `_Historial`;
   - protege las dos pestañas para que solo la dueña de la hoja las edite;
   - agrega a cada pestaña de advisor una columna oculta `_id`.
5. **Implementar → Nueva implementación → Aplicación web.**
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier persona**.
   - Copia la URL que termina en `/exec`. Es tu `API_URL`.

### 3. Conectar el tablero
1. En `web/config.js` pega `API_URL` y `SHEET_URL`, y haz commit.
2. En GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Cada push a `main` publica el tablero. La dirección queda en `https://<usuario>.github.io/<repo>/`.
4. Comparte esa dirección con Mafer, Adriana, Cintia y Andrea.

## Trabajar desde git

- **Tablero (`web/`):** edita, haz commit y push. Se publica solo en ~1 minuto.
- **Servidor (`apps-script/`):** dos opciones.
  - **Manual:** pega el `Code.gs` nuevo en el editor de Apps Script y luego **Implementar → Administrar implementaciones → Editar (lápiz) → Versión: Nueva versión → Implementar**. La URL `/exec` no cambia.
  - **Automático con clasp:**
    1. Una vez, en tu compu: `npm i -g @google/clasp@2.4.2`, `clasp login`, y activa la API en <https://script.google.com/home/usersettings>.
    2. Copia `.clasp.example.json` a `.clasp.json` y pon el ID del proyecto (Apps Script → ⚙ → *ID de la secuencia de comandos*).
    3. Desde tu compu: `clasp push` sube el código; `clasp deploy -i <DEPLOYMENT_ID>` publica la nueva versión.
    4. Para que GitHub lo haga solo en cada push: en **Settings → Secrets and variables → Actions** crea el secret `CLASPRC_JSON` (contenido de `~/.clasprc.json`) y las variables `SCRIPT_ID` y `DEPLOYMENT_ID`.

## Agregar o quitar una advisor

No se toca el código. En la pestaña `_Usuarios` agrega una fila: `Perfil`, `PIN`, `Hoja` (nombre exacto de su pestaña) y `Rol` = `advisor`. Aparece en el tablero al recargar. Para quitarla, borra su fila.

Roles: `advisor` (edita su pestaña), `admin` (ve a todas, descarga Excel), `presentacion` (solo ve y presenta).

## Cómo funciona

- Todas se conectan a la misma hoja. El tablero se actualiza solo cada 45 segundos y al volver a la pestaña.
- **Traspasos:** la campaña sale de la pestaña de una advisor y se agrega al final de la otra con el mismo ID, así que conserva su historial.
- **Seguridad:** el PIN se valida en el servidor, se bloquea 5 minutos tras 5 intentos fallidos y la sesión vence a las 12 horas.

## Cuidados

- No borres ni renombres la columna oculta `_id` ni las pestañas `_Usuarios` y `_Historial`.
- Si renombras una pestaña de advisor, cambia también su columna **Hoja** en `_Usuarios`.
- Las listas desplegables (Status, Propuesta, etc.) se leen de la validación de datos de la fila 2 de cada pestaña.
- Sin costo con Google ni con GitHub Pages para este uso.
