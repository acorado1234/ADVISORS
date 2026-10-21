# Status de campañas · Advisors

Tablero donde cada advisor (Mafer, Adriana, Cintia) lleva el status de sus campañas desde su propia laptop. Todo se guarda al momento en **una sola hoja de Google**, que es la base de datos. Andrea ve a todas y los lunes se presenta desde el tablero.

```
 Laptop de Mafer ─┐
 Laptop de Adriana ┼─▶  tablero (GitHub Pages)  ─▶  Apps Script  ─▶  Hoja de Google
 Laptop de Cintia ─┘          web/                  apps-script/       (base de datos)
 Andrea / Presentación ─┘
```

- **Nadie escribe PIN.** Cada persona tiene su **link personal**: lo abre una vez en su laptop y queda guardada 6 meses. Solo ve y edita sus campañas (su pestaña de la hoja).
- **Es una página web, no hay nada que instalar ni actualizar a mano.** Los cambios de una aparecen en las laptops de las demás en menos de un minuto, y cada mejora que se sube a GitHub llega a todas al recargar. En Chrome se puede instalar como app (ícono ⊕ en la barra de direcciones).
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
4. Si el tablero no va a vivir en `https://acorado1234.github.io/ADVISORS/`, cambia `TABLERO_URL` al inicio de `Code.gs`.
5. Elige la función `configurar` y pulsa **Ejecutar**. Acepta los permisos. Esto:
   - crea la pestaña `_Usuarios` con un **link personal** para cada perfil (columna `Link`). El de Andrea también sale en el registro de ejecución;
   - crea `_Historial`;
   - protege las dos pestañas para que solo la dueña de la hoja las edite;
   - agrega a cada pestaña de advisor una columna oculta `_id`.
6. **Implementar → Nueva implementación → Aplicación web.**
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier persona**.
   - Copia la URL que termina en `/exec`. Es tu `API_URL`.

### 3. Conectar el tablero
1. En `web/config.js` pega `API_URL` y `SHEET_URL`, y haz commit.
2. En GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Cada push a `main` publica el tablero. La dirección queda en `https://<usuario>.github.io/<repo>/`.
4. Andrea abre su link (columna `Link` de `_Usuarios`), entra a **Equipo** y con **Copiar link** le manda a cada una el suyo por WhatsApp.

## Trabajar desde git

- **Tablero (`web/`):** edita, haz commit y push. Se publica solo en ~1 minuto.
- **Servidor (`apps-script/`):** dos opciones.
  - **Manual:** pega el `Code.gs` nuevo en el editor de Apps Script y luego **Implementar → Administrar implementaciones → Editar (lápiz) → Versión: Nueva versión → Implementar**. La URL `/exec` no cambia.
  - **Automático con clasp:**
    1. Una vez, en tu compu: `npm i -g @google/clasp@2.4.2`, `clasp login`, y activa la API en <https://script.google.com/home/usersettings>.
    2. Copia `.clasp.example.json` a `.clasp.json` y pon el ID del proyecto (Apps Script → ⚙ → *ID de la secuencia de comandos*).
    3. Desde tu compu: `clasp push` sube el código; `clasp deploy -i <DEPLOYMENT_ID>` publica la nueva versión.
    4. Para que GitHub lo haga solo en cada push: en **Settings → Secrets and variables → Actions** crea el secret `CLASPRC_JSON` (contenido de `~/.clasprc.json`) y las variables `SCRIPT_ID` y `DEPLOYMENT_ID`.

## Perfiles: crear, compartir y quitar

Desde el tablero, con el perfil de Andrea, en **Equipo**:

- **Crear perfil:** nombre y qué puede hacer.
  - *Advisor:* se le crea su propia pestaña en la hoja, copiando columnas, listas y casillas de la advisor que elijas, pero vacía.
  - *Solo ver y presentar:* ve a todas, no edita.
  - *Jefa:* ve a todas y administra el equipo.
  Al crear el perfil sale su link para copiarlo y mandarlo.
- **Copiar link:** para reenviárselo a alguien que cambió de laptop.
- **Nuevo link:** si un link se compartió por error. El anterior deja de servir en todas las laptops.
- **Quitar a alguien:** en la pestaña `_Usuarios`, cambia su `Rol` a `inactiva` (o borra la fila). Su pestaña con campañas se queda en la hoja.

Columnas de `_Usuarios`: `Perfil`, `Hoja`, `Rol` (`advisor`, `admin`, `presentacion`), `Llave`, `Link`. No edites `Llave` a mano.

## Un solo modelo para todas

Cada advisor arma su pestaña a su manera. El tablero las traduce a **un mismo modelo de campaña**, así todas ven los mismos campos en el mismo orden:

- Los nombres de columna se reconocen sin importar mayúsculas, tildes o espacios (`notas` = `Notas`), y hay sinónimos (`link de asana` = `Link Asana`). La lista está en `CAMPOS` y `ALIAS` en `Code.gs`.
- Si a una pestaña le falta un campo, la columna se crea al final la primera vez que alguien lo llena. Nunca se borra, mueve ni renombra una columna existente.
- Si una pestaña tiene dos columnas para lo mismo (Cintia: `notas` y `Notas`), se leen las dos y se escribe en la que se llama exactamente igual.
- Las listas desplegables son las mismas para todas (`LISTAS_STD`).
- **Por ordenar:** el tablero marca lo que falta o está escrito distinto (sin status, sin puntaje, valores fuera de la lista, fechas escritas como texto) y, cuando puede, propone el arreglo con un clic (por ejemplo, "24/06 al 1/08" → fecha de activación y de finalización). Nadie cambia sus datos sin que ella lo apruebe.

## Bitácora

Cada campaña tiene una bitácora: cada nota se suma arriba con fecha, hora y quién la escribió, sin borrar las anteriores. Se guarda en la columna `Notas` de la hoja, una nota por línea con el formato `[20/10/2026 14:30, Mafer] texto`, así se sigue leyendo bien desde el Excel. Las notas que ya existían quedan como "Antes del tablero". En la lista se ve la última nota y hace cuánto se escribió (en dorado después de 3 días, en rojo después de 7).

## Presentar

Al pulsar **Presentar** primero se eligen las campañas: vienen marcadas las que se movieron en los últimos 7 días y las que se eligieron la vez anterior, y se puede agregar o quitar cualquiera. La selección se guarda en el servidor, así la advisor la arma en su compu y Andrea la presenta desde la suya. Al final sale un resumen con las demás campañas activas.

## Probar sin tocar datos reales

Crea perfiles cuyo nombre empiece con `Prueba` desde **Equipo** y úsalos. Al terminar, en Apps Script ejecuta `borrarPruebas`: elimina esos perfiles, sus pestañas, las campañas que crearon (aunque las hayan pasado a otra advisor) y sus movimientos del historial.

## Cómo funciona

- Todas se conectan a la misma hoja. El tablero se actualiza solo cada 45 segundos y al volver a la pestaña.
- **Traspasos:** la campaña sale de la pestaña de una advisor y se agrega al final de la otra con el mismo ID, así que conserva su historial.
- **Seguridad:** el link personal es como el link de un Google Doc: quien lo tenga entra como esa persona, así que se manda solo a ella. Es una llave al azar de 40 caracteres que se valida en el servidor, imposible de adivinar. Sin link nadie ve nada, aunque el repo sea público.

## Cuidados

- No borres ni renombres la columna oculta `_id` ni las pestañas `_Usuarios` y `_Historial`.
- Si renombras una pestaña de advisor, cambia también su columna **Hoja** en `_Usuarios`.
- Las listas desplegables (Status, Propuesta, etc.) se leen de la validación de datos de la fila 2 de cada pestaña.
- Sin costo con Google ni con GitHub Pages para este uso.
