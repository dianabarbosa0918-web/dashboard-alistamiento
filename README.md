# Dashboard Alistamiento — conectado a OneDrive

Este proyecto reemplaza el botón "Subir Excel" de tu dashboard por una conexión
automática al archivo **Alistamiento Bitácora** en tu OneDrive corporativo,
usando Microsoft Graph.

## Estructura

- `public/index.html` → el dashboard (mismo diseño, ahora carga datos solo).
- `api/data.js` → trae los datos frescos del Excel cada vez que alguien abre el dashboard.
- `api/auth/login.js` y `api/auth/callback.js` → se usan **una sola vez** para autorizar la app y obtener el refresh token.

## Paso 1 — Subir esto a GitHub

1. Crea un repositorio nuevo en GitHub (puede ser privado).
2. Sube todos estos archivos y carpetas tal cual están.

## Paso 2 — Conectar a Vercel

1. Entra a [vercel.com](https://vercel.com) e inicia sesión (puedes usar tu cuenta de GitHub).
2. "Add New" → "Project" → selecciona el repositorio que acabas de subir.
3. Framework Preset: déjalo en **"Other"**.
4. **Antes de darle a Deploy**, ve a la sección "Environment Variables" y agrega:
   - `MS_CLIENT_ID` → `87db862e-72d3-4cd4-a0ab-73e69fc194ae`
   - `MS_TENANT_ID` → `bfb53646-d4ba-4509-aabe-3c0ab4d79291`
   - `MS_CLIENT_SECRET` → (el valor que copiaste en Azure, el que solo se muestra una vez)
   - `MS_FILE_PATH` → la ruta de tu archivo dentro de OneDrive, ej: `Master Kaizen/Sedes/Circunvalar/Alistamiento Bitacora.xlsx`
   - `REDIRECT_URI` → lo completas en el paso 3 (por ahora puedes poner un valor temporal, ej. `https://placeholder.vercel.app/api/auth/callback`)
5. Dale a **Deploy**.

## Paso 3 — Ajustar la URL de redirección

1. Una vez desplegado, Vercel te da una URL real, ej: `https://dashboard-alistamiento-xyz.vercel.app`
2. Ve a **Vercel → Settings → Environment Variables** y actualiza `REDIRECT_URI` a:
   `https://TU-URL-REAL.vercel.app/api/auth/callback`
3. Ve a **Azure Portal → tu app → Autenticación** y agrega esa misma URL como URI de redirección adicional (la de `localhost` puedes dejarla o borrarla, no importa).
4. Vuelve a Vercel y haz **Redeploy** (Deployments → los 3 puntos → Redeploy) para que tome el nuevo valor de `REDIRECT_URI`.

## Paso 4 — Autorizar la app (una sola vez)

1. Entra en tu navegador a: `https://TU-URL-REAL.vercel.app/api/auth/login`
2. Inicia sesión con tu cuenta de Dinissan si te lo pide.
3. Verás una pantalla de "Esta app quiere acceder a tus archivos" → Aceptar.
4. Te va a redirigir a una página que muestra un **refresh token** largo. Cópialo completo.
5. Ve a **Vercel → Settings → Environment Variables**, agrega `MS_REFRESH_TOKEN` con ese valor.
6. Haz **Redeploy** una vez más.

## Paso 5 — Probar

Abre `https://TU-URL-REAL.vercel.app` — el dashboard debería cargar los datos automáticamente,
sin que tengas que subir el Excel manualmente. Cada vez que edites y guardes el Excel en OneDrive,
solo con recargar la página (o dando clic en "🔄 Actualizar ahora") vas a ver los datos más recientes.

## Notas importantes

- **`MS_FILE_PATH` debe ser exacto.** Si tiene tildes o espacios, tal cual como aparece en OneDrive.
- Si cambias el nombre de la hoja de Excel a algo distinto a "Alistamiento" en el nombre de la pestaña, revisa que siga conteniendo esa palabra, o edita `api/data.js` para que busque el nombre correcto.
- El refresh token normalmente dura activo mientras se siga usando. Si en algún momento el dashboard empieza a fallar con error de token, repite el Paso 4 para generar uno nuevo.
- Nunca compartas tu `MS_CLIENT_SECRET` ni tu `MS_REFRESH_TOKEN` fuera de las variables de entorno de Vercel.
