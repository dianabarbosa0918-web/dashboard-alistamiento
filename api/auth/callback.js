// api/auth/callback.js
// Microsoft redirige aquí después de que el usuario autoriza la app.
// Intercambiamos el "code" por un access_token + refresh_token.
// El refresh_token es el que hay que copiar UNA VEZ y guardar en Vercel como MS_REFRESH_TOKEN.
export default async function handler(req, res) {
  const { code, error, error_description } = req.query;

  if (error) {
    res.status(400).send(`<pre>Error de autorización: ${error}\n${error_description || ''}</pre>`);
    return;
  }
  if (!code) {
    res.status(400).send('Falta el parámetro "code" en la URL.');
    return;
  }

  const { MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID, REDIRECT_URI } = process.env;

  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    grant_type: 'authorization_code',
    code: String(code),
    redirect_uri: REDIRECT_URI,
    scope: 'Files.Read offline_access User.Read',
  });

  try {
    const r = await fetch(
      `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body,
      }
    );
    const data = await r.json();

    if (!r.ok) {
      res.status(500).send(`<pre>${JSON.stringify(data, null, 2)}</pre>`);
      return;
    }

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(`
      <html><body style="font-family:sans-serif;max-width:700px;margin:40px auto">
        <h2>✅ Autorización exitosa</h2>
        <p>Copia este <b>refresh token</b> completo y guárdalo en Vercel:</p>
        <p><b>Settings → Environment Variables → MS_REFRESH_TOKEN</b></p>
        <textarea style="width:100%;height:150px;font-size:12px">${data.refresh_token}</textarea>
        <p>Después de guardarlo, haz un <b>redeploy</b> del proyecto en Vercel para que tome el nuevo valor.</p>
        <p style="color:#888;font-size:12px">Puedes cerrar esta pestaña después de copiar el valor. No vuelvas a compartir este token con nadie.</p>
      </body></html>
    `);
  } catch (err) {
    res.status(500).send(`<pre>${String(err)}</pre>`);
  }
}
