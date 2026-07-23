// api/auth/login.js
// Redirige al usuario a la pantalla de autorización de Microsoft.
// Solo se necesita usar ESTA vez, para autorizar la app manualmente y obtener el refresh token.
export default function handler(req, res) {
  const { MS_CLIENT_ID, MS_TENANT_ID, REDIRECT_URI } = process.env;

  if (!MS_CLIENT_ID || !MS_TENANT_ID || !REDIRECT_URI) {
    res.status(500).send('Faltan variables de entorno: MS_CLIENT_ID, MS_TENANT_ID o REDIRECT_URI');
    return;
  }

  const scope = encodeURIComponent('Files.Read offline_access User.Read');
  const redirect = encodeURIComponent(REDIRECT_URI);

  const url =
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${MS_CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${redirect}` +
    `&response_mode=query` +
    `&scope=${scope}`;

  res.writeHead(302, { Location: url });
  res.end();
}
