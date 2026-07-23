// api/data.js
// Cada vez que el dashboard carga, llama a este endpoint.
// Este endpoint: 1) renueva el access token con el refresh token guardado,
// 2) le pide a Microsoft Graph el contenido de la hoja de Excel,
// 3) aplica la MISMA lógica de "aplanado" que usaba tu botón "Subir Excel",
// 4) devuelve el arreglo de vehículos ya listo para pintar en el dashboard.

const MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE","INICIAL"];

function norm(s) {
  return (s == null ? '' : '' + s).replace(/\u00a0/g, ' ').trim();
}
function excelSerialToStr(n) {
  const date = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  const dd = String(date.getUTCDate()).padStart(2, '0');
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = date.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}
function normFecha(v) {
  if (typeof v === 'number') return excelSerialToStr(v);
  return norm(v);
}
function normEstado(v) {
  v = norm(v).toUpperCase();
  if (v.includes("DETENIDO")) return "ALISTAMIENTO DETENIDO";
  if (v.includes("ENTREGADO")) return "9. ENTREGADO A LOGISTICA";
  if (v.includes("EMBELLEC")) return "8. PROCESO EMBELLECIMIENTO";
  if (v.includes("LAVADO") && v.includes("ASIGNAD")) return "7. VEHICULO ASIGNADO EN LAVADO";
  if (v.includes("ESPERAR") || (v.includes("ASIGNAR") && v.includes("LAVADO"))) return "6. A ESPERAR PARA ASIGNAR A LAVADO";
  if (v.includes("INSTALAC")) return "5. INSTALACION ACCESORIOS";
  if (v.includes("MEC")) return "4. PROCESO MECANICA";
  if (v.includes("ANTICORROSIVO")) return "3. APLICACION ANTICORROSIVO";
  if (v.includes("ASIGNAD")) return "2. VEHICULO ASIGNADO";
  if (v.includes("INSPEC")) return "1. INSPECCION VEHICULO";
  return norm(v);
}

// Misma función "flatten" del dashboard original, adaptada para correr en el servidor
function flatten(A) {
  let hr = -1;
  for (let i = 0; i < Math.min(A.length, 15); i++) {
    const up = (A[i] || []).map(x => norm(x).toUpperCase());
    if (up.includes('OT') && up.some(c => c.includes('ESTADO'))) { hr = i; break; }
  }
  if (hr < 0) hr = 0;
  const head = A[hr].map(x => norm(x).toUpperCase());
  const col = (...names) => {
    for (const n of names) { const i = head.findIndex(h => h === n); if (i >= 0) return i; }
    for (const n of names) { const i = head.findIndex(h => h.includes(n)); if (i >= 0) return i; }
    return -1;
  };
  const cOT = col('OT'), cEst = col('ESTADO'), cInv = col('INVENTARIO'), cMod = col('MODELO'),
        cVit = col('VITRINA'), cCol = col('COLOR'), cCar = col('CARGO'), cSla = col('SLA'),
        cReg = col('FECHA REGISTRO', 'REGISTRO'), cEnt = col('FECHA EST. ENTREGA', 'ENTREGA');

  let mes = '';
  for (let i = 0; i < hr; i++) {
    const j = (A[i] || []).map(norm).join(' ').toUpperCase();
    const fm = MESES.find(m => j.includes(m));
    if (fm) mes = fm.charAt(0) + fm.slice(1).toLowerCase();
  }

  let out = [];
  for (let i = hr + 1; i < A.length; i++) {
    const row = A[i]; if (!row) continue;
    const joined = row.map(norm).join(' ').toUpperCase();
    const found = MESES.find(m => joined.includes(m));
    const inv = norm(row[cInv]);
    const ot = norm(row[cOT]);
    const esEncabezado = norm(row[cEst]).toUpperCase() === 'ESTADO' || norm(row[cInv]).toUpperCase() === 'INVENTARIO';
    const isData = inv !== '' && !esEncabezado;
    if (!isData) { if (found) mes = found.charAt(0) + found.slice(1).toLowerCase(); continue; }
    out.push({
      mes: mes || '(sin mes)', ot: ot, inv: norm(row[cInv]), modelo: norm(row[cMod]).toUpperCase(),
      vitrina: norm(row[cVit]).toUpperCase(), color: norm(row[cCol]).toUpperCase(),
      cargo: norm(row[cCar]).toUpperCase(), estado: normEstado(row[cEst]),
      sla: (cSla >= 0 ? norm(row[cSla]) : ''),
      freg: (cReg >= 0 ? normFecha(row[cReg]) : ''), fent: (cEnt >= 0 ? normFecha(row[cEnt]) : '')
       });
  }
  return out;
}

async function getAccessToken() {
  const { MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID, MS_REFRESH_TOKEN } = process.env;
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    grant_type: 'refresh_token',
    refresh_token: MS_REFRESH_TOKEN,
    scope: 'Files.Read offline_access User.Read',
  });
  const r = await fetch(`https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const data = await r.json();
  if (!r.ok) throw new Error('No se pudo renovar el token: ' + JSON.stringify(data));
  return data.access_token;
}

export default async function handler(req, res) {
  try {
    const token = await getAccessToken();
    const filePath = process.env.MS_FILE_PATH; // ej: "Master Kaizen/Sedes/Circunvalar/Alistamiento Bitacora.xlsx"
    const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');

    // 1) listar hojas del archivo para encontrar la que contiene "ALISTAMIENTO"
    const wsRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/workbook/worksheets`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const wsData = await wsRes.json();
    if (!wsRes.ok) throw new Error('Error listando hojas: ' + JSON.stringify(wsData));

    const sheet = wsData.value.find(s => s.name.toUpperCase().includes('ALISTAMIENTO')) || wsData.value[0];

    // 2) traer todos los valores usados de esa hoja
    const rangeRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/workbook/worksheets('${encodeURIComponent(sheet.name)}')/usedRange(valuesOnly=true)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const rangeData = await rangeRes.json();
    if (!rangeRes.ok) throw new Error('Error leyendo rango: ' + JSON.stringify(rangeData));

    const parsed = flatten(rangeData.values || []);

    res.setHeader('Cache-Control', 'no-store');
    res.status(200).json(parsed);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
}
