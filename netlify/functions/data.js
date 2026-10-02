// Lee el Excel de facturación de Palmera en SharePoint/OneDrive con Microsoft Graph
// (credenciales de aplicación) y devuelve la hoja FACTURA como texto,
// con las fechas en dd/mm/aaaa y los montos como números.
const crypto = require("crypto");
const XLSX = require("xlsx");

const DRIVE_ID = process.env.DRIVE_ID || "b!ftpx1PHMU02Xyd_BvRi5gZ4jKtzFMjZHt1vZTICVWd0sPgcGYMRfTZopqpx_5vTB";
// Se direcciona por ruta (no por id) para que siga funcionando si el archivo se vuelve a subir o se reemplaza.
// ITEM_ID es opcional y, si se define, tiene prioridad sobre la ruta.
const FILE_PATH = process.env.FILE_PATH || "MASSIEL/PALMERA/OC/Facturación_Palmera_SEP.xlsx";
const ITEM_ID = process.env.ITEM_ID || "";
const FILE_NAME = process.env.FILE_NAME || "Facturación_Palmera_SEP.xlsx";
const SHEETS = ["FACTURA"]; // por ahora solo se usa la hoja FACTURA

const json = (status, body, extra = {}) => ({
  statusCode: status,
  headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store", ...extra },
  body: JSON.stringify(body),
});

function sameKey(a, b) {
  const x = crypto.createHash("sha256").update(String(a)).digest();
  const y = crypto.createHash("sha256").update(String(b)).digest();
  return crypto.timingSafeEqual(x, y);
}

// Resume el error de Microsoft (código y mensaje) para poder diagnosticar sin ver los logs.
async function detail(r) {
  try {
    const j = await r.json();
    const e = j.error;
    if (e && typeof e === "object") return `${e.code || ""} ${e.message || ""}`.trim().slice(0, 300);
    return `${e || ""} ${j.error_description || ""}`.trim().replace(/\s+/g, " ").slice(0, 300);
  } catch (_) {
    return "sin detalle";
  }
}

async function getToken() {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });
  if (!r.ok) throw new Error(`token ${r.status}: ${await detail(r)}`);
  return (await r.json()).access_token;
}

// La API de libros de Excel de Graph (workbook) rechaza el token de aplicación en OneDrive
// ("Could not obtain a WAC access token"), así que se descarga el archivo y se lee aquí.
async function download(token) {
  const base = ITEM_ID
    ? `items/${ITEM_ID}`
    : `root:/${FILE_PATH.split("/").map(encodeURIComponent).join("/")}:`;
  const r = await fetch(`https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/${base}/content`, {
    headers: { authorization: `Bearer ${token}` },
    redirect: "follow",
  });
  if (!r.ok) throw new Error(`archivo ${r.status}: ${await detail(r)}`);
  return Buffer.from(await r.arrayBuffer());
}

const pad = (n) => String(n).padStart(2, "0");
// Fechas como dd/mm/aaaa (igual que se ven en Excel); el resto de celdas se envía con su valor.
function cell(v) {
  if (v instanceof Date) {
    const d = new Date(v.getTime() + 12 * 36e5);
    return `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  }
  return v;
}

function readSheet(wb, name) {
  const key = wb.SheetNames.find((n) => n.trim().toUpperCase() === name);
  if (!key) throw new Error(`no existe la hoja ${name}`);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[key], { header: 1, raw: true, defval: "" });
  return rows.map((r) => r.map(cell));
}

exports.handler = async (event) => {
  const need = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"].filter((k) => !process.env[k]);
  if (need.length) return json(500, { error: "Faltan variables de entorno: " + need.join(", ") });

  // ACCESS_KEY es opcional: si no está definida, la función responde sin pedir clave.
  if (process.env.ACCESS_KEY) {
    const key = event.headers["x-access-key"] || "";
    if (!sameKey(key, process.env.ACCESS_KEY)) return json(401, { error: "clave incorrecta" });
  }

  try {
    const token = await getToken();
    const wb = XLSX.read(await download(token), { type: "buffer", cellDates: true });
    const sheets = {};
    SHEETS.forEach((n) => (sheets[n] = readSheet(wb, n)));
    return json(200, { name: FILE_NAME, updated: new Date().toISOString(), sheets });
  } catch (e) {
    return json(502, { error: String(e.message || e) });
  }
};
