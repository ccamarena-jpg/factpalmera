// Lee el Excel de facturación de Palmera en SharePoint/OneDrive con Microsoft Graph
// (credenciales de aplicación) y devuelve las hojas CONSOLIDADO y FACTURA como texto,
// con el mismo formato que muestra Excel (fechas dd/mm/aaaa, montos con coma decimal).
const crypto = require("crypto");

const DRIVE_ID = process.env.DRIVE_ID || "b!ftpx1PHMU02Xyd_BvRi5gZ4jKtzFMjZHt1vZTICVWd0sPgcGYMRfTZopqpx_5vTB";
const ITEM_ID = process.env.ITEM_ID || "015HP6ULO5SPMQ35H4ZVCKWQ2QXIQSSAVA";
const FILE_NAME = process.env.FILE_NAME || "Facturación_Palmera_SEP.xlsx";
const SHEETS = ["CONSOLIDADO", "FACTURA"];

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

async function readSheet(token, name) {
  const url = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${ITEM_ID}/workbook/worksheets('${encodeURIComponent(name)}')/usedRange?$select=text`;
  const r = await fetch(url, { headers: { authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`hoja ${name} ${r.status}: ${await detail(r)}`);
  return (await r.json()).text;
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
    const parts = await Promise.all(SHEETS.map((n) => readSheet(token, n)));
    const sheets = {};
    SHEETS.forEach((n, i) => (sheets[n] = parts[i]));
    return json(200, { name: FILE_NAME, updated: new Date().toISOString(), sheets });
  } catch (e) {
    return json(502, { error: String(e.message || e) });
  }
};
