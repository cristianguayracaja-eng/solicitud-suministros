// Capa de almacenamiento: reemplaza window.storage (Claude) por un backend
// propio hecho con Google Apps Script + Google Sheets.
//
// Configura la URL de tu Web App de Google Apps Script en un archivo .env
// en la raíz del proyecto:
//   VITE_API_URL=https://script.google.com/macros/s/XXXXXXXX/exec

const API_URL = import.meta.env.VITE_API_URL;

if (!API_URL) {
  // eslint-disable-next-line no-console
  console.warn(
    "VITE_API_URL no está configurada. Crea un archivo .env con VITE_API_URL=<tu URL de Apps Script>."
  );
}

async function apiGet(key) {
  const res = await fetch(`${API_URL}?action=get&key=${encodeURIComponent(key)}`);
  if (!res.ok) throw new Error("get failed: HTTP " + res.status);
  const data = await res.json();
  // Apps Script a veces responde 200 con un objeto de error: eso NO es "clave vacía".
  if (!data || typeof data !== "object" || data.error) throw new Error("get failed: " + (data && data.error));
  return data.value ?? null; // string | null
}

async function apiSet(key, value) {
  const res = await fetch(API_URL, {
    method: "POST",
    // text/plain evita el preflight CORS que Apps Script no maneja bien
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "set", key, value }),
  });
  if (!res.ok) throw new Error("set failed");
  const data = await res.json().catch(() => null);
  if (!data || data.error) throw new Error("set failed: " + (data && data.error));
  return true;
}

async function apiList(prefix) {
  const res = await fetch(`${API_URL}?action=list&prefix=${encodeURIComponent(prefix || "")}`);
  if (!res.ok) throw new Error("list failed");
  const data = await res.json();
  return data.keys || [];
}

async function apiDelete(key) {
  const res = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "delete", key }),
  });
  if (!res.ok) throw new Error("delete failed");
  const data = await res.json().catch(() => null);
  if (!data || data.error) throw new Error("delete failed: " + (data && data.error));
  return true;
}

// Estas cuatro funciones tienen la MISMA firma que las que usaba el artifact
// de Claude (safeGet/safeSet/safeList/safeDelete), incluido el parámetro
// "shared" que aquí se ignora porque ya no existe el concepto de
// almacenamiento personal vs. compartido (no hay cuentas de usuario).
export async function safeGet(key, _shared) {
  try {
    return await apiGet(key);
  } catch (e) {
    return null;
  }
}
export async function safeSet(key, value, _shared) {
  for (let i = 0; i < 3; i++) {
    try {
      await apiSet(key, value);
      return true;
    } catch (e) {
      await new Promise((r) => setTimeout(r, 400 * (i + 1)));
    }
  }
  return false;
}
export async function safeList(prefix, _shared) {
  try {
    return await apiList(prefix);
  } catch (e) {
    return [];
  }
}
export async function safeDelete(key, _shared) {
  try {
    await apiDelete(key);
    return true;
  } catch (e) {
    return false;
  }
}

// Lectura ESTRICTA: reintenta con espera creciente y distingue
//   { ok: true,  value: string }  -> la clave existe
//   { ok: true,  value: null }    -> la clave de verdad NO existe (vacía)
//   { ok: false, error }          -> no se pudo leer (conexión, 403/500, etc.)
// Úsala siempre que vayas a decidir "sembrar valores por defecto" o
// a hacer leer-modificar-escribir. Nunca sobrescribas nada si ok === false.
export async function safeGetStrict(key, attempts = 4) {
  let lastError;
  for (let i = 0; i < attempts; i++) {
    try {
      const value = await apiGet(key);
      return { ok: true, value };
    } catch (e) {
      lastError = e;
      await new Promise((r) => setTimeout(r, 400 * (i + 1) + Math.random() * 300));
    }
  }
  return { ok: false, error: lastError };
}

const wait_ = (ms) => new Promise((r) => setTimeout(r, ms));

async function withRetry_(fn, attempts) {
  let lastError;
  for (let i = 0; i < attempts; i++) {
    try { return { ok: true, ...(await fn()) }; }
    catch (e) { lastError = e; await wait_(400 * (i + 1) + Math.random() * 300); }
  }
  return { ok: false, error: lastError };
}

async function getJson_(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status);
  const data = await res.json();
  if (!data || typeof data !== "object" || data.error) throw new Error(data && data.error);
  return data;
}

// Varias claves en UNA sola petición (carga inicial). -> { ok, values: {clave: string|null} }
export async function safeGetManyStrict(keys, attempts = 4) {
  return withRetry_(async () => {
    const data = await getJson_(`${API_URL}?action=getMany&keys=${encodeURIComponent(JSON.stringify(keys))}`);
    if (!data.values) throw new Error("respuesta incompleta");
    return { values: data.values };
  }, attempts);
}

// Todos los registros con ese prefijo en UNA petición. -> { ok, items: [{key, value}] }
export async function safeListValuesStrict(prefix, attempts = 3) {
  return withRetry_(async () => {
    const data = await getJson_(`${API_URL}?action=listValues&prefix=${encodeURIComponent(prefix || "")}`);
    if (!Array.isArray(data.items)) throw new Error("respuesta incompleta");
    return { items: data.items };
  }, attempts);
}

// Verificación de correo hecha en el servidor. -> { ok, record: objeto|null }
export async function safeFindByEmailStrict(prefix, email, attempts = 3) {
  return withRetry_(async () => {
    const data = await getJson_(`${API_URL}?action=findByEmail&prefix=${encodeURIComponent(prefix)}&email=${encodeURIComponent(email)}`);
    if (!("found" in data)) throw new Error("respuesta incompleta");
    return { record: data.found || null };
  }, attempts);
}
