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
  if (!res.ok) throw new Error("get failed");
  const data = await res.json();
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
  return true;
}

async function apiList(prefix) {
  const res = await fetch(`${API_URL}?action=list&prefix=${encodeURIComponent(prefix || "")}`);
  if (!res.ok) throw new Error("list failed");
  const data = await res.json();
  return data.keys || [];
}

// Estas tres funciones tienen la MISMA firma que las que usaba el artifact de
// Claude (safeGet/safeSet/safeList), incluido el parámetro "shared" que aquí
// se ignora porque ya no existe el concepto de almacenamiento personal vs.
// compartido (no hay cuentas de usuario).
export async function safeGet(key, _shared) {
  try {
    return await apiGet(key);
  } catch (e) {
    return null;
  }
}
export async function safeSet(key, value, _shared) {
  try {
    await apiSet(key, value);
    return true;
  } catch (e) {
    return false;
  }
}
export async function safeList(prefix, _shared) {
  try {
    return await apiList(prefix);
  } catch (e) {
    return [];
  }
}
