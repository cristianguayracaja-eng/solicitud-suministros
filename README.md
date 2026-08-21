# Solicitud de Suministros — CARTIMEX / COMPUTRON

Mismo sistema que teníamos en Claude (catálogo tipo tarjetas, colores por
marca, barra flotante, informe de administrador, exportación a Excel), pero
ahora es un sitio web independiente que **no requiere cuenta de Claude para
nadie**. Los datos se guardan en una Google Sheet tuya, a través de un
pequeño backend gratuito (Google Apps Script).

Hazlo en este orden: **1) Google Sheet → 2) Vercel**.

---

## Parte 1 — Crear el backend en Google Sheets (10 minutos)

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja de
   cálculo nueva en blanco. Ponle un nombre, por ejemplo
   "Suministros - Base de Datos".
2. En el menú, ve a **Extensiones → Apps Script**.
3. Borra todo el contenido del archivo `Code.gs` que se abre por defecto, y
   pega en su lugar **todo** el contenido del archivo
   `google-apps-script/Code.gs` que viene en esta carpeta.
4. Guarda (ícono de disquete o `Ctrl+S`).
5. Arriba a la derecha, haz clic en **Implementar → Nueva implementación**.
6. En "Selecciona el tipo", elige **Aplicación web**.
7. Configura:
   - **Ejecutar como:** Yo (tu cuenta)
   - **Quién tiene acceso:** Cualquier usuario
8. Haz clic en **Implementar**. Google te pedirá autorizar permisos la
   primera vez — acéptalos (es tu propia hoja, es seguro).
9. Copia la **URL de la aplicación web** que te muestra (termina en
   `/exec`). Esa es tu `VITE_API_URL`.

> Cada vez que edites el código de `Code.gs` en el futuro, tienes que volver
> a **Implementar → Administrar implementaciones → editar (lápiz) → Nueva
> versión → Implementar** para que los cambios se reflejen en la URL.

---

## Parte 2 — Preparar el proyecto localmente

Necesitas tener [Node.js](https://nodejs.org) instalado (versión 18 o
superior). Se instala una sola vez, como cualquier programa.

1. Abre una terminal dentro de esta carpeta del proyecto.
2. Instala las dependencias:
   ```bash
   npm install
   ```
3. Crea tu archivo de variables de entorno copiando el ejemplo:
   ```bash
   cp .env.example .env
   ```
4. Abre `.env` y pega la URL que copiaste de Apps Script en
   `VITE_API_URL`.
5. Prueba localmente:
   ```bash
   npm run dev
   ```
   Se abrirá en `http://localhost:5173`. Verifica que puedas entrar a
   Administración (contraseña inicial: `suministros2026`), abrir un
   periodo, y llenar una solicitud de prueba.

---

## Parte 3 — Subir a Vercel (gratis, sin necesidad de saber programar)

**Opción recomendada: con GitHub (lo más estable para futuras actualizaciones)**

1. Crea un repositorio nuevo en [github.com](https://github.com) y sube esta
   carpeta (puedes arrastrar los archivos desde la web de GitHub si no usas
   git por terminal).
2. Ve a [vercel.com](https://vercel.com), crea una cuenta gratuita (puedes
   entrar con tu cuenta de GitHub directamente).
3. Haz clic en **Add New → Project**, elige el repositorio que subiste.
4. Vercel detecta automáticamente que es un proyecto Vite — no cambies nada
   en la configuración de build.
5. Antes de darle a "Deploy", ve a **Environment Variables** y agrega:
   - **Name:** `VITE_API_URL`
   - **Value:** la URL de tu Apps Script (`.../exec`)
6. Haz clic en **Deploy**. En 1-2 minutos te da el link público, por ejemplo
   `https://solicitud-suministros.vercel.app`. Ese es el link que compartes
   — nadie necesita cuenta de nada para abrirlo y llenarlo.

**Alternativa sin GitHub:** instala la CLI de Vercel (`npm install -g
vercel`) y desde esta carpeta corre `vercel`. Sigue las instrucciones en
pantalla; te va a pedir crear cuenta gratuita y te preguntará por la
variable `VITE_API_URL` durante el proceso.

---

## Dominio propio (opcional)

Una vez desplegado en Vercel, en el panel del proyecto puedes ir a
**Settings → Domains** y agregar tu propio dominio (ej.
`solicitudsuministros.com`) si lo compras en cualquier proveedor
(Namecheap, GoDaddy, etc.). Vercel te da las instrucciones exactas de qué
registro DNS apuntar. Esto es opcional y tiene un costo aparte (el dominio
en sí, usualmente 10-15 USD al año); el hosting en Vercel sigue siendo
gratis.

---

## Actualizar artículos, tiendas o departamentos

Todo eso se administra **dentro del sitio ya desplegado**, en el panel de
Administración → pestañas "Tiendas / Departamentos" y "Artículos" — igual
que en la versión de Claude. No necesitas tocar el código para eso.

## Actualizar el diseño o la lógica en el futuro

Si más adelante quieres que ajuste algo del sistema, pídemelo en Claude, te
doy el archivo `App.jsx` actualizado, lo reemplazas en `src/App.jsx`, subes
el cambio a GitHub (o corres `vercel` de nuevo), y Vercel actualiza el sitio
automáticamente.
