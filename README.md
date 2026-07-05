# Catálogo Manicuría — página estática + Google Sheets + WhatsApp

Página single-file (`index.html`) que lee el catálogo desde un Google Sheet público,
arma un carrito y envía el pedido como texto por WhatsApp. Sin backend.

## Setup

### 1. Google Sheet

Sheet en uso (cuenta maroguromero@gmail.com):
`https://docs.google.com/spreadsheets/d/1VZOHrfKm3vfoz30jzcKSnhVe7Ajo51s4GP8YYFMOyOg/edit`

**Tab `productos`** — columnas en la fila 1 (en minúsculas, sin tildes):

| id | seccion | nombre | descripcion | precio | imagen | disponible |
|----|---------|--------|-------------|--------|--------|------------|
| ESM-MELINE-001 | Esmaltes | Semipermanente rosa | Color intenso, 15ml | 4500 | | si |
| LIMA-100-180 | Herramientas | Lima 100/180 | Doble cara | 2300 | | si |
| HERR-TORNO-001 | Herramientas | Torno 35000 RPM | Recargable | 45000 | | no |

- `id`: identificador del producto — **determinístico, no un contador**, siempre en
  **MAYÚSCULA**. Formato: `SECCION-LINEA-CODIGO`:
  - Esmaltes → línea + código de color: `ESM-MELINE-001`, `ESM-MELINE-002`
    (el `001` es el código de color de esa línea, no un correlativo de fila)
  - Limas → por grado: `LIMA-100-180`, `LIMA-240`
  - Tips → por estilo: `TIP-STILETTO`, `TIP-COFFIN`
  - Resto → abreviatura de sección + nombre corto + código: `HERR-ALICATE-001`,
    `INS-ALGODON-001`, `CAB-UVLED-001`

  Por qué determinístico y no `001, 002, 003...` corrido por fila: si insertás o
  reordenás filas en el sheet, un id secuencial se desincroniza con las fotos ya
  subidas (la fila 3 deja de ser la fila 3). Un id atado al producto (su línea, su
  código de color/grado/estilo) no se mueve nunca — vincula fotos por lo que el
  producto ES, no por dónde está sentado en la planilla.

  Si dejás la celda `id` vacía, la página genera un id automático en mayúscula a
  partir del `nombre` (mismo criterio: determinístico, no secuencial) — pero siempre
  es mejor ponerlo a mano con el criterio de arriba, más corto y prolijo para nombrar
  la carpeta de fotos de ese producto. La comparación id ↔ nombre de carpeta no
  distingue mayúsculas/minúsculas.
- `seccion`: agrupa productos; el orden de aparición define el orden de secciones.
- `precio`: número; acepta `4500`, `4.500`, `$4.500,50`.
- `imagen`: link a la subcarpeta de Drive de ESE producto — **solo referencia visual**,
  para saber de un vistazo dónde subir sus fotos; el código lo ignora (un link de
  carpeta de Drive nunca se trata como foto, no rompe el carrusel). Si en cambio
  pegás ahí la URL de UN archivo de imagen o un link "compartir" de Drive a un
  archivo puntual, ESE sí se toma como foto extra y se muestra primera.
- `disponible`: `no` = muestra "Sin stock" sin botón; cualquier otro valor (o vacío) = disponible.

**Tab `fotos`** — índice de fotos por producto (lo llena solo el Apps Script, no se toca a mano):

| id | file_id |
|----|---------|
| ESM-MELINE-001 | 1JrZFJnqR8... |
| ESM-MELINE-001 | 1uHcJA6Nid... |

Compartir el sheet: **Archivo → Compartir → Cualquier persona con el enlace → Lector**.

### 1b. Fotos de productos (carpeta Drive, una subcarpeta por producto)

Carpeta pública: `https://drive.google.com/drive/folders/1B4DnERnY3ev5dVZKmYu2CXWZzG3lyaOu`

Estructura: **una subcarpeta dentro de esa, nombrada exactamente igual que el `id`
del producto**, con las fotos de ese producto adentro. El nombre de archivo de la
foto es libre (`foto1.jpg`, `IMG_2024.png`, lo que sea) — lo que importa es en qué
subcarpeta está.

```
manicura-fotos/
  ESM-MELINE-001/
    foto1.png
    foto2.png
  LIMA-100-180/
    lima.jpg
```

Orden de las fotos en el carrusel: alfabético por nombre de archivo dentro de la
subcarpeta — si te importa el orden, nombralas `1.jpg`, `2.jpg`, etc.

Flujo para agregar fotos a un producto:
1. Si no existe, crear una subcarpeta con el nombre exacto del `id` (ej. `ESM-MELINE-001`).
   Al estar dentro de la carpeta pública, hereda el permiso — no hay que compartirla de nuevo.
2. Arrastrar las fotos adentro. Nombre de archivo: el que sea.
3. Correr `syncFotos` (o esperar al trigger automático) para que el tab `fotos` se actualice.

Apps Script — pegar en el sheet → Extensiones → Apps Script, y crear un trigger
horario (o "al editar") sobre `syncFotos`:

```js
const FOLDER_ID = "1B4DnERnY3ev5dVZKmYu2CXWZzG3lyaOu";

function syncFotos() {
  const rows = [["id", "file_id"]];
  const subfolders = DriveApp.getFolderById(FOLDER_ID).getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const id = sub.getName().trim().toUpperCase();
    const files = [];
    const it = sub.getFiles();
    while (it.hasNext()) files.push(it.next());
    files.sort((a, b) => a.getName().localeCompare(b.getName()));
    for (const f of files) rows.push([id, f.getId()]);
  }
  const sheet = SpreadsheetApp.getActive().getSheetByName("fotos");
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}
```

### 2. Configurar `index.html`

Editar el bloque `CONFIG` al inicio del `<script>`:

```js
const CONFIG = {
  SHEET_ID: "1BxiMVs0...",        // de la URL del sheet: /d/<ESTO>/edit
  SHEET_GID: "0",                 // pestaña (gid= en la URL; 0 = primera)
  WHATSAPP_NUMBER: "5491122334455", // código país + número, sin + ni espacios
  STORE_NAME: "Mi Tienda",
  CURRENCY: "$",
  REFRESH_SECONDS: 300,           // re-lee el sheet cada 5 min con la página abierta
};
```

Sin `SHEET_ID` la página corre en **modo demo** con productos de muestra.

### 3. Hosting

Cualquier hosting estático sirve: S3 + CloudFront, GitHub Pages, Netlify, Vercel.
Es un solo archivo — subir `index.html` y listo.

> Abrir el archivo en el browser con doble click también funciona para probar
> (el endpoint de Google permite CORS desde `file://`).

## Cómo funciona

- Lee `https://docs.google.com/spreadsheets/d/<ID>/gviz/tq?tqx=out:csv&gid=<GID>`
  (endpoint público de Google, sin API key).
- Cambios en el sheet aparecen al recargar la página (o cada `REFRESH_SECONDS`).
  Google cachea ~1-5 min, no es instantáneo.
- Carrito persiste en `localStorage` (sobrevive recargas/cierres de pestaña).
- "Enviar pedido" abre `https://wa.me/<numero>?text=<pedido>` con el detalle:
  ítems, cantidades, subtotales, total, nombre y comentario del cliente.
  El cliente solo tiene que tocar "Enviar" en WhatsApp.

## Limitaciones conocidas

- El sheet es legible por cualquiera que tenga la URL del endpoint (es público).
  No poner datos sensibles (costos, márgenes, datos de clientes).
- El pedido por WhatsApp lo envía el cliente desde su propio WhatsApp —
  si no lo envía, no llega. No hay registro server-side de pedidos.
- Precios/stock son los del momento de carga de la página; confirmar al tomar el pedido.
