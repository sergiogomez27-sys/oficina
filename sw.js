// Oficina sin conexion: guarda la aplicacion en el telefono la primera vez y despues la abre aunque no haya red.
// Cada vez que hay red se descarga la version nueva en segundo plano y se usa en la siguiente apertura.
const CACHE = 'oficina-v1';
const APP = [
  './', 'index.html', 'styles.css', 'manifest.webmanifest', 'icons/oficina.svg', 'icons/oficina-192.png', 'icons/oficina-512.png',
  'vendor/jszip.min.js',
  'core/ooxml.js', 'core/formula.js', 'core/cf.js', 'core/chart.js', 'core/tablestyles.js', 'core/pivot.js', 'core/funcinfo.js',
  'core/samples.js', 'core/menu.js', 'core/ribbon.js', 'core/shell.js', 'core/touch.js',
  'formats/docx.js', 'formats/xlsx.js', 'formats/pptx.js',
  'apps/writer.js', 'apps/sheet.js', 'apps/slides.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(APP)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Primero lo guardado (abre al instante y sin red); a la vez se pide la copia nueva y se guarda.
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(async (c) => {
    const hit = await c.match(req, { ignoreSearch: true });
    const fresh = fetch(req).then((res) => {
      if (res.ok) c.put(req, res.clone());
      return res;
    });
    if (hit) {
      e.waitUntil(fresh.catch(() => {}));
      return hit;
    }
    return fresh;
  }));
});
