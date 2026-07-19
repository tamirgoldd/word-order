const CACHE = "word-order-v1";
const SHELL = new URL("./", self.location).href;

self.addEventListener("install", (event) => {
  event.waitUntil(
    fetch(SHELL, { cache: "reload" })
      .then((response) => {
        if (!response.ok) throw new Error("Unable to cache the app shell.");
        return caches.open(CACHE).then((cache) => cache.put(SHELL, response));
      })
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then(async (keys) => {
        const oldCaches = keys.filter((key) => key !== CACHE);
        await Promise.all(oldCaches.map((key) => caches.delete(key)));
        await self.clients.claim();
        if (oldCaches.length === 0) return;
        const windows = await self.clients.matchAll({ type: "window" });
        await Promise.all(windows.map((client) => client.navigate(client.url)));
      })
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET" || new URL(event.request.url).origin !== self.location.origin) return;
  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request).then((response) => {
        if (response.ok) void caches.open(CACHE).then((cache) => cache.put(SHELL, response.clone()));
        return response;
      }).catch(() => caches.match(SHELL))
    );
    return;
  }
  event.respondWith(
    caches.match(event.request).then((cached) => cached ?? fetch(event.request).then((response) => {
      if (response.ok) caches.open(CACHE).then((cache) => cache.put(event.request, response.clone()));
      return response;
    }).catch(() => caches.match(SHELL)))
  );
});
