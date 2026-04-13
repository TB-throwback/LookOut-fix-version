const CACHE_VERSION = "lookout-pwa-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./mapi_props.js",
  "./scripts/lookout.mjs",
  "./scripts/tnef.mjs",
  "./manifest.webmanifest",
  "./icons/LOicon-32.png",
  "./icons/LOicon-48.png",
  "./icons/LOicon-64.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(ASSETS)),
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter((key) => key !== CACHE_VERSION)
            .map((key) => caches.delete(key)),
        ),
      ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") {
    return;
  }

  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }

      return fetch(event.request)
        .then((networkResponse) => {
          const copy = networkResponse.clone();
          caches
            .open(CACHE_VERSION)
            .then((cache) => cache.put(event.request, copy));
          return networkResponse;
        })
        .catch(() => caches.match("./index.html"));
    }),
  );
});
