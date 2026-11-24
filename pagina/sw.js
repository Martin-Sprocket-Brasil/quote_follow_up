/* Quote Follow Up · avisos no celular / PC (o servidor entrega este arquivo em /sw.js).
   Só mostra os avisos e abre o programa ao tocar neles: não guarda nada e não funciona offline. */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));

self.addEventListener("push", (e) => {
  let d = {};
  try {
    d = e.data ? e.data.json() : {};
  } catch {
    d = { body: e.data ? e.data.text() : "" };
  }
  e.waitUntil(
    (async () => {
      await self.registration.showNotification(d.title || "Quote Follow Up", {
        body: d.body || "",
        icon: "/icon-192.png",
        badge: "/icon-192.png",
        tag: d.tag || undefined,
        renotify: !!d.tag,
        data: { url: d.url || "/" },
      });
      // programa aberto: busca as novidades na hora (o sino atualiza)
      const list = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
      list.forEach((c) => c.postMessage({ type: "qfu-push", tag: d.tag || "" }));
    })(),
  );
});

self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || "/", self.location.origin).href;
  e.waitUntil(
    (async () => {
      const list = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
      const open = list.find((c) => new URL(c.url).origin === self.location.origin);
      if (open) {
        await open.focus();
        open.postMessage({ type: "qfu-push-open" });
        return;
      }
      await self.clients.openWindow(url);
    })(),
  );
});
