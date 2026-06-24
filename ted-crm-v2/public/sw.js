const CACHE_NAME = 'ted-crm-v1';

self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(clients.claim()));

self.addEventListener('push', e => {
  const data = e.data?.json() || {};
  e.waitUntil(
    self.registration.showNotification(data.title || 'Nouvelle réservation', {
      body: data.body || 'Une nouvelle demande est arrivée',
      icon: '/favicon.png',
      badge: '/favicon.png',
      tag: 'nouvelle-resa',
      renotify: true,
      vibrate: [200, 100, 200],
      data: { url: data.url || 'https://ted-crm.pages.dev' }
    })
  );
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const client of list) {
        if (client.url.includes('ted-crm.pages.dev') && 'focus' in client) {
          return client.focus();
        }
      }
      return clients.openWindow('https://ted-crm.pages.dev');
    })
  );
});
