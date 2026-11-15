importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

// عدد الجديد على أيقونة التطبيق عند وصول إشعار والتطبيق مغلق (الآيفون والحاسوب).
// أندرويد يعرض العدد بنفسه من الإشعارات، فلا يدعم هذه الخاصية ولا يحتاجها.
self.addEventListener('push', event => {
  if (!self.navigator || !self.navigator.setAppBadge) return;
  event.waitUntil((async () => {
    const cache = await caches.open('nour-badge');
    const hit = await cache.match('badge-count');
    const n = (hit ? parseInt(await hit.text(), 10) || 0 : 0) + 1;
    await cache.put('badge-count', new Response(String(n)));
    await self.navigator.setAppBadge(n);
  })());
});
