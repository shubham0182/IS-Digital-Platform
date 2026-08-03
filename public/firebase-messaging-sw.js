/* ============================================================
   IS Digital Platform — Firebase Cloud Messaging service worker
   ------------------------------------------------------------
   Handles background (app-closed) web-push notifications.
   The Firebase config below MUST match the one in index.html
   (public/index.html). The VAPID key is only used on the client
   (notifications.js); the service worker needs the messagingSenderId.
   ============================================================ */
importScripts("https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js");

firebase.initializeApp({
  apiKey: "AIzaSyCmPIRS5LhuavoWp1S9DBrdX3y3TM2YQi0",
  authDomain: "is-digital-platform.firebaseapp.com",
  projectId: "is-digital-platform",
  storageBucket: "is-digital-platform.firebasestorage.app",
  messagingSenderId: "58417697125",
  appId: "1:58417697125:web:1378836c8596b68e9dd18c",
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage(function (payload) {
  const data = payload.data || {};
  const n = payload.notification || {};
  const url = data.url || "/#notifications";

  self.registration.showNotification(n.title || "IS Digital Platform", {
    body: n.body || "",
    icon: "/logo.png",
    badge: "/logo.png",
    data: { url: url },
  });
});

self.addEventListener("notificationclick", function (event) {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || "/";
  event.waitUntil(clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (list) {
    for (const client of list) {
      if ("focus" in client) { client.navigate(url); return client.focus(); }
    }
    return clients.openWindow(url);
  }));
});
