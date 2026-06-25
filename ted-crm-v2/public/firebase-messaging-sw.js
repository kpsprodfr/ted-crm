importScripts('https://www.gstatic.com/firebasejs/10.7.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.7.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyDzIqs5W49wqIve3654rVB9oyDbUG0Jf0Y",
  authDomain: "le-ted.firebaseapp.com",
  projectId: "le-ted",
  storageBucket: "le-ted.firebasestorage.app",
  messagingSenderId: "1009473454163",
  appId: "1:1009473454163:web:f0dd9e6a5b12a80a7506a6"
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage(payload => {
  console.log('Background message reçu:', payload);
  const { title, body } = payload.notification;
  self.registration.showNotification(title, {
    body,
    icon: '/favicon.png',
    badge: '/favicon.png',
    tag: 'nouvelle-resa',
    renotify: true,
    vibrate: [200, 100, 200],
    data: { url: 'https://ted-crm.pages.dev' }
  });
});
