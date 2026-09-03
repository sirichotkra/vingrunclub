/* VING RUN CLUB — service worker
   หน้าที่: ทำให้แอปเปิดเร็ว และเปิดได้แม้เน็ตช้า/หลุดชั่วคราว

   กฎสำคัญ 3 ข้อ
   1. ไม่ยุ่งกับการคุยเซิร์ฟเวอร์เลย (ข้ามโดเมน + POST) — ข้อมูลต้องสดเสมอ
   2. หน้าเว็บใช้ "เน็ตก่อน แคชสำรอง" — แก้โค้ดแล้วคนเห็นของใหม่ทันที ไม่ค้างเวอร์ชันเก่า
   3. ไฟล์คงที่ (ไอคอน/ฟอนต์) ใช้แคชก่อนแล้วค่อยอัปเดตเงียบๆ                       */

var VERSION = "ving-v4";
var SHELL = [
  "./",
  "./index.html",
  "./config.js",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-180.png",
  "./icons/maskable-512.png"
];

self.addEventListener("install", function (e) {
  e.waitUntil(
    caches.open(VERSION)
      .then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
      .catch(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.map(function (k) {
          return k === VERSION ? null : caches.delete(k);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (e) {
  var req = e.request;

  // 1. ไม่แตะคำขอที่ไม่ใช่ GET (เช่น POST ไป Apps Script)
  if (req.method !== "GET") return;

  var url;
  try { url = new URL(req.url); } catch (err) { return; }

  // 2. ไม่แตะอะไรที่อยู่คนละโดเมนกับแอป (Apps Script, Google Fonts)
  if (url.origin !== self.location.origin) return;

  // 3. หน้าเว็บ + ไฟล์ตั้งค่า: เน็ตก่อน ถ้าเน็ตไม่ได้ค่อยหยิบจากแคช
  if (/\/config\.js$/.test(url.pathname)) {
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(VERSION).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () { return caches.match(req); })
    );
    return;
  }

  if (req.mode === "navigate") {
    e.respondWith(
      fetch(req)
        .then(function (res) {
          var copy = res.clone();
          caches.open(VERSION).then(function (c) { c.put("./index.html", copy); });
          return res;
        })
        .catch(function () {
          return caches.match("./index.html").then(function (m) {
            return m || new Response("ออฟไลน์อยู่ — ลองเชื่อมเน็ตแล้วเปิดใหม่", {
              status: 503,
              headers: { "Content-Type": "text/plain; charset=utf-8" }
            });
          });
        })
    );
    return;
  }

  // 4. ไฟล์คงที่: แคชก่อน แล้วอัปเดตเบื้องหลัง
  e.respondWith(
    caches.match(req).then(function (hit) {
      var net = fetch(req).then(function (res) {
        if (res && res.status === 200) {
          var copy = res.clone();
          caches.open(VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      }).catch(function () { return hit; });
      return hit || net;
    })
  );
});
