// 서비스워커: 브라우저 뒤에서 계속 돌면서, 필요한 파일들을 미리 저장해뒀다가
// 인터넷이 느리거나 끊겨도 앱이 열리게 해주는 역할을 함.

const CACHE_NAME = "sigol-sonmat-v4";
const FILES_TO_CACHE = [
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./images/songpyeon.png",
  "./images/injeolmi.png",
  "./images/gaetteok.png",
  "./images/logo.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(FILES_TO_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// 인터넷 되면 항상 최신 파일을 받아오고, 안 될 때만 캐시된 걸로 대체함
// (예전엔 캐시부터 보여줘서 배포해도 손님 폰에 옛날 버전이 계속 남는 문제가 있었음)
self.addEventListener("fetch", (event) => {
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
