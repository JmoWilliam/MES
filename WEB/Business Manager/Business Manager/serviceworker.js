(function () {
    'use strict';

    //#region 快取版本與檔案
    let version = 'cache::1.00';
    let urlsToCache = [
        '/Content/images/logo.png'
    ];

    function addToCache(request, response) {
        if (!response.ok && response.type !== 'opaque')
            return;

        let copy = response.clone();
        caches.open(version)
            .then(function (cache) {
                cache.put(request, copy);
            });
    }

    function updateStaticCache() {
        return caches.open(version)
            .then(function (cache) {
                return cache.addAll(urlsToCache);
            });
    }
    //#endregion

    //#region 監聽Service Worker安裝事件
    self.addEventListener('install', function (event) {
        //console.log('[Service Worker] Installing Service Worker ...', event);

        //#region 設定快取
        event.waitUntil(updateStaticCache());
        //#endregion
    });
    //#endregion

    //#region 監聽Service Worker啟動事件
    self.addEventListener('activate', function (event) {
        //console.log('[Service Worker] Activating Service Worker ...', event);

        //#region 清除舊快取
        event.waitUntil(
            caches.keys()
                .then(function (keys) {
                    return Promise.all(keys
                        .filter(function (key) {
                            return key.indexOf(version) !== 0;
                        })
                        .map(function (key) {
                            return caches.delete(key);
                        })
                    );
                })
        );
        //#endregion

        return self.clients.claim();
    });
    //#endregion

    //#region 監聽Service Worker擷取事件
    self.addEventListener('fetch', function (event) {
        //console.log('[Service Worker] Fetch something ...', event);

        //let request = event.request;

        ////#region Always fetch non-GET requests from the network
        //if (request.method !== 'GET' || request.url.match(/\/browserLink/ig)) {
        //    event.respondWith(
        //        fetch(request)
        //            .catch(function () {
        //                return caches.match(offlineUrl);
        //            })
        //    );
        //    return;
        //}
        ////#endregion

        ////#region For HTML requests, try the network first, fall back to the cache, finally the offline page
        //if (request.headers.get('Accept').indexOf('text/html') !== -1) {
        //    event.respondWith(
        //        fetch(request)
        //            .then(function (response) {
        //                // Stash a copy of this page in the cache
        //                addToCache(request, response);
        //                return response;
        //            })
        //            .catch(function () {
        //                return caches.match(request)
        //                    .then(function (response) {
        //                        return response || caches.match(offlineUrl);
        //                    });
        //            })
        //    );
        //    return;
        //}
        ////#endregion

        ////#region Cache first for fingerprinted resources
        //if (request.url.match(/(\?|&)v=/ig)) {
        //    event.respondWith(
        //        caches.match(request)
        //            .then(function (response) {
        //                return response || fetch(request)
        //                    .then(function (response) {
        //                        addToCache(request, response);
        //                        return response || serveOfflineImage(request);
        //                    })
        //                    .catch(function () {
        //                        return serveOfflineImage(request);
        //                    });
        //            })
        //    );

        //    return;
        //}
        //#endregion

        //#region Network first for non-fingerprinted resources
        //event.respondWith(
        //    fetch(request)
        //        .then(function (response) {
        //            addToCache(request, response);
        //            return response;
        //        })
        //        .catch(function () {
        //            return caches.match(request)
        //                .then(function (response) {
        //                    return response || serveOfflineImage(request);
        //                })
        //                .catch(function () {
        //                    return serveOfflineImage(request);
        //                });
        //        })
        //);
        //#endregion
    });
    //#endregion

    self.addEventListener('push', function (event) {
        let json = JSON.parse(event.data.text());

        console.log('[Service Worker] Push Received.');
        console.log(`[Service Worker] Push had this data is: "${event.data.text()}"`);

        const title = `${json.title}`;
        const options = {
            body: `${json.body}`,
            icon: `${json.icon}`,
            lang: `${json.lang}`,
            vibrate: json.vibrate,
            badge: `${json.badge}`,
            tag: `${json.tag}`,
            renotify: `${json.renotify}`
        };

        if (json.actions) options.actions = json.actions;
        if (json.urls) options.data = json.urls;

        event.waitUntil(self.registration.showNotification(title, options));
    });

    self.addEventListener("notificationclick", function (event) {
        let url;
        if (!event.action) {
            console.log("Notification Click.");
            return;
        }

        let urls = event.notification.data;
        switch (event.action) {
            case "goto":
            case "cancel":
                for (var i = 0; i < urls.length; i++) {
                    if (urls[i].action == event.action) {
                        url = urls[i].url;
                    }
                }
                break;
            default:
                console.log(`Unknown action clicked: '${event.action}'`);
                break;
        }

        event.notification.close(); // Android 需要觸發關閉
        if (url) {
            const urlToOpen = new URL(url, self.location.origin).href;
            const promiseChain = clients
                .matchAll({
                    type: "window",
                    includeUncontrolled: true,
                })
                .then((windowClients) => {
                    let matchingClient = null;

                    for (let i = 0; i < windowClients.length; i++) {
                        const windowClient = windowClients[i];
                        if (windowClient.url === urlToOpen) {
                            matchingClient = windowClient;
                            break;
                        }
                    }

                    if (matchingClient) {
                        return matchingClient.focus();
                    } else {
                        return clients.openWindow(urlToOpen);
                    }
                });
            event.waitUntil(promiseChain);
        }
    });
})();