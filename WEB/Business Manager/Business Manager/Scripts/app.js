let isSubscribed = false;
let swRegistration = null;
let vapidKey = "BKM0tOp6v26V36ll4FmozbA9xWvIg_qpqjOU47t1dXx70gli8rr-o7EcmB98ex65NKA12xT9Olrh6DEPfLBaLM4";
let pushButton = document.querySelector('.subscribe-bell');

//#region 註冊 Service Worker
if ('serviceWorker' in navigator && 'PushManager' in window) {
    console.log('Service Worker and Push are supported');

    navigator.serviceWorker.register('/serviceworker.js')
        .then(function (swReg) {
            console.log('Service Worker is registered', swReg);

            swRegistration = swReg;
            initializeUI();
        })
        .catch(function (error) {
            console.error('Service Worker Error', error);
        });
} else {
    //console.warn('Push messaging is not supported');
    //alert('Push messaging is not supported');
}
//#endregion

function initializeUI() {
    pushButton.addEventListener('click', function () {
        if (Notification.permission === 'denied') {
            alert("請重設瀏覽器通知權限!");
            return;
        }

        if (isSubscribed) {
            unsubscribeUser();
        } else {
            subscribeUser();
        }
    });

    swRegistration.pushManager.getSubscription()
        .then(function (subscription) {
            isSubscribed = !(subscription === null);

            updateSubscriptionOnServer(subscription);

            if (isSubscribed) {
                console.log('User is subscribed.');
            } else {
                console.log('User is NOT subscribed.');
            }

            updateBtn();
        });
}

function updateBtn() {
    if (Notification.permission === 'denied') {
        pushButton.innerHTML = `<i class="fal fa-bell-exclamation"></i>&nbsp;&nbsp;禁止推播`;
        updateSubscriptionOnServer(null);
        return;
    }

    if (isSubscribed) {
        pushButton.innerHTML = `<i class="fal fa-bell-slash"></i>&nbsp;&nbsp;關閉推播`;
    } else {
        pushButton.innerHTML = `<i class="fal fa-bell"></i>&nbsp;&nbsp;訂閱推播`;
    }
}

function subscribeUser() {
    const applicationServerKey = urlB64ToUint8Array(vapidKey);
    swRegistration.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: applicationServerKey
    })
        .then(function (subscription) {
            console.log('User is subscribed.');

            let options = {
                icon: '/Content/images/icons/icon-96x96.png',
                lang: 'zh-TW',
                vibrate: [100, 50, 200],
                badge: '/Content/images/icons/icon-96x96.png',
                tag: 'confirm-notification',
                renotify: true
            };

            navigator.serviceWorker.ready.then(function (swreg) {
                swreg.showNotification('成功訂閱!!', options);
            });

            isSubscribed = true;

            updateSubscriptionOnServer(subscription);
            updateBtn();
        })
        .catch(function (error) {
            console.error('Failed to subscribe the user: ', error);
            updateBtn();
        });
}

function unsubscribeUser() {
    swRegistration.pushManager.getSubscription()
        .then(function (subscription) {
            if (subscription) {
                return subscription.unsubscribe();
            }
        })
        .catch(function (error) {
            console.log('Error unsubscribing', error);
        })
        .then(function () {
            updateSubscriptionOnServer(null);

            console.log('User is unsubscribed.');
            isSubscribed = false;

            updateBtn();
        });
}

function updateSubscriptionOnServer(subscription) {
    if (subscription) {
        var formData = new FormData();
        formData.append('subscription', JSON.stringify(subscription));

        fetch('/SystemSetting/AddPushSubscription', {
            method: 'POST',
            body: formData
        })
            .then((response) => {
                console.log(response);
            })
            .catch((error) => {
                console.log(`Error: ${error}`);
            });

        console.log(JSON.stringify(subscription));
    }
}

//#region 安裝app通知
let deferredPrompt;

window.addEventListener('beforeinstallprompt', function (event) {
    event.preventDefault();
    deferredPrompt = event;

    const installBtn = document.querySelector(".download-icon");
    installBtn.style.display = "block";

    return false;
});

function installApp() {
    if (deferredPrompt) {
        deferredPrompt.prompt();

        deferredPrompt.userChoice.then(function (choiceResult) {
            if (choiceResult.outcome === 'dismissed') {
                console.log('User cancelled installation');
            } else {
                installBtn.style.display = "none";
                console.log('User added to home screen');
            }
        });

        deferredPrompt = null;
    }
}
//#endregion

function urlB64ToUint8Array(base64String) {
    const padding = '='.repeat((4 - base64String.length % 4) % 4);
    const base64 = (base64String + padding)
        .replace(/\-/g, '+')
        .replace(/_/g, '/');

    const rawData = window.atob(base64);
    const outputArray = new Uint8Array(rawData.length);

    for (let i = 0; i < rawData.length; ++i) {
        outputArray[i] = rawData.charCodeAt(i);
    }
    return outputArray;
}