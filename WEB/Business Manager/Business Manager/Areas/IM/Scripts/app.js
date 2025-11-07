let isSubscribed = false;
let swRegistration = null;
let vapidKey = "BKM0tOp6v26V36ll4FmozbA9xWvIg_qpqjOU47t1dXx70gli8rr-o7EcmB98ex65NKA12xT9Olrh6DEPfLBaLM4";
let pushButton = document.querySelector('.subscribe-bell');

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
    console.warn('Push messaging is not supported');
}

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
        pushButton.innerHTML = `<i class="fas fa-bell-exclamation set-icon"></i>`;
        updateSubscriptionOnServer(null);
        return;
    } else {
        if (isSubscribed) {
            pushButton.innerHTML = `<i class="fas fa-bell-on set-icon"></i>`;
        } else {
            pushButton.innerHTML = `<i class="fas fa-bell-slash set-icon"></i>`;
        }
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
                body: '成功訂閱!',
                icon: '/Areas/IM/Content/images/icons/icon-96x96.png',
                lang: 'zh-TW',
                vibrate: [100, 50, 200],
                badge: '/Areas/IM/Content/images/icons/icon-96x96.png',
                tag: 'confirm-notification',
                renotify: true
            };

            navigator.serviceWorker.ready.then(function (swreg) {
                swreg.showNotification('JMO Chat.', options);
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
    }
}

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