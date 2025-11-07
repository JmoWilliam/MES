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