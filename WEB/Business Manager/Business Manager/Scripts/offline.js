(function () {
    'use strict';
    var offline_body = document.querySelector('body');
    var offline_html = document.querySelector('html');
    var offline_img = document.querySelector('img');

    //After DOM Loaded
    document.addEventListener('DOMContentLoaded', function (event) {
        //On initial load to check connectivity
        if (!navigator.onLine) {
            updateNetworkStatus();
        }
        window.addEventListener('online', updateNetworkStatus, false);
        window.addEventListener('offline', updateNetworkStatus, false);
    });

    //To update network status
    function updateNetworkStatus() {
        if (navigator.onLine) {
            toastr.clear();
            offline_body.classList.remove('body_offline');
            offline_html.classList.remove('html_offline');
            offline_img.classList.remove('img_offline');
        }
        else {
            toastr.options = toastrOffline;
            toastr.error("網站離線中，請確認網路狀態");

            offline_body.classList.add('body_offline');
            offline_html.classList.add('html_offline');
            offline_img.classList.add('img_offline');
        }
    }
})();