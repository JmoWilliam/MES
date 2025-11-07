//#region Toastr 全域設定
var toastrDefault = {
    "closeButton": true,
    "debug": false,
    "newestOnTop": true,
    "progressBar": true,
    "positionClass": "toast-top-right",
    "preventDuplicates": false,
    "onclick": null,
    "showDuration": "300",
    "hideDuration": "1000",
    "timeOut": "2500",
    "extendedTimeOut": "1000",
    "showEasing": "swing",
    "hideEasing": "linear",
    "showMethod": "fadeIn",
    "hideMethod": "fadeOut",
    "tapToDismiss": false
}
//#endregion

//#region 右上
var toastrTopRight = JSON.parse(JSON.stringify(toastrDefault));
toastrTopRight.positionClass = "toast-top-right";
//#endregion

//#region 右下
var toastrBottomRight = JSON.parse(JSON.stringify(toastrDefault));
toastrBottomRight.positionClass = "toast-bottom-right";
//#endregion

//#region 左上
var toastrTopLeft = JSON.parse(JSON.stringify(toastrDefault));
toastrTopLeft.positionClass = "toast-top-left";
//#endregion

//#region 左下
var toastrBottomLeft = JSON.parse(JSON.stringify(toastrDefault));
toastrBottomLeft.positionClass = "toast-bottom-left";
//#endregion

//#region 中上
var toastrTopCenter = JSON.parse(JSON.stringify(toastrDefault));
toastrTopCenter.positionClass = "toast-top-center";
//#endregion

//#region 中下
var toastrBottomCenter = JSON.parse(JSON.stringify(toastrDefault));
toastrBottomCenter.positionClass = "toast-bottom-center";
//#endregion

//#region 滿版上
var toastrTopFull = JSON.parse(JSON.stringify(toastrDefault));
toastrTopFull.positionClass = "toast-top-full-width";
//#endregion

//#region 滿版下
var toastrBottomFull = JSON.parse(JSON.stringify(toastrDefault));
toastrBottomFull.positionClass = "toast-bottom-full-width";
//#endregion

//#region 離線用
var toastrOffline = JSON.parse(JSON.stringify(toastrDefault));
toastrOffline.positionClass = "toast-bottom-full-width";
toastrOffline.timeOut = "0";
toastrOffline.extendedTimeOut = "0";
//#endregion