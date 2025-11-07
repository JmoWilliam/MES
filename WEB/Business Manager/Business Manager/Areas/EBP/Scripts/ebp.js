let DataAccess = new DataAccessControl();
let ComboBoxs = new ComboBoxControl();

(function ($) {
    "use strict";

    $(window).on("load", function () {
        //#region Loading特效
        $("div").remove(".loading");
        //#endregion
    });
})(jQuery);

function WeekDay(currentDate, targetElement) {
    var array = ["日", "一", "二", "三", "四", "五", "六"];
    let weekDay = new Date(currentDate).getDay(); //取出今天星期幾
    $(targetElement).html(currentDate + "(" + array[weekDay] + ")");
}