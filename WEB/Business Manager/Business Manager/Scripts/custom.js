let isMobile = false;
let delay;
let cmEditor; //Code Mirror
let rtEditor; //Rich Text
let ComboBoxs = new ComboBoxControl();
let DataAccess = new DataAccessControl();
let FileAccess = new FileAccessControl();
let Suites = new SuiteControl();
let timerInterval;

(function ($) {
    "use strict";

    //#region Init()
    $(function () {
        //#region jQuery Validation 初始化
        $("form").validate();
        //#endregion

        //#region 自定義警告視窗 初始化
        customAlert();
        //#endregion

        //#region 使用裝置偵測
        if (/(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|ipad|iris|kindle|Android|Silk|lge |maemo|midp|mmp|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino/i.test(navigator.userAgent)
            || /1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-/i.test(navigator.userAgent.substr(0, 4))) {
            isMobile = true;
        }
        //#endregion
    });
    //#endregion

    //#region Load()
    $(window).on("load", function () {
        //#region Loading特效
        $("#loaderWrpper").delay(500).fadeOut();
        //#endregion

        //#region 彈跳視窗重複彈出
        MultileModal();
        //#endregion

        //$('[data-toggle="tooltip"]').tooltip();
        $('[data-toggle="popover"]').popover();

        //#region 偵測Dom改變
        let observer = new MutationObserver(mutationRecords => {
            
        });

        observer.observe(elem, {
            childList: true, // observe direct children
            subtree: true, // and lower descendants too
            characterDataOldValue: true // pass old data to callback
        });
        //#endregion
    });
    //#endregion

    //#region Scroll()
    $(window).on("scroll", function () {
        //#region 返回頂部按鈕
        if ($(this).scrollTop() > 100) {
            $("#back-to-top").fadeIn();
        } else {
            $("#back-to-top").fadeOut();
        }
        //#endregion
    });
    //#endregion

    //#region 選單展開
    $(".icon_menu").on("click", function () {
        $("body").toggleClass("mobile_nav");
    });
    //#endregion

    //#region 返回頂部
    $("#back-to-top").on("click", function () {
        $("body,html").animate({
            scrollTop: 0
        }, 800);
    });
    //#endregion

    //#region 表頭隱藏
    $("body").css("padding-top", $(".main-header").outerHeight() + 'px')

    if ($(".smart-scroll").length > 0) { // check if element exists
        var lastScrollTop = 0;

        $(window).on('scroll', function () {
            var scrollTop = $(this).scrollTop();

            if (!$("body").hasClass("mobile_nav")) {
                if (scrollTop < lastScrollTop) {
                    $(".smart-scroll").removeClass("scrolled-down").addClass("scrolled-up");
                    $(".sub-company-header").css("top", lastScrollTop > 60 ? scrollTop : 0);
                } else {
                    if (scrollTop > 0) {
                        $(".smart-scroll").removeClass("scrolled-up").addClass("scrolled-down");
                        $(".sub-company-header").css("top", 0);
                    }
                }
            }

            lastScrollTop = scrollTop;
        });
    }
    //#endregion

    //#region mCustomScrollbar
    if ($(".scroll_auto").length) {
        $(".scroll_auto").mCustomScrollbar({
            setWidth: false,
            setHeight: false,
            setTop: 0,
            setLeft: 0,
            axis: "y",
            scrollbarPosition: "inside",
            scrollInertia: 950,
            autoDraggerLength: true,
            autoHideScrollbar: false,
            autoExpandScrollbar: false,
            alwaysShowScrollbar: 0,
            snapAmount: null,
            snapOffset: 0
        });
    };
    //#endregion

    //#region bootstrap maxlength
    $("[maxlength]").maxlength({
        //warningClass: "btn btn-outline-success btn-sm mt-1",
        //limitReachedClass: "btn btn-outline-danger btn-sm mt-1",
        warningClass: "text-success pr-1",
        limitReachedClass: "text-danger pr-1",
        placement: "bottom-right-inside"
    });
    //#endregion

    //#region textarea autosize
    //autosize($("textarea"));
    //#endregion
})(jQuery);

// Cards
(function ($) {
    $(function () {
        $('.card')
            .on('card:toggle', function () {
                var $this,
                    direction;

                $this = $(this);
                direction = $this.hasClass('card-collapsed') ? 'Down' : 'Up';

                $this.find('.card-body, .card-footer')['slide' + direction](200, function () {
                    $this[(direction === 'Up' ? 'add' : 'remove') + 'Class']('card-collapsed')
                });
            })
            .on('card:dismiss', function () {
                var $this = $(this);

                if (!!($this.parent('div').attr('class') || '').match(/col-(xs|sm|md|lg)/g) && $this.siblings().length === 0) {
                    $row = $this.closest('.row');
                    $this.parent('div').remove();
                    if ($row.children().length === 0) {
                        $row.remove();
                    }
                } else {
                    $this.remove();
                }
            })
            .on('click', '[data-card-toggle]', function (e) {
                e.preventDefault();
                $(this).closest('.card').trigger('card:toggle');
            })
            .on('click', '[data-card-dismiss]', function (e) {
                e.preventDefault();
                $(this).closest('.card').trigger('card:dismiss');
            })
            .on('click', '.card-actions a i.fa-caret-up', function (e) {
                e.preventDefault();
                var $this = $(this);

                $this
                    .removeClass('fa-caret-up')
                    .addClass('fa-caret-down');

                $this.closest('.card').trigger('card:toggle');
            })
            .on('click', '.card-actions a i.fa-caret-down', function (e) {
                e.preventDefault();
                var $this = $(this);

                $this
                    .removeClass('fa-caret-down')
                    .addClass('fa-caret-up');

                $this.closest('.card').trigger('card:toggle');
            })
            .on('click', '.card-actions a.fa-times', function (e) {
                e.preventDefault();
                var $this = $(this);

                $this.closest('.card').trigger('card:dismiss');
            });
    });
})(jQuery);