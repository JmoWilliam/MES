//#region jQuery Validation 全域設定
jQuery.validator.setDefaults({
    errorElement: "small",
    errorClass: "error has-input",
    errorPlacement: function (error, element) {
        error.addClass("form-text");

        if ($(element).is(":radio") || $(element).is(":checkbox")) {
            var eid = $(element).attr("name");
            $("input[name=" + eid + "]:last").parent(".control").parent("div").siblings(".col-form-label").append(error);
            $(element).parent(".control").parent("div").parent(".form-group").addClass("has-danger");
        } else {
            if ($(element).parents(".input-group").length) {
                error.insertAfter($(element).parents(".input-group"));
            } else {
                if ($(element).siblings().length) {
                    error.insertAfter($(element).siblings(":last"));
                } else {
                    error.insertAfter(element);
                }
            }

            $(element).closest(".form-group").addClass("has-danger");
        }
    },
    unhighlight: function (element, errorClass) {
        if ($(element).is(":radio") || $(element).is(":checkbox")) {
            var eid = $(element).attr("name");
            $("input[name=" + eid + "]:last").parent(".control").parent("div").siblings(".col-form-label").find(".form-text").remove();
            this.findByName(element.name).removeClass(errorClass);
            $(element).parent(".control").parent("div").parent(".form-group").removeClass("has-danger");
        } else {
            if ($(element).parent(".input-group").length) {
                $(element).parent(".input-group").siblings(".form-text").remove();

                if (!$(element).closest(".form-group").find(".error.has-input").length) {
                    $(element).closest(".form-group").removeClass("has-danger");
                }
            } else {
                $(element).siblings(".form-text").remove();
                $(element).closest(".form-group").removeClass("has-danger");
            }

            $(element).removeClass(errorClass);
        }
    }
});
//#endregion

//#region 自訂驗證 數字 大於
jQuery.validator.addMethod("extraGreater", function (value, element, params) {
    return this.optional(element) || value > params;
}, jQuery.validator.format("請輸入大於 {0} 的數字"));
//#endregion    

//#region 自訂驗證 數字 大於 或 等於
jQuery.validator.addMethod("extraGreaterEqual", function (value, element, params) {
    return this.optional(element) || value > params || value == params;
}, jQuery.validator.format("請輸入大於或等於 {0} 的數字"));
//#endregion    

//#region 自訂驗證 數字 小於
jQuery.validator.addMethod("extraLesser", function (value, element, params) {
    return this.optional(element) || value < params;
}, jQuery.validator.format("請輸入小於 {0} 的數字"));
//#endregion

//#region 自訂驗證 數字 小於 或 等於
jQuery.validator.addMethod("extraLesserEqual", function (value, element, params) {
    return this.optional(element) || value < params || value == params;
}, jQuery.validator.format("請輸入小於或等於 {0} 的數字"));
//#endregion

//#region 自訂驗證 數字 等於
jQuery.validator.addMethod("extraEqual", function (value, element, params) {
    return this.optional(element) || value == params;
}, jQuery.validator.format("請輸入等於 {0} 的數字"));
//#endregion

//#region 自訂驗證 數字 不等於
jQuery.validator.addMethod("extraNotEqual", function (value, element, params) {
    return this.optional(element) || value != params;
}, jQuery.validator.format("請輸入不等於 {0} 的數字"));
//#endregion

//#region 自訂驗證 數字 介於
jQuery.validator.addMethod("extraBetween", function (value, element, params) {
    return this.optional(element) || (value >= params[0] && value <= params[1]);
}, jQuery.validator.format("請輸入介於 {0} 和 {1} 的數字"));
//#endregion

//#region 自訂驗證 數字 非介於
jQuery.validator.addMethod("extraNotBetween", function (value, element, params) {
    return this.optional(element) || (value < params[0] || value > params[1]);
}, jQuery.validator.format("請輸入非介於 {0} 和 {1} 的數字"));
//#endregion

//#region 自訂驗證 數字 整數
jQuery.validator.addMethod("extraInteger", function (value, element) {
    return this.optional(element) || value % 1 === 0;
}, jQuery.validator.format("請輸入整數"));
//#endregion

//#region 自訂驗證 文件 包含
jQuery.validator.addMethod("extraContain", function (value, element, params) {
    return this.optional(element) || value.includes(params);
}, jQuery.validator.format("內容必須包含: {0}"));
//#endregion

//#region 自訂驗證 文件 不包含
jQuery.validator.addMethod("extraNotContain", function (value, element, params) {
    return this.optional(element) || !value.includes(params);
}, jQuery.validator.format("內容不得包含: {0}"));
//#endregion

//#region 自訂驗證 文字 網址驗證
jQuery.validator.addMethod("extraUrl", function (value, element) {
    return this.optional(element) || /^((?:(https?|ftps?|file|ssh|sftp):\/\/|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}\/)(?:[^\s()<>]+|\((?:[^\s()<>]+|(?:\([^\s()<>]+\)))*\))+(?:\((?:[^\s()<>]+|(?:\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?\xab\xbb\u201c\u201d\u2018\u2019]))$/.test(value);
}, jQuery.validator.format("請輸入正確的網址"));
//#endregion

//#region 自訂驗證 正規表達式
jQuery.validator.addMethod("extraRegex", function (value, element, params) {
    return this.optional(element) || value.match(new RegExp("^" + params + "$"));
}, jQuery.validator.format("請輸入符合規則的值"));
//#endregion

//#region 自訂驗證 核取方塊 選取至少
jQuery.validator.addMethod("extraSelectAtLeast", function (value, element, params) {
    return this.optional(element) || $("input[name=" + $(element).attr("name") + "]:checked").length >= params;
}, jQuery.validator.format("請選取至少 {0} 個選項"));
//#endregion

//#region 自訂驗證 核取方塊 選取最多
jQuery.validator.addMethod("extraSelectAtMost", function (value, element, params) {
    return this.optional(element) || $("input[name=" + $(element).attr("name") + "]:checked").length <= params;
}, jQuery.validator.format("請選取最多 {0} 個選項"));
//#endregion

//#region 自訂驗證 核取方塊 選取剛好
jQuery.validator.addMethod("extraSelectExactly", function (value, element, params) {
    return this.optional(element) || $("input[name=" + $(element).attr("name") + "]:checked").length == params;
}, jQuery.validator.format("請選取 {0} 個選項"));
//#endregion