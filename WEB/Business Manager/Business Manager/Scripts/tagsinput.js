(function ($) {
    $.fn.extend({
        tagsinput: function (options) {
            var target = $(this).attr("id");
            var settings = $.extend({
                splitChar: ";"
            }, options);

            var targetValue = $(`#${target}`).val().split(settings.splitChar);
            var targetClass = $(`#${target}`).attr("class");
            var targetHasSpecificClass = $(`#${target}`).hasClass("input-group-field");
            //var replaceHtml = "<div class=\"input-tags" + (typeof targetClass === "undefined" ? "" : " " + targetClass) + "\">";

            var replaceHtml = `<div class="input-tags-block">
                                 <div class="input-tags">`;
            if ($(`#${target}`).val().length > 0) {
                for (var i = 0; i < targetValue.length; i++) {
                    replaceHtml += `<span class="tag-text">
                                      ${targetValue[i]}
                                      <button type="button" class="tag-delete"><i class="fas fa-times"></i></button>
                                    </span>`;
                }
            }
            replaceHtml += `    <input type="text" class="tag-input" id="${target}" name="${target}" />
                                <input type="hidden" data-target="${target}" value="${$(`#${target}`).val()}" />
                              </div>
                            </div>`;

            if ($(`#${target}`).closest(".input-tags-block").length > 0) {
                $(`#${target}`).closest(".input-tags-block").replaceWith(replaceHtml);
            } else {
                $(`#${target}`).replaceWith(replaceHtml);
            }

            $(`#${target}`).closest(".input-tags-block").parent(".input-group").css("flex-wrap", "unset");
            $(`#${target}`).closest(".input-tags-block").siblings(".input-group-append").find(".btn.btn-success").css("display", "flex").css("align-items", "center");
            
            resizable(document.getElementById(target), 8.1);

            var newInput = $(`#${target}`);
            if (targetHasSpecificClass) newInput.parent(".input-tags").css("margin-bottom", "0");

            newInput.parent(".input-tags").click(function () {
                newInput.focus();
            });

            newInput.change(function () {
                console.log(newInput.val());

                var addStr = newInput.val().split(settings.splitChar);
                for (var i = 0; i < addStr.length; i++) {
                    if (!checkStrExist(addStr[i])) {
                        var innerHtml = `<span class="tag-text">
                                          ${addStr[i]}
                                          <button type="button" class="tag-delete"><i class="fas fa-times"></i></button>
                                        </span>`;
                        newInput.before(innerHtml);
                        var newStr = addStr[i];
                        if ($(`[data-target='${target}']`).val().length > 0) newStr = settings.splitChar + newStr;
                        $(`[data-target='${target}']`).val($(`[data-target='${target}']`).val() + newStr);
                    }
                }
                newInput.val("");
            });

            $("div.input-tags").delegate("button.tag-delete", "click", function () {
                var deleteStr = $(this).parent().text().replace(/\s/g, "");
                var originalStr = $(`[data-target='${target}']`).val().split(settings.splitChar);
                var newStr = "";

                for (var i = 0; i < originalStr.length; i++) {
                    if (originalStr[i] != deleteStr) {
                        newStr = newStr + originalStr[i];
                        if (i < originalStr.length - 1) newStr = newStr + settings.splitChar;
                    }
                }

                if (newStr.slice(-1)) newStr = newStr.slice(0, -1);

                $(`[data-target='${target}']`).val(newStr);
                $(this).parent(".tag-text").remove();
            });

            function checkStrExist(str) {
                var originalStr = $(`[data-target='${target}']`).val().split(settings.splitChar);
                var isExist = false;
                for (var i = 0; i < originalStr.length; i++) {
                    if (originalStr[i] == str) isExist = true;
                }
                return isExist;
            }

            function resizable(el, factor) {
                var int = Number(factor) || 7.7;
                function resize() { el.style.width = ((el.value.length + 1) * int) + 'px' }
                var e = 'keyup,keypress,focus,blur,change'.split(',');
                for (var i in e) el.addEventListener(e[i], resize, false);
                resize();
            }
        }
    });
})(jQuery);