//#region Date格式化
// 月(M)、日(d)、小時(h)、分(m)、秒(s)、季度(q) 可以用 1-2 个占位符， 
// 年(y)可以用 1-4 个占位符，毫秒(S)只能用 1 个占位符(是 1-3 位的数字) 
// 例子： 
// (new Date()).Format("yyyy-MM-dd hh:mm:ss.S") ==> 2006-07-02 08:09:04.423 
// (new Date()).Format("yyyy-M-d h:m:s.S")      ==> 2006-7-2 8:9:4.18 
Date.prototype.Format = function (fmt) { //author: meizz 
    var o = {
        "M+": this.getMonth() + 1,                 //月份 
        "d+": this.getDate(),                    //日 
        "h+": this.getHours(),                   //小时 
        "m+": this.getMinutes(),                 //分 
        "s+": this.getSeconds(),                 //秒 
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度 
        "S": this.getMilliseconds()             //毫秒 
    };
    if (/(y+)/.test(fmt))
        fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt))
            fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}
//#endregion 

//#region 月曆下拉設定
let dtSetting = {
    format: "Y-m-d",
    lang: "zh",
    theme: "", //leaf, vanilla, ryanair
    fx: true,
    fxMobile: true,
    modal: false,
    large: true,
    largeDefault: true,
    largeOnly: false,
    hideDay: false,
    hideMonth: false,
    hideYear: false,
    maxYear: new Date().getFullYear() + 5,
    minYear: new Date().getFullYear() - 80,
    jump: 5,
    lock: false,
    startFromMonday: false,
    maxDate: false,
    minDate: false,
    disabledDays: false, //指定日期不能選擇 12/30/2019,12/31/2019
    enabledDays: false, //指定日期才能選擇 12/30/2019,12/31/2019
    showOnlyEnabledDays: false //只顯示指定日期
}
//#endregion

//#region 時間下拉設定
let tmSetting = {
    format: "HH:mm",
    autoswitch: true,
    meridians: false,
    mousewheel: false,
    init_animation: "fadeIn",
    setCurrentTime: false,
    customClass: "custom-timedropper"
}
//#endregion

//#region CodeMirror設定
let cmSetting = {
    lineNumbers: true,
    mode: "text/html",
    theme: 'monokai',    
    styleActiveLine: true    
}
//#endregion

//#region RichText設定
let rtSetting = {
    mode: "document"
};
//#endregion

//#region Dhtmlx
let calendarSetting = {
    dateFormat: "%Y-%m-%d %H:%i",
    timePicker: true,
    timeFormat: 24,
    css: "dhx_widget--bordered"
};

let calendarNoTimeSetting = {
    dateFormat: "%Y-%m-%d",
    timePicker: false,
    css: "dhx_widget--bordered"
};

let eventChange = new Event("change");
//#endregion

//#region 取得Url特定參數
function getURLParameter(param) {
    if (!window.location.search) {
        return;
    }
    var m = new RegExp(param + '=([^&]*)').exec(window.location.search.substring(1));
    if (!m) {
        return;
    }
    return decodeURIComponent(m[1]);
}
//#endregion

//#region 移除Url特定參數
function removeSearchParam(search, name) {
    if (search[0] === '?') {
        search = search.substring(1);
    }
    var parts = search.split('&');
    var res = [];
    for (var i = 0; i < parts.length; i++) { var pair = parts[i].split('='); if (pair[0] === name) { continue; } res.push(parts[i]); } search = res.join('&'); if (search.length > 0) {
        search = '?' + search;
    }
    return search;
}
//#endregion

//#region 設定Url特定參數
function setSearchParam(search, name, value) {
    search = removeSearchParam(search, name);
    if (search === '') {
        search = '?';
    }
    if (search.length > 1) {
        search += '&';
    }
    return search + name + '=' + encodeURIComponent(value);
}
//#endregion

//#region 自定義警告視窗
function customAlert() {
    if (window._alert) {
        return;
    }
    window._alert = window.alert;
    window.alert = function alert(message, alertType, timeOut, position) {
        alertType = (typeof alertType !== 'undefined') ? alertType : "error";
        timeOut = (typeof timeOut !== 'undefined') ? timeOut : 1000;
        position = (typeof position !== 'undefined') ? position : toastrTopRight;

        message.replace(new RegExp("\r\n", "g"), "<br />");

        toastr.clear();
        toastr.options.timeOut = timeOut;
        toastr.options = position;

        switch (alertType) {
            case "success":
                toastr.success(message);
                break;
            case "warning":
                toastr.warning(message);
                break;
            case "info":
                toastr.info(message);
                break;
            default:
                toastr.error(message);
                break;
        }
    };
}
//#endregion
//#region JS Overload Method
var overLoadingList = new Array();

function addMethod(object, name, fn) {
    if (object[name] == undefined) {
        object[name] = function () {
            for (n in overLoadingList) {
                if (overLoadingList[n]["Name"] == name) {
                    for (f in overLoadingList[n]["Func"]) {
                        if (overLoadingList[n]["Func"][f].length == arguments.length) {
                            return overLoadingList[n]["Func"][f].apply(null, arguments);
                        }
                    }
                }
            }
        }
    }

    var isExist = false;
    for (n in overLoadingList) {
        if (overLoadingList[n]["Name"] == name) {
            overLoadingList[n]["Func"].push(fn);
            isExist = true;
            break;
        }
    }

    if (!isExist) {
        var fnList = new Array();
        fnList.push(fn);
        overLoadingList.push(overLoading(name, fnList));
    }
}

function overLoading(fname, fnList) {
    return { Name: fname, Func: fnList };
}
//#endregion

//#region 下拉選單
function ComboBoxControl() {
    //預設值
    addMethod(this, "Default", function (targetSelector, targetUrl, postData, selectedValue, selectedType, textFlag, optionText, optionValue) {
        $.ajax({
            url: targetUrl,
            type: "POST",
            async: false,
            dataType: "json",
            data: postData,
            beforeSend: function (XMLHttpRequest) {
            },
            success: function (data) {
                $(targetSelector).empty();

                switch (selectedType) {
                    case "ALL":
                        $(targetSelector).append(new Option("全部", "-1"));
                        break;
                    case "NA":
                        $(targetSelector).append(new Option("請選擇", ""));
                        break;
                    case "CHOOSE":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, ""));
                        break;
                    case "CHOOSE-NUMBER":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, "-1"));
                        break;
                    case "SP":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, "-999"));
                        break;
                    case "HIDDEN":
                        if (textFlag) {
                            let innerHtml = '<option value="" hidden>' + textFlag + '</option>';
                            $(targetSelector).append(innerHtml);
                        }
                        break;
                }

                if (data.status == "success") {
                    $.each(data.result, function (i, item) {
                        let innerHtml = '<option value="' + item[optionValue] + '" data-text="' + item[optionText] + '">' + item[optionText] + '</option>';
                        $(targetSelector).append(innerHtml);
                    });
                } else {
                    alert(data.msg, "alert", 0);
                }

                if (selectedValue) $(targetSelector).val(selectedValue);
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
            }
        });
    });

    //顯示文字加上值
    addMethod(this, "WithText", function (targetSelector, targetUrl, postData, selectedValue, selectedType, textFlag, optionText, optionValue) {
        $.ajax({
            url: targetUrl,
            type: "POST",
            async: false,
            dataType: "json",
            data: postData,
            beforeSend: function (XMLHttpRequest) {
            },
            success: function (data, textStatus, jqXHR) {
                $(targetSelector).empty();

                switch (selectedType) {
                    case "ALL":
                        $(targetSelector).append(new Option("全部", "-1"));
                        break;
                    case "NA":
                        $(targetSelector).append(new Option("請選擇", ""));
                        break;
                    case "CHOOSE":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, ""));
                        break;
                    case "CHOOSE-NUMBER":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, "-1"));
                        break;
                    case "SP":
                        if (textFlag) $(targetSelector).append(new Option(textFlag, "-999"));
                        break;
                    case "HIDDEN":
                        if (textFlag) {
                            let innerHtml = '<option value="" hidden>' + textFlag + '</option>';
                            $(targetSelector).append(innerHtml);
                        }
                        break;
                }

                if (data.status == "success") {
                    $.each(data.result, function (i, item) {
                        let innerHtml = '<option value="' + item[optionValue] + '" data-text="' + item[optionText] + '">' + item[optionValue] + " " + item[optionText] + '</option>';
                        $(targetSelector).append(innerHtml);
                    });
                } else {
                    alert(data.msg, "alert", 0);
                }

                if (selectedValue) $(targetSelector).val(selectedValue);
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
            }
        });
    });
}
//#endregion

//#region 分頁
function Pager(targetSelector, pageCount, objPage, func) {
    $(targetSelector).pagination(pageCount, {
        callback: function (page_id, jq) {
            objPage.pageIndex = page_id;
            func(objPage);
        },
        first_text: "",
        prev_text: "上一頁",
        next_text: "下一頁",
        last_text: "",
        items_per_page: objPage.pageSize,
        num_display_entries: 3,
        current_page: objPage.pageIndex,
        num_edge_entries: 1
    });
}
//#endregion

//#region Form 重置
function ResetForm(targetSelector) {
    var form = $(targetSelector);
    form.validate().resetForm();
    form[0].reset();
}
//#endregion

//#region 登出
function Logout() {
    $.ajax({
        url: "/User/Logout",
        type: "POST",
        async: false,
        dataType: "json",
        data: {},
        beforeSend: function (XMLHttpRequest) {
            $("#loaderWrpper").show();
        },
        success: function (data, textStatus, jqXHR) {
            if (data.status != "success") {
                alert(data.msg);
            }
            else {
                location.reload();
            }
        },
        complete: function (XMLHttpRequest, textStatus) {
            if (textStatus == 'timeout') {
                var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                xmlhttp.abort();
                alert("網路逾時!");
            }

            $("#loaderWrpper").hide();
        },
        error: function (XMLHttpRequest, textStatus) {
            alert(XMLHttpRequest.statusText);
        }
    });
}
//#endregion

//#region element元素切換
function Switch(showSelector, hideSelector) {
    $(showSelector).show();
    $(hideSelector).hide();
}
//#endregion
//#region Table元件初始化
function TableComponentInit() {
    //$("[data-toggle='tooltip']").tooltip({
    //    html: false,
    //    placement: "right",
    //    trigger: "click hover",
    //    sanitize: false
    //});

    PageAuthorityControl();
}
//#endregion

//#region 資料頁數計算
function PageCaculate(rowCount, pageObject) {
    let totalPage = parseInt(rowCount / pageObject.pageSize);
    if ((rowCount % pageObject.pageSize) != 0) totalPage++;
    if (pageObject.pageIndex >= totalPage) pageObject.pageIndex = (totalPage - 1);

    return pageObject;
}
//#endregion

//#region 資料存取
function DataAccessControl() {
    addMethod(this, "Get", function (targetUrl, postData, asynchronous) {
        let result;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料載入中...");
            },
            success: function (data) {
                result = data;
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "GetAPI", function (targetUrl, postData, asynchronous) {
        let result;

        $.ajax({
            url: targetUrl,
            type: "Get",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料載入中...");
            },
            success: function (data) {
                result = data;
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "Add", function (targetUrl, postData, asynchronous, successAlert) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料寫入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = true;

                    if (successAlert) {
                        alert("新增完成", "success");
                    }
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });    

    addMethod(this, "Update", function (targetUrl, postData, asynchronous, successAlert) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料寫入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = true;

                    if (successAlert) {
                        alert("更新完成", "success");
                    }
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "Delete", function (targetUrl, postData, asynchronous, successAlert) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料刪除中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = true;

                    if (successAlert) {
                        alert("刪除完成", "success");
                    }
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "EditContinued", function (targetUrl, postData, asynchronous) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料載入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = data;
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) { 
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "Other", function (targetUrl, postData, asynchronous, successAlert) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
            },
            success: function (data) {
                if (data.status == "success") {
                    result = true;

                    if (successAlert) {
                        alert(data.msg, "success");
                    }
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });

    addMethod(this, "EditPicking", function (targetUrl, postData, asynchronous, alertMsg, passUrl) {
        let result = false;

        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料寫入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = true;

                    swal.fire({
                        icon: "success",
                        title: alertMsg,
                        confirmButtonColor: '#0041c4',
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = passUrl;
                        }
                    });
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });
}
//#endregion

//#region 複製剪貼簿
function CopyToClipboardFF(text) {
    window.prompt("Copy to clipboard: Ctrl C, Enter", text);
}

function CopyToClipboard(targetSelector) {
    var success = true,
        range = document.createRange(),
        selection;

    // For IE.
    if (window.clipboardData) {
        window.clipboardData.setData("Text", $(targetSelector).val());
    } else {
        // Create a temporary element off screen.
        var tmpElem = $('<div>');
        tmpElem.css({
            position: "absolute",
            left: "-1000px",
            top: "-1000px",
        });
        // Add the input value to the temp element.
        tmpElem.text($(targetSelector).val());
        $("body").append(tmpElem);
        // Select temp element.
        range.selectNodeContents(tmpElem.get(0));
        selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        // Lets copy.
        try {
            success = document.execCommand("copy", false, null);
        }
        catch (e) {
            CopyToClipboardFF($(targetSelector).val());
        }
        if (success) {
            alert("已複製至剪貼簿", "success");
            // remove temp element.
            tmpElem.remove();
        }
    }
}
//#endregion

//#region 檔案存取
function FileAccessControl() {
    addMethod(this, "Download", function (targetFile) {
        if (targetFile) {
            let files = targetFile.split(";");

            for (var i = 0; i < files.length; i++) {
                let virtualLink = document.createElement("a");
                virtualLink.download = '';
                virtualLink.style.display = 'none';
                virtualLink.href = files[i];
                document.body.appendChild(virtualLink);
                virtualLink.click();
                document.body.removeChild(virtualLink);
            }
        } else {
            alert("沒有檔案可以下載!");
        }
    });

    addMethod(this, "AjaxDownload", function (targetUrl, postData, asynchronous) {
        $.ajax({
            url: targetUrl,
            type: "POST",
            async: asynchronous,
            data: postData,
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料載入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    let parameter = {
                        fileGuid: data.fileGuid,
                        fileName: data.fileName,
                        fileExtension: data.fileExtension,
                    };

                    let virtualLink = document.createElement("a");
                    virtualLink.download = '';
                    virtualLink.style.display = 'none';
                    virtualLink.href = "/Web/Download?" + new URLSearchParams(parameter).toString();
                    document.body.appendChild(virtualLink);
                    virtualLink.click();
                    document.body.removeChild(virtualLink);
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });
    });

    addMethod(this, "Preview", function (targetFile) {
        let result;

        $.ajax({
            url: "/Web/Preview",
            type: "POST",
            async: false,
            data: {
                previewFile: targetFile
            },
            dataType: "json",
            beforeSend: function (XMLHttpRequest) {
                //console.log("資料載入中...");
            },
            success: function (data) {
                if (data.status == "success") {
                    result = data;
                } else {
                    alert(data.msg, "error", 0);
                }
            },
            complete: function (XMLHttpRequest, textStatus) {
                if (textStatus == 'timeout') {
                    var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                    xmlhttp.abort();
                    alert("網路逾時!");
                }
            },
            error: function (XMLHttpRequest, textStatus) {
                alert(XMLHttpRequest.statusText, "error", 0);
            }
        });

        return result;
    });
}
//#endregion

//#region 狀態名稱
function StatusName(statusSchema, statusNo) {
    let statusName = "";

    $.ajax({
        url: "/BasicInformation/GetStatus",
        type: "POST",
        async: false,
        data: {
            "StatusSchema": statusSchema,
            "StatusNo": statusNo
        },
        dataType: "json",
        beforeSend: function (XMLHttpRequest) {
            //console.log("資料載入中...");
        },
        success: function (data) {
            if (data.status == "success") {
                $.each(data.result, function (i, item) {
                    statusName = item.ApplyStatusName;
                });
            } else {
                statusName = "系統錯誤：" + data.msg;
            }
        },
        complete: function (XMLHttpRequest, textStatus) {
            if (textStatus == 'timeout') {
                var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                xmlhttp.abort();
                alert("網路逾時!");
            }
        },
        error: function (XMLHttpRequest, textStatus) {
            alert(XMLHttpRequest.statusText, "error", 0);
        }
    });

    return statusName;
}
//#endregion

//#region 類別名稱
function TypeName(typeSchema, typeNo) {
    let typeName = "";

    $.ajax({
        url: "/BasicInformation/GetType",
        type: "POST",
        async: false,
        data: {
            "TypeSchema": typeSchema,
            "TypeNo": typeNo
        },
        dataType: "json",
        beforeSend: function (XMLHttpRequest) {
            //console.log("資料載入中...");
        },
        success: function (data) {
            if (data.status == "success") {
                $.each(data.result, function (i, item) {
                    typeName = item.ApplyTypeName;
                });
            } else {
                statusName = "系統錯誤：" + data.msg;
            }
        },
        complete: function (XMLHttpRequest, textStatus) {
            if (textStatus == 'timeout') {
                var xmlhttp = window.XMLHttpRequest ? new window.XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHttp");
                xmlhttp.abort();
                alert("網路逾時!");
            }
        },
        error: function (XMLHttpRequest, textStatus) {
            alert(XMLHttpRequest.statusText, "error", 0);
        }
    });

    return typeName;
}
//#endregion

//#region 頁面權限控管
function PageAuthorityControl() {
    let temp = $("#pageAuthority").val();

    $("[class*='auth-']:not([class*='proj-auth-'])").hide();

    if (temp) {
        let currentAuthority = JSON.parse(temp);
        $(".un-auth").show();

        for (var i = 0; i < currentAuthority.length; i++) {
            let className = ".auth-" + currentAuthority[i].DetailCode;

            if (currentAuthority[i].Authority == 1) {
                $(className).show();
                $(".un-auth" + className).hide();
            } else {
                $(className).hide();
                $(".un-auth" + className).show();
            }
        }
    }
}

function ProjectAuthorityControl() {
    let temp = $("#projectAuthority").val();

    $("[class*='proj-auth-']").hide();

    if (temp) {
        let currentAuthority = JSON.parse(temp);

        for (var i = 0; i < currentAuthority.length; i++) {
            let className = ".proj-auth-" + currentAuthority[i].AuthorityCode;

            if (currentAuthority[i].Authority == 1) {
                $(className).show();
                $(".un-auth" + className).hide();
            } else {
                $(className).hide();
                $(".un-auth" + className).show();
            }
        }
    }
}
//#endregion

//#region 彈跳視窗重複彈出設定
function MultileModal() {
    $(document).off("show.bs.modal").on("show.bs.modal", ".modal", function () {
        let currentZIndex = 1050 + 10 * $(".modal:visible").length;

        $(this).css("z-index", currentZIndex);

        setTimeout(() => $(".modal-backdrop").not(".modal-stack").css("z-index", currentZIndex - 5).addClass("modal-stack"));
    });
}
//#endregion

//#region 套件初始化
function SuiteControl() {
    addMethod(this, "CodeMirror", function (targetSelector) {
        var cmSelector = document.getElementById(targetSelector);

        var targetInputContent = "";
        if ($("#" + targetSelector).val()) {
            targetInputContent = $("#" + targetSelector).val();
        }

        if (cmEditor) {
            cmEditor.toTextArea();
            cmEditor.setValue(targetInputContent);
        }

        cmEditor = CodeMirror.fromTextArea(cmSelector, cmSetting);
        cmEditor.on("change", function () {
            cmEditor.save();
            clearTimeout(delay);
            delay = setTimeout(UpdatePreview, 300);
        });

        UpdatePreview();
    });

    addMethod(this, "FileUpload", function (config) {
        let uploadSetting = {
            theme: "fas",
            language: "zh-TW",
            showCaption: true,
            showPreview: true,
            showRemove: true,
            showUpload: true,
            uploadUrl: "/Web/Upload",
            uploadAsync: true,
            minFileCount: (config.required ? 1 : 0),
            maxFileCount: config.maxFileCount,
            validateInitialCount: true,
            overwriteInitial: false,
            initialPreviewAsData: true,
            browseOnZoneClick: true,
            allowedFileExtensions: config.allowedFileExtensions,
            uploadExtraData: {
                savePath: config.uploadPath,
                isPreview: false
            },
            previewFileIconSettings: {
                doc: '<i class="fas fa-file-word" style="color: #0078d7;"></i>',
                docx: '<i class="fas fa-file-word" style="color: #0078d7;"></i>',
                xls: '<i class="fas fa-file-excel" style="color: #1d6f42;"></i>',
                xlsx: '<i class="fas fa-file-excel" style="color: #1d6f42;"></i>',
                csv: '<i class="fas fa-file-csv" style="color: #1d6f42;"></i>',
                ppt: '<i class="fas fa-file-powerpoint" style="color: #d04423"></i>',
                pptx: '<i class="fas fa-file-powerpoint" style="color: #d04423"></i>'
            }
        };

        let previewData;
        if (config.defaultFile) {
            previewData = FileAccess.Preview(config.defaultFile);
        }

        if (previewData) {
            uploadSetting.initialPreview = previewData.initialPreview;
            uploadSetting.initialPreviewConfig = previewData.initialPreviewConfig;
        }

        $(config.fileSelector).fileinput('destroy');
        $(config.fileSelector).fileinput(uploadSetting)
            .off('fileuploaded')
            .on('fileuploaded', function (event, previewId, index, fileId) {
                let uploadPath = previewId.response.initialPreviewConfig[0]["key"];

                tempFile = [];
                if ($(config.targetSelector).val()) tempFile = $(config.targetSelector).val().split(',');
                tempFile.push(uploadPath);

                //$(config.targetSelector).val(tempFile.join(","));

                ElementSetValue(config.targetSelector, tempFile.join(","), "change");
            })
            .off('filedeleted')
            .on('filedeleted', function (event, key, jqXHR, data) {
                tempFile = [];
                if ($(config.targetSelector).val()) tempFile = $(config.targetSelector).val().split(',');

                let deletePathIndex = tempFile.indexOf(key.toString());
                if (deletePathIndex > -1) tempFile.splice(deletePathIndex, 1);

                //$(config.targetSelector).val(tempFile.join(","));

                ElementSetValue(config.targetSelector, tempFile.join(","), "change");
            });
    });

    addMethod(this, "FileAutoUpload", function (config) {
        let uploadSetting = {
            theme: "fas",
            language: "zh-TW",
            showCaption: true,
            showPreview: true,
            showRemove: false,
            showUpload: false,
            showCancel: false,
            uploadUrl: "/Web/Upload",
            uploadAsync: true,
            minFileCount: (config.required ? 1 : 0),
            maxFileCount: config.maxFileCount,
            validateInitialCount: true,
            overwriteInitial: false,
            initialPreviewAsData: true,
            browseOnZoneClick: true,
            uploadExtraData: {
                savePath: config.uploadPath,
                isPreview: false
            },
            previewFileIconSettings: {
                doc: '<i class="fas fa-file-word" style="color: #0078d7;"></i>',
                docx: '<i class="fas fa-file-word" style="color: #0078d7;"></i>',
                xls: '<i class="fas fa-file-excel" style="color: #1d6f42;"></i>',
                xlsx: '<i class="fas fa-file-excel" style="color: #1d6f42;"></i>',
                csv: '<i class="fas fa-file-csv" style="color: #1d6f42;"></i>',
                ppt: '<i class="fas fa-file-powerpoint" style="color: #d04423"></i>',
                pptx: '<i class="fas fa-file-powerpoint" style="color: #d04423"></i>'
            }
        };

        $(config.fileSelector).fileinput('destroy');
        $(config.fileSelector).fileinput(uploadSetting)
            .off('filebatchselected')
            .on('filebatchselected', function (event, files) {
                $(config.fileSelector).fileinput("upload");
            })
            .off('fileuploaded')
            .on('fileuploaded', function (event, previewId, index, fileId) {
                let uploadPath = previewId.response.initialPreviewConfig[0]["key"];

                tempFile = [];
                if ($(config.targetSelector).val()) tempFile = $(config.targetSelector).val().split(',');
                tempFile.push(uploadPath);

                $(config.targetSelector).val(tempFile.join(","));

                $(config.fileSelector).fileinput('clear');
            });
    });

    addMethod(this, "RichText", function (targetSelector, inputSelector, defaultValue) {
        dhx.i18n.setLocale("richtext", richTextLocales);
        if (rtEditor) rtEditor.destructor();

        rtEditor = new dhx.Richtext(targetSelector, {
            mode: "document",
            toolbarBlocks: [
                "undo", "style", "decoration", "colors", "align", "link", "fullscreen"
            ]
        });

        rtEditor.setValue(defaultValue, "html");

        rtEditor.events.on("Change", function (action) {
            $("#" + inputSelector).val(rtEditor.getValue("html"));
        });
    });

    addMethod(this, "TimeDropper", function (targetSelector) {
        $(".td-wrap").remove();
        $(targetSelector).timeDropper(tmSetting);
    });

    addMethod(this, "DateDropper", function (targetSelector) {
        $(targetSelector).dateDropper('destroy');
        $(targetSelector).dateDropper(dtSetting);
    });

    addMethod(this, "SetDateDropper", function (targetSelector, setDate) {
        $(targetSelector).val(setDate.Format("yyyy-MM-dd"));
        $(targetSelector).dateDropper("setDate", { d: setDate.Format("dd"), m: setDate.Format("MM"), y: setDate.Format("yyyy") });
    });

    addMethod(this, "DatePicker", function (targetSelector, defaultDate) {
        dhx.i18n.setLocale("calendar", calendarLocales);

        let calendar = new dhx.Calendar(null, calendarNoTimeSetting);
        let popup = new dhx.Popup();
        let dateInput = document.getElementById(targetSelector);

        popup.attach(calendar);

        if (defaultDate) {
            calendar.setValue(defaultDate);
            dateInput.value = calendar.getValue();
        }

        calendar.events.on("change", function () {
            dateInput.value = calendar.getValue();
            dateInput.dispatchEvent(eventChange);
            popup.hide();
        });

        let func = function () {
            popup.show(dateInput);
        };

        dateInput.removeEventListener("click", func, false);
        dateInput.addEventListener("click", func, false);
    });

    addMethod(this, "TimePicker", function (targetSelector, defaultTime) {
        dhx.i18n.setLocale("timepicker", timepickerLocales);

        let clock = new dhx.Timepicker();
        let popup = new dhx.Popup();
        let timeInput = document.getElementById(targetSelector);
        
        popup.attach(clock);

        if (defaultTime) {
            clock.setValue(defaultTime);
            timeInput.value = clock.getValue();
        }

        clock.events.on("change", function () {
            timeInput.value = clock.getValue();
            timeInput.dispatchEvent(eventChange);
        });

        let func = function () {
            popup.show(timeInput);
        };

        timeInput.removeEventListener("click", func, false);
        timeInput.addEventListener("click", func, false);
    });

    function UpdatePreview() {
        var previewFrame = document.getElementById('iframePreview');
        var preview = previewFrame.contentDocument || previewFrame.contentWindow.document;

        var templateHtml = '<link href="/Content/fontawesome.pro/5.13.0/css/all.min.css" rel="stylesheet" />';

        preview.open();
        preview.write(templateHtml + '<span style="font-size: 2.5rem; color: #fff; text-align: center; display: block;">' + cmEditor.getValue() + '</span>');
        preview.close();

        preview.body.style.backgroundColor = "#1d5697";
    }
}
//#endregion

//#region 檔案大小位元轉換
function FormatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];

    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}
//#endregion

//#region 彈跳視窗初始化
function InitPopView(popConfig) {
    $(popConfig.targetSelector).siblings().find(".pop-view").off("click").on("click", function () {
        if (typeof window[popConfig.popFunction] === "function") window[popConfig.popFunction](popConfig);
    });

    $(popConfig.targetSelector).siblings().find(".pop-clear").off("click").on("click", function () {
        $(popConfig.targetSelector).parent().find(".help-block").hide();
        $(popConfig.targetSelector).parent().siblings().find(".help-block").hide();
        ResetElement(popConfig.targetSelector);
        ResetElement($(popConfig.targetSelector).siblings(".help-block").find("dd"));
    });
}
//#endregion

//#region 單一element清除
function ResetElement(targetSelector) {
    switch ($(targetSelector).prop("nodeName")) {
        case "INPUT":
            if ($(targetSelector).prop("type") == "number") {
                $(targetSelector).val(0);
            } else {
                $(targetSelector).val("");
            }
            break;
        case "SELECT":
            $(targetSelector).prop("checked", false);
            if ($(targetSelector).prop("multiple")) {
                $(targetSelector).prop("selectedIndex", -1).trigger("change");
            } else {
                $(targetSelector).prop("selectedIndex", 0);
            }
            break;
        default:
            $(targetSelector).html("");
            break;
    }
}
//#endregion

//#region element設定值
function ElementSetValue(targetSelector, value, eventType) {
    switch ($(targetSelector).prop("nodeName")) {
        case "INPUT":
        case "SELECT":
            $(targetSelector).val(value).trigger(eventType);
            break;
        default:
            $(targetSelector).html(value);
            break;
    }
}
//#endregion

//#region function延遲執行
function DelayFn(passFunction, millisecond) {
    setTimeout(() => {
        passFunction();
    }, millisecond);
}

function DelayTrigger(passFunction, millisecond) {
    let timer = 0
    return function (...args) {
        clearTimeout(timer)
        timer = setTimeout(passFunction.bind(this, ...args), millisecond || 0)
    }
}
//#endregion

//#region 取得元件位置
function GetElementPosition(targetSelector) {
    return $(targetSelector)[0].getBoundingClientRect();
}
//#endregion

//#region 移動特定位置
function ScrollPosition(xCoordinate, yCoordinate) {
    window.scroll(xCoordinate, yCoordinate + $(window).scrollTop());
}
//#endregion

//#region Json routing
function getObjects(obj, key, val) {
    var objects = [];
    for (var i in obj) {
        if (!obj.hasOwnProperty(i)) continue;
        if (typeof obj[i] == 'object') {
            objects = objects.concat(getObjects(obj[i], key, val));
        } else
            if (i == key && obj[i] == val || i == key && val == '') { //
                objects.push(obj);
            } else if (obj[i] == val && key == '') {
                if (objects.lastIndexOf(obj) == -1) {
                    objects.push(obj);
                }
            }
    }
    return objects;
}

function getValues(obj, key) {
    var objects = [];
    for (var i in obj) {
        if (!obj.hasOwnProperty(i)) continue;
        if (typeof obj[i] == 'object') {
            objects = objects.concat(getValues(obj[i], key));
        } else if (i == key) {
            objects.push(obj[i]);
        }
    }
    return objects;
}

function getKeys(obj, val) {
    var objects = [];
    for (var i in obj) {
        if (!obj.hasOwnProperty(i)) continue;
        if (typeof obj[i] == 'object') {
            objects = objects.concat(getKeys(obj[i], val));
        } else if (obj[i] == val) {
            objects.push(i);
        }
    }
    return objects;
}
//#endregion

//#region 倒數
function Countdown(distance, target) {
    let timePassed = 0;
    timeLeft = distance * 60;

    clearInterval(timerInterval);
    timerInterval = setInterval(function () {
        timePassed = timePassed + 1;
        timeLeft = (distance * 60) - timePassed;

        var days = Math.floor(timeLeft / (60 * 60 * 24));
        var hours = Math.floor(timeLeft % (60 * 60 * 24) / (60 * 60));
        var minutes = Math.floor(timeLeft % (60 * 60) / 60);
        var seconds = Math.floor(timeLeft % 60);

        document.getElementById(target).innerHTML = days + "天 " + hours + "小時 " + minutes + "分 " + seconds + "秒 ";

        if (timeLeft < 0) {
            clearInterval(timerInterval);
            document.getElementById(target).innerHTML = "已到期";
        }
    }, 1000);
}
//#endregion

//#region ERP單據狀態
function ErpDocStatus(status) {
    switch (status) {
        case "Y":
            return `<span class="erp-status"><span>核</span></span>`;
            break;
        case "N":
            return ``;
            break;
        case "V":
            return `<span class="erp-status"><span>廢</span></span>`;
            break;
        default:
            return ``;
    }
}
//#endregion

//#region File Preview
async function FilePreview(url, target) {
    let html = ``;

    await fetch(url)
        .then(res => {
            test(res.headers.get('Content-Type'));
        });

    //let response = await fetch(url);
    //let contentType = response.headers.get('Content-Type');
    //console.log($(this));

    //test(contentType);

    function test(str) {
        switch (true) {
            case /image/.test(str):
                html = `<img src="${url}" />`;
                break;
            default:
                html = `<a class="active-link" href="${url}"><i class="fas fa-download"></i><span>檔案下載</span></a>`
                break;
        }
    }

    $(`span[data-target=${target}]`).html(html);
}
//#endregion

//#region 時間轉換
function TimeTransform(distance) {
    var days = Math.floor(distance / 60 / 60 / 24);
    var hours = Math.floor(distance / (60 * 60) % 24);
    var minutes = Math.floor(distance / 60 % 60);
    var seconds = Math.floor(distance % 60);

    return `${days > 0 ? `${days}天 ` : ``}${hours > 0 ? `${hours}小時 ` : ``}${minutes > 0 ? `${minutes}分 ` : ``}${seconds > 0 ? `${seconds}秒 ` : ``}`;
}
//#endregion