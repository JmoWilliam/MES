function customAlert() {
    if (window._alert) {
        return;
    }
    window._alert = window.alert;
    window.alert = function alert(message, alertType, timeOut) {
        alertType = (typeof alertType !== 'undefined') ? alertType : "error";
        timeOut = (typeof timeOut !== 'undefined') ? timeOut : 1000;

        message.replace(new RegExp("\r\n", "g"), "<br />");

        switch (alertType) {
            case "success":
                swal.fire({
                    icon: "success",
                    text: message,
                    showConfirmButton: false,
                    timer: timeOut
                });
                break;
            case "warning":
                swal.fire({
                    icon: "success",
                    text: message,
                    showConfirmButton: false,
                    timer: timeOut
                });
                break;
            case "info":
                swal.fire({
                    icon: "info",
                    text: message,
                    showConfirmButton: false,
                    timer: timeOut
                });
                break;
            default:
                swal.fire({
                    icon: "error",
                    text: message,
                    showConfirmButton: false,
                    timer: timeOut
                });
                break;
        }
    };
}

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
                        swal.fire({
                            icon: "success",
                            title: "新增完成"
                        });
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
                        swal.fire({
                            icon: "success",
                            title: '更新完成'
                        });
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
                        swal.fire({
                            icon: "success",
                            title: '刪除完成'
                        });
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
}

function LetterEncrypt(phrase, key, iv) {
    var encrypted = CryptoJS.AES.encrypt(phrase, key, {
        iv: iv,
        mode: CryptoJS.mode.CBC,
        padding: CryptoJS.pad.Pkcs7
    });

    return encrypted.ciphertext.toString(CryptoJS.enc.Hex);
}

function LetterDecrypt(phrase, key, iv) {
    let decrypted = CryptoJS.AES.decrypt({ ciphertext: CryptoJS.enc.Hex.parse(phrase) }, key, {
        iv: iv,
        mode: CryptoJS.mode.CBC,
        padding: CryptoJS.pad.Pkcs7
    });

    return decrypted.toString(CryptoJS.enc.Utf8);
}