using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ChatRoomController : WebController
    {
        // GET: ChatRoom
        #region// UC API test
        private readonly string ezucServerUrl = "https://192.168.110.203:8443";

        #region //SendMessage
        [HttpPost]
        [Route("api/MES/apiSendMessage")]
        public void ApiSendMessage(
            string Company = "",
            string SecretKey = "",
            string apiKey = "",
            string rcvType = "employee",
            string rcvList = "",
            string empRcvListItemType = "",
            string deptRcvListItemType = "",
            string message = "")
        {
            JObject jsonResponse;

            try
            {
                // 設定回應類型
                Response.ContentType = "application/json";

                // 記錄開始處理
                if (logger != null)
                {
                    logger.Info($"開始處理 ApiSendMessage - Company: {Company}");
                }

                #region //Api金鑰驗證
                try
                {
                    ApiKeyVerify(Company, SecretKey, "SendMessage");
                }
                catch (Exception ex)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = $"API 金鑰驗證失敗: {ex.Message}"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }
                #endregion

                #region //組成 para 參數
                // 解析 rcvList (預期是 JSON 陣列格式的字串)
                JArray rcvListArray;
                try
                {
                    if (!string.IsNullOrEmpty(rcvList))
                    {
                        rcvListArray = JArray.Parse(rcvList);
                    }
                    else
                    {
                        rcvListArray = new JArray();
                    }
                }
                catch (Exception ex)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = $"rcvList 格式不正確，請提供 JSON 陣列格式: {ex.Message}"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 組成 para JObject
                var paraObj = new JObject
                {
                    ["apiKey"] = apiKey,
                    ["rcvType"] = rcvType,
                    ["rcvList"] = rcvListArray,
                    ["message"] = message
                };

                // 根據 rcvType 加入對應的 ItemType
                if (rcvType == "employee" && !string.IsNullOrEmpty(empRcvListItemType))
                {
                    paraObj["empRcvListItemType"] = empRcvListItemType;
                }
                else if (rcvType == "department" && !string.IsNullOrEmpty(deptRcvListItemType))
                {
                    paraObj["deptRcvListItemType"] = deptRcvListItemType;
                }
                #endregion

                // 驗證必要參數
                if (string.IsNullOrEmpty(apiKey))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "缺少必要參數：apiKey"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                if (rcvListArray == null || rcvListArray.Count == 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "缺少必要參數：rcvList 或收件者名單為空"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 檢查是否有上傳檔案
                HttpPostedFileBase uploadFile = null;
                if (Request.Files.Count > 0)
                {
                    uploadFile = Request.Files["upload"];
                }

                // 驗證訊息內容或檔案必須有其中一個（也可以同時提供）
                if (string.IsNullOrEmpty(message) && (uploadFile == null || uploadFile.ContentLength == 0))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "訊息內容(message)或檔案(upload)必須提供其中一個，也可以同時提供"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 檔案大小檢查（50MB 限制，根據 EZUC+ API 文件）
                if (uploadFile != null && uploadFile.ContentLength > 50 * 1024 * 1024)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "檔案大小超過限制（50MB）"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 訊息字數檢查（根據 EZUC+ API 文件，上限 4000 字）
                if (!string.IsNullOrEmpty(message) && message.Length > 4000)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "訊息內容超過字數限制（4000字）"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                if (rcvType == "employee" && string.IsNullOrEmpty(empRcvListItemType))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "當 rcvType 為 employee 時，empRcvListItemType 為必要參數"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 呼叫支援檔案上傳的 EZUC+ API
                var ezucResult = CallEzucSendMessageApiWithUpload(paraObj, uploadFile);

                // 根據 EZUC+ API 回應處理結果
                if (ezucResult != null)
                {
                    var returnCode = ezucResult["returnCode"]?.Value<int>() ?? 999;
                    var returnInfo = ezucResult["returnInfo"]?.ToString() ?? "未知錯誤";

                    if (returnCode == 0)
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "訊息發送成功",
                            data = ezucResult
                        });
                    }
                    else
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = $"EZUC+ API 錯誤 (代碼: {returnCode}): {returnInfo}",
                            errorCode = returnCode,
                            data = ezucResult
                        });
                    }
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "呼叫 EZUC+ API 失敗，未收到回應"
                    });
                }
            }
            catch (Exception e)
            {
                // 詳細錯誤記錄
                if (logger != null)
                {
                    logger.Error($"ApiSendMessage 發生異常: {e.Message}\n堆疊追蹤: {e.StackTrace}");
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = $"系統錯誤: {e.Message}",
                    stackTrace = e.StackTrace
                });
            }

            try
            {
                Response.Write(jsonResponse.ToString());
            }
            catch (Exception ex)
            {
                Response.Write($"{{\"status\":\"error\",\"msg\":\"Response write error: {ex.Message}\"}}");
            }
        }

        // 支援檔案上傳的 API 呼叫方法
        private JObject CallEzucSendMessageApiWithUpload(JObject parameters, HttpPostedFileBase uploadFile)
        {
            try
            {
                // 設定 EZUC+ 伺服器資訊                
                var apiUrl = $"{ezucServerUrl}/ucrm/api/chat/sendMessage";

                // 建立 HttpClientHandler 來處理 SSL 憑證問題
                var handler = new HttpClientHandler()
                {
                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
                };

                using (var httpClient = new HttpClient(handler))
                {
                    // 設定請求超時
                    httpClient.Timeout = TimeSpan.FromSeconds(30);

                    // 建立 MultipartFormDataContent 來同時傳送參數和檔案
                    using (var formContent = new MultipartFormDataContent())
                    {
                        // 加入 para 參數
                        var paraContent = new StringContent(parameters.ToString(Formatting.None));
                        formContent.Add(paraContent, "para");

                        // 如果有檔案，加入檔案內容
                        if (uploadFile != null && uploadFile.ContentLength > 0)
                        {
                            // 確保檔案流從開始位置讀取
                            uploadFile.InputStream.Seek(0, SeekOrigin.Begin);

                            // 讀取檔案內容
                            byte[] fileData = new byte[uploadFile.ContentLength];
                            uploadFile.InputStream.Read(fileData, 0, uploadFile.ContentLength);

                            var fileContent = new ByteArrayContent(fileData);
                            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(uploadFile.ContentType ?? "application/octet-stream");
                            formContent.Add(fileContent, "upload", uploadFile.FileName ?? "uploaded_file");
                        }

                        // 發送請求
                        var response = httpClient.PostAsync(apiUrl, formContent).Result;

                        if (response.IsSuccessStatusCode)
                        {
                            var responseContent = response.Content.ReadAsStringAsync().Result;

                            // 嘗試解析 JSON 回應
                            try
                            {
                                return JObject.Parse(responseContent);
                            }
                            catch (Exception)
                            {
                                // 如果不是 JSON 格式，返回原始內容
                                return JObject.FromObject(new
                                {
                                    returnCode = 999,
                                    returnInfo = "API 回應格式不正確",
                                    rawResponse = responseContent
                                });
                            }
                        }
                        else
                        {
                            var errorContent = response.Content.ReadAsStringAsync().Result;
                            return JObject.FromObject(new
                            {
                                returnCode = 999,
                                returnInfo = $"HTTP 請求失敗: {response.StatusCode} - {response.ReasonPhrase}",
                                errorContent = errorContent
                            });
                        }
                    }
                }
            }
            catch (Exception e)
            {
                if (logger != null)
                {
                    logger.Error($"呼叫 EZUC+ API 異常: {e.Message}");
                }

                return JObject.FromObject(new
                {
                    returnCode = 999,
                    returnInfo = $"系統錯誤: {e.Message}"
                });
            }
        }
        #endregion

        #region//SendBoardMessage

        private static readonly HttpClient httpClient = new HttpClient(new HttpClientHandler()
        {
            ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
        })
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        #region 整合公告建立 API - 改寫版本
        [HttpPost]
        [Route("api/MES/CreateBulletin")]
        public void CreateBulletin(
            string Company = "",
            string SecretKey = "",
            string apiKey = "",
            string type = "",
            string subject = "",
            string startTime = "",
            string endTime = "",
            string content = "")
        {
            JObject jsonResponse;

            try
            {
                // 設定回應類型
                Response.ContentType = "application/json";

                #region API金鑰驗證
                try
                {
                    ApiKeyVerify(Company, SecretKey, "CreateBulletin");
                }
                catch (Exception ex)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = $"API 金鑰驗證失敗: {ex.Message}"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }
                #endregion

                #region 組成 para 參數
                // 組成 para JObject
                var paraObj = new JObject
                {
                    ["apiKey"] = apiKey,
                    ["type"] = type,
                    ["subject"] = subject,
                    ["startTime"] = startTime,
                    ["endTime"] = endTime,
                    ["content"] = content
                };
                #endregion

                // 驗證必要參數
                var validationResult = ValidateBulletinParameters(paraObj);
                if (!validationResult.IsValid)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = validationResult.ErrorMessage
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 處理檔案上傳
                var uploadResult = ProcessFileUploads(paraObj);
                if (!uploadResult.IsSuccess)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = uploadResult.ErrorMessage
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                // 更新參數物件，加入上傳檔案的資訊
                if (uploadResult.SubjectPhotoCode != null)
                {
                    paraObj["subjectPhoto"] = JObject.FromObject(new
                    {
                        code = uploadResult.SubjectPhotoCode,
                        fileName = uploadResult.SubjectPhotoFileName
                    });
                }

                if (uploadResult.AttachmentCodes.Count > 0)
                {
                    var attachmentsArray = new JArray();
                    for (int i = 0; i < uploadResult.AttachmentCodes.Count; i++)
                    {
                        attachmentsArray.Add(JObject.FromObject(new
                        {
                            code = uploadResult.AttachmentCodes[i],
                            fileName = uploadResult.AttachmentFileNames[i]
                        }));
                    }
                    paraObj["attachments"] = attachmentsArray;
                }

                // 建立公告
                var ezucResult = CallEzucAddBulletinApi(paraObj);

                // 根據 EZUC+ API 回應處理結果
                if (ezucResult != null)
                {
                    var returnCode = ezucResult["returnCode"]?.Value<int>() ?? 999;
                    var returnInfo = ezucResult["returnInfo"]?.ToString() ?? "未知錯誤";

                    if (returnCode == 0)
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "公告建立成功",
                            data = new
                            {
                                bulletin = ezucResult,
                                uploadedFiles = new
                                {
                                    subjectPhoto = uploadResult.SubjectPhotoCode != null ? new
                                    {
                                        code = uploadResult.SubjectPhotoCode,
                                        fileName = uploadResult.SubjectPhotoFileName
                                    } : null,
                                    attachments = uploadResult.AttachmentCodes.Count > 0 ?
                                        uploadResult.AttachmentCodes.Select((code, index) => new {
                                            code = code,
                                            fileName = uploadResult.AttachmentFileNames[index]
                                        }).ToArray() : null
                                }
                            }
                        });
                    }
                    else
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = $"EZUC+ API 錯誤 (代碼: {returnCode}): {returnInfo}",
                            errorCode = returnCode,
                            data = ezucResult
                        });
                    }
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "呼叫 EZUC+ API 失敗，未收到回應"
                    });
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = $"系統錯誤: {e.Message}",
                    stackTrace = e.StackTrace
                });
            }

            try
            {
                Response.Write(jsonResponse.ToString());
            }
            catch (Exception ex)
            {
                Response.Write($"{{\"status\":\"error\",\"msg\":\"Response write error: {ex.Message}\"}}");
            }
        }

        #endregion

        #region 參數驗證
        private ValidationResult ValidateBulletinParameters(JObject paraObj)
        {
            var apiKey = paraObj["apiKey"]?.ToString();
            var type = paraObj["type"]?.ToString();
            var subject = paraObj["subject"]?.ToString();
            var startTime = paraObj["startTime"]?.ToString();
            var endTime = paraObj["endTime"]?.ToString();
            var content = paraObj["content"]?.ToString();

            if (string.IsNullOrEmpty(apiKey))
                return new ValidationResult { IsValid = false, ErrorMessage = "缺少必要參數：apiKey" };

            if (string.IsNullOrEmpty(type))
                return new ValidationResult { IsValid = false, ErrorMessage = "缺少必要參數：type (公告類別)" };

            if (string.IsNullOrEmpty(subject))
                return new ValidationResult { IsValid = false, ErrorMessage = "缺少必要參數：subject (公告主題)" };

            if (string.IsNullOrEmpty(startTime))
                return new ValidationResult { IsValid = false, ErrorMessage = "缺少必要參數：startTime (公告開始時間)" };

            if (string.IsNullOrEmpty(endTime))
                return new ValidationResult { IsValid = false, ErrorMessage = "缺少必要參數：endTime (公告結束時間)" };

            // 檢查是否有主題圖片檔案或內容
            bool hasSubjectPhoto = Request.Files["subjectPhoto"] != null && Request.Files["subjectPhoto"].ContentLength > 0;
            bool hasContent = !string.IsNullOrEmpty(content);

            if (!hasSubjectPhoto && !hasContent)
                return new ValidationResult { IsValid = false, ErrorMessage = "主題圖片或公告內容必須提供其中一個" };

            // 驗證主題長度
            if (subject.Length > 500)
                return new ValidationResult { IsValid = false, ErrorMessage = "公告主題不能超過 500 字元" };

            // 驗證內容長度
            if (!string.IsNullOrEmpty(content) && content.Length > 4000)
                return new ValidationResult { IsValid = false, ErrorMessage = "公告內容不能超過 4000 字元" };

            // 驗證時間格式
            DateTime startDateTime, endDateTime;
            if (!DateTime.TryParseExact(startTime, "yyyy-MM-dd HH:mm", null, System.Globalization.DateTimeStyles.None, out startDateTime))
                return new ValidationResult { IsValid = false, ErrorMessage = "開始時間格式不正確，應為 yyyy-MM-dd HH:mm" };

            if (!DateTime.TryParseExact(endTime, "yyyy-MM-dd HH:mm", null, System.Globalization.DateTimeStyles.None, out endDateTime))
                return new ValidationResult { IsValid = false, ErrorMessage = "結束時間格式不正確，應為 yyyy-MM-dd HH:mm" };

            if (startDateTime >= endDateTime)
                return new ValidationResult { IsValid = false, ErrorMessage = "開始時間必須早於結束時間" };

            return new ValidationResult { IsValid = true };
        }

        private class ValidationResult
        {
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; }
        }
        #endregion

        #region 檔案上傳處理
        private FileUploadResult ProcessFileUploads(JObject paraObj)
        {
            var result = new FileUploadResult { IsSuccess = true };

            try
            {
                // 處理主題圖片上傳
                var subjectPhoto = Request.Files["subjectPhoto"];
                if (subjectPhoto != null && subjectPhoto.ContentLength > 0)
                {
                    // 檢查主題圖片大小限制 (2MB)
                    if (subjectPhoto.ContentLength > 2 * 1024 * 1024)
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "主題圖片檔案大小不能超過 2MB";
                        return result;
                    }

                    // 檢查圖片格式
                    var allowedImageTypes = new[] { "image/jpeg", "image/jpg", "image/png", "image/gif" };
                    if (!allowedImageTypes.Contains(subjectPhoto.ContentType.ToLower()))
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "主題圖片格式不支援，僅支援 JPG、PNG、GIF 格式";
                        return result;
                    }

                    var photoUploadResult = CallEzucUploadSubjectPhotoApi(paraObj, subjectPhoto);
                    if (photoUploadResult != null && photoUploadResult["returnCode"]?.Value<int>() == 0)
                    {
                        result.SubjectPhotoCode = photoUploadResult["code"]?.ToString();
                        result.SubjectPhotoFileName = subjectPhoto.FileName;
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "主題圖片上傳失敗：" + (photoUploadResult?["returnInfo"]?.ToString() ?? "未知錯誤");
                        return result;
                    }
                }

                // 處理附件上傳（支援多個附件）
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    var file = Request.Files[i];
                    var fileKey = Request.Files.AllKeys[i];

                    // 跳過主題圖片，只處理附件
                    if (fileKey == "subjectPhoto" || file.ContentLength == 0)
                        continue;

                    // 檢查附件檔案名稱是否為attachment開頭
                    if (!fileKey.StartsWith("attachment"))
                        continue;

                    // 檢查附件大小限制 (15MB)
                    if (file.ContentLength > 15 * 1024 * 1024)
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = $"附件 {file.FileName} 檔案大小不能超過 15MB";
                        return result;
                    }

                    var attachmentUploadResult = CallEzucUploadAttachmentApi(paraObj, file);
                    if (attachmentUploadResult != null && attachmentUploadResult["returnCode"]?.Value<int>() == 0)
                    {
                        result.AttachmentCodes.Add(attachmentUploadResult["code"]?.ToString());
                        result.AttachmentFileNames.Add(file.FileName);
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = $"附件 {file.FileName} 上傳失敗：" + (attachmentUploadResult?["returnInfo"]?.ToString() ?? "未知錯誤");
                        return result;
                    }
                }

                // 檢查附件數量限制
                if (result.AttachmentCodes.Count > 5)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = "公告附件數量不能超過 5 個";
                    return result;
                }

                return result;
            }
            catch (Exception e)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"檔案上傳處理錯誤: {e.Message}";
                return result;
            }
        }

        private class FileUploadResult
        {
            public bool IsSuccess { get; set; }
            public string ErrorMessage { get; set; }
            public string SubjectPhotoCode { get; set; }
            public string SubjectPhotoFileName { get; set; }
            public List<string> AttachmentCodes { get; set; } = new List<string>();
            public List<string> AttachmentFileNames { get; set; } = new List<string>();
        }
        #endregion

        #region EZUC+ API 呼叫方法
        private JObject CallEzucUploadSubjectPhotoApi(JObject parameters, HttpPostedFileBase uploadFile)
        {
            try
            {
                var apiUrl = $"{ezucServerUrl}/ucrm/api/bulletin/uploadSubjectPhoto";

                using (var formContent = new MultipartFormDataContent())
                {
                    // 加入 para 參數
                    var paraContent = new StringContent(parameters.ToString(Formatting.None));
                    formContent.Add(paraContent, "para");

                    // 加入圖片檔案
                    uploadFile.InputStream.Seek(0, SeekOrigin.Begin);
                    byte[] fileData = new byte[uploadFile.ContentLength];
                    uploadFile.InputStream.Read(fileData, 0, uploadFile.ContentLength);

                    var fileContent = new ByteArrayContent(fileData);
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(uploadFile.ContentType ?? "image/jpeg");
                    formContent.Add(fileContent, "upload", uploadFile.FileName ?? "subject_photo.jpg");

                    // 發送請求
                    var response = httpClient.PostAsync(apiUrl, formContent).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = response.Content.ReadAsStringAsync().Result;
                        try
                        {
                            return JObject.Parse(responseContent);
                        }
                        catch (Exception)
                        {
                            return JObject.FromObject(new
                            {
                                returnCode = 999,
                                returnInfo = "API 回應格式不正確",
                                rawResponse = responseContent
                            });
                        }
                    }
                    else
                    {
                        var errorContent = response.Content.ReadAsStringAsync().Result;
                        return JObject.FromObject(new
                        {
                            returnCode = 999,
                            returnInfo = $"HTTP 請求失敗: {response.StatusCode} - {response.ReasonPhrase}",
                            errorContent = errorContent
                        });
                    }
                }
            }
            catch (Exception e)
            {
                return JObject.FromObject(new
                {
                    returnCode = 999,
                    returnInfo = $"系統錯誤: {e.Message}"
                });
            }
        }

        private JObject CallEzucUploadAttachmentApi(JObject parameters, HttpPostedFileBase uploadFile)
        {
            try
            {
                var apiUrl = $"{ezucServerUrl}/ucrm/api/bulletin/uploadAttachment";

                using (var formContent = new MultipartFormDataContent())
                {
                    // 加入 para 參數
                    var paraContent = new StringContent(parameters.ToString(Formatting.None));
                    formContent.Add(paraContent, "para");

                    // 加入附件檔案
                    uploadFile.InputStream.Seek(0, SeekOrigin.Begin);
                    byte[] fileData = new byte[uploadFile.ContentLength];
                    uploadFile.InputStream.Read(fileData, 0, uploadFile.ContentLength);

                    var fileContent = new ByteArrayContent(fileData);
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(uploadFile.ContentType ?? "application/octet-stream");
                    formContent.Add(fileContent, "upload", uploadFile.FileName ?? "attachment");

                    // 發送請求
                    var response = httpClient.PostAsync(apiUrl, formContent).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = response.Content.ReadAsStringAsync().Result;
                        try
                        {
                            return JObject.Parse(responseContent);
                        }
                        catch (Exception)
                        {
                            return JObject.FromObject(new
                            {
                                returnCode = 999,
                                returnInfo = "API 回應格式不正確",
                                rawResponse = responseContent
                            });
                        }
                    }
                    else
                    {
                        var errorContent = response.Content.ReadAsStringAsync().Result;
                        return JObject.FromObject(new
                        {
                            returnCode = 999,
                            returnInfo = $"HTTP 請求失敗: {response.StatusCode} - {response.ReasonPhrase}",
                            errorContent = errorContent
                        });
                    }
                }
            }
            catch (Exception e)
            {
                return JObject.FromObject(new
                {
                    returnCode = 999,
                    returnInfo = $"系統錯誤: {e.Message}"
                });
            }
        }

        private JObject CallEzucAddBulletinApi(JObject parameters)
        {
            try
            {
                var apiUrl = $"{ezucServerUrl}/ucrm/api/bulletin/addBulletin";

                using (var formContent = new MultipartFormDataContent())
                {
                    var paraContent = new StringContent(parameters.ToString(Formatting.None));
                    formContent.Add(paraContent, "para");

                    // 發送請求
                    var response = httpClient.PostAsync(apiUrl, formContent).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = response.Content.ReadAsStringAsync().Result;
                        try
                        {
                            return JObject.Parse(responseContent);
                        }
                        catch (Exception)
                        {
                            return JObject.FromObject(new
                            {
                                returnCode = 999,
                                returnInfo = "API 回應格式不正確",
                                rawResponse = responseContent
                            });
                        }
                    }
                    else
                    {
                        var errorContent = response.Content.ReadAsStringAsync().Result;
                        return JObject.FromObject(new
                        {
                            returnCode = 999,
                            returnInfo = $"HTTP 請求失敗: {response.StatusCode} - {response.ReasonPhrase}",
                            errorContent = errorContent
                        });
                    }
                }
            }
            catch (Exception e)
            {
                return JObject.FromObject(new
                {
                    returnCode = 999,
                    returnInfo = $"系統錯誤: {e.Message}"
                });
            }
        }
        #endregion

        #region 輔助方法
        private void ApiKeyVerify(string company, string secretKey, string functionName)
        {
            if (string.IsNullOrEmpty(company) || string.IsNullOrEmpty(secretKey))
            {
                throw new Exception("API 金鑰驗證失敗：公司代碼或密鑰為空");
            }
            // 這裡應該加入實際的驗證邏輯
        }
        #endregion

        #endregion

        #endregion
    }
}