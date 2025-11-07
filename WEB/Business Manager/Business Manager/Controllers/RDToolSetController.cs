using System.Net.Http;
using System.Web.Mvc;
using System.Threading.Tasks;
using System;
using Newtonsoft.Json;
using Helpers;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Compression;
using PDMDA;
using System.Configuration;

namespace Business_Manager.Controllers
{
    public class RDToolSetController : WebController
    {
        public string PMDServerConnectionStrings = "";

        public int CurrentUser = -1;
        public int CurrentCompany = -1;
        public string CurrentUserNo = "";
        public string CurrentCompanyNo = "";

        public HttpClient PMDHttpClient = new HttpClient();
        public RDToolSetDA rDToolSetDA = new RDToolSetDA();

        public RDToolSetController()
        {
            GetPMDHttpClientBaseAddress();
        }

        #region //View
        public ActionResult ZemaxToolSet()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UA3PContourf()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult JmoCheckToolSet()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Basic
        #region //GetPMDHttpClientBaseAddress
        private void GetPMDHttpClientBaseAddress()
        {
            try
            {
                PMDServerConnectionStrings = ConfigurationManager.AppSettings["PMDServerPath"];
                PMDHttpClient.BaseAddress = new Uri(PMDServerConnectionStrings);
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //GetUserInfo
        private void GetUserInfo()
        {
            try
            {
                CurrentUser = Convert.ToInt32(HttpContext.Session["UserId"]);
                CurrentCompany = Convert.ToInt32(HttpContext.Session["UserCompany"]);

                string resultUser = rDToolSetDA.GetUser(CurrentUser, -1, CurrentCompany, "", "");
                var jsonResult = JObject.Parse(resultUser);

                if (jsonResult["status"].ToString() == "success")
                {
                    if (!(jsonResult["data"] is JArray) || !jsonResult["data"].HasValues) throw new SystemException("無法取得使用者資訊");
                    CurrentUserNo = jsonResult["data"][0]["UserNo"].ToString();
                    CurrentCompanyNo = jsonResult["data"][0]["CompanyNo"].ToString();
                }
                else
                {
                    throw new SystemException("無法取得使用者資訊");
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //ValidateProgramPermission
        private void ValidateProgramPermission(string programName)
        {
            switch (programName)
            {
                case "zemax_asp":
                    WebApiLoginCheck("ZemaxToolSet", "file-add");
                    break;
                case "UA3P_Contourf":
                    WebApiLoginCheck("UA3PContourf", "file-add");
                    break;
                case "core_jmo_check":
                case "insert_to_jmo":
                    WebApiLoginCheck("JmoCheckToolSet", "file-add");
                    break;
                default:
                    throw new SystemException($"未知的程式名稱: {programName}");
            }
        }
        #endregion

        #region //ValidateDownloadPermission
        private void ValidateDownloadPermission(string programName)
        {
            switch (programName)
            {
                case "zemax_asp":
                    WebApiLoginCheck("ZemaxToolSet", "file-read");
                    break;
                case "UA3P_Contourf":
                    WebApiLoginCheck("UA3PContourf", "file-read");
                    break;
                default:
                    throw new SystemException($"未知的程式名稱: {programName}");
            }
        }
        #endregion
        #endregion

        #region //API
        #region //Get
        #region //GetZemaxToolSetHistory -- 取得zemax參數轉DXF歷史紀錄 -- Andrew 2025.02.10
        [HttpPost]
        public async Task<ActionResult> GetZemaxToolSetHistory(string FileName = "", string StartDate = "", string EndDate = ""
            , int PageIndex = 1, int PageSize = 10)
        {
            try
            {
                WebApiLoginCheck("ZemaxToolSet", "read,constrained-data");

                //取得使用者資料
                GetUserInfo();

                #region //Request
                // 準備 PMDHttpClient 請求參數
                var postData = new Dictionary<string, string>
                {
                    { "program_name", "zemax_asp" },
                    { "account", CurrentUserNo }
                };
                var content = new FormUrlEncodedContent(postData);

                // 呼叫 PMDHttpClient API
                var response = await PMDHttpClient.PostAsync("/GetHistoryForBM", content);

                if (!response.IsSuccessStatusCode)
                {
                    return Json(new { status = "error", msg = "無法取得歷史紀錄" });
                }

                // 讀取回應內容
                var historyData = await response.Content.ReadAsStringAsync();
                var historyList = JsonConvert.DeserializeObject<List<dynamic>>(historyData);
                var filteredData = historyList;

                // 套用篩選條件
                if (!string.IsNullOrEmpty(FileName))
                {
                    filteredData = filteredData.Where(x =>
                    {
                        var requestData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.request_data.ToString());
                        var responseData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.response_data.ToString());

                        // 取得txt檔名
                        string txtUrl = requestData["txt_url"]?.ToString() ?? "";
                        string txtFileName = Path.GetFileName(txtUrl);

                        return txtFileName.IndexOf(FileName, StringComparison.OrdinalIgnoreCase) >= 0;
                    }).ToList();
                }

                if (!string.IsNullOrEmpty(StartDate))
                {
                    var startDateTime = DateTime.Parse(StartDate);
                    filteredData = filteredData.Where(x =>
                        DateTime.Parse(x.request_time.ToString()) >= startDateTime).ToList();
                }

                if (!string.IsNullOrEmpty(EndDate))
                {
                    var endDateTime = DateTime.Parse(EndDate).AddDays(1);
                    filteredData = filteredData.Where(x =>
                        DateTime.Parse(x.request_time.ToString()) < endDateTime).ToList();
                }

                // 排序
                filteredData = filteredData.OrderByDescending(x =>
                    DateTime.Parse(x.request_time.ToString())).ToList();

                // 計算總筆數
                var totalCount = filteredData.Count();

                // 分頁
                var pagedData = filteredData
                    .Skip((PageIndex - 1) * PageSize)
                    .Take(PageSize)
                    .Select(x => {
                        var requestData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.request_data.ToString());
                        var responseData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.response_data.ToString());

                        return new
                        {
                            id = requestData["_id"]?.ToString(),
                            fileName = Path.GetFileName(requestData["txt_url"]?.ToString() ?? ""),
                            parameter = requestData.ContainsKey("point_count") ? requestData["point_count"]?.ToString() ?? "" : "",
                            uploadTime = x.request_time.ToString(),
                            account = x.account.ToString(),
                            outputFileDir = responseData["output_file_dir"]?.ToString(),
                        };
                    })
                    .ToList();

                var result = new
                {
                    status = "success",
                    data = new
                    {
                        total = totalCount,
                        rows = pagedData
                    }
                };
                #endregion

                #region //Response
                return Json(result);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                logger.Error(e.Message);
                return Json(new { status = "error", msg = e.Message });
                #endregion
            }
        }
        #endregion

        #region //GetUA3PContourfHistory -- 取得UA3P等高分析歷史紀錄 -- Andrew 2025.02.13
        [HttpPost]
        public async Task<ActionResult> GetUA3PContourfHistory(string FileName = "", string StartDate = "", string EndDate = ""
            , int PageIndex = 1, int PageSize = 10)
        {
            try
            {
                WebApiLoginCheck("UA3PContourf", "read,constrained-data");

                //取得使用者資料
                GetUserInfo();

                #region //Request
                // 準備 PMDHttpClient 請求參數
                var postData = new Dictionary<string, string>
                {
                    { "program_name", "UA3P_Contourf" },
                    { "account", CurrentUserNo }
                };
                var content = new FormUrlEncodedContent(postData);

                // 呼叫 PMDHttpClient API
                var response = await PMDHttpClient.PostAsync("/GetHistoryForBM", content);

                if (!response.IsSuccessStatusCode)
                {
                    return Json(new { status = "error", msg = "無法取得歷史紀錄" });
                }

                // 讀取回應內容
                var historyData = await response.Content.ReadAsStringAsync();
                var historyList = JsonConvert.DeserializeObject<List<dynamic>>(historyData);
                var filteredData = historyList;

                // 套用篩選條件
                if (!string.IsNullOrEmpty(FileName))
                {
                    filteredData = filteredData.Where(x =>
                    {
                        var requestData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.request_data.ToString());
                        var responseData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.response_data.ToString());

                        // 取得(alx, aly)csv檔名
                        string alxUrl = requestData["path_x"]?.ToString() ?? "";
                        string alxFileName = Path.GetFileName(alxUrl);
                        string alyUrl = requestData["path_y"]?.ToString() ?? "";
                        string alyFileName = Path.GetFileName(alyUrl);

                        return (alxFileName.IndexOf(FileName, StringComparison.OrdinalIgnoreCase) >= 0 || alyFileName.IndexOf(FileName, StringComparison.OrdinalIgnoreCase) >= 0);
                    }).ToList();
                }

                if (!string.IsNullOrEmpty(StartDate))
                {
                    var startDateTime = DateTime.Parse(StartDate);
                    filteredData = filteredData.Where(x =>
                        DateTime.Parse(x.request_time.ToString()) >= startDateTime).ToList();
                }

                if (!string.IsNullOrEmpty(EndDate))
                {
                    var endDateTime = DateTime.Parse(EndDate).AddDays(1);
                    filteredData = filteredData.Where(x =>
                        DateTime.Parse(x.request_time.ToString()) < endDateTime).ToList();
                }

                // 排序
                filteredData = filteredData.OrderByDescending(x =>
                    DateTime.Parse(x.request_time.ToString())).ToList();

                // 計算總筆數
                var totalCount = filteredData.Count();

                // 分頁
                var pagedData = filteredData
                    .Skip((PageIndex - 1) * PageSize)
                    .Take(PageSize)
                    .Select(x => {
                        var requestData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.request_data.ToString());
                        var responseData = JsonConvert.DeserializeObject<Dictionary<string, object>>(x.response_data.ToString());

                        return new
                        {
                            id = requestData["_id"]?.ToString(),
                            fileName = $"{Path.GetFileName(requestData["path_x"]?.ToString() ?? "")}<br>{Path.GetFileName(requestData["path_y"]?.ToString() ?? "")}",
                            parameter = requestData["post_message"]?.ToString(),
                            uploadTime = x.request_time.ToString(),
                            account = x.account.ToString(),
                            outputFileDir = responseData["output_file_dir"]?.ToString(),
                        };
                    })
                    .ToList();

                var result = new
                {
                    status = "success",
                    data = new
                    {
                        total = totalCount,
                        rows = pagedData
                    }
                };
                #endregion

                #region //Response
                return Json(result);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                logger.Error(e.Message);
                return Json(new { status = "error", msg = e.Message });
                #endregion
            }
        }
        #endregion

        #endregion

        #region //PythonExecution

        #region //SpecifyInputParameters
        private Dictionary<string, string> SpecifyInputParameters(string programName)
        {
            var parameters = new Dictionary<string, string>();

            switch (programName)
            {
                case "zemax_asp":
                    string pointCount = Request.Form["pointCount"];
                    if (!string.IsNullOrEmpty(pointCount))
                    {
                        parameters.Add("point_count", pointCount);
                    }
                    break;

                case "UA3P_Contourf":
                    string radius = Request.Form["radius"];
                    string maxZ = Request.Form["maxZ"];
                    string minZ = Request.Form["minZ"];

                    string postMessage = string.Join(",", new[] { radius, maxZ, minZ }
                        .Where(x => !string.IsNullOrEmpty(x)));
                    parameters.Add("post_message", postMessage);
                    break;

                case "core_jmo_check":
                    break;

                default:
                    throw new SystemException($"未知的程式名稱: {programName}");
            }

            return parameters;
        }
        #endregion

        #region //ExecutePythonScript
        [HttpPost]
        public async Task<JsonResult> ExecutePythonScript()
        {
            try
            {
                //取得參數
                string folderName = Request.Form["folderName"];
                string programName = Request.Form["programName"];

                //取得使用者資料
                GetUserInfo();

                //權限驗證
                //ValidateProgramPermission(programName);

                using (var content = new MultipartFormDataContent())
                {
                    // 基本參數
                    content.Add(new StringContent(CurrentUserNo), "account");
                    content.Add(new StringContent(Request.UserHostAddress), "pc_ip");
                    content.Add(new StringContent(folderName), "folder_name");

                    // 取得程式特定參數
                    var programParameters = SpecifyInputParameters(programName);
                    foreach (var param in programParameters)
                    {
                        content.Add(new StringContent(param.Value), param.Key);
                    }

                    // 檔案處理
                    if (Request.Files.Count > 0)
                    {
                        string fileKeys = Request.Form["fileKeys"];
                        var fileKeysDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(fileKeys);

                        foreach (var key in fileKeysDict.Keys)
                        {
                            var file = Request.Files[key];
                            if (file != null)
                            {
                                var fileContent = new StreamContent(file.InputStream);
                                content.Add(fileContent, key, fileKeysDict[key]);
                            }
                        }
                    }

                    // 呼叫46主機 Python API
                    var response = await PMDHttpClient.PostAsync($"/PmdWebSystem/{programName}", content);

                    if (response.IsSuccessStatusCode)
                    {
                        var result = await response.Content.ReadAsStringAsync();
                        var logData = JsonConvert.DeserializeObject<dynamic>(result);
                        var executionResult = await CheckExecutionStatus(logData._id.ToString());

                        var finalResult = new
                        {
                            account = ((IDictionary<string, JToken>)executionResult)["account"]?.ToString(),
                            output_file_dir = ((IDictionary<string, JToken>)executionResult)["output_file_dir"]?.ToString(),
                        };

                        // 建立回傳物件
                        var responseData = new
                        {
                            success = true,
                            message = "檔案上傳成功",
                            data = finalResult
                        };

                        return Json(responseData);
                    }
                    else
                    {
                        throw new SystemException("執行失敗");
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }
        #endregion

        #region //CheckExecutionStatus
        private async Task<object> CheckExecutionStatus(string logId)
        {
            try
            {
                for (int i = 0; i < 10; i++)
                {
                    var statusResponse = await PMDHttpClient.GetAsync($"/GetStatusY/{logId}");

                    if (statusResponse.IsSuccessStatusCode)
                    {
                        var status = await statusResponse.Content.ReadAsStringAsync();
                        var statusData = JsonConvert.DeserializeObject<dynamic>(status);

                        if (statusData.status.ToString() == "Y")
                        {
                            return statusData.response_data;
                        }
                    }

                    await Task.Delay(1000);
                }

                return new { error = "Request timeout" };
            }
            catch (Exception ex)
            {
                return new { error = ex.Message };
            }
        }
        #endregion
        #endregion

        #region //Download
        #region //DownloadRDToolSet
        [HttpPost]
        public async Task<ActionResult> DownloadRDToolSet(string programName, string fileType, string originalFileName, string outputFileDir)
        {
            try
            {
                //權限驗證
                ValidateDownloadPermission(programName);

                // 取得檔案清單
                var allFiles = await GetFilesList(outputFileDir);
                if (allFiles == null)
                {
                    return Json(new { success = false, message = "取得檔案清單失敗" }, JsonRequestBehavior.AllowGet);
                }

                // 篩選目標檔案
                var (targetFiles, fileName) = FilterFiles(allFiles, fileType, originalFileName);

                // 檢查是否有找到檔案
                if (!targetFiles.Any())
                {
                    return Json(new { success = false, message = "未找到符合的檔案" }, JsonRequestBehavior.AllowGet);
                }

                // 下載檔案
                return await DownloadFiles(targetFiles, fileName, fileType);
            }
            catch (ArgumentException ex)
            {
                return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "下載過程發生錯誤：" + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region //GetFilesList
        private async Task<string[]> GetFilesList(string outputFileDir)
        {
            var postData = new Dictionary<string, string>
            {
                { "dir_url", outputFileDir }
            };
            var filesListContent = new FormUrlEncodedContent(postData);
            var filesListResponse = await PMDHttpClient.PostAsync("/GetFilesList", filesListContent);

            if (!filesListResponse.IsSuccessStatusCode)
            {
                return null;
            }

            var filesListResult = await filesListResponse.Content.ReadAsStringAsync();
            var filesList = JsonConvert.DeserializeObject<dynamic>(filesListResult).files_ls;
            return filesList.ToObject<string[]>();
        }
        #endregion

        #region //FilterFiles
        private (List<string> targetFiles, string fileName) FilterFiles(string[] allFiles, string fileType, string originalFileName)
        {
            var targetFiles = new List<string>();
            string fileName;

            switch (fileType.ToLower())
            {
                case "all":
                case "history":
                    targetFiles = new List<string>(allFiles);
                    fileName = $"{originalFileName}.zip";
                    break;
                case "excel":
                    targetFiles = allFiles.Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).ToList();
                    fileName = $"{originalFileName}.xlsx";
                    break;
                case "dxf":
                    targetFiles = allFiles.Where(f => f.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase)).ToList();
                    fileName = $"{originalFileName}.dxf";
                    break;
                case "png":
                    targetFiles = allFiles.Where(f => f.EndsWith(".png", StringComparison.OrdinalIgnoreCase)).ToList();
                    fileName = $"{originalFileName}.png";
                    break;
                default:
                    throw new ArgumentException("無效的檔案類型");
            }

            return (targetFiles, fileName);
        }
        #endregion

        #region //DownloadFiles
        private async Task<ActionResult> DownloadFiles(List<string> targetFiles, string fileName, string fileType)
        {
            // 單一檔案下載
            if (fileType.ToLower() != "all" && targetFiles.Count == 1)
            {
                return await DownloadSingleFile(targetFiles.First(), fileName, fileType);
            }

            // 多檔案壓縮下載
            return await DownloadMultipleFiles(targetFiles, fileName);
        }
        #endregion

        #region //DownloadSingleFile
        private async Task<ActionResult> DownloadSingleFile(string filePath, string fileName, string fileType)
        {
            var downloadData = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                { "file_path", filePath }
            });

            var fileResponse = await PMDHttpClient.PostAsync("/DownloadFile", downloadData);

            if (!fileResponse.IsSuccessStatusCode)
            {
                return Json(new { success = false, message = "檔案下載失敗" }, JsonRequestBehavior.AllowGet);
            }

            var fileContent = await fileResponse.Content.ReadAsByteArrayAsync();
            var contentType = fileType.ToLower() == "excel" ?
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" :
                "application/octet-stream";

            return File(fileContent, contentType, fileName);
        }
        #endregion

        #region //DownloadMultipleFiles
        private async Task<ActionResult> DownloadMultipleFiles(List<string> filePaths, string fileName)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var filePath in filePaths)
                    {
                        var downloadData = new FormUrlEncodedContent(new Dictionary<string, string>
                        {
                            { "file_path", filePath }
                        });

                        var fileResponse = await PMDHttpClient.PostAsync("/DownloadFile", downloadData);

                        if (fileResponse.IsSuccessStatusCode)
                        {
                            var fileContent = await fileResponse.Content.ReadAsByteArrayAsync();
                            var entryName = Path.GetFileName(filePath);
                            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);

                            using (var entryStream = entry.Open())
                            {
                                await entryStream.WriteAsync(fileContent, 0, fileContent.Length);
                                await entryStream.FlushAsync();
                            }
                        }
                    }
                }

                return File(memoryStream.ToArray(), "application/zip", fileName);
            }
        }
        #endregion
        #endregion

        #endregion
    }
}