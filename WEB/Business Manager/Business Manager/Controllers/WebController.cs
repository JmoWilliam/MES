using System;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using NLog.Targets;
using System.Configuration;
using NLog;
using NLog.Targets.Wrappers;
using Newtonsoft.Json.Linq;
using Helpers;
using BASDA;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using EIPDA;

namespace Business_Manager
{
    public class WebController : Controller
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public JObject jsonResponse = new JObject();
        public string dataRequest = "";

        #region //DA初始化
        #region //BASDA
        public BasicInformationDA basicInformationDA = new BasicInformationDA();
        public SystemSettingDA systemSettingDA = new SystemSettingDA();
        #endregion
        #endregion

        public WebController()
        {
        }

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            string userNo = "";
            HttpCookie Login = Request.Cookies.Get("Login");
            if (Login != null)
            {
                userNo = Login.Value;
            }

            CreateLogger(userNo);

            ViewBag.Title = "";
            ViewBag.Description = "";
            ViewBag.Keywords = "";
            ViewBag.Footer = "";
            ViewBag.PageTitle = "";

            base.OnActionExecuting(filterContext);
        }

        #region //檔案上傳
        [HttpPost]
        public void Upload(string savePath, bool isPreview = true)
        {
            try
            {
                List<FileModel> files = FileHelper.FileSave(Request.Files);
                List<string> initialPreview = new List<string>();
                List<FileConfig> fileConfigs = new List<FileConfig>();
                string controllerName = Convert.ToString(Request.RequestContext.RouteData.Values["Controller"]);
                string actionName = Convert.ToString(Request.RequestContext.RouteData.Values["Action"]);

                foreach (var file in files)
                {
                    #region //Request
                    dataRequest = systemSettingDA.AddFile(file.FileName, file.FileContent, file.FileExtension, file.FileSize, BaseHelper.ClientIP(), savePath);
                    #endregion

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    int fileId = Convert.ToInt32(jsonResponse["result"][0]["FileId"]);
                    string type = "";
                    bool filePreview = false;

                    FileHelper.FileType(ref type, ref filePreview, file.FileExtension);

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = file.FileContent;

                    string downloadPath = string.Format("/Web/Download?fileGuid={0}&fileName={1}&fileExtension={2}", fileGuid, file.FileName, file.FileExtension);

                    initialPreview.Add(downloadPath);

                    FileConfig fileConfig = new FileConfig
                    {
                        key = fileId.ToString(),
                        caption = file.FileName + file.FileExtension,
                        type = type,
                        filetype = FileHelper.GetMime(file.FileExtension),
                        size = file.FileSize,
                        previewAsData = filePreview,
                        url = "/Web/Delete",
                        downloadUrl = downloadPath
                    };

                    fileConfigs.Add(fileConfig);
                }

                if (isPreview)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        initialPreview,
                        initialPreviewConfig = fileConfigs,
                        initialPreviewAsData = true
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        initialPreview = new List<string>(),
                        initialPreviewConfig = fileConfigs,
                        initialPreviewAsData = true
                    });
                }
                
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //檔案刪除
        [HttpPost]
        public void Delete(string key)
        {
            try
            {
                #region //Request
                systemSettingDA.DeleteFile(Convert.ToInt32(key));
                #endregion
                
                jsonResponse = JObject.FromObject(new
                {
                    message = "檔案刪除成功!"
                });
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    message = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //檔案預覽
        [HttpPost]
        public void Preview(string previewFile)
        {
            try
            {
                string[] filePath = previewFile.Split(',');
                List<string> initialPreview = new List<string>();
                List<FileConfig> fileConfigs = new List<FileConfig>();

                if (filePath.Length > 0 && previewFile != "-1")
                {
                    for (int i = 0; i < filePath.Length; i++)
                    {
                        int fileId = Convert.ToInt32(filePath[i]);

                        #region //Request
                        dataRequest = systemSettingDA.GetFile(fileId, -1, "", "", "", "", -1, -1);
                        #endregion

                        #region //Response
                        jsonResponse = BaseHelper.DAResponse(dataRequest);
                        #endregion

                        if (jsonResponse["status"].ToString() == "success")
                        {
                            string fileName = jsonResponse["result"][0]["FileName"].ToString();
                            byte[] fileContent = (byte[])jsonResponse["result"][0]["FileContent"];
                            string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();
                            int fileSize = Convert.ToInt32(jsonResponse["result"][0]["FileSize"]);
                            
                            string type = "";
                            bool filePreview = false;

                            FileHelper.FileType(ref type, ref filePreview, fileExtension);

                            string fileGuid = Guid.NewGuid().ToString();
                            Session[fileGuid] = fileContent;

                            string downloadPath = string.Format("/Web/Download?fileGuid={0}&fileName={1}&fileExtension={2}", fileGuid, fileName, fileExtension);

                            initialPreview.Add(downloadPath);

                            FileConfig fileConfig = new FileConfig
                            {
                                key = fileId.ToString(),
                                caption = fileName + fileExtension,
                                type = type,
                                filetype = FileHelper.GetMime(fileExtension),
                                size = fileSize,
                                previewAsData = filePreview,
                                url = "/Web/Delete",
                                downloadUrl = downloadPath
                            };

                            fileConfigs.Add(fileConfig);

                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "OK",
                                initialPreview,
                                initialPreviewConfig = fileConfigs,
                                initialPreviewAsData = true
                            });
                        }
                    }                    
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "沒有可預覽的檔案!"
                    });
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //取得檔案
        public virtual ActionResult GetFile(int fileId)
        {
            try
            {
                #region //Request
                
                dataRequest = systemSettingDA.GetFile(fileId, -1, "", "", "", "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    string fileName = jsonResponse["result"][0]["FileName"].ToString();
                    byte[] fileContent = (byte[])jsonResponse["result"][0]["FileContent"];
                    string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();

                    return File(fileContent, FileHelper.GetMime(fileExtension), fileName + fileExtension);
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                } 
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                return new EmptyResult();
            }
        }
        #endregion

        #region //檔案下載
        public virtual ActionResult Download(string fileGuid, string fileName, string fileExtension)
        {
            if (Session[fileGuid] != null)
            {
                byte[] data = Session[fileGuid] as byte[];
                return File(data, FileHelper.GetMime(fileExtension), fileName + fileExtension);
            }
            else
            {
                return new EmptyResult();
            }
        }
        #endregion

        [NonAction]
        public void ViewLoginCheck(string functionCode = "", string detailCode = "")
        {
            bool verify = LoginVerify();

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "User");

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
            else
            {
                basicInformationDA = new BasicInformationDA();
                systemSettingDA = new SystemSettingDA();

                AuthoritySetting();

                if (functionCode.Length > 0)
                {
                    ViewAuthorityVerify(functionCode, detailCode);
                }
            }
        }

        [NonAction]
        public void ViewIMLoginCheck()
        {
            bool verify = LoginVerify("IM");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "User", new { Area = "IM" });

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }

        [NonAction]
        public void WebApiLoginCheck(string functionCode = "", string detailCode = "")
        {
            bool verify = LoginVerify();

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
            else
            {
                basicInformationDA = new BasicInformationDA();
                systemSettingDA = new SystemSettingDA();

                if (functionCode.Length > 0)
                {
                    AuthorityVerify(functionCode, detailCode);
                }
            }
        }

        [NonAction]
        public void ApiKeyVerify(string company, string secretKey, string functionCode)
        {
            dataRequest = systemSettingDA.GetApiKeyVerify(company, secretKey, functionCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("公司別或金鑰錯誤");
                }
             }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }

        [NonAction]
        protected void ViewAuthorityVerify(string functionCode, string detailCode)
        {
            dataRequest = systemSettingDA.GetAuthorityVerify(functionCode, detailCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            string redirectLoginUrl = Url.Action("Gate", "Home");
            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    HttpContext.Response.Redirect(redirectLoginUrl);
                }
            }
            else
            {
                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }

        [NonAction]
        protected void AuthorityVerify(string functionCode, string detailCode)
        {
            dataRequest = systemSettingDA.GetAuthorityVerify(functionCode, detailCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("無【" + functionCode + "】的【" + detailCode + "】權限");
                }
            }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }

        [NonAction]
        protected bool LoginVerify(string platform = "BM")
        {
            bool verify = false;

            if (Session["UserLock"] == null)
            {
                HttpCookie Login = Request.Cookies.Get("Login");
                HttpCookie LoginKey = Request.Cookies.Get("LoginKey");

                if (Login != null && LoginKey != null)
                {
                    #region //檢查使用者登入IP
                    bool autoLogin = false;
                    #region //Request
                    dataRequest = basicInformationDA.GetLoginIPCheck(Login.Value);
                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                    {
                        var result = JObject.Parse(dataRequest)["data"];

                        if (result.Count() == 1)
                        {
                            for (int i = 0; i < result.Count(); i++)
                            {
                                autoLogin = BaseHelper.ClientIP() == result[i]["LoginIP"].ToString();
                            }
                        }
                    }
                    #endregion

                    if (autoLogin)
                    {
                        #region //Request
                        dataRequest = basicInformationDA.GetLoginByKey(Login.Value, LoginKey.Value);
                        #endregion

                        if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                        {
                            var result = JObject.Parse(dataRequest)["data"];

                            if (result.Count() == 1)
                            {
                                for (int i = 0; i < result.Count(); i++)
                                {
                                    verify = true;
                                    Session["UserLock"] = true;
                                    Session["UserCompany"] = Convert.ToInt32(result[i]["CompanyId"]);
                                    Session["UserDepartmentId"] = Convert.ToInt32(result[i]["DepartmentId"]);
                                    Session["UserId"] = Convert.ToInt32(result[i]["UserId"]);
                                    Session["UserNo"] = result[i]["UserNo"].ToString();

                                    LoginLog(Convert.ToInt32(result[i]["UserId"]), result[i]["UserNo"].ToString(), platform);
                                }
                            }
                            else
                            {
                                Session.Clear();
                                Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                                Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
                            }
                        }
                    }
                    else
                    {
                        Session.Clear();
                        Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                        Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
                    }
                }
            }
            else
            {
                HttpCookie Login = Request.Cookies.Get("Login");

                if (Login != null)
                {
                    if (Session["UserNo"].ToString() != Login.Value)
                    {
                        Session.Clear();
                        Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                        Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
                    }
                    else
                    {
                        verify = true;
                    }
                }
                else
                {
                    Session.Clear();
                    Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                    Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
                }
            }

            return verify;
        }

        [NonAction]
        public void LoginLog(int userId = -1, string account = "", string platform = "BM")
        {
            basicInformationDA.AddUserLoginLog(userId, platform);
            basicInformationDA.UpdateUserLoginKey(userId, account);
        }

        [NonAction]
        public void AuthoritySetting()
        {
            ViewBag.authority = "";
            string functionCode = this.RouteData.Values["action"].ToString();

            dataRequest = systemSettingDA.GetAuthorityVerify(functionCode, "");
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() >= 0)
                {
                    ViewBag.authority = jsonResponse["result"].ToString(Formatting.None);
                }
            }
        }

        private void CreateLogger(string userNo)
        {
            DatabaseTarget databaseTarget = null;
            Target targetWrapper = LogManager.Configuration.FindTargetByName("databaseTarget");

            if (targetWrapper is AsyncTargetWrapper)
            {
                databaseTarget = (targetWrapper as AsyncTargetWrapper).WrappedTarget as DatabaseTarget;
            }
            else if (targetWrapper is DatabaseTarget)
            {
                databaseTarget = targetWrapper as DatabaseTarget;
            }

            if (databaseTarget != null)
            {
                databaseTarget.ConnectionString = ConfigurationManager.AppSettings["MainDb"].ToString();
                LogManager.ReconfigExistingLoggers();
            }

            GlobalDiagnosticsContext.Set("userNo", userNo);
        }

        #region //FOR EIP API
        #region //取得檔案

        [HttpPost]
        [Route("api/BAS/GetFileEIP")]
        public void GetFileEIP(int fileId)
        {
            try
            {
                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetFile(fileId, -1, "", "", "", "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #endregion

    }
}