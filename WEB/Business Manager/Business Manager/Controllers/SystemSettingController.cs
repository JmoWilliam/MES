using BASDA;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Org.BouncyCastle.Asn1.Ocsp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Web.Mvc;
using ZXing;
using ZXing.Common;

namespace Business_Manager.Controllers
{
    public class SystemSettingController : WebController
    {
        private string MamoUrl = "http://192.168.20.46:2536/Mattermost/";

        #region //View
        public ActionResult ApiKeyManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AuthorityManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SystemLog()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MailServer()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MailTemplate()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult Calendar()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult FileManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UserLoginLog()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SubSystemManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PasswordSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QRCodeGenerator()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PushMessage()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ChatRoom()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult VideoChat()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AudioChat()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MultipleVideoChat()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult KeyenceTest()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UserInfo()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRandomApiKey 取得隨機Api金鑰
        [HttpPost]
        public void GetRandomApiKey()
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "add,update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRandomApiKey();
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

        #region //GetApiKey 取得Api金鑰資料
        [HttpPost]
        public void GetApiKey(int KeyId = -1, string KeyText = "", string Purpose = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetApiKey(KeyId, KeyText, Purpose, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetApiKeyFunction 取得Api適用功能
        [HttpPost]
        public void GetApiKeyFunction(int FunctionId = -1, int KeyId = -1, string FunctionCode = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "function");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetApiKeyFunction(FunctionId, KeyId, FunctionCode
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetRole 取得角色資料
        [HttpPost]
        public void GetRole(int RoleId = -1, string RoleName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRole(RoleId, RoleName, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetRoleAuthority 取得角色權限資料
        [HttpPost]
        public void GetRoleAuthority(int RoleId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRoleAuthority(RoleId, SearchKey);
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

        #region //GetRoleDetailAuthority 取得角色詳細權限資料
        [HttpPost]
        public void GetRoleDetailAuthority(int ModuleId = -1, int RoleId = -1, string SearchKey = "", bool Grant = false)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRoleDetailAuthority(ModuleId, RoleId, SearchKey, Grant);
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

        #region //GetAuthorityUser 取得權限使用者資料
        [HttpPost]
        public void GetAuthorityUser(string Roles = "", int CompanyId = -1, string Departments = "", string UserNo = "", string UserName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityUser(Roles, CompanyId, Departments, UserNo, UserName, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetAuthorityUserCompany 取得使用者公司權限資料
        [HttpPost]
        public void GetAuthorityUserCompany(int CompanyId = -1, string Departments = "", string UserNo = "", string UserName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityUserCompany(CompanyId, Departments, UserNo, UserName, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetRoleUser 取得角色使用者資料
        [HttpPost]
        public void GetRoleUser(int RoleId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRoleUser(RoleId, SearchKey);
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

        #region //GetUserRole 取得使用者角色資料
        [HttpPost]
        public void GetUserRole(string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserRole(SearchKey);
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

        #region //GetUserAuthority 取得使用者權限資料
        [HttpPost]
        public void GetUserAuthority(int UserId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserAuthority(UserId, SearchKey);
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

        #region //GetUserDetailAuthority 取得使用者詳細權限資料
        [HttpPost]
        public void GetUserDetailAuthority(int ModuleId = -1, int UserId = -1, string SearchKey = "", bool Grant = false)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserDetailAuthority(ModuleId, UserId, SearchKey, Grant);
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

        #region //GetUserCompanyAuthority 取得使用者公司權限
        [HttpPost]
        public void GetUserCompanyAuthority()
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserCompanyAuthority();
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

        #region //GetAuthorityVerify 取得權限驗證資料
        [HttpPost]
        public void GetAuthorityVerify(string FunctionCode = "", string DetailCode = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityVerify(FunctionCode, DetailCode);
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

        #region //GetLog 取得系統日誌資料
        [HttpPost]
        public void GetLog(int LogId = -1, string UserNo = "", string Message = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SystemLog", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetLog(LogId, UserNo, Message, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetMailServer 取得郵件伺服器資料
        [HttpPost]
        public void GetMailServer(int ServerId = -1, string Host = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MailServer", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailServer(ServerId, Host
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetMailContact 取得郵件聯絡人資料
        [HttpPost]
        public void GetMailContact(string Contact = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "contact");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailContact(Contact
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetContact 取得聯絡人資料
        [HttpPost]
        public void GetContact(int CompanyId = -1, int DepartmentId = -1, string Contact = "", string Mode = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetContact(CompanyId, DepartmentId, Contact, Mode
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetMailUser 取得聯絡人資料
        [HttpPost]
        public void GetMailUser()
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailUser();
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

        #region //GetMailTemplate 取得郵件樣板資料
        [HttpPost]
        public void GetMailTemplate(int MailId = -1, int ServerId = -1, string MailName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailTemplate(MailId, ServerId, MailName
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetMailSendSetting 取得郵件寄送設定
        [HttpPost]
        public void GetMailSendSetting(int SettingId = -1, int MailId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "send-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailSendSetting(SettingId, MailId, SearchKey
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetCalendar 取得行事曆資料
        [HttpPost]
        public void GetCalendar(string CalendarDate = "")
        {
            try
            {
                WebApiLoginCheck("Calendar", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetCalendar(CalendarDate);
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

        #region //GetCalendarInfo 取得行事曆資料
        [HttpPost]
        public void GetCalendarInfo(int CalendarId = -1, string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetCalendarInfo(CalendarId, StartDate, EndDate);
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

        #region //GetFile 取得檔案資料
        [HttpPost]
        public void GetFile(int FileId = -1, int UploadUser = -1, string DeleteStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetFile(FileId, UploadUser, DeleteStatus, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetUserLogin 取得使用者登入紀錄
        [HttpPost]
        public void GetUserLogin(int LogId = -1, int CompanyId = -1, int DepartmentId = -1
            , string UserNo = "", string UserName = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserLoginLog", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserLogin(LogId, CompanyId, DepartmentId
                    , UserNo, UserName, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetSubSystem 取得子系統資料
        [HttpPost]
        public void GetSubSystem(int SubSystemId = -1, string SubSystemCode = "", string SubSystemName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetSubSystem(SubSystemId, SubSystemCode, SubSystemName, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetRandomSystemKey 取得隨機系統金鑰
        [HttpPost]
        public void GetRandomSystemKey()
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "add,update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetRandomSystemKey();
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

        #region //GetSubSystemUser 取得子系統使用者資料
        [HttpPost]
        public void GetSubSystemUser(int SubSystemUserId = -1, int SubSystemId = -1, int CompanyId = -1
            , string Departments = "", string UserNo = "", string UserName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "allow-user");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetSubSystemUser(SubSystemUserId, SubSystemId, CompanyId
                    , Departments, UserNo, UserName
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetPasswordSetting 取得密碼參數設定資料
        [HttpPost]
        public void GetPasswordSetting()
        {
            try
            {
                WebApiLoginCheck("PasswordSetting", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetPasswordSetting();
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

        #region //GetNotificationSetting 取得系統通知個人設定資料
        [HttpPost]
        public void GetNotificationSetting()
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetNotificationSetting();
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

        #region //GetNotificationLog 取得系統通知紀錄資料
        [HttpPost]
        public void GetNotificationLog(string TriggerFunction = "", string ReadStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetNotificationLog(TriggerFunction, ReadStatus
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetUser 取得新版使用者資料
        [HttpPost]
        public void GetUser(int UserId = -1, string UserType = "", int CompanyId = -1, string Departments = ""
            , string InnerCode = "", string DisplayName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserInfo", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUser(UserId, UserType, CompanyId, Departments
                    , InnerCode, DisplayName, Status
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetUserEmployee 取得使用者內部員工設定檔資料
        [HttpPost]
        public void GetUserEmployee(int UserId = -1, string EmployeeStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetUserEmployee(UserId, EmployeeStatus
                    , OrderBy, PageIndex, PageSize);
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

        #region //Add
        #region //AddApiKey Api金鑰新增
        [HttpPost]
        public void AddApiKey(string KeyText = "", string Purpose = "")
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddApiKey(KeyText, Purpose);
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

        #region //AddApiKeyFunction Api適用功能新增
        [HttpPost]
        public void AddApiKeyFunction(int KeyId = -1, string FunctionCode = "", string Purpose = "")
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "function");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddApiKeyFunction(KeyId, FunctionCode, Purpose);
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

        #region //AddRole 角色資料新增
        [HttpPost]
        public void AddRole(string RoleName = "", string AdminStatus = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddRole(RoleName, AdminStatus);
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

        #region //AddMailServer 郵件伺服器新增
        [HttpPost]
        public void AddMailServer(string Host = "", string SendMode = "", int Port = -1, string Account = "", string Password = "")
        {
            try
            {
                WebApiLoginCheck("MailServer", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddMailServer(Host, SendMode, Port, Account, Password);
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

        #region //AddMailContact 郵件聯絡人資料新增
        [HttpPost]
        public void AddMailContact(string ContactName = "", string Email = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "contact");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddMailContact(ContactName, Email);
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

        #region //AddMailTemplate 郵件樣板新增
        [HttpPost]
        public void AddMailTemplate(int ServerId = -1, string MailName = "", string MailFrom = "", string MailTo = ""
            , string MailCc = "", string MailBcc = "", string MailSubject = "", string MailContent = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddMailTemplate(ServerId, MailName, MailFrom, MailTo, MailCc, MailBcc, MailSubject, MailContent);
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

        #region //AddMailSendSetting 郵件寄送設定新增
        [HttpPost]
        public void AddMailSendSetting(int MailId = -1, string SettingSchema = "", string SettingNo = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "send-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddMailSendSetting(MailId, SettingSchema, SettingNo);
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

        #region //AddCalendar 行事曆資料新增
        [HttpPost]
        public void AddCalendar(string Year = "")
        {
            try
            {
                WebApiLoginCheck("Calendar", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddCalendar(Year);
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

        #region //AddSubSystem 子系統資料新增
        [HttpPost]
        public void AddSubSystem(string SubSystemCode = "", string SubSystemName = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddSubSystem(SubSystemCode, SubSystemName, KeyText);
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

        #region //AddSubSystemUser 子系統使用者資料新增
        [HttpPost]
        public void AddSubSystemUser(int SubSystemId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "allow-user");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddSubSystemUser(SubSystemId, Users);
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

        #region //AddPushSubscription 推播訂閱資料新增
        [HttpPost]
        public void AddPushSubscription(string subscription)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddPushSubscription(subscription);
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

        #region //AddAuthorityUserCompany 使用者公司權限資料新增
        [HttpPost]
        public void AddAuthorityUserCompany(string Users = "", string Companys = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddAuthorityUserCompany(Users, Companys);
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

        #region //AddUser 新版使用者資料新增
        [HttpPost]
        public void AddUser(string UserJson)
        {
            try
            {
                WebApiLoginCheck("UserInfo", "add");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddUser(UserJson);
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

        #region //Update
        #region //UpdateApiKey Api金鑰更新
        [HttpPost]
        public void UpdateApiKey(int KeyId = -1, string KeyText = "", string Purpose = "")
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateApiKey(KeyId, KeyText, Purpose);
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

        #region //UpdateApiKeyStatus Api金鑰狀態更新
        [HttpPost]
        public void UpdateApiKeyStatus(int KeyId = -1)
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "status-switch");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateApiKeyStatus(KeyId);
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

        #region //UpdateApiKeyFunction Api適用功能更新
        [HttpPost]
        public void UpdateApiKeyFunction(int FunctionId = -1, string FunctionCode = "", string Purpose = "")
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "function");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateApiKeyFunction(FunctionId, FunctionCode, Purpose);
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

        #region //UpdateRole 角色資料更新
        [HttpPost]
        public void UpdateRole(int RoleId = -1, string RoleName = "", string AdminStatus = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateRole(RoleId, RoleName, AdminStatus);
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

        #region //UpdateRoleStatus 角色狀態更新
        [HttpPost]
        public void UpdateRoleStatus(int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateRoleStatus(RoleId);
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

        #region //UpdateRoleFunctionDetail 角色權限更新
        [HttpPost]
        public void UpdateRoleFunctionDetail(int RoleId = -1, int DetailId = -1, bool Checked = false, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateRoleFunctionDetail(RoleId, DetailId, Checked, SearchKey);
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

        #region //UpdateRoleUser 角色人員更新
        [HttpPost]
        public void UpdateRoleUser(int RoleId = -1, string UserList = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateRoleUser(RoleId, UserList);
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

        #region //UpdateUserFunctionDetail 使用者權限更新
        [HttpPost]
        public void UpdateUserFunctionDetail(int UserId = -1, int DetailId = -1, bool Checked = false, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateUserFunctionDetail(UserId, DetailId, Checked, SearchKey);
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

        #region //UpdateMailServer 郵件伺服器更新
        [HttpPost]
        public void UpdateMailServer(int ServerId = -1, string Host = "", int Port = -1
            , string SendMode = "", string Account = "", string Password = "")
        {
            try
            {
                WebApiLoginCheck("MailServer", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateMailServer(ServerId, Host, Port, SendMode, Account, Password);
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

        #region //UpdateMailContact 郵件聯絡人資料更新
        [HttpPost]
        public void UpdateMailContact(int ContactId = -1, string ContactName = "", string Email = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "contact");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateMailContact(ContactId, ContactName, Email);
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

        #region //UpdateMailTemplate 郵件樣板更新
        [HttpPost]
        public void UpdateMailTemplate(int MailId = -1, int ServerId = -1, string MailName = "", string MailFrom = "", string MailTo = ""
            , string MailCc = "", string MailBcc = "", string MailSubject = "", string MailContent = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateMailTemplate(MailId, ServerId, MailName, MailFrom, MailTo, MailCc, MailBcc, MailSubject, MailContent);
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

        #region //UpdateMailSendSetting 郵件寄送設定更新
        [HttpPost]
        public void UpdateMailSendSetting(int SettingId = -1, string SettingSchema = "", string SettingNo = "")
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "send-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateMailSendSetting(SettingId, SettingSchema, SettingNo);
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

        #region //UpdateCalendar 行事曆資料更新
        [HttpPost]
        public void UpdateCalendar(string CalendarDate = "", string CalendarDesc = "", string DateType = "")
        {
            try
            {
                WebApiLoginCheck("Calendar", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateCalendar(CalendarDate, CalendarDesc, DateType);
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

        #region //UpdateSubSystem 子系統資料更新
        [HttpPost]
        public void UpdateSubSystem(int SubSystemId = -1, string SubSystemCode = "", string SubSystemName = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateSubSystem(SubSystemId, SubSystemCode, SubSystemName, KeyText);
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

        #region //UpdateSubSystemStatus 子系統狀態更新
        [HttpPost]
        public void UpdateSubSystemStatus(int SubSystemId = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "status-switch");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateSubSystemStatus(SubSystemId);
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

        #region //UpdateSubSystemUserStatus 子系統使用者狀態更新
        [HttpPost]
        public void UpdateSubSystemUserStatus(int SubSystemUserId = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "allow-user");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateSubSystemUserStatus(SubSystemUserId);
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

        #region //UpdatePasswordSetting 密碼參數設定資料更新
        [HttpPost]
        public void UpdatePasswordSetting(int PasswordExpiration = -1, string PasswordFormat = "", int PasswordWrongCount = -1)
        {
            try
            {
                WebApiLoginCheck("PasswordSetting", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdatePasswordSetting(PasswordExpiration, PasswordFormat, PasswordWrongCount);
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

        #region //UpdateNotificationSetting 系統通知個人設定更新
        [HttpPost]
        public void UpdateNotificationSetting(string MailAdviceStatus = "", string PushAdviceStatus = "", string WorkWeixinStatus = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateNotificationSetting(MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus);
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

        #region //UpdateNotificationLogReadStatus 系統通知讀取狀態更新
        [HttpPost]
        public void UpdateNotificationLogReadStatus(int LogId = -1, string TriggerFunction = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateNotificationLogReadStatus(LogId, TriggerFunction);
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

        #region //UpdateUser 新版使用者資料更新
        [HttpPost]
        public void UpdateUser(int UserId = -1, string UserJson = "")
        {
            try
            {
                WebApiLoginCheck("UserInfo", "update");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateUser(UserId, UserJson);
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

        #region //UpdateUserStatus 新版使用者狀態更新
        [HttpPost]
        public void UpdateUserStatus(int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("UserInfo", "status-switch");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateUserStatus(UserId);
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

        #region //UpdateUserPasswordReset 新版使用者密碼重置
        [HttpPost]
        public void UpdateUserPasswordReset(int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("UserInfo", "password-reset");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateUserPasswordReset(UserId);
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

        #region //UpdateUserInfoCopy 新舊使用者介接程式
        public void UpdateUserInfoCopy()
        {
            try
            {
                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateUserInfoCopy();
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

        #region //UpdateDepartmentCopy 新舊部門介接程式
        public void UpdateDepartmentCopy()
        {
            try
            {
                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.UpdateDepartmentCopy();
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

        #region //Delete
        #region //DeleteApiKey Api金鑰刪除
        [HttpPost]
        public void DeleteApiKey(int KeyId = -1)
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteApiKey(KeyId);
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

        #region //DeleteApiKeyFunction Api適用功能刪除
        [HttpPost]
        public void DeleteApiKeyFunction(int FunctionId = -1)
        {
            try
            {
                WebApiLoginCheck("ApiKeyManagement", "function");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteApiKeyFunction(FunctionId);
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

        #region //DeleteRoleUser 角色使用者刪除
        [HttpPost]
        public void DeleteRoleUser(int UserId = -1, int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteRoleUser(UserId, RoleId);
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

        #region //DeleteMailServer 郵件伺服器刪除
        [HttpPost]
        public void DeleteMailServer(int ServerId = -1)
        {
            try
            {
                WebApiLoginCheck("MailServer", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteMailServer(ServerId);
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

        #region //DeleteMailContact 郵件聯絡人資料刪除
        [HttpPost]
        public void DeleteMailContact(int ContactId = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "contact");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteMailContact(ContactId);
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

        #region //DeleteMailTemplate 郵件樣板刪除
        [HttpPost]
        public void DeleteMailTemplate(int MailId = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteMailTemplate(MailId);
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

        #region //DeleteCalendar 行事曆資料刪除
        [HttpPost]
        public void DeleteCalendar(string CalendarDate = "")
        {
            try
            {
                WebApiLoginCheck("Calendar", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteCalendar(CalendarDate);
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

        #region //DeleteFile 檔案資料刪除
        [HttpPost]
        public void DeleteFile(int FileId = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteFile(FileId);
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

        #region //DeleteFilePermanent 檔案資料永久刪除
        [HttpPost]
        public void DeleteFilePermanent(int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("FileManagement", "permanent-delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteFilePermanent(FileId);
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

        #region //DeleteLotFilePermanent 檔案資料永久刪除(批次)
        [HttpGet]
        public void DeleteLotFilePermanent()
        {
            try
            {
                WebApiLoginCheck("FileManagement", "permanent-delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteLotFilePermanent();
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

        #region //DeleteSubSystem 子系統資料刪除
        [HttpPost]
        public void DeleteSubSystem(int SubSystemId = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteSubSystem(SubSystemId);
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

        #region //DeleteSubSystemUser 子系統使用者資料刪除
        [HttpPost]
        public void DeleteSubSystemUser(int SubSystemUserId = -1)
        {
            try
            {
                WebApiLoginCheck("SubSystemManagement", "allow-user");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteSubSystemUser(SubSystemUserId);
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

        #region //DeleteAuthorityUserCompany 使用者公司權限資料刪除
        [HttpPost]
        public void DeleteAuthorityUserCompany(int UserId = -1, int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteAuthorityUserCompany(UserId, CompanyId);
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

        #region //DeleteUser -- 刪除新版使用者資料
        [HttpPost]
        public void DeleteUser(int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("UserInfo", "delete");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.DeleteUser(UserId);
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

        #region //Api
        #region //GetNotificationMail 取得系統通知信件資料
        [HttpPost]
        [Route("api/BAS/GetNotificationMail")]
        public void GetNotificationMail(string Company, string SecretKey, string UserNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                dataRequest = systemSettingDA.GetNotificationMail(UserNo);
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

        #region //GetPushSubscription 取得推播訂閱資料
        [HttpPost]
        [Route("api/BAS/GetSubscription")]
        public void GetPushSubscription(string Company, string SecretKey, string UserNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetPushSubscription");
                #endregion

                #region //Request
                dataRequest = systemSettingDA.GetPushSubscription(UserNo, "");
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

        #region //AddNotificationLog 系統通知紀錄新增
        [HttpPost]
        [Route("api/BAS/AddNotification")]
        public void AddNotificationLog(string Company, string SecretKey, string Notification)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "AddNotificationLog");
                #endregion

                #region //Request
                dataRequest = systemSettingDA.AddNotificationLog(JsonConvert.DeserializeObject<Notification>(Notification));
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

        #region //DeletePushSubscription 推播訂閱資料刪除
        [HttpPost]
        [Route("api/BAS/DeleteSubscription")]
        public void DeletePushSubscription(string Company, string SecretKey, int PushSubscriptionId, string ApiEndpoint)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DeletePushSubscription");
                #endregion

                #region //Request
                dataRequest = systemSettingDA.DeletePushSubscription(PushSubscriptionId, ApiEndpoint);
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

        #region //DeleteNLog 刪除NLog資料
        [HttpPost]
        [Route("api/SYS/DeleteNLog")]
        public void DeleteNLog(string Company, string SecretKey, string ActualStart)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DeleteNLog");
                #endregion

                #region //Request
                dataRequest = systemSettingDA.DeleteNLog(ActualStart);
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

        #region //MAMO相關API
        #region //團隊
        #region //CreateTeams 建立團隊 
        [HttpPost]
        [Route("api/MAMO/CreateTeams")]
        public void CreateTeams(string Company, int UserId, string SecretKey, string TeamName, string Remark)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Teams/Create");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(TeamName), "name");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //取得回傳訊息
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                var responseJson = JObject.Parse(response.Result);

                                string MamoTeamId = "";
                                string TeamNo = "";
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "id")
                                    {
                                        MamoTeamId = item.Value.ToString();
                                    }
                                    else if (item.Key == "name")
                                    {
                                        TeamNo = item.Value.ToString();
                                    }
                                }
                                #endregion

                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoTeamId, "", "CreateTeams", "", Remark, UserId);
                                #endregion

                                #region //新增團隊
                                dataRequest = systemSettingDA.AddTeams(MamoTeamId, TeamName, TeamNo, Remark, Company, UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //DeleteTeams 刪除團隊
        [HttpPost]
        [Route("api/MAMO/DeleteTeams")]
        public void DeleteTeams(string Company, int UserId, string SecretKey, int TeamId)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Teams/Delete");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得團隊相關資料
                dataRequest = systemSettingDA.GetMamoTeams(TeamId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoTeamId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoTeamId = item["MamoTeamId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoTeamId), "team_id");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoTeamId, "", "DeleteTeams", "", "", UserId);
                                #endregion

                                #region //刪除團隊(非真正刪除，僅更改狀態為停用)
                                dataRequest = systemSettingDA.UpdateTeamStatus(TeamId, "S", UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //RestoreTeams 重啟團隊
        [HttpPost]
        [Route("api/MAMO/RestoreTeams")]
        public void RestoreTeams(string Company, int UserId, string SecretKey, int TeamId)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Teams/Restore");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得團隊相關資料
                dataRequest = systemSettingDA.GetMamoTeams(TeamId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoTeamId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoTeamId = item["MamoTeamId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoTeamId), "team_id");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoTeamId, "", "RestoreTeams", "", "", UserId);
                                #endregion

                                #region //重啟團隊(將團隊狀態改為啟用)
                                dataRequest = systemSettingDA.UpdateTeamStatus(TeamId, "A", UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //AddTeamMembers 新增團隊成員
        [HttpPost]
        [Route("api/MAMO/AddTeamMembers")]
        public void AddTeamMembers(string Company, int UserId, string SecretKey, int TeamId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Teams/AddMembers");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得團隊相關資料
                dataRequest = systemSettingDA.GetMamoTeams(TeamId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoTeamId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoTeamId = item["MamoTeamId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //處理UserNo格式
                string userNoListString = JsonConvert.SerializeObject(UserNo);
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoTeamId), "team_id");
                        multipart.Add(new StringContent(userNoListString), "usernames");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoTeamId, "", "AddTeamMembers", String.Join(", ", UserNo), "", UserId);
                                #endregion

                                #region //新增團隊成員
                                dataRequest = systemSettingDA.AddTeamMembers(TeamId, UserNo, UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                JObject responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "error")
                                    {
                                        throw new SystemException(item.Value.ToString());
                                    }
                                }

                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //DeleteTeamMembers 刪除團隊成員
        [HttpPost]
        [Route("api/MAMO/DeleteTeamMembers")]
        public void DeleteTeamMembers(string Company, int UserId, string SecretKey, int TeamId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Teams/RemoveMembers");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得團隊相關資料
                dataRequest = systemSettingDA.GetMamoTeams(TeamId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoTeamId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoTeamId = item["MamoTeamId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //處理UserNo格式
                string userNoListString = JsonConvert.SerializeObject(UserNo);
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoTeamId), "team_id");
                        multipart.Add(new StringContent(userNoListString), "usernames");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoTeamId, "", "DeleteTeamMembers", String.Join(", ", UserNo), "", UserId);
                                #endregion

                                #region //刪除團隊成員
                                dataRequest = systemSettingDA.DeleteTeamMembers(TeamId, UserNo);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                JObject responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "error")
                                    {
                                        throw new SystemException(item.Value.ToString());
                                    }
                                }

                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //頻道
        #region //CreateChannels 建立頻道
        [HttpPost]
        [Route("api/MAMO/CreateChannels")]
        public void CreateChannels(string Company, int UserId, string SecretKey, int TeamId, string ChannelName, string Remark)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Channels/Create");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得團隊相關資料
                dataRequest = systemSettingDA.GetMamoTeams(TeamId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoTeamId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoTeamId = item["MamoTeamId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoTeamId), "team_id");
                        multipart.Add(new StringContent(ChannelName), "name");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //取得回傳訊息
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                var responseJson = JObject.Parse(response.Result);

                                string MamoChannelId = "";
                                string ChannelNo = "";
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "id")
                                    {
                                        MamoChannelId = item.Value.ToString();
                                    }
                                    else if (item.Key == "name")
                                    {
                                        ChannelNo = item.Value.ToString();
                                    }
                                }
                                #endregion

                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoChannelId, MamoTeamId, "CreateChannels", "", Remark, UserId);
                                #endregion

                                #region //新增頻道
                                dataRequest = systemSettingDA.AddChannels(TeamId, MamoChannelId, ChannelName, ChannelNo, UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //DeleteChannels 刪除頻道
        [HttpPost]
        [Route("api/MAMO/DeleteChannels")]
        public void DeleteChannels(string Company, int UserId, string SecretKey, int ChannelId)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Channels/Delete");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得頻道相關資料
                dataRequest = systemSettingDA.GetMamoChannels(ChannelId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoChannelId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoChannelId = item["MamoChannelId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoChannelId), "channel_id");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                var responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    #region //紀錄LOG
                                    dataRequest = systemSettingDA.AddMamoLog(MamoChannelId, "", "DeleteChannels", "", "", UserId);
                                    #endregion

                                    #region //刪除頻道(非真正刪除，僅更改狀態為停用)
                                    dataRequest = systemSettingDA.UpdateChannelStatus(ChannelId, "S", UserId);
                                    #endregion

                                    #region //Response
                                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //RestoreChannels 重啟頻道
        [HttpPost]
        [Route("api/MAMO/RestoreChannels")]
        public void RestoreChannels(string Company, int UserId, string SecretKey, int ChannelId)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Channels/Restore");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得頻道相關資料
                dataRequest = systemSettingDA.GetMamoChannels(ChannelId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoChannelId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoChannelId = item["MamoChannelId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoChannelId), "channel_id");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                var responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    #region //紀錄LOG
                                    dataRequest = systemSettingDA.AddMamoLog(MamoChannelId, "", "RestoreChannels", "", "", UserId);
                                    #endregion

                                    #region //重啟頻道(將頻道狀態改為啟用)
                                    dataRequest = systemSettingDA.UpdateChannelStatus(ChannelId, "A", UserId);
                                    #endregion

                                    #region //Response
                                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //AddChannelMembers 新增頻道成員
        [HttpPost]
        [Route("api/MAMO/AddChannelMembers")]
        public void AddChannelMembers(string Company, string SecretKey, int ChannelId, List<string> UserNo, int UserId)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Channels/AddMembers");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得頻道相關資料
                dataRequest = systemSettingDA.GetMamoChannels(ChannelId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoChannelId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoChannelId = item["MamoChannelId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //抓所有使用者
                //UserNo = systemSettingDA.GetAllUser(Company);
                #endregion

                #region //處理UserNo格式
                string userNoListString = JsonConvert.SerializeObject(UserNo);
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoChannelId), "channel_id");
                        multipart.Add(new StringContent(userNoListString), "usernames");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoChannelId, "", "AddChannelMembers", String.Join(", ", UserNo), "", UserId);
                                #endregion

                                #region //新增頻道成員
                                dataRequest = systemSettingDA.AddChannelMembers(ChannelId, UserNo, UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                JObject responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "error")
                                    {
                                        throw new SystemException(item.Value.ToString());
                                    }
                                }

                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //DeleteChannelMembers 刪除頻道成員
        [HttpPost]
        [Route("api/MAMO/DeleteChannelMembers")]
        public void DeleteChannelMembers(string Company, int UserId, string SecretKey, int ChannelId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Channels/RemoveMembers");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //取得頻道相關資料
                dataRequest = systemSettingDA.GetMamoChannels(ChannelId, Company);
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string MamoChannelId = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        MamoChannelId = item["MamoChannelId"].ToString();
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                #region //處理UserNo格式
                string userNoListString = JsonConvert.SerializeObject(UserNo);
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(MamoChannelId), "channel_id");
                        multipart.Add(new StringContent(userNoListString), "usernames");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //紀錄LOG
                                dataRequest = systemSettingDA.AddMamoLog(MamoChannelId, "", "DeleteChannelMembers", String.Join(", ", UserNo), "", UserId);
                                #endregion

                                #region //刪除頻道成員
                                dataRequest = systemSettingDA.DeleteChannelMembers(ChannelId, UserNo);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                JObject responseJson = JObject.Parse(response.Result);
                                foreach (var item in responseJson)
                                {
                                    if (item.Key == "error")
                                    {
                                        throw new SystemException(item.Value.ToString());
                                    }
                                }

                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }

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

        #region //SendTest 寄送訊息 BM段測試
        [HttpPost]
        [Route("api/MAMO/SendTest")]
        public void SendTest()
        {
            try
            {
                string Company = "JMO";
                string SecretKey = "683BE6D57D5157C7D286D75071BBBEE1CD1A5C3BB14978EF0AA1F87C6B8A1094";
                string SendId = "4";
                string Content = @"為您推薦Nexflix本週精彩影集:
                NO.1 兄弟之道
                https://youtu.be/_wGPjnwwahY
                NO.2 京城怪物
                https://youtu.be/T3roUXqoces";
                List<string> Tags = new List<string>();
                Tags.Add("ZY01001");

                List<int> Files = new List<int>();
                Files.Add(19060);
                Files.Add(19061);

                SendMessage(Company, 986, SecretKey, "Channel", SendId, Content, Tags, Files);
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

            //Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //SendMessage 寄送訊息(單人/頻道)
        [HttpPost]
        [Route("api/MAMO/SendMessage")]
        public void SendMessage(string Company, int UserId, string SecretKey, string PushType, string SendId, string Content, List<string> Tags, List<int> Files)
        {
            try
            {
                string apiUrl = Path.Combine(MamoUrl, "Posts/Send");

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //處理UserNo格式
                string TagString = "";
                if (Tags != null && Tags.Count() > 0)
                {
                    TagString = JsonConvert.SerializeObject(Tags);
                }
                #endregion

                if (PushType == "Channel")
                {
                    #region //取得頻道相關資料
                    dataRequest = systemSettingDA.GetMamoChannels(Convert.ToInt32(SendId), Company);
                    jsonResponse = BaseHelper.DAResponse(dataRequest);

                    if (jsonResponse["status"].ToString() == "success")
                    {
                        foreach (var item in jsonResponse["result"])
                        {
                            SendId = item["MamoChannelId"].ToString();
                        }
                    }
                    else
                    {
                        throw new SystemException(jsonResponse["msg"].ToString());
                    }
                    #endregion
                }

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent("bm_sys"), "account");
                        multipart.Add(new StringContent(SendId), "to");
                        multipart.Add(new StringContent(Content), "content");

                        #region //處理標記部分
                        if (TagString.Length > 0)
                        {
                            multipart.Add(new StringContent(TagString), "tags");
                        }
                        #endregion

                        #region //處理檔案部分
                        if (Files != null && Files.Count() > 0)
                        {
                            int count = 1;
                            foreach (var fileId in Files)
                            {
                                #region //取得檔案資料
                                dataRequest = systemSettingDA.GetFile(fileId, -1, "", "", "", "", -1, -1);
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion

                                if (jsonResponse["status"].ToString() == "success")
                                {
                                    foreach (var item in jsonResponse["result"])
                                    {
                                        multipart.Add(new ByteArrayContent((byte[])item["FileContent"]), "file" + count, item["FileName"].ToString() + item["FileExtension"].ToString());
                                    }
                                }

                                count++;
                            }
                        }
                        #endregion

                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //新增推播紀錄
                                dataRequest = systemSettingDA.AddMamoPushNotification(PushType, SendId, Content, Tags, Files, UserId);
                                #endregion

                                #region //Response
                                jsonResponse = BaseHelper.DAResponse(dataRequest);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }
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
        #endregion

        #region //Other
        #region //SendMailTemplate 郵件樣板發送
        [HttpPost]
        public void SendMailTemplate(int MailId = -1)
        {
            try
            {
                WebApiLoginCheck("MailTemplate", "test-mail");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetMailList(MailId);
                #endregion

                JObject data = BaseHelper.DAResponse(dataRequest);

                MailConfig mailConfig = new MailConfig
                {
                    Host = data["result"][0]["Host"].ToString(),
                    Port = Convert.ToInt32(data["result"][0]["Port"]),
                    SendMode = Convert.ToInt32(data["result"][0]["SendMode"]),
                    From = data["result"][0]["MailFrom"].ToString(),
                    Subject = data["result"][0]["MailSubject"].ToString(),
                    Account = data["result"][0]["Account"].ToString(),
                    Password = data["result"][0]["Password"].ToString(),
                    MailTo = data["result"][0]["MailTo"].ToString(),
                    MailCc = data["result"][0]["MailCc"].ToString(),
                    MailBcc = data["result"][0]["MailBcc"].ToString(),
                    HtmlBody = data["result"][0]["MailContent"].ToString(),
                    TextBody = "-"
                };

                BaseHelper.MailSend(mailConfig);

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "發送成功"
                });
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

        #region //QRCodeOutput Qrcode產生
        public ActionResult QRCodeOutput(string inputString = "default", int dimension = 250)
        {
            BarcodeWriter writer = new BarcodeWriter()
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Height = dimension,
                    Width = dimension,
                    PureBarcode = true,
                    Margin = 2
                },
            };

            FileContentResult img = null;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                Bitmap bitmap = writer.Write(inputString);

                bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                img = this.File(memoryStream.GetBuffer(), "image/Jpeg");
            }

            return img;
        }
        #endregion

        #region //PushNotification 推播訊息
        [HttpPost]
        public void PushNotification(string Users = "", string Content = "")
        {
            try
            {
                WebApiLoginCheck("PushMessage", "read");

                #region //Request
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetPushSubscription("", Users);
                #endregion

                JObject data = BaseHelper.DAResponse(dataRequest);
                for (int i = 0; i < data["result"].Count(); i++)
                {
                    JObject pushInfo = JObject.Parse(data["result"][i]["PushInfo"].ToString());

                    List<PushNotificationUser> notificationUsers = JsonConvert.DeserializeObject<List<PushNotificationUser>>(pushInfo["data"].ToString());

                    //List<NotificationActions> notificationActions = new List<NotificationActions>
                    //{
                    //    new NotificationActions
                    //    {
                    //        action = "goto",
                    //        title = "前往"
                    //    },
                    //    new NotificationActions
                    //    {
                    //        action = "cancel",
                    //        title = "取消"
                    //    }
                    //};

                    //List<NotificationActionUrls> notificationActionUrls = new List<NotificationActionUrls>
                    //{
                    //    new NotificationActionUrls
                    //    {
                    //        action = "goto",
                    //        url = "/Home/Gate"
                    //    },
                    //    new NotificationActionUrls
                    //    {
                    //        action = "cancel",
                    //        url = "/SystemSetting/PushMessage"
                    //    }
                    //};

                    WebPushHelper webPushHelper = new WebPushHelper();
                    webPushHelper.SendPush(notificationUsers, "JMO BM", Content);
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "推送成功"
                });
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

        #region//DownloadBmFile
        public ActionResult DownloadProjectFile(string FunctionCode, int FileId)
        {
            try
            {
                WebApiLoginCheck(FunctionCode, "download");

                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.AddDownloadLog(FileId);
                return RedirectToAction("GetFile", "Web", new { fileId = FileId });
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
                Response.Write(jsonResponse.ToString());
                return new EmptyResult();
            }
        }
        #endregion

        #region //KeyenceTest
        public string returnMsg = "";
        [HttpPost]
        public void KeyenceScan(string DeviceIP = "", int DevicePort = -1)
        {
            try
            {
                //DeviceIP = "192.168.7.240";
                if (DevicePort <= 0) DevicePort = 9004;
                string message = "4C4F4E0D";
                byte[] data = HexStringToByteArray(message);
                // 创建TCP客户端套接字
                TcpClient clientSocket = new TcpClient();
                clientSocket.Connect(DeviceIP, DevicePort);
                if (clientSocket.Connected)
                {
                    NetworkStream stream = clientSocket.GetStream();
                    stream.Write(data, 0, data.Length);
                    // 接收服务器回复
                    byte[] responseData = new byte[256];
                    StringBuilder responseMessage = new StringBuilder();
                    int bytes = stream.Read(responseData, 0, responseData.Length);
                    responseMessage.Append(Encoding.ASCII.GetString(responseData, 0, bytes));

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = responseMessage
                    });

                    // 关闭套接字
                    stream.Dispose();
                    stream.Close();
                    clientSocket.Close();
                }
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

        private static byte[] HexStringToByteArray(string hexString)
        {
            int len = hexString.Length;
            byte[] data = new byte[len / 2];
            for (int i = 0; i < len; i += 2)
            {
                data[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }
            return data;
        }

        public static bool IsConnet = true;
        public void Connet(string Iptxt, int Port)//接收参数是目标ip地址和目标端口号。客户端无须关心本地端口号
        {
            //创建一个新的Socket对象
            Socket client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IsConnet = true;//注意，此处是全局变量，将其设置为true
                    //将方法写进线程中
            Thread thread = new Thread(() =>
            {
                while (IsConnet)//循环
                {
                    try
                    {
                        client.Connect(IPAddress.Parse(Iptxt), Port);//尝试连接，失败则会跳去catch
                        IsConnet = false;//成功连接后修改bool值为false,这样下一步循环就不再执行。
                        break;//在此处加上break，成功就跳出循环，避免死循环
                    }
                    catch
                    {
                        client.Close();//先关闭
                                       /*使用新的客户端资源覆盖，上一个已经废弃。如果继续使用以前的资源进行连接，
                                       即使参数正确， 服务器全部打开也会无法连接*/
                        client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                        Thread.Sleep(1000);//等待1s再去重连
                    }
                }
                /*这里不一样就是放接收线程，在连接上后break出来，执行。
                因为需要带参数，所以要用到特别的ParameterizedThreadStart，
                然后开始线程。↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓*/
                Thread thread2 = new Thread(new ParameterizedThreadStart(ClientReceiveData));//接收线程方法
                thread2.IsBackground = true;//该值指示某个线程是否为后台线程。
                thread2.Start(client);//参数是用我们自建的Socket对象，就是上面的Socket client=new……

            });
            thread.IsBackground = true;//设置为后台线程，在程序退出时自己会自动释放
            thread.Start();//开始执行线程
        }

        public void ClientReceiveData(object socket)//TCPClient消息的方法
        {
            var ProxSocket = socket as Socket;//处理上一步传过来的Socket函数
            byte[] data = new byte[1024 * 1024];//接收消息的缓冲区
            while (!IsConnet)//同样循环中止的条件
            {
                int len = 0;//记录消息长度，以及判断是否连接
                try
                {
                    //连接函数Receive会将数据放入data,从0开始放，之后返回数据长度。
                    len = ProxSocket.Receive(data, 0, data.Length, SocketFlags.None);
                }
                catch (Exception)
                {
                    //异常退出
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    Connet("192.168.7.240", 9004);
                    IsConnet = false;
                    return;
                }

                if (len <= 0)
                {
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    Connet("192.168.7.240", 9004);
                    IsConnet = false;
                    return;
                }

                //这里做你想要对消息做的处理
                returnMsg = Encoding.Default.GetString(data, 0, len);//二进制数组转换成字符串……
            }
        }
        #endregion
        #endregion
    }
}