using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using QMSDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class QcPlanningProcessController : WebController
    {
        private QcPlanningProcessDA qcPlanningProcessDA = new QcPlanningProcessDA();

        #region //View
        public ActionResult QcPlanningPlatform()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult Login()
        {
            return View();
        }

        public ActionResult Index()
        {
            QcPlanningProcessLoginCheck();

            return View();
        }

        public ActionResult QcPlanningHistory()
        {
            QcPlanningProcessLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetLoginUserInfo 取得使用者資料
        [HttpPost]
        public void GetLoginUserInfo()
        {
            try
            {
                WebApiLoginCheck("QcPlanningPlatform", "read");

                #region //Request
                qcPlanningProcessDA = new QcPlanningProcessDA();
                dataRequest = qcPlanningProcessDA.GetLoginUserInfo();
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

        #region //GetUserIP 取得使用者資料
        [HttpPost]
        public void GetUserIP()
        {
            try
            {
                WebApiLoginCheck("QcPlanningPlatform", "read");

                string DeviceIdentifierCode = BaseHelper.ClientIP();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    data = DeviceIdentifierCode
                });
                dataRequest = jsonResponse.ToString();
                jsonResponse = BaseHelper.DAResponse(dataRequest);
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

        #region //GetQcMachineModePlanning 取得量測機台排程資料
        [HttpPost]
        public void GetQcMachineModePlanning(int QmmpId = -1, int QrpId = -1, int QmmDetailId = -1, int QcMachineModeId = -1, string StartDate = "", string Status = "", string MtlItemName = "", int QcRecordId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = 1)
        {
            try
            {
                WebApiLoginCheck("QcPlanningPlatform", "read");

                #region //Request
                qcPlanningProcessDA = new QcPlanningProcessDA();
                dataRequest = qcPlanningProcessDA.GetQcMachineModePlanning(QmmpId, QrpId, QmmDetailId, QcMachineModeId, StartDate, Status, MtlItemName, QcRecordId
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

        #endregion

        #region //Update
        #region //UpdateQcPlanningProcess 更新量測排程過站狀態
        [HttpPost]
        public void UpdateQcPlanningProcess(int QmmpId = -1, string PlanningStatus = "")
        {
            try
            {
                WebApiLoginCheck("QcPlanningPlatform", "update");

                #region //Request
                qcPlanningProcessDA = new QcPlanningProcessDA();
                dataRequest = qcPlanningProcessDA.UpdateQcPlanningProcess(QmmpId, PlanningStatus);
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

        #endregion

        #region //Custom
        #region //登入
        [HttpPost]
        public void QcPlanningProcessLogin(string SystemKey, string Account)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetSubSystemLogin(SystemKey, Account, "QcPlanningProcess");
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), Account, "QcPlanningProcess");
                }

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

        #region //View檢查登入
        [NonAction]
        public void QcPlanningProcessLoginCheck()
        {
            bool verify = LoginVerify("QcPlanningProcess");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "QcPlanningProcess");

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }
        #endregion

        #region //Api檢查登入
        [NonAction]
        public void QcPlanningProcessApiLoginCheck()
        {
            bool verify = LoginVerify("QcPlanningProcess");

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
        }
        #endregion
        #endregion
    }
}