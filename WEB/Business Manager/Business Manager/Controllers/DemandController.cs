using Helpers;
using Newtonsoft.Json.Linq;
using SSODA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class DemandController : WebController
    {
        private DemandDA demandDA = new DemandDA();

        #region //View
        public ActionResult DemandPlatform()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameFlowPreview()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameDemandTrack()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetDemand 取得需求資料
        [HttpPost]
        public void GetDemand(int DemandId = -1, int SourceId = -1, string DemandNo = "", int DemandDepartment = -1
            , int DemandCustomer = -1, string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemand(DemandId, SourceId, DemandNo, DemandDepartment
                    , DemandCustomer, StartDate, EndDate, Status
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

        #region //GetDemandCertificate 取得需求單憑證資料
        [HttpPost]
        public void GetDemandCertificate(int CertificateId = -1, int FileId = -1, int DemandId = -1, string CertificateDesc = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "cert");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandCertificate(CertificateId, FileId, DemandId, CertificateDesc
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

        #region //GetDemandFlowSettingDiagram 取得需求流程設定資料(流程圖)
        [HttpPost]
        public void GetDemandFlowSettingDiagram(int DemandId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandFlowSettingDiagram(DemandId);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                #region //Response
                if (jsonResponse["status"].ToString() == "error")
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = jsonResponse["msg"].ToString()
                    });
                }
                else if (jsonResponse["status"].ToString() == "errorForDA")
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "errorForDA",
                        msg = jsonResponse["msg"].ToString()
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        dataShapes = jsonResponse["dataShapes"],
                        dataLines = jsonResponse["dataLines"]
                    });
                }
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

        #region //GetDemandFlowSetting 取得需求流程設定資料
        [HttpPost]
        public void GetDemandFlowSetting(int SettingId = -1, int DemandId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandFlowSetting(SettingId, DemandId);
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

        #region //GetDemandFlowUser 取得需求流程使用者對應資料
        [HttpPost]
        public void GetDemandFlowUser(int SettingId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandFlowUser(SettingId);
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

        #region //GetDemandItemNo 取得需求流程對應項目資料
        [HttpPost]
        public void GetDemandItemNo(int SettingId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandItemNo(SettingId);
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

        #region //GetDemandNotificationLog 取得需求平台通知紀錄
        [HttpPost]
        public void GetDemandNotificationLog(int SettingId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.GetDemandNotificationLog(SettingId);
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
        #region //AddDemand 需求資料新增
        [HttpPost]
        public void AddDemand(int SourceId = -1, string DemandDesc = "", string DemandDate = "", string DemandDeadline = ""
            , string DepCus = "", int DemandDepartment = -1, int DemandCustomer = -1, int DemandUser = -1, int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandSource", "add");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.AddDemand(SourceId, DemandDesc, DemandDate, DemandDeadline
                    , DepCus, DemandDepartment, DemandCustomer, DemandUser, ItemTypeId);
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

        #region //AddDemandCertificate 需求單憑證新增
        [HttpPost]
        public void AddDemandCertificate(int FileId = -1, int DemandId = -1, string CertificateDesc = "")
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "cert");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.AddDemandCertificate(FileId, DemandId, CertificateDesc);
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

        #region //AddDemandFlowUser 需求流程使用者對應新增
        [HttpPost]
        public void AddDemandFlowUser(int SettingId = -1, string UserRole = "", string Users = "", string Roles = "")
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "user-setting");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.AddDemandFlowUser(SettingId, UserRole, Users, Roles);
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
        #region //UpdateDemand 需求資料更新
        [HttpPost]
        public void UpdateDemand(int DemandId = -1, int SourceId = -1, string DemandDesc = "", string DemandDate = "", string DemandDeadline = ""
            , string DepCus = "", int DemandDepartment = -1, int DemandCustomer = -1, int DemandUser = -1, int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandSource", "update");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.UpdateDemand(DemandId, SourceId, DemandDesc, DemandDate, DemandDeadline
                    , DepCus, DemandDepartment, DemandCustomer, DemandUser, ItemTypeId);
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

        #region //UpdateDemandCertificate 需求單憑證更名
        [HttpPost]
        public void UpdateDemandCertificate(int CertificateId = -1, string NewFileName = "", string CertificateDesc = "")
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "cert");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.UpdateDemandCertificate(CertificateId, NewFileName, CertificateDesc);
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

        #region //UpdateDemandConfirm 需求確認
        [HttpPost]
        public void UpdateDemandConfirm(int DemandId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "confirm");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.UpdateDemandConfirm(DemandId);
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

        #region //UpdateDemandReConfirm 需求反確認
        [HttpPost]
        public void UpdateDemandReConfirm(int DemandId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "reconfirm");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.UpdateDemandReConfirm(DemandId);
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

        #region //UpdateDemandFlowUserStatus 需求流程使用者狀態更新
        [HttpPost]
        public void UpdateDemandFlowUserStatus(int SettingId = -1, int RoleId = -1, int UserId = -1
            , string MailAdviceStatus = "", string PushAdviceStatus = "", string WorkWeixinStatus = "")
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "user-setting");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.UpdateDemandFlowUserStatus(SettingId, RoleId, UserId
                    , MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus);
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
        #region //DeleteDemand 需求資料刪除
        [HttpPost]
        public void DeleteDemand(int DemandId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "delete");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.DeleteDemand(DemandId);
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

        #region //DeleteDemandCertificate 需求單憑證刪除
        [HttpPost]
        public void DeleteDemandCertificate(int CertificateId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "cert");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.DeleteDemandCertificate(CertificateId);
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

        #region //DeleteDemandFlowUser 需求流程使用者對應刪除
        [HttpPost]
        public void DeleteDemandFlowUser(int SettingId = -1, int RoleId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "user-setting");

                #region //Request
                demandDA = new DemandDA();
                dataRequest = demandDA.DeleteDemandFlowUser(SettingId, RoleId, UserId);
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
        #region //UpdateDemandFlowTrigger 需求流程狀態觸發
        [HttpPost]
        [Route("api/SSO/DemandFlowTrigger")]
        public void UpdateDemandFlowTrigger(string Company, string SecretKey, string DemandNo, string FlowNo, string ItemNo, string UserNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DemandFlowTrigger");
                #endregion

                #region //Request
                dataRequest = demandDA.UpdateDemandFlowTrigger(DemandNo, FlowNo, ItemNo, UserNo);
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

        #region //UpdateDemandItemBind 需求流程項目綁定
        [HttpPost]
        [Route("api/SSO/DemandItemBind")]
        public void UpdateDemandItemBind(string Company, string SecretKey, string DemandNo, string FlowNo, string ItemNo, string UserNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DemandItemBind");
                #endregion

                #region //Request
                dataRequest = demandDA.UpdateDemandItemBind(DemandNo, FlowNo, ItemNo, UserNo);
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