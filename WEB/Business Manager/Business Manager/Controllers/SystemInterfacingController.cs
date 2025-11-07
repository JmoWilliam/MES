using BASDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Net;
using System.Diagnostics;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;

namespace Business_Manager.Controllers
{
    public class SystemInterfacingController : WebController
    {
        private SystemInterfacingDA systemInterfacingDA = new SystemInterfacingDA();

        #region //View
        public ActionResult Interfacing()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult Operation()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult InterfacingControlPanel()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetInterfacing 取得系統介接設定資料
        [HttpPost]
        public void GetInterfacing(int InterfacingId = -1, int CompanyId = -1, string InterfacingType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandPlatform", "read");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.GetInterfacing(InterfacingId, CompanyId, InterfacingType
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

        #region //GetOperation 取得作業模式資料
        [HttpPost]
        public void GetOperation(int OperationId = -1, int CompanyId = -1, int InterfacingId = -1, string OperationNo = ""
            , string OperationName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Operation", "read");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.GetOperation(OperationId, CompanyId, InterfacingId, OperationNo
                    , OperationName, Status
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
        #region //AddInterfacing 系統介接設定資料新增
        [HttpPost]
        public void AddInterfacing(string InterfacingJson = "")
        {
            try
            {
                WebApiLoginCheck("Interfacing", "add");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.AddInterfacing(InterfacingJson);
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

        #region //AddOperation 作業模式資料新增
        [HttpPost]
        public void AddOperation(string OperationJson = "")
        {
            try
            {
                WebApiLoginCheck("Operation", "add");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.AddOperation(OperationJson);
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
        #region //UpdateInterfacing 系統介接設定資料更新
        [HttpPost]
        public void UpdateInterfacing(int InterfacingId = -1, string InterfacingJson = "")
        {
            try
            {
                WebApiLoginCheck("Interfacing", "update");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.UpdateInterfacing(InterfacingId, InterfacingJson);
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

        #region //UpdateInterfacingConnectionStatus 介接連線狀態更新
        [HttpPost]
        public void UpdateInterfacingConnectionStatus(int ConnectionId = -1)
        {
            try
            {
                WebApiLoginCheck("Interfacing", "status-switch");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.UpdateInterfacingConnectionStatus(ConnectionId);
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

        #region //UpdateOperation 作業模式資料新增
        [HttpPost]
        public void UpdateOperation(int OperationId = -1, string OperationJson = "")
        {
            try
            {
                WebApiLoginCheck("Operation", "update");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.UpdateOperation(OperationId, OperationJson);
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

        #region //UpdateOperationStatus 作業模式狀態更新
        [HttpPost]
        public void UpdateOperationStatus(int OperationId = -1)
        {
            try
            {
                WebApiLoginCheck("Operation", "status-switch");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.UpdateOperationStatus(OperationId);
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
        #region //DeleteInterfacing 系統介接設定資料刪除
        [HttpPost]
        public void DeleteInterfacing(int InterfacingId = -1)
        {
            try
            {
                WebApiLoginCheck("Interfacing", "delete");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.DeleteInterfacing(InterfacingId);
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

        #region //DeleteInterfacingConnection 介接連線資料刪除
        [HttpPost]
        public void DeleteInterfacingConnection(int ConnectionId = -1)
        {
            try
            {
                WebApiLoginCheck("Interfacing", "delete");

                #region //Request
                systemInterfacingDA = new SystemInterfacingDA();
                dataRequest = systemInterfacingDA.DeleteInterfacingConnection(ConnectionId);
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