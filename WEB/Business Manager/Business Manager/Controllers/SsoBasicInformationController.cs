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
    public class SsoBasicInformationController : WebController
    {
        private SsoBasicInformationDA ssoBasicInformationDA = new SsoBasicInformationDA();

        #region //View
        public ActionResult DemandSource()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DemandFlow()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DemandRole()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetDemandSource 取得需求來源資料
        [HttpPost]
        public void GetDemandSource(int SourceId = -1, string SearchKey = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandSource", "read,constrained-data");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.GetDemandSource(SourceId, SearchKey, Status
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

        #region //GetDemandFlow 取得流程資料
        [HttpPost]
        public void GetDemandFlow(int FlowId = -1, string SearchKey = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandFlow", "read,constrained-data");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.GetDemandFlow(FlowId, SearchKey, Status
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

        #region //GetDemandRole 取得角色資料
        [HttpPost]
        public void GetDemandRole(int RoleId = -1, string SearchKey = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandRole", "read,constrained-data");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.GetDemandRole(RoleId, SearchKey, Status
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
        public void GetRoleUser(int RoleId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DemandRole", "role-user");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.GetRoleUser(RoleId, CompanyId, DepartmentId, SearchKey
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
        #region //AddDemandSource 需求來源資料新增
        [HttpPost]
        public void AddDemandSource(string SourceNo = "", string SourceName = "", string SourceDesc = "")
        {
            try
            {
                WebApiLoginCheck("DemandSource", "add");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.AddDemandSource(SourceNo, SourceName, SourceDesc);
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

        #region //AddDemandFlow 流程資料新增
        [HttpPost]
        public void AddDemandFlow(string FlowNo = "", string FlowName = "", string FlowImage = "")
        {
            try
            {
                WebApiLoginCheck("DemandFlow", "add");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.AddDemandFlow(FlowNo, FlowName, FlowImage);
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

        #region //AddDemandRole 角色資料新增
        [HttpPost]
        public void AddDemandRole(string RoleName = "")
        {
            try
            {
                WebApiLoginCheck("DemandRole", "add");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.AddDemandRole(RoleName);
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

        #region //AddRoleUser 角色使用者資料新增
        [HttpPost]
        public void AddRoleUser(int RoleId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("DemandRole", "role-user");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.AddRoleUser(RoleId, Users);
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
        #region //UpdateDemandSource 需求來源資料更新
        [HttpPost]
        public void UpdateDemandSource(int SourceId = -1, string SourceName = "", string SourceDesc = "")
        {
            try
            {
                WebApiLoginCheck("DemandSource", "update");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandSource(SourceId, SourceName, SourceDesc);
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

        #region //UpdateDemandSourceStatus 需求來源狀態更新
        [HttpPost]
        public void UpdateDemandSourceStatus(int SourceId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandSource", "status-switch");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandSourceStatus(SourceId);
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

        #region //UpdateDemandFlow 流程資料更新
        [HttpPost]
        public void UpdateDemandFlow(int FlowId = -1, string FlowNo = "", string FlowName = "", string FlowImage = "")
        {
            try
            {
                WebApiLoginCheck("DemandFlow", "update");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandFlow(FlowId, FlowNo, FlowName, FlowImage);
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

        #region //UpdateDemandFlowStatus 流程狀態更新
        [HttpPost]
        public void UpdateDemandFlowStatus(int FlowId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandFlow", "status-switch");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandFlowStatus(FlowId);
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

        #region //UpdateDemandRole 角色資料更新
        [HttpPost]
        public void UpdateDemandRole(int RoleId = -1, string RoleName = "")
        {
            try
            {
                WebApiLoginCheck("DemandRole", "update");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandRole(RoleId, RoleName);
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

        #region //UpdateDemandRoleStatus 角色狀態更新
        [HttpPost]
        public void UpdateDemandRoleStatus(int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandRole", "status-switch");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.UpdateDemandRoleStatus(RoleId);
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
        #region DeleteRoleUser 角色使用者資料刪除
        [HttpPost]
        public void DeleteRoleUser(int RoleId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("DemandRole", "role-user");

                #region //Request
                ssoBasicInformationDA = new SsoBasicInformationDA();
                dataRequest = ssoBasicInformationDA.DeleteRoleUser(RoleId, UserId);
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