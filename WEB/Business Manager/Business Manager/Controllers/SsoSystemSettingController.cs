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
    public class SsoSystemSettingController : WebController
    {
        private SsoSystemSettingDA ssoSystemSettingDA = new SsoSystemSettingDA();

        #region //View
        public ActionResult SourceFlowSetting()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetSourceFlowSetting 取得需求來源流程設定資料
        [HttpPost]
        public void GetSourceFlowSetting(int SourceId = -1, int ExcludeSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.GetSourceFlowSetting(SourceId, ExcludeSettingId);
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

        #region //GetSourceFlowUser 取得需求來源流程使用者對應資料
        [HttpPost]
        public void GetSourceFlowUser(int SettingId)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.GetSourceFlowUser(SettingId);
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

        #region //GetSourceFlowSettingDiagram 取得需求來源流程設定資料(流程圖)
        [HttpPost]
        public void GetSourceFlowSettingDiagram(int SourceId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.GetSourceFlowSettingDiagram(SourceId);
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
        #endregion

        #region //Add
        #region //AddSourceFlowSetting 需求來源流程設定資料新增
        [HttpPost]
        public void AddSourceFlowSetting(int SourceId = -1, int FlowId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.AddSourceFlowSetting(SourceId, FlowId);
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

        #region //AddSourceFlowLink 需求來源流程設定連結新增
        [HttpPost]
        public void AddSourceFlowLink(string SourceSettings = "", int TargetSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.AddSourceFlowLink(SourceSettings, TargetSettingId);
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

        #region //AddSourceFlowUser 需求來源流程使用者對應新增
        [HttpPost]
        public void AddSourceFlowUser(int SettingId = -1, string UserRole = "", string Users = "", string Roles = "")
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.AddSourceFlowUser(SettingId, UserRole, Users, Roles);
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
        #region //UpdateSourceFlowUserStatus 需求來源流程使用者狀態更新
        [HttpPost]
        public void UpdateSourceFlowUserStatus(int SettingId = -1, int RoleId = -1, int UserId = -1
            , string MailAdviceStatus = "", string PushAdviceStatus = "", string WorkWeixinStatus = "")
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.UpdateSourceFlowUserStatus(SettingId, RoleId, UserId
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

        #region //UpdateSourceFlowSettingCoordinates 需求來源流程設定座標更新
        [HttpPost]
        public void UpdateSourceFlowSettingCoordinates(int SourceId = -1, string DiagramData = "")
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.UpdateSourceFlowSettingCoordinates(SourceId, DiagramData);
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
        #region //DeleteSourceFlowUser 需求來源流程使用者對應刪除
        [HttpPost]
        public void DeleteSourceFlowUser(int SettingId = -1, int RoleId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.DeleteSourceFlowUser(SettingId, RoleId, UserId);
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

        #region //DeleteSourceFlowSetting 需求來源流程設定資料刪除
        [HttpPost]
        public void DeleteSourceFlowSetting(int SettingId = -1)
        {
            try
            {
                WebApiLoginCheck("SourceFlowSetting", "flow-setting");

                #region //Request
                ssoSystemSettingDA = new SsoSystemSettingDA();
                dataRequest = ssoSystemSettingDA.DeleteSourceFlowSetting(SettingId);
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