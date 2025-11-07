using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NiceLabel.SDK;
using System.IO;
using System.Text;
using System.Reflection;
using MESDA;
using ClosedXML.Excel;

namespace Business_Manager.Controllers
{
    public class EventController : WebController
    {
        private EventDA eventDA = new EventDA();
        #region //View
        // GET: Event
        
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult UserEventManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MachineEventManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessEventManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region//Get

        #region //GetUserEvent 取得人員事件
        [HttpPost]
        public void GetUserEvent(int UserEventId =-1, int UserEventItemId = -1
            , string StartDateBegin = "", string StartDateEnd = "", int Operator = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventManagement", "read,constrained-data");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.GetUserEvent(UserEventId,UserEventItemId
                    , StartDateBegin, StartDateEnd, Operator
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

        #region //GetMachineEvent 取得機台事件
        [HttpPost]
        public void GetMachineEvent(int MachineEventId = -1, int MachineEventItemId = -1, int ShopId = -1,int MachineId = -1, int TypeNo = -1
            , string StartDateBegin = "", string StartDateEnd = "", int Operator = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineEventManagement", "read,constrained-data");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.GetMachineEvent(MachineEventId, MachineEventItemId, ShopId, MachineId
                    , StartDateBegin, StartDateEnd, Operator
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

        #region //GetProcessEvent 取得製程參數事件
        [HttpPost]
        public void GetProcessEvent(int BarcodeEventId = -1, int ProcessEventItemId = -1, int ModeId = -1, int ParameterId = -1, int ProcessId = -1, int TypeNo = -1
            , string StartDateBegin = "", string StartDateEnd = "", int OperatorId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessEventManagement", "read,constrained-data");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.GetProcessEvent(BarcodeEventId, ProcessEventItemId, ModeId, ParameterId, ProcessId, TypeNo
                    , StartDateBegin, StartDateEnd, OperatorId
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

        #region//Add

        #region //AddUserEvent 人員事件新增 -- Ellie 2023.11.02
        [HttpPost]
        public void AddUserEvent(int UserEventId = -1, int UserEventItemId = -1
            , string StartYmd = "", string StartTime= ""
            , string FinishYmd = "", string FinishTime=""
            , int Operator = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("UserEventManagement", "add");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.AddUserEvent( UserEventId,  UserEventItemId
                , StartYmd, StartTime
                , FinishYmd, FinishTime
                , Operator, Status);
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

        #region //AddMachineEvent  機台事件新增 -- Ellie 2023.11.08
        [HttpPost]
        public void AddMachineEvent(int MachineEventId = -1, int MachineEventItemId = -1
            , int Operator = -1
            , string StartYmd = "", string StartTime= ""
            , string FinishYmd = "", string FinishTime=""
            )
        {
            try
            {
                WebApiLoginCheck("MachineEventManagement", "add");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.AddMachineEvent(MachineEventId, MachineEventItemId
                , Operator
                , StartYmd, StartTime
                , FinishYmd, FinishTime
                );
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

        #region //AddProcessEvent  製程事件新增 -- Ellie 2023.11.14
        [HttpPost]
        public void AddProcessEvent(int BarcodeEventId = -1, int ProcessEventItemId = -1
            , int OperatorId = -1
            , string StartYmd = "", string StartTime = ""
            , string FinishYmd = "", string FinishTime = ""
            )
        {
            try
            {
                WebApiLoginCheck("ProcessEventManagement", "add");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.AddProcessEvent(BarcodeEventId, ProcessEventItemId
                , OperatorId
                , StartYmd, StartTime
                , FinishYmd, FinishTime
                );
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

        #region//Update

        #region //UpdateUserEvent 人員事件資料更新
        [HttpPost]
        public void UpdateUserEvent(int UserEventId = -1, int UserEventItemId = -1, string StartYmd = "", string StartTime = ""
            , string FinishYmd = "", string FinishTime = "", int Operator = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventManagement", "update");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.UpdateUserEvent(UserEventId,  UserEventItemId,  StartYmd,  StartTime
                ,  FinishYmd,  FinishTime, Operator);
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

        #region //UpdateMachineEvent 機台事件資料更新
        [HttpPost]
        public void UpdateMachineEvent(int MachineEventId = -1, int MachineEventItemId = -1
            , int ShopId = -1, int MachineId = -1
            , int Operator = -1, string StartYmd = "", string StartTime = ""
            , string FinishYmd = "", string FinishTime = "")
        {
            try
            {
                WebApiLoginCheck("MachineEventManagement", "update");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.UpdateMachineEvent(MachineEventId, MachineEventItemId, ShopId, MachineId, Operator, StartYmd, StartTime
                , FinishYmd, FinishTime);
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

        #region //UpdateProcessEvent 製程事件資料更新
        [HttpPost]
        public void UpdateProcessEvent(int BarcodeEventId = -1, int ProcessEventItemId = -1
            , int OperatorId = -1, string StartYmd = "", string StartTime = ""
            , string FinishYmd = "", string FinishTime = "")
        {
            try
            {
                WebApiLoginCheck("ProcessEventManagement", "update");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.UpdateProcessEvent(BarcodeEventId, ProcessEventItemId, OperatorId, StartYmd, StartTime
                , FinishYmd, FinishTime);
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

        #region//Delete

        #region //DeleteUserEvent 人員事件刪除
        [HttpPost]
        public void DeleteUserEvent(int UserEventId = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventManagement", "delete");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.DeleteUserEvent(UserEventId);
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

        #region //DeleteMachineEvent 機台事件刪除
        [HttpPost]
        public void DeleteMachineEvent(int MachineEventId = -1)
        {
            try
            {
                WebApiLoginCheck("MachineEventManagement", "delete");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.DeleteMachineEvent(MachineEventId);
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

        #region //DeleteProcessEvent 製程事件刪除
        [HttpPost]
        public void DeleteProcessEvent(int BarcodeEventId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessEventManagement", "delete");

                #region //Request
                eventDA = new EventDA();
                dataRequest = eventDA.DeleteProcessEvent(BarcodeEventId);
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