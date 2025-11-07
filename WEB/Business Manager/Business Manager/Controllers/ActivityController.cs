using EBPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ActivityController : WebController
    {
        private ActivityDA activityDA = new ActivityDA();

        #region //View
        public ActionResult ActivityManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetActivity 取得活動資料
        [HttpPost]
        public void GetActivity(int ActivityId = -1, int AnnualId = -1, string ActivityName = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "read");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetActivity(ActivityId, AnnualId, ActivityName, StartDate, EndDate, Status
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

        #region //GetAwards 取得活動獎項資料
        [HttpPost]
        public void GetAwards(int AwardsId = -1, int ActivityId = -1, string AwardsName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetAwards(AwardsId, ActivityId, AwardsName
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

        #region //GetAwardsUser 取得獎項人員資料
        [HttpPost]
        public void GetAwardsUser(int AwardsId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetAwardsUser(AwardsId, SearchKey);
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

        #region //GetPoints 取得活動積分資料
        [HttpPost]
        public void GetPoints(int ActivityPointId = -1, int ActivityId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "points");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetPoints(ActivityPointId, ActivityId
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

        #region //GetParticipants 取得參與人員資料
        [HttpPost]
        public void GetParticipants(int ActivityId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "participants");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetParticipants(ActivityId, SearchKey
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
        #region //AddActivity 活動資料新增
        [HttpPost]
        public void AddActivity(int AnnualId = -1, string ActivityName = "", string ActivityStartDate = "", string ActivityEndDate = "", string ActivityContent = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "add");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.AddActivity(AnnualId, ActivityName, ActivityStartDate, ActivityEndDate, ActivityContent);
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

        #region //AddAwards 活動獎項資料新增
        [HttpPost]
        public void AddAwards(int ActivityId = -1, string AwardsName = "", int AwardsPrice = -1, int AwardsQuantity = -1
            , string AwardsUnit = "", string AwardsRemark = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.AddAwards(ActivityId, AwardsName, AwardsPrice, AwardsQuantity
                    , AwardsUnit, AwardsRemark);
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

        #region //AddAwardsUser 獎項人員資料新增
        [HttpPost]
        public void AddAwardsUser(int AwardsId = -1, string GroupName = "", string Users = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.AddAwardsUser(AwardsId, GroupName, Users, Remark);
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

        #region //AddPoints 活動積分資料新增
        [HttpPost]
        public void AddPoints(int ActivityId = -1, string PointType = "", int PointQty = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "points");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.AddPoints(ActivityId, PointType, PointQty, Remark);
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

        #region //AddParticipants 參與人員資料新增
        [HttpPost]
        public void AddParticipants(int ActivityId = -1, string Users = "", string ActivityPoints = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "participants");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.AddParticipants(ActivityId, Users, ActivityPoints);
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
        #region //UpdateActivity 活動資料更新
        [HttpPost]
        public void UpdateActivity(int ActivityId = -1, int AnnualId = -1, string ActivityName = "", string ActivityStartDate = "", string ActivityEndDate = "", string ActivityContent = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "update");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateActivity(ActivityId, AnnualId, ActivityName, ActivityStartDate, ActivityEndDate, ActivityContent);
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

        #region //UpdateActivityStatus 活動狀態更新
        [HttpPost]
        public void UpdateActivityStatus(int ActivityId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "status-switch");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateActivityStatus(ActivityId);
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

        #region //UpdateAwards 活動獎項資料更新
        [HttpPost]
        public void UpdateAwards(int AwardsId = -1, string AwardsName = "", int AwardsPrice = -1, int AwardsQuantity = -1
            , string AwardsUnit = "", string AwardsRemark = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateAwards(AwardsId, AwardsName, AwardsPrice, AwardsQuantity
                    , AwardsUnit, AwardsRemark);
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

        #region //UpdateAwardsStatus 活動獎項狀態更新
        [HttpPost]
        public void UpdateAwardsStatus(int AwardsId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateAwardsStatus(AwardsId);
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

        #region //UpdateAwardsSort 活動獎項順序調整
        [HttpPost]
        public void UpdateAwardsSort(int ActivityId = -1, string AwardsList = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "awards");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateAwardsSort(ActivityId, AwardsList);
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

        #region //UpdatePoints 活動積分資料更新
        [HttpPost]
        public void UpdatePoints(int ActivityPointId = -1, string PointType = "", int PointQty = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "points");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdatePoints(ActivityPointId, PointType, PointQty, Remark);
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

        #region //UpdatePointStatus 活動獎項狀態更新
        [HttpPost]
        public void UpdatePointStatus(int ActivityPointId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "points");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdatePointStatus(ActivityPointId);
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

        #region //UpdateRedeemPoint 參與人員狀態更新
        [HttpPost]
        public void UpdateRedeemPoint(int ActivityId = -1, int ParticipantId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "redeem");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.UpdateRedeemPoint(ActivityId, ParticipantId);
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
        #region //DeleteActivity 活動資料刪除
        [HttpPost]
        public void DeleteActivity(int ActivityId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "delete");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.DeleteActivity(ActivityId);
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

        #region //DeleteAwards 活動獎項資料刪除
        [HttpPost]
        public void DeleteAwards(int AwardsId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "delete");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.DeleteAwards(AwardsId);
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

        #region //DeleteAwardsUser 刪除獎項人員資料
        [HttpPost]
        public void DeleteAwardsUser(int AwardsId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "delete");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.DeleteAwardsUser(AwardsId, UserId);
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

        #region //DeletePoints 刪除活動積分資料
        [HttpPost]
        public void DeletePoints(int ActivityPointId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "points");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.DeletePoints(ActivityPointId);
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

        #region //DeleteParticipants 刪除參與人員資料
        [HttpPost]
        public void DeleteParticipants(int ParticipantId = -1)
        {
            try
            {
                WebApiLoginCheck("ActivityManagement", "participants");

                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.DeleteParticipants(ParticipantId);
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