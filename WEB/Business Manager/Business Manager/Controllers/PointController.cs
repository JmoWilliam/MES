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
    public class PointController : WebController
    {
        private PointDA pointDA = new PointDA();

        #region //View
        public ActionResult PointsReward()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PersonalPoints()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRedeemPrize 取得兌換獎勵資料
        [HttpPost]
        public void GetRedeemPrize(int RedeemPrizeId = -1, int AnnualId = -1, string PrizeName = "", int RedeemPoint = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PointsReward", "read");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.GetRedeemPrize(RedeemPrizeId, AnnualId, PrizeName, RedeemPoint, Status
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

        #region //GetUserPoint 取得人員積分資料
        [HttpPost]
        public void GetUserPoint(int DepartmentId = -1, int UserId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PersonalPoints", "read");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.GetUserPoint(DepartmentId, UserId
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

        #region //GetPersonalPoints 取得個人積分資料
        [HttpPost]
        public void GetPersonalPoints(int UserId = -1, int AnnualId = -1, string PointType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PersonalPoints", "read");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.GetPersonalPoints(UserId, AnnualId, PointType
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
        #region //AddRedeemPrize 兌換獎勵資料新增
        [HttpPost]
        public void AddRedeemPrize(int AnnualId = -1, string PrizeName = "", int RedeemPoint = -1)
        {
            try
            {
                WebApiLoginCheck("PointsReward", "add");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.AddRedeemPrize(AnnualId, PrizeName, RedeemPoint);
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

        #region //AddPoint 人員積分資料新增
        [HttpPost]
        public void AddPoint(int AnnualId = -1, string UserId = "", string PointType = "", int RedeemPrizeId = -1, int PointQty = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("PersonalPoints", "add");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.AddPoint(AnnualId, UserId, PointType, RedeemPrizeId, PointQty, Remark);
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
        #region //UpdateRedeemPrize 兌換獎勵資料更新
        [HttpPost]
        public void UpdateRedeemPrize(int RedeemPrizeId = -1, int AnnualId = -1, string PrizeName = "", int RedeemPoint = -1)
        {
            try
            {
                WebApiLoginCheck("PointsReward", "update");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.UpdateRedeemPrize(RedeemPrizeId, AnnualId, PrizeName, RedeemPoint);
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

        #region //UpdateRedeemPrizeStatus 兌換獎勵狀態更新
        [HttpPost]
        public void UpdateRedeemPrizeStatus(int RedeemPrizeId = -1)
        {
            try
            {
                WebApiLoginCheck("PointsReward", "status-switch");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.UpdateRedeemPrizeStatus(RedeemPrizeId);
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
        #region //DeleteRedeemPrize 刪除兌換獎勵資料
        [HttpPost]
        public void DeleteRedeemPrize(int RedeemPrizeId = -1)
        {
            try
            {
                WebApiLoginCheck("PointsReward", "delete");

                #region //Request
                pointDA = new PointDA();
                dataRequest = pointDA.DeleteRedeemPrize(RedeemPrizeId);
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