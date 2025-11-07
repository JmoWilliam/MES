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
    public class GamingController : WebController
    {
        private GamingDA gamingDA = new GamingDA();

        #region //View
        public ActionResult BingoManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult Bingo()
        {
            return View();
        }

        public ActionResult BingoVerification()
        {
            ViewLoginCheck("BingoRound", "lottery");

            return View();
        }

        public ActionResult BingoRound()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BingoLottery()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetBingo 取得賓果資料
        [HttpPost]
        public void GetBingo(int BingoId = -1, string BingoNo = "", string Status = "", string WinStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BingoManagement", "read");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.GetBingo(BingoId, BingoNo, Status, WinStatus
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

        #region //GetBingoMap 取得賓果分佈圖資料
        [HttpPost]
        public void GetBingoMap(int BingoId = -1, string BingoNo = "")
        {
            try
            {
                #region //Request
                dataRequest = gamingDA.GetBingoMap(BingoId, BingoNo);
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

        #region //GetBingoRound 取得賓果開局資料
        [HttpPost]
        public void GetBingoRound(int RoundId = -1, string RoundName = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BingoRound", "read");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.GetBingoRound(RoundId, RoundName, StartDate, EndDate, Status
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

        #region //GetBingoLotteryRecord 取得賓果開獎記錄資料
        [HttpPost]
        public void GetBingoLotteryRecord(int RoundId = -1)
        {
            try
            {
                WebApiLoginCheck("BingoRound", "lottery-record");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.GetBingoLotteryRecord(RoundId);
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
        #region //AddBingo 賓果資料新增
        [HttpPost]
        public void AddBingo(string BingoMap = "")
        {
            try
            {
                WebApiLoginCheck("BingoManagement", "add");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.AddBingo(BingoMap);
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

        #region //AddBingoRound 賓果開局資料新增
        [HttpPost]
        public void AddBingoRound(string RoundName = "", string RoundDate = "")
        {
            try
            {
                WebApiLoginCheck("BingoRound", "add");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.AddBingoRound(RoundName, RoundDate);
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

        #region //AddBingoLotteryRecord 賓果開獎紀錄新增
        [HttpPost]
        public void AddBingoLotteryRecord(int RoundId = -1, string RecordContent = "")
        {
            try
            {
                WebApiLoginCheck("BingoRound", "lottery");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.AddBingoLotteryRecord(RoundId, RecordContent);
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
        #region //UpdateBingoStatus 賓果狀態更新
        [HttpPost]
        public void UpdateBingoStatus(int BingoId = -1)
        {
            try
            {
                WebApiLoginCheck("BingoManagement", "status-switch");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.UpdateBingoStatus(BingoId);
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

        #region //UpdateBingoWinStatus 賓果中獎狀態更新
        [HttpPost]
        public void UpdateBingoWinStatus(int RoundId = -1, string BingoNo = "")
        {
            try
            {
                WebApiLoginCheck("BingoManagement", "status-switch");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.UpdateBingoWinStatus(RoundId, BingoNo);
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

        #region //UpdateBingoRound 賓果開局資料更新
        [HttpPost]
        public void UpdateBingoRound(int RoundId = -1, string RoundName = "", string RoundDate = "")
        {
            try
            {
                WebApiLoginCheck("BingoRound", "update");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.UpdateBingoRound(RoundId, RoundName, RoundDate);
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

        #region //UpdateBingoRoundStatus 賓果開局狀態更新
        [HttpPost]
        public void UpdateBingoRoundStatus(int RoundId = -1)
        {
            try
            {
                WebApiLoginCheck("BingoRound", "status-switch");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.UpdateBingoRoundStatus(RoundId);
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

        #region //UpdateBingoRoundReset 賓果開局重置
        [HttpPost]
        public void UpdateBingoRoundReset(int RoundId = -1)
        {
            try
            {
                WebApiLoginCheck("BingoRound", "lottery-reset");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.UpdateBingoRoundReset(RoundId);
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
        #region //DeleteBingo 賓果資料刪除
        [HttpPost]
        public void DeleteBingo(int BingoId = -1)
        {
            try
            {
                WebApiLoginCheck("BingoManagement", "delete");

                #region //Request
                gamingDA = new GamingDA();
                dataRequest = gamingDA.DeleteBingo(BingoId);
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