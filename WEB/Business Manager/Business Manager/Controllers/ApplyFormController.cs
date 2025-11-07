using ClosedXML.Excel;
using EBPDA;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using iTextSharp.tool.xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using Xceed.Words.NET;
using System.Dynamic;
using Newtonsoft.Json.Converters;

namespace Business_Manager.Controllers
{
    public class ApplyFormController : WebController
    {
        private ApplyFormDA applyFormDA = new ApplyFormDA();

        #region //View
        public ActionResult DinnerSubsidy()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ClubBudgetSubsidy()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DinerList()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetSubsidyRange 取得聚餐區間資料
        [HttpPost]
        public void GetSubsidyRange(int SubsidyRangeId= -1, int AnnualId = -1, string RangeType = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = applyFormDA.GetSubsidyRange(SubsidyRangeId, AnnualId, RangeType, Status
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

        #region //GetDinnerSubsidy 取得聚餐補助資料
        [HttpPost]
        public void GetDinnerSubsidy(int DinnerSubsidyId = -1, int SubsidyRangeId = -1, int ApplyId = -1, int PayeeId = -1
            , string Status = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "read");

                #region //Request
                dataRequest = applyFormDA.GetDinnerSubsidy(DinnerSubsidyId, SubsidyRangeId, ApplyId, PayeeId
                    , Status, StartDate, EndDate
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

        #region //GetSubsidyCertificate 取得補助憑證資料
        [HttpPost]
        public void GetSubsidyCertificate(string CertificateType = "", int DinnerSubsidyId = -1, int ClubBudgetSubsidyId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                dataRequest = applyFormDA.GetSubsidyCertificate(CertificateType, DinnerSubsidyId, ClubBudgetSubsidyId
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
        public void GetParticipants(int DinnerSubsidyId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "read");

                #region //Request
                dataRequest = applyFormDA.GetParticipants(DinnerSubsidyId, SearchKey
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

        #region //GetClubBudgetSubsidy 取得社團經費補助資料
        [HttpPost]
        public void GetClubBudgetSubsidy(int ClubBudgetSubsidyId = -1, int SubsidyRangeId = -1, int ApplyId = -1, int PayeeId = -1
            , string Status = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "read");

                #region //Request
                dataRequest = applyFormDA.GetClubBudgetSubsidy(ClubBudgetSubsidyId, SubsidyRangeId, ApplyId, PayeeId
                    , Status, StartDate, EndDate
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

        #region //GetClubBudgetDetail 取得社團經費明細資料
        [HttpPost]
        public void GetClubBudgetDetail(int ClubBudgetDetailId = -1, int ClubBudgetSubsidyId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "read");

                #region //Request
                dataRequest = applyFormDA.GetClubBudgetDetail(ClubBudgetDetailId, ClubBudgetSubsidyId
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

        #region //GetClubBudgetParticipant 取得社團經費補助人員資料
        [HttpPost]
        public void GetClubBudgetParticipant(int ClubBudgetSubsidyId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "read");

                #region //Request
                dataRequest = applyFormDA.GetClubBudgetParticipant(ClubBudgetSubsidyId, SearchKey
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

        #region //GetDinerInfo 取得用餐明細-點餐資料
        [HttpPost]
        public void GetDinerInfo(int UmoId = -1, int CompanyId = -1, string CompanyName = "", int UserId = -1, string UserName = ""
            , int RestaurantId = -1, string RestaurantName = "", int MealId = -1, string MealName = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                dataRequest = applyFormDA.GetDinerInfo(UmoId, CompanyId, CompanyName, UserId, UserName
                    , RestaurantId, RestaurantName, MealId, MealName, StartDate, EndDate
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

        #region //GetDailyDinerInfo 取得用餐明細-每日訂餐明細匯出 -- (目前未使用)
        [HttpPost]
        public void GetDailyDinerInfo(int UmoId = -1, int CompanyId = -1, int UserId = -1, int RestaurantId = -1
            , string MealName = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                dataRequest = applyFormDA.GetDailyDinerInfo(UmoId, CompanyId, UserId, RestaurantId
                    , MealName, StartDate, EndDate
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

        #region //GetMonthDinerInfo 取得用餐明細-每月訂餐明細匯出 -- (目前未使用)
        [HttpPost]
        public void GetMonthDinerInfo(string Year = "", string Month = "", string Company = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                dataRequest = applyFormDA.GetMonthDinerInfo(Year, Month, Company);
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

        #region //GetPickupInfo 取得取餐區資訊 -- (目前未使用)
        [HttpPost]
        public void GetPickupInfo(int CompanyId = -1, int UserId = -1, int RestaurantId = -1
            , string MealName = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                dataRequest = applyFormDA.GetPickupInfo(CompanyId, UserId, RestaurantId
                    , MealName, StartDate, EndDate);
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

        #region //GetTotalMealList 取得餐點統計資訊 -- (目前未使用)
        [HttpPost]
        public void GetTotalMealList(string Year = "", string Month = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                dataRequest = applyFormDA.GetTotalMealList(Year, Month);
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

        #region //GetMcDetail 取得客製化項目資訊
        public void GetMcDetail(int MealId = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                applyFormDA = new ApplyFormDA();
                dataRequest = applyFormDA.GetMcDetail(MealId);
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

        #region //GetReportData 取得用餐明細-單頭報表明細匯出 -- (目前未使用)
        public void GetReportData(int CompanyId = -1, int UserId = -1, string Year = "", string Month = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                applyFormDA = new ApplyFormDA();
                dataRequest = applyFormDA.GetReportData(CompanyId, UserId, Year, Month);
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

        #region //GetDetailReportData 取得用餐明細-單身報表明細匯出
        public void GetDetailReportData(int CompanyId = -1, int UserId = -1, int RestaurantId = -1, string MealName = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "read");

                #region //Request
                applyFormDA = new ApplyFormDA();
                dataRequest = applyFormDA.GetDetailReportData(CompanyId, UserId, RestaurantId, MealName, StartDate, EndDate);
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
        #region //AddSubsidyRange 補助區間資料新增
        [HttpPost]
        public void AddSubsidyRange(int AnnualId = -1, string RangeType = "", string RangeName = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = applyFormDA.AddSubsidyRange(AnnualId, RangeType, RangeName, StartDate, EndDate);
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

        #region //AddDinnerSubsidy 聚餐補助資料新增
        [HttpPost]
        public void AddDinnerSubsidy(int SubsidyRangeId = -1, int Amount = -1, string ApplyDate = "", int ApplyId = -1, int PayeeId = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "add");

                #region //Request
                dataRequest = applyFormDA.AddDinnerSubsidy(SubsidyRangeId, Amount, ApplyDate, ApplyId, PayeeId);
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

        #region //AddSubsidyCertificate 補助憑證新增
        [HttpPost]
        public void AddSubsidyCertificate(string CertificateType = "", int FileId = -1, int DinnerSubsidyId = -1, int ClubBudgetSubsidyId = -1)
        {
            try
            {
                #region //Request
                dataRequest = applyFormDA.AddSubsidyCertificate(CertificateType, FileId, DinnerSubsidyId, ClubBudgetSubsidyId);
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

        #region //AddParticipants 聚餐補助人員資料新增
        [HttpPost]
        public void AddParticipants(int DinnerSubsidyId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "add");

                #region //Request
                dataRequest = applyFormDA.AddParticipants(DinnerSubsidyId, Users);
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

        #region //AddClubBudgetSubsidy 社團經費補助資料新增
        [HttpPost]
        public void AddClubBudgetSubsidy(int ClubId = -1, int SubsidyRangeId = -1, int Amount = -1, string ApplyDate = ""
            , int ApplyId = -1, int PayeeId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "add");

                #region //Request
                dataRequest = applyFormDA.AddClubBudgetSubsidy(ClubId, SubsidyRangeId, Amount, ApplyDate, ApplyId, PayeeId);
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

        #region //AddClubBudgetDetail 社團經費明細資料新增
        [HttpPost]
        public void AddClubBudgetDetail(int ClubBudgetSubsidyId = -1, string OccurrenceDate = "", string DetailDesc = ""
            , int Player = -1, int Amount = -1, string ActiveRegion = "")
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "add");

                #region //Request
                dataRequest = applyFormDA.AddClubBudgetDetail(ClubBudgetSubsidyId, OccurrenceDate, DetailDesc
                    , Player, Amount, ActiveRegion);
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

        #region //AddClubBudgetParticipant 社團經費補助人員資料新增
        [HttpPost]
        public void AddClubBudgetParticipant(int ClubBudgetSubsidyId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "add");

                #region //Request
                dataRequest = applyFormDA.AddClubBudgetParticipant(ClubBudgetSubsidyId, Users);
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

        #region //AddDinerInfo 團膳用餐明細新增
        [HttpPost]
        public void AddDinerInfo(int CompanyId = -1, int UserId = -1, string UmoDate = "", double UmoDiscount = -1, double UmoAmount = -1
            , string UmoDetailRemark = "", int UmoDetailQty = -1, double UmoDetailPrice = -1, string MealDetails = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "add");

                #region //Request
                dataRequest = applyFormDA.AddDinerInfo(CompanyId, UserId, UmoDate, UmoDiscount, UmoAmount
                    , UmoDetailRemark, UmoDetailQty, UmoDetailPrice, MealDetails);
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
        #region //UpdateSubsidyRange 補助區間資料更新
        [HttpPost]
        public void UpdateSubsidyRange(int SubsidyRangeId = -1, int AnnualId = -1, string RangeName = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = applyFormDA.UpdateSubsidyRange(SubsidyRangeId, AnnualId, RangeName, StartDate, EndDate);
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

        #region //UpdateSubsidyRangeStatus 補助區間狀態更新
        [HttpPost]
        public void UpdateSubsidyRangeStatus(int SubsidyRangeId = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = applyFormDA.UpdateSubsidyRangeStatus(SubsidyRangeId);
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

        #region //UpdateDinnerSubsidy 聚餐補助資料更新
        [HttpPost]
        public void UpdateDinnerSubsidy(int DinnerSubsidyId = -1, int SubsidyRangeId = -1, int Amount = -1, string ApplyDate = "", int ApplyId = -1, int PayeeId = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "update");

                #region //Request
                dataRequest = applyFormDA.UpdateDinnerSubsidy(DinnerSubsidyId, SubsidyRangeId, Amount, ApplyDate, ApplyId, PayeeId);
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

        #region //UpdateSubsidyStatus 聚餐補助狀態更新
        [HttpPost]
        public void UpdateSubsidyStatus(int DinnerSubsidyId = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "check");

                #region //Request
                dataRequest = applyFormDA.UpdateSubsidyStatus(DinnerSubsidyId);
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

        #region //UpdateSubsidyCertificateRename 補助憑證更名
        [HttpPost]
        public void UpdateSubsidyCertificateRename(int CertificateId = -1, string NewFileName = "")
        {
            try
            {
                //WebApiLoginCheck("DinnerSubsidy", "file-rename");

                #region //Request
                dataRequest = applyFormDA.UpdateSubsidyCertificateRename(CertificateId, NewFileName);
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

        #region //UpdateClubBudgetSubsidy 社團經費補助資料更新
        [HttpPost]
        public void UpdateClubBudgetSubsidy(int ClubBudgetSubsidyId = -1, int ClubId = -1, int SubsidyRangeId = -1, int Amount = -1
            , string ApplyDate = "", int ApplyId = -1, int PayeeId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "update");

                #region //Request
                dataRequest = applyFormDA.UpdateClubBudgetSubsidy(ClubBudgetSubsidyId, ClubId, SubsidyRangeId, Amount
                    , ApplyDate, ApplyId, PayeeId);
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

        #region //UpdateClubBudgetSubsidyStatus 社團經費補助狀態更新
        [HttpPost]
        public void UpdateClubBudgetSubsidyStatus(int ClubBudgetSubsidyId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "check");

                #region //Request
                dataRequest = applyFormDA.UpdateClubBudgetSubsidyStatus(ClubBudgetSubsidyId);
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

        #region //UpdateClubBudgetDetail 社團經費明細資料更新
        [HttpPost]
        public void UpdateClubBudgetDetail(int ClubBudgetDetailId = -1, string OccurrenceDate = "", string DetailDesc = ""
            , int Player = -1, int Amount = -1, string ActiveRegion = "")
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "update");

                #region //Request
                dataRequest = applyFormDA.UpdateClubBudgetDetail(ClubBudgetDetailId, OccurrenceDate, DetailDesc
                    , Player, Amount, ActiveRegion);
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

        #region //UpdateDinerInfo 團膳用餐明細更新
        [HttpPost]
        public void UpdateDinerInfo(int UmoId = -1, int CompanyId = -1, int UserId = -1, string UmoDate = "", double UmoDiscount = -1, string MealDetails = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "update");

                #region //Request
                dataRequest = applyFormDA.UpdateDinerInfo(UmoId, CompanyId, UserId, UmoDate, UmoDiscount, MealDetails);
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
        #region //DeleteSubsidyRange 補助區間資料刪除
        [HttpPost]
        public void DeleteSubsidyRange(int SubsidyRangeId = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = applyFormDA.DeleteSubsidyRange(SubsidyRangeId);
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

        #region //DeleteDinnerSubsidy 聚餐補助資料刪除
        [HttpPost]
        public void DeleteDinnerSubsidy(int DinnerSubsidyId = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteDinnerSubsidy(DinnerSubsidyId);
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

        #region //DeleteSubsidyCertificate 聚餐補助憑證刪除
        [HttpPost]
        public void DeleteSubsidyCertificate(int CertificateId = -1, string CertificateType = "")
        {
            try
            {
                #region //Request
                dataRequest = applyFormDA.DeleteSubsidyCertificate(CertificateId, CertificateType);
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

        #region //DeleteParticipants 聚餐補助人員資料刪除
        [HttpPost]
        public void DeleteParticipants(int ParticipantId = -1)
        {
            try
            {
                WebApiLoginCheck("DinnerSubsidy", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteParticipants(ParticipantId);
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

        #region //DeleteClubBudgetSubsidy 社團經費補助資料刪除
        [HttpPost]
        public void DeleteClubBudgetSubsidy(int ClubBudgetSubsidyId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteClubBudgetSubsidy(ClubBudgetSubsidyId);
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

        #region //DeleteClubBudgetDetail 社團經費明細資料刪除
        [HttpPost]
        public void DeleteClubBudgetDetail(int ClubBudgetDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteClubBudgetDetail(ClubBudgetDetailId);
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

        #region //DeleteClubBudgetParticipant 社團經費補助人員資料刪除
        [HttpPost]
        public void DeleteClubBudgetParticipant(int ParticipantId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubBudgetSubsidy", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteClubBudgetParticipant(ParticipantId);
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

        #region //DeleteDinerInfo 用餐明細-點餐資料刪除
        [HttpPost]
        public void DeleteDinerInfo(int UmoId = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "delete");

                #region //Request
                dataRequest = applyFormDA.DeleteDinerInfo(UmoId);
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

        #region //Download
        #region //Word
        #region //ClubBudgetSubsidyWordDownload 社團經費申請表下載
        public void ClubBudgetSubsidyWordDownload(int ClubBudgetSubsidyId = -1)
        {
            try
            {
                #region //產生Word
                #region //Request
                dataRequest = applyFormDA.GetClubBudgetSubsidyWord(ClubBudgetSubsidyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                byte[] wordFile;
                string wordFileName = "{0}社團經費申請表-{1}";

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("經費明細不完整!");

                    using (DocX doc = DocX.Load(Server.MapPath("~/WordTemplate/EBP/社團經費申請表.docx")))
                    {
                        doc.ReplaceText("[ClubName]", result[0]["ClubName"].ToString());
                        doc.ReplaceText("[ApplyDate]", result[0]["ApplyDate"].ToString());
                        doc.ReplaceText("[ApplyUser]", result[0]["ApplyUser"].ToString());
                        doc.ReplaceText("[PayeeUser]", result[0]["PayeeUser"].ToString());

                        wordFileName = string.Format(wordFileName, Convert.ToDateTime(result[0]["ApplyDate"]).ToString("[yyMMdd]"), result[0]["ClubName"].ToString());
                        
                        if (result[0]["ClubBudgetDetail"].ToString().Length > 0)
                        {
                            var detail = JObject.Parse(result[0]["ClubBudgetDetail"].ToString())["data"];

                            for (int i = 0; i < detail.Count(); i++)
                            {
                                doc.ReplaceText(string.Format("[Date{0:00}]", i + 1), detail[i]["OccurrenceDate"].ToString());
                                doc.ReplaceText(string.Format("[Desc{0:00}]", i + 1), detail[i]["DetailDesc"].ToString());
                                doc.ReplaceText(string.Format("[Player{0:00}]", i + 1), detail[i]["Player"].ToString());
                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), detail[i]["Amount"].ToString());
                                doc.ReplaceText(string.Format("[ActiveRegion{0:00}]", i + 1), detail[i]["ActiveRegion"].ToString());
                            }

                            if (detail.Count() < 20)
                            {
                                for (int i = detail.Count(); i < 20; i++)
                                {
                                    doc.ReplaceText(string.Format("[Date{0:00}]", i + 1), "");
                                    doc.ReplaceText(string.Format("[Desc{0:00}]", i + 1), "");
                                    doc.ReplaceText(string.Format("[Player{0:00}]", i + 1), "");
                                    doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                    doc.ReplaceText(string.Format("[ActiveRegion{0:00}]", i + 1), "");
                                }
                            }
                        }

                        using (MemoryStream output = new MemoryStream())
                        {
                            doc.SaveAs(output);

                            wordFile = output.ToArray();
                        }
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = ".docx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "目前尚無資料"
                    });
                    #endregion
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
        #endregion

        #region //Excel
        #region //DailyDinerExcelDownload 每日訂餐明細
        public void DailyDinerExcelDownload(int UmoId = -1, int CompanyId = -1, string CompanyName = "", int UserId = -1, string UserName = ""
            , int RestaurantId = -1, string RestaurantName = "", int MealId = -1, string MealName = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "excel");

                #region //Request
                dataRequest = applyFormDA.GetDinerInfo(-1, CompanyId, CompanyName, UserId, UserName
                    , RestaurantId, RestaurantName, MealId, MealName, StartDate, EndDate
                    , "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "_-* #,##0.00_-;-* #,##0.00_-;_-* \" - \"??_-;_-@_-";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0";
                    #endregion
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "每日點餐明細Excel檔";
                    string excelsheetName = "每日點餐明細";

                    dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    string[] header = new string[] { "訂單編號", "公司", "部門", "員工", "餐點日期", "訂購日期", "總金額", "餐廳", "餐點", "份數", "價格", "備註", "客製化項目($)", "取餐地點" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 7)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 14).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //每日點餐明細分頁
                        int startIndex = 0;
                        foreach (var item in json.data)
                        {
                            startIndex = rowIndex + 1;

                            if (item.UmoDetail != null)
                            {
                                dynamic umoDetail = JsonConvert.DeserializeObject<ExpandoObject>(item.UmoDetail, new ExpandoObjectConverter());

                                foreach (var detail in umoDetail.data)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.UmoNo.ToString(); //訂單編號
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CompanyName.ToString(); //公司
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DepartmentName.ToString(); //部門
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.UserInfo.ToString(); //員工
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.UmoDate.ToString(); //餐點日期
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.CreateDate.ToString(); //訂購日期
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = (item.UmoAmount- item.UmoDiscount); //總金額
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Style = currencyStyle;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = detail.RestaurantName.ToString(); //餐廳
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = detail.MealName.ToString(); //餐點
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = detail.UmoDetailQty; //份數
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Style = currencyStyle;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = detail.UmoDetailPrice; //價格
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style = currencyStyle;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = detail.UmoDetailRemark.ToString(); //備註
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = detail.UmoAdditional.ToString(); //客製化項目($)
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = detail.PickupName.ToString(); //取餐地點
                                }
                            }
                            else
                            {
                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.UmoNo.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CompanyName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DepartmentName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.UserNo.ToString() + " " + item.UserName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.UmoDate.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.CreateDate.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = "";
                            }
                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            worksheet.Range(startIndex, 6, rowIndex, 6).Merge();
                            worksheet.Range(startIndex, 7, rowIndex, 7).Merge();
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //置中
                        //worksheet.Columns("A:C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");
                        //worksheet.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        //worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        //worksheet.Cell(4, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(3, 1, 8, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        #endregion

                        #region //合併儲存格
                        //worksheet.Range(8, 1, 8, 9).Merge();
                        #endregion

                        #region //公式
                        ////R=Row、C=Column
                        ////[]指相對位置，[-]向前，[+]向後
                        ////未[]指絕對位置
                        //worksheet.Cell(10, 3).FormulaR1C1 = "RC[-2]+R3C2";
                        //worksheet.Cell(11, 3).FormulaR1C1 = "RC[-2]+RC[2]";
                        //worksheet.Cell(12, 3).FormulaR1C1 = "RC[-2]+RC2";
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #region //隱藏
                        //worksheet.Column(1).Hide();
                        #endregion

                        #region //格式化
                        //string conditionalFormatRange = "G:G";
                        //string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        //{
                        //    range.AddConditionalFormat().WhenEquals("M")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.BabyBlue);

                        //    range.AddConditionalFormat().WhenEquals("F")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.Red);
                        //}

                        //string conditionalFormatRange2 = "E:E";
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange2))
                        //{
                        //    range.AddConditionalFormat().WhenContains("1")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.RedNcs);
                        //}
                        #endregion

                        #region //複製
                        //worksheet.CopyTo("複製");
                        #endregion

                        #endregion

                        #region //表格
                        //var fakeData = Enumerable.Range(1, 5)
                        //.Select(x => new FakeData
                        //{
                        //    Time = TimeSpan.FromSeconds(x * 123.667),
                        //    X = x,
                        //    Y = -x,
                        //    Address = "a" + x,
                        //    Distance = x * 100
                        //}).ToArray();

                        //var table = worksheet.Cell(10, 1).InsertTable(fakeData);
                        //table.Style.Font.FontSize = 9;
                        //var tableData = worksheet.Cell(17, 1).InsertData(fakeData);
                        //tableData.Style.Font.FontSize = 9;
                        //worksheet.Range(11, 1, 21, 1).Style.DateFormat.Format = "HH:mm:ss.000";
                        #endregion

                        #region //取餐區分頁
                        dataRequest = applyFormDA.GetDetailReportData(CompanyId, UserId, RestaurantId, MealName, StartDate, EndDate);

                        if (jsonResponse["status"].ToString() == "success")
                        {
                            List<DetailReportData> detailReportDatas = JsonConvert.DeserializeObject<List<DetailReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                            var places = detailReportDatas.Select(x => x.Pickup).Distinct().ToList();

                            foreach (var place in places)
                            {
                                string[] placeheader = new string[] { "序號", "工號", "姓名", "部門", "餐點", "份數", "領取便當" };
                                colIndex = "";
                                rowIndex = 1;

                                worksheet = workbook.Worksheets.Add(place.ToString());
                                worksheet.RowHeight = 15;
                                worksheet.Style = defaultStyle;

                                #region //HEADER
                                for (int i = 0; i < placeheader.Length; i++)
                                {
                                    colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                    worksheet.Cell(colIndex).Value = placeheader[i];
                                    worksheet.Cell(colIndex).Style = headerStyle;
                                }
                                #endregion

                                #region //BODY
                                var mealData = detailReportDatas
                                                    .Where(x => x.Pickup == place.ToString())
                                                    .Select(x => x)
                                                    .OrderBy(x => x.UserNo)
                                                    .ToList();

                                var placeMeal = detailReportDatas
                                                    .Where(x => x.Pickup == place.ToString())
                                                    .Select(x => x.MealName)
                                                    .Distinct()
                                                    .ToList();

                                int index = 1;
                                foreach (var meal in mealData)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = index; //序號
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = meal.UserNo; //工號
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = meal.UserName; //姓名
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = meal.DepartmentName; //部門
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = meal.MealName; //餐點
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = meal.UmoDetailQty; //份數
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = ""; //領取便當

                                    index++;
                                }

                                worksheet.Cell(1, 7).Style.Fill.BackgroundColor = XLColor.BabyPink; //領取便當 設定底色

                                #region //用餐樓層
                                int pickupRow = 3, pickupCol = 9;

                                worksheet.Cell(BaseHelper.MergeNumberToChar(pickupCol, pickupRow)).Value = "用餐樓層";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(pickupCol, pickupRow)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 設定水平對齊為左對齊
                                worksheet.Cell(BaseHelper.MergeNumberToChar(pickupCol + 1, pickupRow)).Value = place.ToString();

                                #region //設定底色
                                string conditionalFormatRange = "I3:J3";
                                foreach (var range in worksheet.Ranges(conditionalFormatRange))
                                {
                                    range.AddConditionalFormat().WhenNotEquals("")
                                        .Fill.SetBackgroundColor(XLColor.BabyBlue);
                                }
                                #endregion

                                #endregion

                                #region //合計便當數量
                                int mealRow = 5, mealCol = 9;
                                double totalMeal = 0;
                                foreach (var mealInfo in placeMeal)
                                {
                                    double countMeal = detailReportDatas
                                                        .Where(x => x.Pickup == place.ToString() && x.MealName == mealInfo)
                                                        .Sum(x => x.UmoDetailQty);

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol, mealRow)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 設定水平對齊為左對齊
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol, mealRow)).Value = mealInfo;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol + 1, mealRow)).Value = countMeal.ToString();
                                    mealRow++;
                                    totalMeal = totalMeal + countMeal;
                                }

                                worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol, mealRow)).Value = "合計便當數量";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol, mealRow)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 設定水平對齊為左對齊
                                worksheet.Cell(BaseHelper.MergeNumberToChar(mealCol + 1, mealRow)).Value = totalMeal.ToString();
                                #endregion

                                #endregion

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion
                            }
                        }
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
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
        #endregion

        #region //MonthDinerExcelDownload 每月訂餐總表
        public void MonthDinerExcelDownload(string Year = "", string Month = "", int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "excel");

                #region //Request
                dataRequest = applyFormDA.GetReportData(CompanyId, -1, Year, Month);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //discountStyle
                    var discountStyle = XLWorkbook.DefaultStyle;
                    discountStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    discountStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    discountStyle.Border.TopBorder = XLBorderStyleValues.None;
                    discountStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    discountStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    discountStyle.Border.RightBorder = XLBorderStyleValues.None;
                    discountStyle.Border.TopBorderColor = XLColor.NoColor;
                    discountStyle.Border.BottomBorderColor = XLColor.NoColor;
                    discountStyle.Border.LeftBorderColor = XLColor.NoColor;
                    discountStyle.Border.RightBorderColor = XLColor.NoColor;
                    discountStyle.Fill.BackgroundColor = XLColor.LemonChiffon;
                    discountStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    discountStyle.Font.FontSize = 12;
                    discountStyle.Font.Bold = false;
                    //discountStyle.Font.FontColor = (XLColor.Red);
                    discountStyle.Protection.SetLocked(false);
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //headerStyleDate
                    var headerStyleDate = XLWorkbook.DefaultStyle;
                    headerStyleDate.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyleDate.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyleDate.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.TopBorderColor = XLColor.Black;
                    headerStyleDate.Border.BottomBorderColor = XLColor.Black;
                    headerStyleDate.Border.LeftBorderColor = XLColor.Black;
                    headerStyleDate.Border.RightBorderColor = XLColor.Black;
                    headerStyleDate.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyleDate.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyleDate.Font.FontSize = 14;
                    headerStyleDate.Font.Bold = true;
                    headerStyleDate.DateFormat.Format = "MM/dd";
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0";
                    #endregion

                    #region //totalQtyStyle
                    var totalQtyStyle = XLWorkbook.DefaultStyle;
                    totalQtyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    totalQtyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    totalQtyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.TopBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.RightBorderColor = XLColor.NoColor;
                    totalQtyStyle.Fill.BackgroundColor = XLColor.PowderBlue;
                    totalQtyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    totalQtyStyle.Font.FontSize = 12;
                    totalQtyStyle.Font.Bold = false;
                    totalQtyStyle.NumberFormat.Format = "#,##0";
                    //totalQtyStyle.Font.FontColor = (XLColor.Red);
                    totalQtyStyle.Protection.SetLocked(false);
                    #endregion

                    #region //totalPriceStyle
                    var totalPriceStyle = XLWorkbook.DefaultStyle;
                    totalPriceStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    totalPriceStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    totalPriceStyle.Border.TopBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.RightBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.TopBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.BottomBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.LeftBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.RightBorderColor = XLColor.NoColor;
                    totalPriceStyle.Fill.BackgroundColor = XLColor.PowderBlue;
                    totalPriceStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    totalPriceStyle.Font.FontSize = 12;
                    totalPriceStyle.Font.Bold = false;
                    totalPriceStyle.NumberFormat.Format = "#,##0";
                    //totalPriceStyle.Font.FontColor = (XLColor.Red);
                    totalPriceStyle.Protection.SetLocked(false);
                    #endregion
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "每月訂餐總表Excel檔";
                    string excelsheetName = "每月訂餐總表";

                    List<ReportData> reportDataForMonths = JsonConvert.DeserializeObject<List<ReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                    List<string> header = new List<string> { "工號", "中文名", "部門" };

                    DateTime startDate = Convert.ToDateTime(Year + "/" + Month + "/01");
                    DateTime endDate = Convert.ToDateTime(Year + "/" + Month + "/01").AddMonths(1).AddDays(-1);

                    int day = Convert.ToInt32(endDate.Subtract(startDate).TotalDays); //日期格式的相減

                    //判斷日期所屬位置
                    for (var i = 0; i <= day; i++)
                    {
                        header.Add(startDate.AddDays(i).ToString("MM/dd"));
                    }

                    header.Add("用餐份數");
                    header.Add("總額");
                    header.Add("應付總額");
                    header.Add("補助款");
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //每月訂餐總表分頁

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 15)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, header.Count()).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Count(); i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];

                            if (i > 2 && i != (header.Count() - 1))
                            {
                                worksheet.Cell(colIndex).Style = headerStyleDate;
                            }
                            else
                            {
                                worksheet.Cell(colIndex).Style = headerStyle;
                            }
                        }
                        #endregion

                        #region //BODY
                        var userList = reportDataForMonths.Select(x => new { x.UserNo, x.UserName, x.DepartmentName }).Distinct().OrderBy(x => x.UserNo);

                        int startIndex = 0;
                        foreach (var user in userList)
                        {
                            startIndex = rowIndex + 1;

                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = user.UserNo; //工號
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = user.UserName; //中文名
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = user.DepartmentName; //部門
                            for (var i = 0; i <= day; i++)
                            {
                                string currentDate = startDate.AddDays(i).ToString("MM-dd");

                                double price = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserNo == user.UserNo).Sum(x => x.UmoAmount);
                                double discount = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserNo == user.UserNo).Sum(x => x.UmoDiscount);

                                if (price != 0)
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, rowIndex)).Value = (price - discount); //各日期之用餐金額
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, rowIndex)).Style = currencyStyle;

                                    if (discount > 0)
                                    {
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, rowIndex)).Style = discountStyle;
                                    }
                                    else
                                    {
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, rowIndex)).Style = defaultStyle;
                                    }
                                }
                            }
                            double qty = reportDataForMonths.Where(x => x.UserNo == user.UserNo).Sum(x => x.UmoDetailQty);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, rowIndex)).Value = qty; //用餐份數
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, rowIndex)).Style = currencyStyle;

                            double totalPrice = reportDataForMonths.Where(x => x.UserNo == user.UserNo).Sum(x => x.UmoAmount);

                            double totalDiscount = reportDataForMonths.Where(x => x.UserNo == user.UserNo).Sum(x => x.UmoDiscount);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, rowIndex)).Value = (totalPrice - totalDiscount); //總額
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, rowIndex)).Style = currencyStyle;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, rowIndex)).Value = totalPrice; //應付總額
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, rowIndex)).Style = currencyStyle;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), rowIndex)).Value = totalDiscount; //補助款
                            worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), rowIndex)).Style = currencyStyle;
                        }

                        #region //員工用餐數
                        int totalQtyRowIndex = rowIndex + 1;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex)).Value = "員工用餐數";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex)).Style = totalQtyStyle;

                        #region //合併儲存格
                        worksheet.Range(totalQtyRowIndex, 1, totalQtyRowIndex, 3).Merge();
                        //1, first row (0-based)
                        //1, first column  (0-based)
                        //1, last row (0-based)
                        //2  last column  (0-based)
                        #endregion

                        for (var i = 0; i <= day; i++)
                        {
                            string currentDate = startDate.AddDays(i).ToString("MM-dd");

                            double totalStaffQty = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserName != "來賓").Sum(x => x.UmoDetailQty);

                            string totalQtyCell = BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex)).Value = totalStaffQty; //每日員工用餐數
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex)).Style = totalQtyStyle;
                        }

                        double totalStaffAmountQty = reportDataForMonths.Where(x => x.UserName != "來賓").Sum(x => x.UmoDetailQty);
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex)).Value = totalStaffAmountQty; //員工用餐份數總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex)).Style = totalQtyStyle;

                        double totalStaffAmountPrice = reportDataForMonths.Where(x => x.UserName != "來賓").Sum(x => x.UmoAmount);
                        double totalStaffAmountDiscount = reportDataForMonths.Where(x => x.UserName != "來賓").Sum(x => x.UmoDiscount);
                        double totalStaffAmount = totalStaffAmountPrice - totalStaffAmountDiscount;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex)).Value = totalStaffAmount; //員工總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex)).Value = totalStaffAmountPrice; //員工應付總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex)).Value = totalStaffAmountDiscount; //員工補助款總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex)).Style = totalQtyStyle;
                        #endregion

                        #region //來賓用餐數
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex + 1)).Value = "來賓用餐數";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex + 1)).Style = totalQtyStyle;

                        #region //合併儲存格
                        worksheet.Range(totalQtyRowIndex + 1, 1, totalQtyRowIndex + 1, 3).Merge();
                        //1, first row (0-based)
                        //1, first column  (0-based)
                        //1, last row (0-based)
                        //2  last column  (0-based)
                        #endregion

                        for (var i = 0; i <= day; i++)
                        {
                            string currentDate = startDate.AddDays(i).ToString("MM-dd");

                            double totalGuestQty = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserName == "來賓").Sum(x => x.UmoDetailQty);

                            string totalQtyCell = BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex + 1)).Value = totalGuestQty; //每日來賓用餐數
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex + 1)).Style = totalQtyStyle;
                        }

                        double totalGuestAmountQty = reportDataForMonths.Where(x => x.UserName == "來賓").Sum(x => x.UmoDetailQty);
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex + 1)).Value = totalGuestAmountQty; //來賓用餐數總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex + 1)).Style = totalQtyStyle;

                        double totalGuestAmountPrice = reportDataForMonths.Where(x => x.UserName == "來賓").Sum(x => x.UmoAmount);
                        double totalGuestAmountDiscount = reportDataForMonths.Where(x => x.UserName == "來賓").Sum(x => x.UmoDiscount);
                        double totalGuestAmount = totalGuestAmountPrice - totalGuestAmountDiscount;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex + 1)).Value = totalGuestAmount; //來賓總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex + 1)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex + 1)).Value = totalGuestAmountPrice; //來賓應付總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex + 1)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex + 1)).Value = totalGuestAmountDiscount; //來賓補助款總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex + 1)).Style = totalQtyStyle;
                        #endregion

                        #region //總計份數
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex + 2)).Value = "總計份數";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex + 2)).Style = totalQtyStyle;

                        #region //合併儲存格
                        worksheet.Range(totalQtyRowIndex + 2, 1, totalQtyRowIndex + 2, 3).Merge();
                        //1, first row (0-based)
                        //1, first column  (0-based)
                        //1, last row (0-based)
                        //2  last column  (0-based)
                        #endregion

                        for (var i = 0; i <= day; i++)
                        {
                            string currentDate = startDate.AddDays(i).ToString("MM-dd");

                            double totalStaffQty = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserName != "來賓").Sum(x => x.UmoDetailQty);
                            double totalGuestQty = reportDataForMonths.Where(x => x.UmoDate == currentDate && x.UserName == "來賓").Sum(x => x.UmoDetailQty);

                            string totalQtyCell = BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex + 2)).Value = (totalStaffQty + totalGuestQty); //每日用餐數
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4 + i, totalQtyRowIndex + 2)).Style = totalQtyStyle;
                        }

                        double StaffAmountQty = reportDataForMonths.Where(x => x.UserName != "來賓").Sum(x => x.UmoDetailQty);
                        double GuestAmountQty = reportDataForMonths.Where(x => x.UserName == "來賓").Sum(x => x.UmoDetailQty);
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex + 2)).Value = (StaffAmountQty + GuestAmountQty); //用餐數總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 3, totalQtyRowIndex + 2)).Style = totalQtyStyle;

                        double totalAmountPrice = totalStaffAmount + totalGuestAmount;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex + 2)).Value = totalAmountPrice; //總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 2, totalQtyRowIndex + 2)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex + 2)).Value = totalStaffAmountPrice + totalGuestAmountPrice; //應付總額總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count() - 1, totalQtyRowIndex + 2)).Style = totalQtyStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex + 2)).Value = totalStaffAmountDiscount + totalGuestAmountDiscount; //補助款總計
                        worksheet.Cell(BaseHelper.MergeNumberToChar(header.Count(), totalQtyRowIndex + 2)).Style = totalQtyStyle;
                        #endregion

                        #region //找到需要隱藏的列
                        List<int> rowsToHide = new List<int>();
                        for (int i = 4; i < header.Count(); i++) //第四列(D3位置算起)
                        {
                            bool hasData = false;
                            for (var j = 2; j <= rowIndex; j++) //第三欄(D3位置算起)
                            {
                                if (int.TryParse(worksheet.Cell(BaseHelper.MergeNumberToChar(i, j)).Value.ToString(), out int tempInt))
                                {
                                    hasData = true;
                                    break;
                                }
                            }

                            if (!hasData)
                            {
                                rowsToHide.Add(i);
                            }
                        }

                        // 隱藏列
                        foreach (var row in rowsToHide)
                        {
                            worksheet.Column(row).Hide(); // 隱藏該列
                        }
                        #endregion

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #endregion

                        #endregion

                        #region //各部門費用分頁
                        dataRequest = applyFormDA.GetReportData(CompanyId, -1, Year, Month);

                        if (jsonResponse["status"].ToString() == "success")
                        {
                            List<ReportData> reportDatas = JsonConvert.DeserializeObject<List<ReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                            var segments = reportDatas.Select(x => new { x.DepartmentNo, x.DepartmentName, x.DepartmentCategory }).Distinct().OrderBy(x => x.DepartmentNo);

                            #region //員工用餐

                            #region //參數初始化
                            string sheetName = "各部門費用";
                            string[] staffheader = new string[] { "分類", "部門代號", "部門", "應付總額", "員工應付", "員工補助款" };

                            colIndex = "";
                            rowIndex = 1;

                            worksheet = workbook.Worksheets.Add(sheetName);
                            worksheet.RowHeight = 15;
                            worksheet.Style = defaultStyle;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = "員工用餐"; //Title
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Style = headerStyle;

                            #region //合併儲存格
                            worksheet.Range(rowIndex, 1, rowIndex, 6).Merge();
                            #endregion

                            #endregion

                            #region //HEADER
                            rowIndex++;
                            for (int i = 0; i < staffheader.Length; i++)
                            {
                                colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                worksheet.Cell(colIndex).Value = staffheader[i];
                                worksheet.Cell(colIndex).Style = headerStyle;
                            }
                            #endregion

                            #region //BODY

                            int index = 1;
                            foreach (var item in segments)
                            {
                                double segmentStaffAmount = reportDatas.Where(x => x.DepartmentNo == item.DepartmentNo && x.UserName != "來賓").Sum(x => x.UmoAmount);
                                double segmentStaffDiscount = reportDatas.Where(x => x.DepartmentNo == item.DepartmentNo && x.UserName != "來賓").Sum(x => x.UmoDiscount);
                                double segmentStaffAmountPrice = segmentStaffAmount - segmentStaffDiscount;

                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DepartmentCategory; //分類
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DepartmentNo; //部門代號
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DepartmentName; //部門名稱

                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = segmentStaffAmount; //應付總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Style = numberStyle;

                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = segmentStaffAmountPrice; //員工應付
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Style = numberStyle;

                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = segmentStaffDiscount; //員工補助款
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Style = numberStyle;

                                index++;

                                #region //合計金額
                                int totalRowIndex = rowIndex + 1;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, totalRowIndex)).Value = "合計";

                                string totalPriceCell = BaseHelper.MergeNumberToChar(4, totalRowIndex);
                                var sumStaffAmount = worksheet.Evaluate("SUM(D:D)");
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, totalRowIndex)).Value = sumStaffAmount; //應付總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, totalRowIndex)).Style = totalPriceStyle;

                                var sumStaffAmountPrice = worksheet.Evaluate("SUM(E:E)");
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, totalRowIndex)).Value = sumStaffAmountPrice; //員工應付
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, totalRowIndex)).Style = totalPriceStyle;

                                var sumStaffDiscount = worksheet.Evaluate("SUM(F:F)");
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, totalRowIndex)).Value = sumStaffDiscount; //員工補助款
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, totalRowIndex)).Style = totalPriceStyle;
                                #endregion
                            }
                            #endregion

                            #endregion

                            #region //來賓用餐

                            #region //參數初始化
                            string[] guestheader = new string[] { "部門", "應付總額" };

                            colIndex = "";
                            rowIndex = 1;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = "來賓用餐"; //Title
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Style = headerStyle;

                            #region //合併儲存格
                            worksheet.Range(rowIndex, 8, rowIndex, 9).Merge();
                            #endregion

                            #endregion

                            #region //HEADER
                            rowIndex++;
                            for (int i = 0; i < guestheader.Length; i++)
                            {
                                colIndex = BaseHelper.MergeNumberToChar(i + 8, rowIndex);
                                worksheet.Cell(colIndex).Value = guestheader[i];
                                worksheet.Cell(colIndex).Style = headerStyle;
                            }
                            #endregion

                            #region //BODY
                            foreach (var item in segments)
                            {
                                double segmentGuestAmount = reportDatas.Where(x => x.DepartmentNo == item.DepartmentNo && x.UserName == "來賓" && x.UmoAmount > 0).Sum(x => x.UmoAmount);
                                double segmentGuestDiscount = reportDatas.Where(x => x.DepartmentNo == item.DepartmentNo && x.UserName == "來賓").Sum(x => x.UmoDiscount);
                                double segmentGuestAmountPrice = segmentGuestAmount - segmentGuestDiscount;

                                //僅顯示應付總額有值的部門別
                                if (segmentGuestAmount <= 0)
                                    continue;

                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.DepartmentName; //部門名稱

                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = segmentGuestAmount; //應付總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = numberStyle;

                                index++;

                                #region //合計金額
                                int totalRowIndex = rowIndex + 1;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, totalRowIndex)).Value = "合計";

                                var sumGuestAmount = worksheet.Evaluate("SUM(I:I)");
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, totalRowIndex)).Value = sumGuestAmount; //應付總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, totalRowIndex)).Style = totalPriceStyle;
                                #endregion
                            }
                            #endregion

                            #endregion

                            #region //自適應欄寬
                            worksheet.Columns().AdjustToContents();
                            #endregion
                        }

                        #endregion

                        #region //各公司餐廳總額分頁
                        dataRequest = applyFormDA.GetDetailReportData(CompanyId, -1, -1, "", Convert.ToDateTime(Year + "-" + Month + "-01").ToString("yyyy-MM-dd"), Convert.ToDateTime(Year + "-" + Month + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"));

                        if (jsonResponse["status"].ToString() == "success")
                        {
                            List<DetailReportData> detailReportDatas = JsonConvert.DeserializeObject<List<DetailReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                            var companyRestaurantList = detailReportDatas.Select(x => new { x.CompanyName, x.RestaurantName }).Distinct().OrderBy(x => x.CompanyName);

                            #region //參數初始化
                            string sheetName = "各公司餐廳總額";
                            string[] dataheader = new string[] { "公司", "餐廳", "總計金額" };

                            colIndex = "";
                            rowIndex = 1;

                            worksheet = workbook.Worksheets.Add(sheetName);
                            worksheet.RowHeight = 15;
                            worksheet.Style = defaultStyle;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = "每月各餐廳總額"; //Title
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Style = headerStyle;

                            #region //合併儲存格
                            worksheet.Range(rowIndex, 1, rowIndex, 3).Merge();
                            #endregion

                            #endregion

                            #region //HEADER
                            rowIndex++;
                            for (int i = 0; i < dataheader.Length; i++)
                            {
                                colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                worksheet.Cell(colIndex).Value = dataheader[i];
                                worksheet.Cell(colIndex).Style = headerStyle;
                            }
                            #endregion

                            #region //BODY

                            int index = 1;
                            foreach (var item in companyRestaurantList)
                            {
                                double price = detailReportDatas.Where(x => x.CompanyName == item.CompanyName && x.RestaurantName == item.RestaurantName).Sum(x => x.UmoDetailPrice);

                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.CompanyName; //公司
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.RestaurantName; //餐廳

                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = price; //總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Style = numberStyle;

                                index++;
                            }
                            #endregion

                            #region //自適應欄寬
                            worksheet.Columns().AdjustToContents();
                            #endregion
                        }

                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
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
        #endregion

        #region TotalMealListExcelDownload 中揚+紘立統計報表
        public void TotalMealListExcelDownload(string Year = "", string Month = "")
        {
            try
            {
                WebApiLoginCheck("DinerList", "excel");

                #region //Request
                dataRequest = applyFormDA.GetDetailReportData(-1, -1, -1, "", Convert.ToDateTime(Year + "-" + Month + "-01").ToString("yyyy-MM-dd"), Convert.ToDateTime(Year + "-" + Month + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"));
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "$#,##0";
                    #endregion
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "中揚+紘立統計報表Excel檔";
                    #endregion

                    #region //Excel
                    using (var workbook = new XLWorkbook())
                    {
                        #region //分頁模式
                        if (jsonResponse["status"].ToString() == "success")
                        {
                            List<DetailReportData> detailReportDatas = JsonConvert.DeserializeObject<List<DetailReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                            var restaurants = detailReportDatas.Select(x => x.RestaurantName).Distinct().ToList();

                            foreach (var restaurant in restaurants)
                            {
                                List<string> header = new List<string> { "日期", "廠商", "訂購份數" };

                                DateTime startDate = Convert.ToDateTime(Year + "/" + Month + "/01");
                                DateTime endDate = Convert.ToDateTime(Year + "/" + Month + "/01").AddMonths(1).AddDays(-1);

                                int day = Convert.ToInt32(endDate.Subtract(startDate).TotalDays); //日期格式的相減

                                string colIndex = "";
                                int rowIndex = 1;

                                var worksheet = workbook.Worksheets.Add(restaurant.ToString()); //分頁名稱依照餐廳命名
                                worksheet.RowHeight = 15;
                                worksheet.Style = defaultStyle;

                                #region //HEADER
                                var companyList = detailReportDatas.Select(x => new { x.CompanyId, x.CompanyName }).Distinct().OrderBy(x => x.CompanyId).ToList();
                                foreach (var companyData in companyList)
                                {
                                    header.Add(companyData.CompanyName + "用餐份數"); //公司
                                }

                                for (int i = 0; i < header.Count(); i++)
                                {
                                    colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                    worksheet.Cell(colIndex).Value = header[i];
                                    worksheet.Cell(colIndex).Style = headerStyle;
                                }
                                #endregion

                                #region //BODY
                                var dates = detailReportDatas
                                                    .Where(x => x.RestaurantName == restaurant.ToString())
                                                    .Select(x => x.UmoDate).Distinct()
                                                    .OrderBy(x => x).ToList();

                                foreach (var date in dates)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = date.ToString("yyyy-MM-dd"); //日期
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = restaurant.ToString(); //廠商

                                    double totalQty = detailReportDatas.Where(x => x.UmoDate == date && x.RestaurantName == restaurant).Sum(x => x.UmoDetailQty);
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = totalQty; //訂購總份數
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Style = numberStyle;

                                    foreach (var companyData in companyList)
                                    {
                                        double companyQty = detailReportDatas.Where(x => x.UmoDate == date && x.RestaurantName == restaurant && x.CompanyId == companyData.CompanyId).Sum(x => x.UmoDetailQty);

                                        int columnIndex = 4 + companyList.IndexOf(companyData);
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, rowIndex)).Value = companyQty; //各公司用餐份數
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, rowIndex)).Style = numberStyle;
                                    }
                                }

                                #region //合計份數
                                int totalQtyRowIndex = rowIndex + 1;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex)).Value = "合計份數";

                                #region //合併儲存格
                                worksheet.Range(totalQtyRowIndex, 1, totalQtyRowIndex, 2).Merge();
                                //1, first row (0-based)
                                //1, first column  (0-based)
                                //1, last row (0-based)
                                //2  last column  (0-based)
                                #endregion

                                string totalQtyCell = BaseHelper.MergeNumberToChar(3, totalQtyRowIndex);
                                worksheet.Cell(totalQtyCell).Value = detailReportDatas.Where(x => x.RestaurantName == restaurant && dates.Contains(x.UmoDate)).Sum(x => x.UmoDetailQty);
                                worksheet.Cell(totalQtyCell).Style = numberStyle;

                                foreach (var companyData in companyList)
                                {
                                    double totalCompanyQty = detailReportDatas.Where(x => x.RestaurantName == restaurant && dates.Contains(x.UmoDate) && x.CompanyId == companyData.CompanyId).Sum(x => x.UmoDetailQty);
                                    int columnIndex = 4 + companyList.IndexOf(companyData);
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, totalQtyRowIndex)).Value = totalCompanyQty; //各公司用餐份數總計
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, totalQtyRowIndex)).Style = numberStyle;
                                }

                                #endregion

                                #region //合計總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, totalQtyRowIndex + 1)).Value = "合計總額";

                                #region //合併儲存格
                                worksheet.Range(totalQtyRowIndex + 1, 1, totalQtyRowIndex + 1, 2).Merge();
                                #endregion

                                double Amount = detailReportDatas.Where(x => x.RestaurantName == restaurant).Sum(x => x.UmoDetailPrice);

                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, totalQtyRowIndex + 1)).Value = Amount; //訂購份數合計總額
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, totalQtyRowIndex + 1)).Style = currencyStyle;

                                foreach (var companyData in companyList)
                                {
                                    double CompanyAmount = detailReportDatas.Where(x => x.RestaurantName == restaurant && x.CompanyId == companyData.CompanyId).Sum(x => x.UmoDetailPrice);

                                    int columnIndex = 4 + companyList.IndexOf(companyData);
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, totalQtyRowIndex + 1)).Value = CompanyAmount; //各公司用餐合計總額
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(columnIndex, totalQtyRowIndex + 1)).Style = currencyStyle;
                                }

                                #endregion
                                #endregion

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion

                                #region //凍結
                                //窗格、首欄、頂端列
                                //worksheet.SheetView.Freeze(1, 1);
                                //worksheet.SheetView.FreezeColumns(1);
                                worksheet.SheetView.FreezeRows(1);
                                #endregion
                            }
                        }

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion

                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            fileGuid,
                            fileName = excelFileName,
                            fileExtension = ".xlsx"
                        });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Request
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

        #region //DailyReportExcelDownload 外食餐點日報
        public void DailyReportExcelDownload(string StartDate = "", string EndDate = "", int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("DinerList", "excel");

                #region //Request
                dataRequest = applyFormDA.GetDetailReportData(CompanyId, -1, -1, "", StartDate, EndDate);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    List<DetailReportData> detailReportDatas = JsonConvert.DeserializeObject<List<DetailReportData>>(JObject.Parse(dataRequest)["data"].ToString());

                    var dates = detailReportDatas.Select(x => x).Distinct().ToList();

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("查無當日訂餐資料!");

                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //discountStyle
                    var discountStyle = XLWorkbook.DefaultStyle;
                    discountStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    discountStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    discountStyle.Border.TopBorder = XLBorderStyleValues.None;
                    discountStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    discountStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    discountStyle.Border.RightBorder = XLBorderStyleValues.None;
                    discountStyle.Border.TopBorderColor = XLColor.NoColor;
                    discountStyle.Border.BottomBorderColor = XLColor.NoColor;
                    discountStyle.Border.LeftBorderColor = XLColor.NoColor;
                    discountStyle.Border.RightBorderColor = XLColor.NoColor;
                    discountStyle.Fill.BackgroundColor = XLColor.LemonChiffon;
                    discountStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    discountStyle.Font.FontSize = 12;
                    discountStyle.Font.Bold = false;
                    //discountStyle.Font.FontColor = (XLColor.Red);
                    discountStyle.Protection.SetLocked(false);
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //headerStyleDate
                    var headerStyleDate = XLWorkbook.DefaultStyle;
                    headerStyleDate.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyleDate.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyleDate.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyleDate.Border.TopBorderColor = XLColor.Black;
                    headerStyleDate.Border.BottomBorderColor = XLColor.Black;
                    headerStyleDate.Border.LeftBorderColor = XLColor.Black;
                    headerStyleDate.Border.RightBorderColor = XLColor.Black;
                    headerStyleDate.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyleDate.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyleDate.Font.FontSize = 14;
                    headerStyleDate.Font.Bold = true;
                    headerStyleDate.DateFormat.Format = "MM/dd";
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0";
                    #endregion

                    #region //totalQtyStyle
                    var totalQtyStyle = XLWorkbook.DefaultStyle;
                    totalQtyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    totalQtyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    totalQtyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    totalQtyStyle.Border.TopBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    totalQtyStyle.Border.RightBorderColor = XLColor.NoColor;
                    totalQtyStyle.Fill.BackgroundColor = XLColor.PowderBlue;
                    totalQtyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    totalQtyStyle.Font.FontSize = 12;
                    totalQtyStyle.Font.Bold = false;
                    totalQtyStyle.NumberFormat.Format = "#,##0";
                    //totalQtyStyle.Font.FontColor = (XLColor.Red);
                    totalQtyStyle.Protection.SetLocked(false);
                    #endregion

                    #region //totalPriceStyle
                    var totalPriceStyle = XLWorkbook.DefaultStyle;
                    totalPriceStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    totalPriceStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    totalPriceStyle.Border.TopBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.RightBorder = XLBorderStyleValues.None;
                    totalPriceStyle.Border.TopBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.BottomBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.LeftBorderColor = XLColor.NoColor;
                    totalPriceStyle.Border.RightBorderColor = XLColor.NoColor;
                    totalPriceStyle.Fill.BackgroundColor = XLColor.PowderBlue;
                    totalPriceStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    totalPriceStyle.Font.FontSize = 12;
                    totalPriceStyle.Font.Bold = false;
                    totalPriceStyle.NumberFormat.Format = "#,##0";
                    //totalPriceStyle.Font.FontColor = (XLColor.Red);
                    totalPriceStyle.Protection.SetLocked(false);
                    #endregion
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "外食餐點日報Excel檔";
                    string excelsheetName = "外食餐點日報";
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var dailyInfos = detailReportDatas.Select(x => new { x.UmoDate, x.CompanyName, x.RestaurantName, x.MealName }).Distinct().OrderBy(x => x.UmoDate).ThenBy(x => x.CompanyName).ThenBy(x => x.RestaurantName).ThenBy(x => x.MealName);

                        List<string> header = new List<string> { "日期", "公司", "餐廳", "餐點", "總份數", "總計金額" };

                        string colIndex = "";
                        int rowIndex = 1;

                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 3)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, header.Count()).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Count(); i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        int index = 1;
                        foreach (var item in dailyInfos)
                        {
                            double qty = detailReportDatas.Where(x => x.UmoDate == item.UmoDate && x.CompanyName == item.CompanyName && x.RestaurantName == item.RestaurantName && x.MealName == item.MealName).Sum(x => x.UmoDetailQty);
                            double price = detailReportDatas.Where(x => x.UmoDate == item.UmoDate && x.CompanyName == item.CompanyName && x.RestaurantName == item.RestaurantName && x.MealName == item.MealName).Sum(x => x.UmoDetailPrice);

                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.UmoDate;         //日期
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CompanyName;     //公司
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.RestaurantName;  //餐廳
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MealName;        //餐點


                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = qty;                  //餐點總份數
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Style = currencyStyle;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = price;                //總計金額
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Style = currencyStyle;

                            index++;
                        }

                        #endregion

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
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
        #endregion
        #endregion
    }
}