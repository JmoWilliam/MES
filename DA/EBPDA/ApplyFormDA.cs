using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace EBPDA
{
    public class ApplyFormDA
    {
        public string MainConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();

        public string StartTime = "17:10:00";
        public string EndTime = "09:00:00";
        public string DiscountStartTime = "17:10:00";
        public string DiscountEndTime = "08:00:00";

        public ApplyFormDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //GetUserInfo 取得使用者資訊
        private void GetUserInfo()
        {
            try
            {
                CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["UserCompany"]);
                CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);

                if (HttpContext.Current.Session["CompanySwitch"] != null)
                {
                    if (HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                    {
                        CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["CompanyId"]);
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //Get
        #region //GetSubsidyRange -- 取得聚餐區間資料 -- Zoey 2023.02.15
        public string GetSubsidyRange(int SubsidyRangeId, int AnnualId, string RangeType, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SubsidyRangeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RangeName, a.AnnualId, a.RangeType, a.Status
                        , FORMAT(a.StartDate, 'yyyy-MM-dd') StartDate
                        , FORMAT(a.EndDate, 'yyyy-MM-dd') EndDate
                        , (CAST(b.Annual AS nvarchar) + '年度 ' + a.RangeName) AnnualWithRangeName
                        , b.Annual";
                    sqlQuery.mainTables =
                        @"FROM EBP.SubsidyRange a
                        INNER JOIN EBP.Annual b ON a.AnnualId = b.AnnualId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubsidyRangeId", @" AND a.SubsidyRangeId = @SubsidyRangeId", SubsidyRangeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnualId", @" AND a.AnnualId = @AnnualId", AnnualId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RangeType", @" AND a.RangeType = @RangeType", RangeType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AnnualId DESC, a.SubsidyRangeId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDinnerSubsidy -- 取得聚餐補助資料 -- Zoey 2023.02.15
        public string GetDinnerSubsidy(int DinnerSubsidyId, int SubsidyRangeId, int ApplyId, int PayeeId
            , string Status, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DinnerSubsidyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SubsidyRangeId, a.Amount, FORMAT(a.ApplyDate, 'yyyy-MM-dd') ApplyDate
                        , a.ApplyId, a.PayeeId, a.Status
                        , b.RangeName
                        , c.UserNo ApplyNo, c.UserName ApplyName
                        , d.UserNo PayeeNo, d.UserName PayeeName
                        , (CAST(e.Annual AS nvarchar) + '年度 ' + b.RangeName) AnnualWithRangeName";
                    sqlQuery.mainTables =
                        @"FROM EBP.DinnerSubsidy a
                        INNER JOIN EBP.SubsidyRange b ON b.SubsidyRangeId = a.SubsidyRangeId
                        INNER JOIN BAS.[User] c ON c.UserId = a.ApplyId
                        INNER JOIN BAS.[User] d ON d.UserId = a.PayeeId
                        INNER JOIN EBP.Annual e ON b.AnnualId = e.AnnualId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DinnerSubsidyId", @" AND a.DinnerSubsidyId = @DinnerSubsidyId", DinnerSubsidyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubsidyRangeId", @" AND a.SubsidyRangeId = @SubsidyRangeId", SubsidyRangeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ApplyId", @" AND a.ApplyId = @ApplyId", ApplyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PayeeId", @" AND a.PayeeId = @PayeeId", PayeeId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Status, a.ApplyDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetSubsidyCertificate -- 取得補助憑證資料 -- Zoey 2023.02.16
        public string GetSubsidyCertificate(string CertificateType, int DinnerSubsidyId, int ClubBudgetSubsidyId
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CertificateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CertificateType, a.FileId , a.DinnerSubsidyId, a.ClubBudgetSubsidyId
                          , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                          , b.FileName, b.FileContent, b.FileExtension, b.FileSize";
                    sqlQuery.mainTables =
                        @"FROM EBP.SubsidyCertificate a
                          INNER JOIN BAS.[File] b ON b.FileId = a.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CertificateType", @" AND a.CertificateType = @CertificateType", CertificateType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DinnerSubsidyId", @" AND a.DinnerSubsidyId = @DinnerSubsidyId", DinnerSubsidyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubBudgetSubsidyId", @" AND a.ClubBudgetSubsidyId = @ClubBudgetSubsidyId", ClubBudgetSubsidyId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CertificateId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetParticipants -- 取得參與人員資料 -- Zoey 2023.02.16
        public string GetParticipants(int DinnerSubsidyId, string SearchKey
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ParticipantId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserId, a.DinnerSubsidyId
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentName";
                    sqlQuery.mainTables =
                        @"FROM EBP.DinnerParticipant a
                        INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DinnerSubsidyId", @" AND a.DinnerSubsidyId = @DinnerSubsidyId", DinnerSubsidyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.UserNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetClubBudgetSubsidy -- 取得社團經費補助資料 -- Ben Ma 2023.02.20
        public string GetClubBudgetSubsidy(int ClubBudgetSubsidyId, int SubsidyRangeId, int ApplyId, int PayeeId
            , string Status, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ClubBudgetSubsidyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ClubId, a.SubsidyRangeId, a.Amount
                        , FORMAT(a.ApplyDate, 'yyyy-MM-dd') ApplyDate, a.ApplyId, a.PayeeId, a.[Status]
                        , b.ClubName
                        , c.RangeName
                        , d.UserNo ApplyNo, d.UserName ApplyName
                        , e.UserNo PayeeNo, e.UserName PayeeName
                        , (CAST(f.Annual AS nvarchar) + '年度 ' + c.RangeName) AnnualWithRangeName";
                    sqlQuery.mainTables =
                        @"FROM EBP.ClubBudgetSubsidy a
                        INNER JOIN EBP.Club b ON a.ClubId = b.ClubId
                        INNER JOIN EBP.SubsidyRange c ON a.SubsidyRangeId = c.SubsidyRangeId
                        INNER JOIN BAS.[User] d ON a.ApplyId = d.UserId
                        INNER JOIN BAS.[User] e ON a.PayeeId = e.UserId
                        INNER JOIN EBP.Annual f ON c.AnnualId = f.AnnualId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubBudgetSubsidyId", @" AND a.ClubBudgetSubsidyId = @ClubBudgetSubsidyId", ClubBudgetSubsidyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubsidyRangeId", @" AND a.SubsidyRangeId = @SubsidyRangeId", SubsidyRangeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ApplyId", @" AND a.ApplyId = @ApplyId", ApplyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PayeeId", @" AND a.PayeeId = @PayeeId", PayeeId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Status, a.ApplyDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetClubBudgetDetail -- 取得社團經費明細資料 -- Ben Ma 2023.02.20
        public string GetClubBudgetDetail(int ClubBudgetDetailId, int ClubBudgetSubsidyId
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ClubBudgetDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ClubBudgetSubsidyId, FORMAT(a.OccurrenceDate, 'yyyy-MM-dd') OccurrenceDate
                        , a.DetailDesc, a.Player, a.Amount, a.ActiveRegion";
                    sqlQuery.mainTables =
                        @"FROM EBP.ClubBudgetDetail a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubBudgetDetailId", @" AND a.ClubBudgetDetailId = @ClubBudgetDetailId", ClubBudgetDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubBudgetSubsidyId", @" AND a.ClubBudgetSubsidyId = @ClubBudgetSubsidyId", ClubBudgetSubsidyId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.OccurrenceDate DESC, a.ClubBudgetDetailId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetClubBudgetParticipant -- 取得社團經費補助人員資料 -- Ben Ma 2023.02.20
        public string GetClubBudgetParticipant(int ClubBudgetSubsidyId, string SearchKey
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ParticipantId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserId, a.ClubBudgetSubsidyId
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentName";
                    sqlQuery.mainTables =
                        @"FROM EBP.ClubBudgetParticipant a
                        INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubBudgetSubsidyId", @" AND a.ClubBudgetSubsidyId = @ClubBudgetSubsidyId", ClubBudgetSubsidyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.UserNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetClubBudgetSubsidyWord -- 取得社團補助單據資料 -- Ben Ma 2023.03.02
        public string GetClubBudgetSubsidyWord(int ClubBudgetSubsidyId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.Amount, FORMAT(a.ApplyDate, 'yyyy-MM-dd') ApplyDate
                            , b.ClubName
                            , c.UserName ApplyUser
                            , (e.DepartmentName + ' ' + d.UserNo + ' ' + d.UserName) PayeeUser
                            , (
                                SELECT FORMAT(z.OccurrenceDate, 'yyyy-MM-dd') OccurrenceDate
                                , z.DetailDesc, z.Player, z.Amount, z.ActiveRegion
                                FROM EBP.ClubBudgetDetail z
                                WHERE z.ClubBudgetSubsidyId = a.ClubBudgetSubsidyId
                                ORDER BY z.ClubBudgetSubsidyId
                                FOR JSON PATH, ROOT('data')
                            ) ClubBudgetDetail
                            FROM EBP.ClubBudgetSubsidy a
                            INNER JOIN EBP.Club b ON a.ClubId = b.ClubId
                            INNER JOIN BAS.[User] c ON a.ApplyId = c.UserId
                            INNER JOIN BAS.[User] d ON a.PayeeId = d.UserId
                            INNER JOIN BAS.Department e ON d.DepartmentId = e.DepartmentId
                            WHERE a.ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                    dynamicParameters.Add("@ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDinerInfo -- 取得用餐明細-點餐資料 -- Yi 2023.05.11
        public string GetDinerInfo(int UmoId, int CompanyId, string CompanyName, int UserId, string UserName
            , int RestaurantId, string RestaurantName, int MealId, string MealName, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UmoId";
                    sqlQuery.auxKey = "";
                    string sqlColumns = @"
                        , d.CompanyId, a.UserId, a.UmoNo, d.DepartmentName, c.UserNo + ' ' + c.UserName UserInfo, c.UserNo, c.UserName
                        , FORMAT(a.UmoDate, 'yyyy-MM-dd') UmoDate, a.UmoDiscount, a.UmoAmount
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate, b.CompanyName
                        , (
                            SELECT aa.UmoDetailId, ac.RestaurantId, ac.RestaurantName
                            , aa.MealId, aa.MealName, aa.UmoDetailRemark, aa.UmoDetailQty, aa.UmoDetailPrice, aa.Pickup, ad.TypeName PickupName
                            , ISNULL(ae.UmoAdditional, '') UmoAdditional
                            , (
                                SELECT aaa.McDetailId, aaa.McDetailName, aaa.UmoAdditionalPrice
                                FROM EBP.UmoAdditional aaa
                                WHERE aaa.UmoDetailId = aa.UmoDetailId
                                FOR JSON PATH
                            ) UmoAdditionalInfo
                            FROM EBP.UmoDetail aa
                            INNER JOIN EBP.RestaurantMeal ab ON ab.MealId = aa.MealId
                            INNER JOIN EBP.Restaurant ac ON ac.RestaurantId = ab.RestaurantId
                            INNER JOIN BAS.[Type] ad ON ad.TypeNo = aa.Pickup AND ad.TypeSchema = 'UmoDetail.Pickup'
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT '、' + aaa.McDetailName + ' ' + '($' + CAST(aaa.UmoAdditionalPrice AS NVARCHAR) + ')'
                                    FROM EBP.UmoAdditional aaa
                                    WHERE aaa.UmoDetailId = aa.UmoDetailId
                                    FOR XML PATH('')
                                ), 1, 1, '') UmoAdditional
                            ) ae
                            WHERE aa.UmoId = a.UmoId";
                    BaseHelper.SqlParameter(ref sqlColumns, ref dynamicParameters, "RestaurantId", @" AND ab.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref sqlColumns, ref dynamicParameters, "MealName", @" AND ab.MealName LIKE '%' + @MealName + '%'", MealName);
                    sqlColumns += @"
                            FOR JSON PATH, ROOT('data')
                        ) UmoDetail";
                    sqlQuery.columns = sqlColumns;
                    sqlQuery.mainTables =
                        @"FROM EBP.UserMealOrder a
                        INNER JOIN BAS.[User] c ON c.UserId = a.UserId
                        INNER JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company b ON b.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UmoId", @" AND a.UmoId = @UmoId", UmoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyName", @" AND b.CompanyName LIKE '%' + @CompanyName + '%'", CompanyName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND c.UserName LIKE '%' + @UserName + '%'", UserName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND ab.RestaurantId = @RestaurantId
                                                                                                        )", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantName", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                                                                                            INNER JOIN EBP.Restaurant ac ON ab.RestaurantId = ac.RestaurantId
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND ac.RestaurantName LIKE '%' + @RestaurantName + '%'
                                                                                                        )", RestaurantName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND aa.MealId = @MealId
                                                                                                        )", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealName", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND ab.MealName LIKE '%' + @MealName + '%'
                                                                                                        )", MealName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", EndDate);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UmoId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDailyDinerInfo -- 取得用餐明細-每日訂餐明細匯出 -- Yi 2023.06.15(目前未使用)
        public string GetDailyDinerInfo(int UmoId, int CompanyId, int UserId, int RestaurantId, string MealName, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UmoId";
                    sqlQuery.auxKey = "";

                    string sqlColumns = @"
                        , d.CompanyId, a.UserId, a.UmoNo, d.DepartmentName, c.UserNo + ' ' + c.UserName UserInfo, c.UserNo, c.UserName
                        , FORMAT(a.UmoDate, 'yyyy-MM-dd') UmoDate, a.UmoDiscount, a.UmoAmount
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate, b.CompanyName
                        , (
                            SELECT aa.UmoDetailId, ac.RestaurantId, ac.RestaurantName
                            , aa.MealId, aa.MealName, aa.UmoDetailRemark, aa.UmoDetailQty, aa.UmoDetailPrice, ad.TypeName Pickup
                            , ISNULL(ae.UmoAdditional, '') UmoAdditional
                            FROM EBP.UmoDetail aa
                            INNER JOIN EBP.RestaurantMeal ab ON ab.MealId = aa.MealId
                            INNER JOIN EBP.Restaurant ac ON ac.RestaurantId = ab.RestaurantId
                            INNER JOIN BAS.[Type] ad ON ad.TypeNo = aa.Pickup AND ad.TypeSchema = 'UmoDetail.Pickup'
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT '、' + aaa.McDetailName + ' ' + '($' + CAST(aaa.UmoAdditionalPrice AS NVARCHAR) + ')'
                                    FROM EBP.UmoAdditional aaa
                                    WHERE aaa.UmoDetailId = aa.UmoDetailId
                                    FOR XML PATH('')
                                ), 1, 1, '') UmoAdditional
                            ) ae
                            WHERE aa.UmoId = a.UmoId";
                    BaseHelper.SqlParameter(ref sqlColumns, ref dynamicParameters, "RestaurantId", @" AND ab.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref sqlColumns, ref dynamicParameters, "MealName", @" AND ab.MealName LIKE '%' + @MealName + '%'", MealName);
                    sqlColumns += @"
                            FOR JSON PATH, ROOT('data')
                        ) UmoDetail";
                    sqlQuery.columns = sqlColumns;
                    sqlQuery.mainTables =
                        @"FROM EBP.UserMealOrder a
                        INNER JOIN BAS.[User] c ON c.UserId = a.UserId
                        INNER JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company b ON b.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UmoId", @" AND a.UmoId = @UmoId", UmoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND ab.RestaurantId = @RestaurantId
                                                                                                        )", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealName", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM EBP.UmoDetail aa
                                                                                                            INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                                                                                            WHERE aa.UmoId = a.UmoId
                                                                                                            AND ab.MealName LIKE '%' + @MealName + '%'
                                                                                                        )", MealName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", EndDate);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UmoId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetMonthDinerInfo -- 取得用餐明細-每月訂餐總表匯出 -- Yi 2023.06.19(目前未使用)
        public string GetMonthDinerInfo(string Year, string Month, string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DateTime startDate = Convert.ToDateTime(Year + "/" + Month + "/01");
                    DateTime endDate = Convert.ToDateTime(Year + "/" + Month + "/01").AddMonths(1).AddDays(-1);

                    sql = @"SELECT a.UmoId, d.CompanyId, b.CompanyName, d.DepartmentName, c.UserNo, c.UserName
                            , FORMAT(a.UmoDate, 'MM-dd') UmoDate, a.UmoDiscount, a.UmoAmount, d.DepartmentNo, d.DepartmentName DepartmentCategory
                            , (
                                SELECT SUM(x.UmoDetailQty)
                                FROM EBP.UmoDetail x
                                WHERE x.UmoId = a.UmoId
                            ) UmoDetail
                            FROM EBP.UserMealOrder a
                            INNER JOIN BAS.[User] c ON c.UserId = a.UserId
                            INNER JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                            INNER JOIN BAS.Company b ON b.CompanyId = d.CompanyId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT '、' + aaa.McDetailName + ' ' + '($' + CAST(aaa.UmoAdditionalPrice AS NVARCHAR) + ')'
                                    FROM EBP.UmoAdditional aaa
                                    WHERE aaa.UmoDetailId = aa.UmoDetailId
                                    FOR XML PATH('')
                                ), 1, 1, '') UmoAdditional
                            ) ae";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", startDate.ToString("yyyy-MM-dd"));
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", endDate.ToString("yyyy-MM-dd"));
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND d.CompanyId LIKE '%' + @CompanyId + '%'", Company);
                    sql += @"
                            ORDER BY a.UmoId";
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetPickupInfo -- 取得取餐區資訊 -- Yi 2023.06.27(目前未使用)
        public string GetPickupInfo(int CompanyId, int UserId, int RestaurantId
            , string MealName, string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UmoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", d.CompanyId, ac.RestaurantId, c.UserNo, c.UserName, d.DepartmentName, a.MealName, e.TypeName Pickup, a.UmoDetailQty";
                    sqlQuery.mainTables =
                        @"FROM EBP.UmoDetail a
                        INNER JOIN EBP.UserMealOrder b ON b.UmoId = a.UmoId
                        INNER JOIN BAS.[User] c ON c.UserId = b.UserId
                        INNER JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.[Type] e ON e.TypeNo = a.Pickup AND e.TypeSchema = 'UmoDetail.Pickup'
                        INNER JOIN EBP.RestaurantMeal ab ON ab.MealId = a.MealId
                        INNER JOIN EBP.Restaurant ac ON ac.RestaurantId = ab.RestaurantId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND ac.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND b.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealName", @" AND a.MealName LIKE '%' + @MealName + '%'", MealName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.UmoDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.UmoDate <= @EndDate", EndDate);
                    sqlQuery.conditions = queryCondition;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetTotalMealList -- 取得餐點統計資訊 -- Yi 2023.06.29(目前未使用)
        public string GetTotalMealList(string Year, string Month)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DateTime startDate = Convert.ToDateTime(Year + "/" + Month + "/01");
                    DateTime endDate = Convert.ToDateTime(Year + "/" + Month + "/01").AddMonths(1).AddDays(-1);

                    sqlQuery.mainKey = "a.UmoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UmoDate, d.CompanyId, b.CompanyName, c.UserNo, c.UserName, e.RestaurantName, e.UmoDetailQty, a.UmoAmount, a.UmoDiscount";
                    sqlQuery.mainTables =
                          @"FROM EBP.UserMealOrder a
                            INNER JOIN BAS.[User] c ON c.UserId = a.UserId
                            INNER JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                            INNER JOIN BAS.Company b ON b.CompanyId = d.CompanyId
                            OUTER APPLY (
                                SELECT DISTINCT ec.RestaurantName, SUM(ea.UmoDetailQty) UmoDetailQty
                                FROM EBP.UmoDetail ea
                                INNER JOIN EBP.RestaurantMeal eb ON eb.MealId = ea.MealId
                                INNER JOIN EBP.Restaurant ec ON ec.RestaurantId = eb.RestaurantId
                                WHERE ea.UmoId = a.UmoId
                                GROUP BY ec.RestaurantName
                            ) e";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", startDate.ToString("yyyy-MM-dd"));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", endDate.ToString("yyyy-MM-dd"));
                    sqlQuery.conditions = queryCondition;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetMcDetail -- 取得客製化項目資訊 -- Yi 2023.07.18
        public string GetMcDetail(int MealId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    //判斷日期，當日日期超過17:10就會跳下一天餐點
                    string mealDate = "";

                    if (DateTime.Now.TimeOfDay >= Convert.ToDateTime(StartTime).TimeOfDay && DateTime.Now.TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                    {
                        mealDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        mealDate = DateTime.Now.ToString("yyyy-MM-dd");
                    }

                    sql = @"SELECT a.MealId, a.MealName, a.MealPrice, ISNULL(a.MealImage, -1) MealImage, b.RestaurantName
                            , (
                                SELECT aa.MealCgId, aa.CategoryName, aa.CategoryType
                                , (
                                    SELECT aaa.McDetailId, aaa.McDetailName, aaa.McDetailPrice
                                    FROM EBP.McDetail aaa
                                    WHERE aaa.MealCgId = aa.MealCgId
                                    FOR JSON PATH, ROOT('data')
                                ) McDetail
                                FROM EBP.MealCategory aa
                                WHERE aa.MealId = a.MealId
                                FOR JSON PATH, ROOT('data')
                            ) MealCategory
                            FROM EBP.RestaurantMeal a
                            INNER JOIN EBP.Restaurant b ON b.RestaurantId = a.RestaurantId
                            WHERE 1=1
                            AND a.MealId IN (
                                SELECT DISTINCT y.MealId
                                FROM EBP.DailyMeal y
                                INNER JOIN BAS.Calendar z ON z.CalendarId = y.CalendarId
                                WHERE FORMAT(z.CalendarDate, 'yyyy-MM-dd') = @MealDate
                            )";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MealId", @" AND a.MealId = @MealId", MealId);
                    dynamicParameters.Add("MealDate", mealDate);
                    sql += @"
                            ORDER BY a.MealId";

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDailyReport -- 取得用餐明細(外食餐點日報用) -- Yi 2023.09.08(目前未使用)
        public string GetDailyReport(string StartDate, string EndDate, string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.CompanyId, h.CompanyName
                            , a.UmoDate, e.RestaurantName, d.MealName
                            , i.UmoDetailQty, i.UmoAmount
                            FROM EBP.UserMealOrder a
                            INNER JOIN EBP.UmoDetail b ON b.UmoId = a.UmoId
                            LEFT JOIN EBP.UmoAdditional c ON c.UmoDetailId = b.UmoDetailId
                            INNER JOIN EBP.RestaurantMeal d ON d.MealId = b.MealId
                            INNER JOIN EBP.Restaurant e ON e.RestaurantId = d.RestaurantId
                            INNER JOIN BAS.[User] f ON f.UserId = a.UserId
                            INNER JOIN BAS.Department g ON g.DepartmentId = f.DepartmentId
                            INNER JOIN BAS.Company h ON h.CompanyId = g.CompanyId
                            OUTER APPLY(
                                SELECT SUM(ia.UmoDetailQty) UmoDetailQty, SUM(ia.UmoDetailPrice) UmoAmount
                                FROM EBP.UmoDetail ia
                                INNER JOIN EBP.UserMealOrder ib ON ib.UmoId = ia.UmoId
                                INNER JOIN EBP.RestaurantMeal ic ON ic.MealId = ia.MealId
                                INNER JOIN EBP.Restaurant id ON id.RestaurantId = ic.RestaurantId
                                INNER JOIN BAS.[User] ie ON ie.UserId = ib.UserId
                                INNER JOIN BAS.Department ig ON ig.DepartmentId = ie.DepartmentId
                                INNER JOIN BAS.Company ih ON ih.CompanyId = ig.CompanyId
                                WHERE ih.CompanyId = h.CompanyId
                                AND ib.UmoDate = a.UmoDate
                                AND id.RestaurantId = e.RestaurantId
                                AND ic.MealId = d.MealId
                            ) i
                            WHERE FORMAT(a.UmoDate, 'yyyy-MM-dd') >= @StartDate
                            AND FORMAT(a.UmoDate, 'yyyy-MM-dd') <= @EndDate";
                    //dynamicParameters.Add("UmoDate", StartDate);
                    //dynamicParameters.Add("UmoDate", EndDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", EndDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND h.CompanyId LIKE '%' + @CompanyId + '%'", Company);
                    sql += @"
                            ORDER BY a.CompanyId, h.CompanyName, a.UmoDate, e.RestaurantName, d.MealName, i.UmoDetailQty, i.UmoAmount";
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetReportData -- 取得用餐明細-單頭報表明細匯出 -- Yi 2023.09.28
        public string GetReportData(int CompanyId, int UserId, string Year, string Month)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DateTime startDate = Convert.ToDateTime(Year + "/" + Month + "/01");
                    DateTime endDate = Convert.ToDateTime(Year + "/" + Month + "/01").AddMonths(1).AddDays(-1);

                    sql = @"SELECT a.UmoId, c.CompanyId, d.CompanyName, c.DepartmentName, b.UserNo, b.UserName, FORMAT(a.UmoDate, 'MM-dd') UmoDate
                            , a.UmoDiscount, a.UmoAmount, c.DepartmentNo, c.DepartmentCategory, e.UmoDetailQty, e.UmoAmount
                            FROM EBP.UserMealOrder a
                            INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                            INNER JOIN BAS.Department c ON c.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company d ON d.CompanyId = c.CompanyId
                            OUTER APPLY (
                                SELECT SUM(x.UmoDetailQty) UmoDetailQty, SUM(x.UmoDetailPrice) UmoAmount
                                FROM EBP.UmoDetail x
                                WHERE x.UmoId = a.UmoId
                            ) e
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.UmoDate >= @StartDate", startDate.ToString("yyyy-MM-dd"));
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.UmoDate <= @EndDate", endDate.ToString("yyyy-MM-dd"));
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    sql += @" ORDER BY a.UmoId";
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDetailReportData -- 取得用餐明細-單身報表明細匯出(不含折扣價格) -- Yi 2023.09.28
        public string GetDetailReportData(int CompanyId, int UserId, int RestaurantId, string MealName, string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UmoId, f.CompanyId, g.CompanyName, b.UmoDate, d.RestaurantId, d.RestaurantName, c.MealName, e.UserNo, e.UserName
                            , a.UmoDetailQty, a.UmoDetailPrice, f.DepartmentNo, f.DepartmentName, f.DepartmentCategory
                            , h.TypeName Pickup
                            FROM EBP.UmoDetail a
                            INNER JOIN EBP.UserMealOrder b ON a.UmoId = b.UmoId
                            INNER JOIN EBP.RestaurantMeal c ON a.MealId = c.MealId
                            INNER JOIN EBP.Restaurant d ON c.RestaurantId = d.RestaurantId
                            INNER JOIN BAS.[User] e ON b.UserId = e.UserId
                            INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                            INNER JOIN BAS.Company g ON f.CompanyId = g.CompanyId
                            INNER JOIN BAS.[Type] h ON a.Pickup =  h.TypeNo AND h.TypeSchema = 'UmoDetail.Pickup'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND b.UmoDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND b.UmoDate <= @EndDate", EndDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND f.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RestaurantId", @" AND d.RestaurantId = @RestaurantId", RestaurantId);
                    sql += @" ORDER BY a.UmoId DESC";

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Add
        #region //AddSubsidyRange -- 補助區間資料新增 -- Zoey 2023.02.15
        public string AddSubsidyRange(int AnnualId, string RangeType, string RangeName, string StartDate, string EndDate)
        {
            try
            {
                if (AnnualId <= 0) throw new SystemException("【年份】不能為空");
                if (RangeType.Length <= 0) throw new SystemException("【區間類型】不能為空!");
                if (RangeName.Length <= 0) throw new SystemException("【區間名稱】不能為空!");
                if (RangeName.Length > 100) throw new SystemException("【區間名稱】長度錯誤!");
                if (StartDate.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (!DateTime.TryParse(StartDate, out DateTime tempDate)) throw new SystemException("【開始日期】格式錯誤!");
                if (EndDate.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (!DateTime.TryParse(EndDate, out tempDate)) throw new SystemException("【結束日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷區間類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "SubsidyRange.RangeType");
                        dynamicParameters.Add("TypeNo", RangeType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("區間類型資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.SubsidyRange (AnnualId, RangeType, RangeName, StartDate, EndDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SubsidyRangeId
                                VALUES (@AnnualId, @RangeType, @RangeName, @StartDate, @EndDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AnnualId,
                                RangeType,
                                RangeName,
                                StartDate,
                                EndDate,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddDinnerSubsidy -- 聚餐補助資料新增 -- Zoey 2023.02.16
        public string AddDinnerSubsidy(int SubsidyRangeId, int Amount, string ApplyDate, int ApplyId, int PayeeId)
        {
            try
            {
                if (SubsidyRangeId <= 0) throw new SystemException("【聚餐補助】不能為空!");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ApplyDate.Length <= 0) throw new SystemException("【申請時間】不能為空");
                if (!DateTime.TryParse(ApplyDate, out DateTime tempDate)) throw new SystemException("【申請時間】格式錯誤!");
                if (ApplyId <= 0) throw new SystemException("【申請人】不能為空");
                if (PayeeId <= 0) throw new SystemException("【領款人】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷補助區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE RangeType = @RangeType
                                AND SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("RangeType", "Dinner");
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("補助區間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.DinnerSubsidy (SubsidyRangeId, Amount, ApplyDate, ApplyId, PayeeId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DinnerSubsidyId, INSERTED.Status
                                VALUES (@SubsidyRangeId, @Amount, @ApplyDate, @ApplyId, @PayeeId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SubsidyRangeId,
                                Amount,
                                ApplyDate,
                                ApplyId,
                                PayeeId,
                                Status = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddSubsidyCertificate -- 補助憑證新增 -- Zoey 2023.02.16
        public string AddSubsidyCertificate(string CertificateType, int FileId, int DinnerSubsidyId, int ClubBudgetSubsidyId)
        {
            try
            {
                if (!Regex.IsMatch(CertificateType, "^(Dinner|Club)$", RegexOptions.IgnoreCase)) throw new SystemException("【憑證類型】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        switch (CertificateType)
                        {
                            case "Dinner":
                                #region //判斷聚餐補助資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.DinnerSubsidy
                                        WHERE DinnerSubsidyId = @DinnerSubsidyId";
                                dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                                var resultDinner = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDinner.Count() <= 0) throw new SystemException("聚餐補助資料錯誤!");
                                #endregion
                                break;
                            case "Club":
                                #region //判斷社團經費補助資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.ClubBudgetSubsidy
                                        WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                                dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                                var resultClubBudget = sqlConnection.Query(sql, dynamicParameters);
                                if (resultClubBudget.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                                #endregion
                                break;
                        }

                        #region //判斷檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.SubsidyCertificate (CertificateType, FileId, DinnerSubsidyId, ClubBudgetSubsidyId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CertificateId
                                VALUES (@CertificateType, @FileId, @DinnerSubsidyId, @ClubBudgetSubsidyId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CertificateType,
                                FileId,
                                DinnerSubsidyId = DinnerSubsidyId <= 0 ? (int?)null : DinnerSubsidyId,
                                ClubBudgetSubsidyId = ClubBudgetSubsidyId <= 0 ? (int?)null : ClubBudgetSubsidyId,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddParticipants -- 聚餐補助人員資料新增 -- Zoey 2023.02.16
        public string AddParticipants(int DinnerSubsidyId, string Users)
        {
            try
            {
                if(Users.Length <= 0) throw new SystemException("【參與人員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.DinnerSubsidy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐補助資料錯誤!");
                        #endregion

                        #region //判斷參與人員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("參與人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            #region //判斷該人員該補助區間是否已補助
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 c.UserNo, c.UserName
                                    FROM EBP.DinnerParticipant a
                                    INNER JOIN EBP.DinnerSubsidy b ON a.DinnerSubsidyId = b.DinnerSubsidyId
                                    INNER JOIN BAS.[User] c ON a.UserId = c.UserId
                                    WHERE a.UserId = @UserId
                                    AND b.SubsidyRangeId IN (
                                        SELECT TOP 1 SubsidyRangeId
                                        FROM EBP.DinnerSubsidy
                                        WHERE DinnerSubsidyId = @DinnerSubsidyId
                                    )";
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));
                            dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                            var resultSubsidy = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSubsidy.Count() > 0)
                            {
                                foreach (var item in resultSubsidy)
                                {
                                    throw new SystemException(string.Format("{0} {1} 重複申請!", item.UserNo, item.UserName));
                                }
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.DinnerParticipant (DinnerSubsidyId, UserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ParticipantId
                                    VALUES (@DinnerSubsidyId, @UserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DinnerSubsidyId,
                                    UserId = Convert.ToInt32(user),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddClubBudgetSubsidy -- 社團經費補助資料新增 -- Ben Ma 2023.02.20
        public string AddClubBudgetSubsidy(int ClubId, int SubsidyRangeId, int Amount, string ApplyDate, int ApplyId, int PayeeId)
        {
            try
            {
                if (ClubId <= 0) throw new SystemException("【社團】不能為空!");
                if (SubsidyRangeId <= 0) throw new SystemException("【聚餐補助】不能為空!");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ApplyDate.Length <= 0) throw new SystemException("【申請時間】不能為空");
                if (!DateTime.TryParse(ApplyDate, out DateTime tempDate)) throw new SystemException("【申請時間】格式錯誤!");
                if (ApplyId <= 0) throw new SystemException("【申請人】不能為空");
                if (PayeeId <= 0) throw new SystemException("【領款人】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Club
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團資料錯誤!");
                        #endregion

                        #region //判斷補助區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE RangeType = @RangeType
                                AND SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("RangeType", "Club");
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("補助區間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.ClubBudgetSubsidy (ClubId, SubsidyRangeId, Amount, ApplyDate, ApplyId, PayeeId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClubBudgetSubsidyId, INSERTED.Status
                                VALUES (@ClubId, @SubsidyRangeId, @Amount, @ApplyDate, @ApplyId, @PayeeId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClubId,
                                SubsidyRangeId,
                                Amount,
                                ApplyDate,
                                ApplyId,
                                PayeeId,
                                Status = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddClubBudgetDetail -- 社團經費明細資料新增 -- Ben Ma 2023.02.20
        public string AddClubBudgetDetail(int ClubBudgetSubsidyId, string OccurrenceDate, string DetailDesc
            , int Player, int Amount, string ActiveRegion)
        {
            try
            {
                if (OccurrenceDate.Length <= 0) throw new SystemException("【發生日期】不能為空");
                if (!DateTime.TryParse(OccurrenceDate, out DateTime tempDate)) throw new SystemException("【發生日期】格式錯誤!");
                if (DetailDesc.Length <= 0) throw new SystemException("【內容說明】不能為空");
                if (Player < 5) throw new SystemException("【參與人數】至少需要5人");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ActiveRegion.Length <= 0) throw new SystemException("【活動地點】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.ClubBudgetDetail (ClubBudgetSubsidyId, OccurrenceDate, DetailDesc
                                , Player, Amount, ActiveRegion
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClubBudgetDetailId
                                VALUES (@ClubBudgetSubsidyId, @OccurrenceDate, @DetailDesc
                                , @Player, @Amount, @ActiveRegion
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClubBudgetSubsidyId,
                                OccurrenceDate,
                                DetailDesc,
                                Player,
                                Amount,
                                ActiveRegion,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddClubBudgetParticipant -- 社團經費補助人員資料新增 -- Ben Ma 2023.02.20
        public string AddClubBudgetParticipant(int ClubBudgetSubsidyId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【參與人員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                        #endregion

                        #region //判斷參與人員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("參與人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            #region //判斷人員是否重複新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EBP.ClubBudgetParticipant a
                                    WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId
                                    AND UserId = @UserId";
                            dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));

                            var resultUser = sqlConnection.Query(sql, dynamicParameters);
                            if (resultUser.Count() > 0) throw new SystemException("參與人員已重複新增!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.ClubBudgetParticipant (ClubBudgetSubsidyId, UserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ParticipantId
                                    VALUES (@ClubBudgetSubsidyId, @UserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ClubBudgetSubsidyId,
                                    UserId = Convert.ToInt32(user),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddDinerInfo -- 團膳用餐明細新增 -- Yi 2023.05.18
        public string AddDinerInfo(int CompanyId, int UserId, string UmoDate, double UmoDiscount, double UmoAmount
            , string UmoDetailRemark, int UmoDetailQty, double UmoDetailPrice, string MealDetails)
        {
            try
            {
                #region //判斷餐點資料長度
                if (CompanyId <= 0) throw new SystemException("【公司】不能為空!");
                if (UserId <= 0) throw new SystemException("【人員】不能為空!");
                if (UmoDate.Length <= 0) throw new SystemException("【日期】不能為空!");
                #endregion

                if (!MealDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("客製化資料格式錯誤");

                JObject mealJson = JObject.Parse(MealDetails);
                //if (!mealJson.ContainsKey("data")) throw new SystemException("客製化資料錯誤");

                //JToken mealData = mealJson["data"];
                //if (mealData.Count() < 0) throw new SystemException("查無客製化選項內容");

                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //訂單編號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(UmoNo), '000'), 3)) + 1 CurrentNum
                                FROM EBP.UserMealOrder
                                WHERE UmoNo LIKE @UmoNo";
                        dynamicParameters.Add("UmoNo", string.Format("{0}{1}___", "FD", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string umoNo = string.Format("{0}{1}{2}", "FD", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        #region//新增餐點訂單
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.UserMealOrder (CompanyId, UserId, UmoDate, UmoNo, UmoDiscount, UmoAmount
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UmoId
                                VALUES (@CompanyId, @UserId, @UmoDate, @UmoNo, @UmoDiscount, @UmoAmount
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                UserId,
                                UmoDate,
                                UmoNo = umoNo,
                                UmoDiscount,
                                UmoAmount,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int umoId = -1;
                        foreach (var item in insertResult)
                        {
                            umoId = Convert.ToInt32(item.UmoId);
                        }
                        #endregion

                        #region //UserId取出UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultoUserNo = sqlConnection.Query(sql, dynamicParameters);

                        string userNo = "";
                        foreach (var item in resultoUserNo)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        double totalPrice = 0;
                        for (int i = 0; i < mealJson["UmoDetail"].Count(); i++)
                        {
                            int price = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoDetailPrice"]);
                            int qty = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoDetailQty"]);

                            #region//新增餐點明細
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.UmoDetail (UmoId, MealId, MealName, UmoDetailRemark, UmoDetailQty, UmoDetailPrice, Pickup
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.UmoDetailId
                                    VALUES (@UmoId, @MealId, @MealName, @UmoDetailRemark, @UmoDetailQty, @UmoDetailPrice, @Pickup
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UmoId = umoId,
                                    MealId = Convert.ToInt32(mealJson["UmoDetail"][i]["MealId"]),
                                    MealName = mealJson["UmoDetail"][i]["MealName"].ToString(),
                                    UmoDetailRemark = mealJson["UmoDetail"][i]["UmoDetailRemark"].ToString(),
                                    UmoDetailQty = mealJson["UmoDetail"][i]["UmoDetailQty"].ToString(),
                                    UmoDetailPrice = price,
                                    Pickup = mealJson["UmoDetail"][i]["Pickup"].ToString(),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int UmoDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                UmoDetailId = item.UmoDetailId;
                            }
                            #endregion

                            if (mealJson["UmoDetail"][i]["UmoAdditional"] != null)
                            {
                                totalPrice += Caculate(price, mealJson["UmoDetail"][i]["UmoAdditional"]);

                                for (int j = 0; j < mealJson["UmoDetail"][i]["UmoAdditional"].Count(); j++)
                                {
                                    #region //新增餐點加購明細
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO EBP.UmoAdditional (UmoDetailId, McDetailId, McDetailName, UmoAdditionalPrice
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@UmoDetailId, @McDetailId, @McDetailName, @UmoAdditionalPrice
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            UmoDetailId,
                                            McDetailId = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoAdditional"][j]["McDetailId"]),
                                            McDetailName = mealJson["UmoDetail"][i]["UmoAdditional"][j]["McDetailName"].ToString(),
                                            UmoAdditionalPrice = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoAdditional"][j]["UmoAdditionalPrice"]),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }
                            }
                            else
                            {
                                totalPrice += price;
                            }
                        }

                        #region//更新餐點明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.UserMealOrder SET
                                CompanyId = @CompanyId,
                                UserId = @UserId,
                                UmoDate = @UmoDate,
                                UmoDiscount = @UmoDiscount,
                                UmoAmount = @UmoAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UmoId = @UmoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            CompanyId,
                            UserId,
                            UmoDate,
                            UmoDiscount,
                            UmoAmount = totalPrice,
                            LastModifiedDate,
                            LastModifiedBy,
                            umoId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增餐點訂單Log
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.UserMealOrderLog (UmoId, UserNo, UmoNo, MotionDate, MotionRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UmoId, @UserNo, @UmoNo, @MotionDate, @MotionRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UmoId = umoId,
                                UserNo = userNo,
                                UmoNo = umoNo,
                                MotionDate = currentDate,
                                MotionRemark = "後台—【團膳用餐明細表】資料【新增】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }

                #region //計算餐點費用總計
                int Caculate(int unitPrice, JToken meal)
                {
                    int totalprice = unitPrice;

                    for (int x = 0; x < meal.Count(); x++)
                    {
                        totalprice += Convert.ToInt32(meal[x]["UmoAdditionalPrice"]);

                    }

                    return totalprice;
                }
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Update
        #region //UpdateSubsidyRange -- 補助區間資料更新 -- Zoey 2023.02.15
        public string UpdateSubsidyRange(int SubsidyRangeId, int AnnualId, string RangeName, string StartDate, string EndDate)
        {
            try
            {
                if (AnnualId <= 0) throw new SystemException("【年份】不能為空");
                if (RangeName.Length <= 0) throw new SystemException("【區間名稱】不能為空!");
                if (RangeName.Length > 100) throw new SystemException("【區間名稱】長度錯誤!");
                if (StartDate.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (!DateTime.TryParse(StartDate, out DateTime tempDate)) throw new SystemException("【開始日期】格式錯誤!");
                if (EndDate.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (!DateTime.TryParse(EndDate, out tempDate)) throw new SystemException("【結束日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷補助區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐區間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.SubsidyRange SET
                                AnnualId = @AnnualId,
                                RangeName = @RangeName,
                                StartDate = @StartDate,
                                EndDate = @EndDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AnnualId,
                                RangeName,
                                StartDate,
                                EndDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                SubsidyRangeId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateSubsidyRangeStatus -- 聚餐區間狀態更新 -- Zoey 2023.02.15
        public string UpdateSubsidyRangeStatus(int SubsidyRangeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.SubsidyRange
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐區間資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.SubsidyRange SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SubsidyRangeId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateDinnerSubsidy -- 聚餐補助資料更新 -- Zoey 2023.02.16
        public string UpdateDinnerSubsidy(int DinnerSubsidyId, int SubsidyRangeId, int Amount, string ApplyDate, int ApplyId, int PayeeId)
        {
            try
            {
                if (SubsidyRangeId <= 0) throw new SystemException("【聚餐補助】不能為空!");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ApplyDate.Length <= 0) throw new SystemException("【申請時間】不能為空");
                if (!DateTime.TryParse(ApplyDate, out DateTime tempDate)) throw new SystemException("【申請時間】格式錯誤!");
                if (ApplyId <= 0) throw new SystemException("【申請人】不能為空");
                if (PayeeId <= 0) throw new SystemException("【領款人】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.DinnerSubsidy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐補助資料錯誤!");
                        #endregion

                        #region //判斷補助區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE RangeType = @RangeType
                                AND SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("RangeType", "Dinner");
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("補助區間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.DinnerSubsidy SET
                                SubsidyRangeId = @SubsidyRangeId,
                                Amount = @Amount,
                                ApplyDate = @ApplyDate,
                                ApplyId = @ApplyId,
                                PayeeId = @PayeeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SubsidyRangeId,
                                Amount,
                                ApplyDate,
                                ApplyId,
                                PayeeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                DinnerSubsidyId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateSubsidyStatus -- 聚餐補助狀態更新 -- Zoey 2023.02.16
        public string UpdateSubsidyStatus(int DinnerSubsidyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.DinnerSubsidy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐補助資料錯誤!");
                        #endregion

                        #region //判斷是否有上傳憑證
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyCertificate
                                WHERE CertificateType = @CertificateType
                                AND DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("CertificateType", "Dinner");
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("無上傳憑證資料!");
                        #endregion

                        #region //判斷補助人數是否有大於5人
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM (
                                    SELECT COUNT(1) TotalParticipant
                                    FROM EBP.DinnerParticipant
                                    WHERE DinnerSubsidyId = @DinnerSubsidyId
                                ) a
                                WHERE a.TotalParticipant >= 5";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("單次補助人數需大於5人!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.DinnerSubsidy SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                DinnerSubsidyId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateSubsidyCertificateRename -- 補助憑證更名 -- Ben Ma 2023.03.01
        public string UpdateSubsidyCertificateRename(int CertificateId, string NewFileName)
        {
            try
            {
                if (NewFileName.Length <= 0) throw new SystemException("【新檔名】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷補助憑證資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyCertificate
                                WHERE CertificateId = @CertificateId";
                        dynamicParameters.Add("CertificateId", CertificateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("補助憑證資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.FileName = @FileName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM BAS.[File] a
                                INNER JOIN EBP.SubsidyCertificate b ON a.FileId = b.FileId
                                WHERE b.CertificateId = @CertificateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileName = NewFileName,
                                LastModifiedDate,
                                LastModifiedBy,
                                CertificateId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateClubBudgetSubsidy -- 社團經費補助資料更新 -- Ben Ma 2023.02.20
        public string UpdateClubBudgetSubsidy(int ClubBudgetSubsidyId, int ClubId, int SubsidyRangeId, int Amount, string ApplyDate, int ApplyId, int PayeeId)
        {
            try
            {
                if (ClubId <= 0) throw new SystemException("【社團】不能為空!");
                if (SubsidyRangeId <= 0) throw new SystemException("【聚餐補助】不能為空!");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ApplyDate.Length <= 0) throw new SystemException("【申請時間】不能為空");
                if (!DateTime.TryParse(ApplyDate, out DateTime tempDate)) throw new SystemException("【申請時間】格式錯誤!");
                if (ApplyId <= 0) throw new SystemException("【申請人】不能為空");
                if (PayeeId <= 0) throw new SystemException("【領款人】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                        #endregion

                        #region //判斷社團資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Club
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("社團資料錯誤!");
                        #endregion

                        #region //判斷補助區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE RangeType = @RangeType
                                AND SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("RangeType", "Club");
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("補助區間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.ClubBudgetSubsidy SET
                                ClubId = @ClubId,
                                SubsidyRangeId = @SubsidyRangeId,
                                Amount = @Amount,
                                ApplyDate = @ApplyDate,
                                ApplyId = @ApplyId,
                                PayeeId = @PayeeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClubId,
                                SubsidyRangeId,
                                Amount,
                                ApplyDate,
                                ApplyId,
                                PayeeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubBudgetSubsidyId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateClubBudgetSubsidyStatus -- 社團經費補助狀態更新 -- Ben Ma 2023.02.20
        public string UpdateClubBudgetSubsidyStatus(int ClubBudgetSubsidyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                        #endregion

                        #region //判斷是否有上傳憑證
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyCertificate
                                WHERE CertificateType = @CertificateType
                                AND ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("CertificateType", "Club");
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("無上傳憑證資料!");
                        #endregion

                        #region //判斷是否有經費明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetDetail
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("無經費明細資料!");
                        #endregion

                        #region //判斷補助人數是否有大於5人
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM (
                                    SELECT COUNT(1) TotalParticipant
                                    FROM EBP.ClubBudgetParticipant
                                    WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId
                                ) a
                                WHERE a.TotalParticipant >= 5";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("單次補助人數需大於5人!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.ClubBudgetSubsidy SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubBudgetSubsidyId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateClubBudgetDetail -- 社團經費明細資料更新 -- Ben Ma 2023.02.20
        public string UpdateClubBudgetDetail(int ClubBudgetDetailId, string OccurrenceDate, string DetailDesc
            , int Player, int Amount, string ActiveRegion)
        {
            try
            {
                if (OccurrenceDate.Length <= 0) throw new SystemException("【發生日期】不能為空");
                if (!DateTime.TryParse(OccurrenceDate, out DateTime tempDate)) throw new SystemException("【發生日期】格式錯誤!");
                if (DetailDesc.Length <= 0) throw new SystemException("【內容說明】不能為空");
                if (Player < 5) throw new SystemException("【參與人數】至少需要5人");
                if (Amount <= 0) throw new SystemException("【金額】不能小於0");
                if (ActiveRegion.Length <= 0) throw new SystemException("【活動地點】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費明細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetDetail
                                WHERE ClubBudgetDetailId = @ClubBudgetDetailId";
                        dynamicParameters.Add("ClubBudgetDetailId", ClubBudgetDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費明細資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.ClubBudgetDetail SET
                                OccurrenceDate = @OccurrenceDate,
                                DetailDesc = @DetailDesc,
                                Player = @Player,
                                Amount = @Amount,
                                ActiveRegion = @ActiveRegion,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubBudgetDetailId = @ClubBudgetDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OccurrenceDate,
                                DetailDesc,
                                Player,
                                Amount,
                                ActiveRegion,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubBudgetDetailId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateDinerInfo -- 團膳用餐明細更新 -- Yi 2023.05.18
        public string UpdateDinerInfo(int UmoId, int CompanyId, int UserId, string UmoDate, double UmoDiscount, string MealDetails)
        {
            try
            {
                #region //判斷餐點資料長度
                if (CompanyId <= 0) throw new SystemException("【公司】不能為空!");
                if (UserId <= 0) throw new SystemException("【人員】不能為空!");
                if (UmoDate.Length <= 0) throw new SystemException("【日期】不能為空!");
                #endregion

                if (!MealDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("客製化資料格式錯誤");

                JObject mealJson = JObject.Parse(MealDetails);
                //if (!mealJson.ContainsKey("data")) throw new SystemException("客製化資料錯誤");

                //JToken mealData = mealJson["data"];
                //if (mealData.Count() < 0) throw new SystemException("查無客製化選項內容");

                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐點訂單是否正確
                        sql = @"SELECT TOP 1 UmoId, UmoNo
                                FROM EBP.UserMealOrder
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("團膳用餐明細資料錯誤!");

                        int umoId = -1;
                        string umoNo = "";
                        foreach (var item in result)
                        {
                            umoId = item.UmoId;
                            umoNo = item.UmoNo;
                        }
                        #endregion

                        #region //判斷餐點明細是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UmoDetailId
                                FROM EBP.UserMealOrder a
                                INNER JOIN EBP.UmoDetail b ON b.UmoId = a.UmoId
                                WHERE a.UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetail.Count() <= 0) throw new SystemException("餐點明細資料錯誤!");

                        int UmoDetailId = -1;
                        foreach (var item in resultDetail)
                        {
                            UmoDetailId = item.UmoDetailId;

                            #region //加購資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE EBP.UmoAdditional
                                WHERE UmoDetailId = @UmoDetailId";
                            dynamicParameters.Add("UmoDetailId", UmoDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //餐點明細資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.UmoDetail
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //UserId取出UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultoUserNo = sqlConnection.Query(sql, dynamicParameters);

                        string userNo = "";
                        foreach (var item in resultoUserNo)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        double totalPrice = 0;

                        for (int i = 0; i < mealJson["UmoDetail"].Count(); i++)
                        {
                            int price = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoDetailPrice"]);
                            int qty = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoDetailQty"]);

                            #region//新增餐點明細
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.UmoDetail (UmoId, MealId, MealName, UmoDetailRemark, UmoDetailQty, UmoDetailPrice, Pickup
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.UmoDetailId
                                    VALUES (@UmoId, @MealId, @MealName, @UmoDetailRemark, @UmoDetailQty, @UmoDetailPrice, @Pickup
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UmoId = umoId,
                                    MealId = Convert.ToInt32(mealJson["UmoDetail"][i]["MealId"]),
                                    MealName = mealJson["UmoDetail"][i]["MealName"].ToString(),
                                    UmoDetailRemark = mealJson["UmoDetail"][i]["UmoDetailRemark"].ToString(),
                                    UmoDetailQty = mealJson["UmoDetail"][i]["UmoDetailQty"].ToString(),
                                    UmoDetailPrice = price,
                                    Pickup = mealJson["UmoDetail"][i]["Pickup"].ToString(),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            UmoDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                UmoDetailId = item.UmoDetailId;
                            }
                            #endregion

                            if (mealJson["UmoDetail"][i]["UmoAdditional"] != null)
                            {
                                totalPrice += Caculate(price, mealJson["UmoDetail"][i]["UmoAdditional"]);

                                for (int j = 0; j < mealJson["UmoDetail"][i]["UmoAdditional"].Count(); j++)
                                {
                                    #region //新增餐點加購明細
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO EBP.UmoAdditional (UmoDetailId, McDetailId, McDetailName, UmoAdditionalPrice
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@UmoDetailId, @McDetailId, @McDetailName, @UmoAdditionalPrice
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            UmoDetailId,
                                            McDetailId = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoAdditional"][j]["McDetailId"]),
                                            McDetailName = mealJson["UmoDetail"][i]["UmoAdditional"][j]["McDetailName"].ToString(),
                                            UmoAdditionalPrice = Convert.ToInt32(mealJson["UmoDetail"][i]["UmoAdditional"][j]["UmoAdditionalPrice"]),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                totalPrice += price;
                            }
                        }

                        #region//更新餐點明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.UserMealOrder SET
                                CompanyId = @CompanyId,
                                UserId = @UserId,
                                UmoDate = @UmoDate,
                                UmoDiscount = @UmoDiscount,
                                UmoAmount = @UmoAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UmoId = @UmoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            CompanyId,
                            UserId,
                            UmoDate,
                            UmoDiscount,
                            UmoAmount = totalPrice,
                            LastModifiedDate,
                            LastModifiedBy,
                            UmoId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增餐點訂單Log
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.UserMealOrderLog (UmoId, UserNo, UmoNo, MotionDate, MotionRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UmoId, @UserNo, @UmoNo, @MotionDate, @MotionRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UmoId = umoId,
                                UserNo = userNo,
                                UmoNo = umoNo,
                                MotionDate = currentDate,
                                MotionRemark = "後台—【團膳用餐明細表】資料【修改】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }

                #region //計算餐點費用總計
                int Caculate(int unitPrice, JToken meal)
                {
                    int totalprice = unitPrice;

                    for (int x = 0; x < meal.Count(); x++)
                    {
                        totalprice += Convert.ToInt32(meal[x]["UmoAdditionalPrice"]);

                    }

                    return totalprice;
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Delete
        #region //DeleteSubsidyRange -- 聚餐區間資料新增 -- Zoey 2023.02.15
        public string DeleteSubsidyRange(int SubsidyRangeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐區間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyRange
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐區間資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.SubsidyRange
                                WHERE SubsidyRangeId = @SubsidyRangeId";
                        dynamicParameters.Add("SubsidyRangeId", SubsidyRangeId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteDinnerSubsidy -- 聚餐補助資料刪除 -- Zoey 2023.02.16
        public string DeleteDinnerSubsidy(int DinnerSubsidyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.DinnerSubsidy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聚餐補助資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除補助憑證
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.SubsidyCertificate
                                WHERE CertificateType = @CertificateType
                                AND DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("CertificateType", "Dinner");
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除聚餐補助人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.DinnerParticipant
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.DinnerSubsidy
                                WHERE DinnerSubsidyId = @DinnerSubsidyId";
                        dynamicParameters.Add("DinnerSubsidyId", DinnerSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteSubsidyCertificate -- 補助憑證刪除 -- Zoey 2023.02.16
        public string DeleteSubsidyCertificate(int CertificateId, string CertificateType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聚餐補助憑證是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.SubsidyCertificate
                                WHERE CertificateId = @CertificateId
                                AND CertificateType = @CertificateType";
                        dynamicParameters.Add("CertificateId", CertificateId);
                        dynamicParameters.Add("CertificateType", CertificateType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("補助憑證資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.SubsidyCertificate
                                WHERE CertificateId = @CertificateId
                                AND CertificateType = @CertificateType";
                        dynamicParameters.Add("CertificateId", CertificateId);
                        dynamicParameters.Add("CertificateType", CertificateType);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteParticipants -- 聚餐補助人員資料刪除 -- Zoey 2023.02.16
        public string DeleteParticipants(int ParticipantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷參與人員資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.DinnerParticipant
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("參與人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.DinnerParticipant
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteClubBudgetSubsidy -- 社團經費補助資料刪除 -- Ben Ma 2023.02.20
        public string DeleteClubBudgetSubsidy(int ClubBudgetSubsidyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除補助憑證
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.SubsidyCertificate
                                WHERE CertificateType = @CertificateType
                                AND ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("CertificateType", "Club");
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除社團經費明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubBudgetDetail
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除社團經費補助人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubBudgetParticipant
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubBudgetSubsidy
                                WHERE ClubBudgetSubsidyId = @ClubBudgetSubsidyId";
                        dynamicParameters.Add("ClubBudgetSubsidyId", ClubBudgetSubsidyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteClubBudgetDetail -- 社團經費明細資料刪除 -- Ben Ma 2023.02.20
        public string DeleteClubBudgetDetail(int ClubBudgetDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetDetail
                                WHERE ClubBudgetDetailId = @ClubBudgetDetailId";
                        dynamicParameters.Add("ClubBudgetDetailId", ClubBudgetDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費明細資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubBudgetDetail
                                WHERE ClubBudgetDetailId = @ClubBudgetDetailId";
                        dynamicParameters.Add("ClubBudgetDetailId", ClubBudgetDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteClubBudgetParticipant -- 社團經費補助人員資料刪除 -- Ben Ma 2023.02.20
        public string DeleteClubBudgetParticipant(int ParticipantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團經費補助人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubBudgetParticipant
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團經費補助人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubBudgetParticipant
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //DeleteDinerInfo -- 用餐明細-點餐資料刪除 -- Yi 2023.05.15
        public string DeleteDinerInfo(int UmoId)
        {
            try
            {
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷點餐資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UmoId, UserId, UmoNo
                                FROM EBP.UserMealOrder
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("用餐明細資料錯誤!");

                        int umoId = -1;
                        int userId = -1;
                        string umoNo = "";
                        foreach (var item in result)
                        {
                            umoId = item.UmoId;
                            userId = item.UserId;
                            umoNo = item.UmoNo;
                        }
                        #endregion

                        #region //UserId取出UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", userId);

                        var resultoUserNo = sqlConnection.Query(sql, dynamicParameters);

                        string userNo = "";
                        foreach (var item in resultoUserNo)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除加購項目
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.UmoAdditional a
                                INNER JOIN EBP.UmoDetail b ON a.UmoDetailId = b.UmoDetailId
                                INNER JOIN EBP.UserMealOrder c ON b.UmoId = c.UmoId
                                WHERE c.UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除餐點資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.UmoDetail a
                                INNER JOIN EBP.UserMealOrder b ON a.UmoId = b.UmoId
                                WHERE b.UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #endregion

                        #region //刪除主要table-餐點訂單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.UserMealOrder
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增餐點訂單Log
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.UserMealOrderLog (UmoId, UserNo, UmoNo, MotionDate, MotionRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.UmoLogId
                                    VALUES (@UmoId, @UserNo, @UmoNo, @MotionDate, @MotionRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UmoId = umoId,
                                UserNo = userNo,
                                UmoNo = umoNo,
                                MotionDate = currentDate,
                                MotionRemark = "後台—【團膳用餐明細表】資料【刪除】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion
    }
}
