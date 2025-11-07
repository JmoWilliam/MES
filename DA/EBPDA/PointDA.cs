using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace EBPDA
{
    public class PointDA
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

        public PointDA()
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
        #region //GetRedeemPrize -- 取得兌換獎勵資料 -- Zoey 2023.02.17
        public string GetRedeemPrize(int RedeemPrizeId, int AnnualId, string PrizeName, int RedeemPoint, string Status
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RedeemPrizeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PrizeName, a.RedeemPoint, a.AnnualId, a.Status
                          , b.Annual";
                    sqlQuery.mainTables =
                        @"FROM EBP.RedeemPrize a
                          INNER JOIN EBP.Annual b ON b.AnnualId = a.AnnualId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RedeemPrizeId", @" AND a.RedeemPrizeId = @RedeemPrizeId", RedeemPrizeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnualId", @" AND a.AnnualId = @AnnualId", AnnualId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrizeName", @" AND a.PrizeName LIKE '%' + @PrizeName + '%'", PrizeName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RedeemPoint", @" AND a.RedeemPoint = @RedeemPoint", RedeemPoint);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AnnualId DESC, a.RedeemPrizeId";
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

        #region //GetUserPoint -- 取得人員積分資料 -- Zoey 2023.02.22
        public string GetUserPoint(int DepartmentId, int UserId
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender, a.UserNo + a.UserName UserFullNo
                          , b.DepartmentNo, b.DepartmentName
                          , c.LogoIcon
                          , z.PointQty";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                          INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                          INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                          OUTER APPLY(
                            SELECT ISNULL(SUM(x.PointQty), 0) PointQty
                            FROM EBP.Point x
                            INNER JOIN EBP.Annual y ON y.AnnualId = x.AnnualId
                            WHERE y.Status = 'A'
                            AND x.UserId = a.UserId
                          ) z";
                    sqlQuery.auxTables = "";
                    string queryCondition = "AND c.CompanyId = @CompanyId " +
                                            "AND a.Status = 'A'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserNo";
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

        #region //GetPersonalPoints -- 取得個人積分資料 -- Zoey 2023.02.20
        public string GetPersonalPoints(int UserId, int AnnualId, string PointType
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.PointId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PointQty, a.Remark, a.ActivityId, a.ClubId, a.RedeemPrizeId
                          , c.Annual
                          , d.TypeName
                          , e.ActivityName
                          , f.ClubName
                          , g.PrizeName";
                    sqlQuery.mainTables =
                        @"FROM EBP.Point a
                          INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                          INNER JOIN EBP.Annual c ON c.AnnualId = a.AnnualId
                          INNER JOIN BAS.[Type] d ON d.TypeNo = a.PointType AND d.TypeSchema = 'Point.PointType'
                          LEFT JOIN EBP.Activity e ON e.ActivityId = a.ActivityId
                          LEFT JOIN EBP.Club f ON f.ClubId = a.ClubId
                          LEFT JOIN EBP.RedeemPrize g ON g.RedeemPrizeId = a.RedeemPrizeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND b.UserId = @UserId";
                    dynamicParameters.Add("UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnualId", @" AND a.AnnualId = @AnnualId", AnnualId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PointType", @" AND a.PointType = @PointType", PointType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PointId";
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
        #endregion

        #region //Add
        #region //AddRedeemPrize -- 兌換獎勵資料新增 -- Zoey 2023.02.17
        public string AddRedeemPrize(int AnnualId, string PrizeName, int RedeemPoint)
        {
            try
            {
                if (AnnualId <= 0) throw new SystemException("【年份】不能為空!");
                if (PrizeName.Length <= 0) throw new SystemException("【獎勵名稱】不能為空!");
                if (PrizeName.Length > 100) throw new SystemException("【獎勵名稱】長度錯誤!");
                if (RedeemPoint <= 0) throw new SystemException("兌換點數不可小於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷獎勵名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RedeemPrize
                                WHERE PrizeName = @PrizeName";
                        dynamicParameters.Add("PrizeName", PrizeName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【獎勵名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.RedeemPrize (PrizeName, RedeemPoint, AnnualId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RedeemPrizeId
                                VALUES (@PrizeName, @RedeemPoint, @AnnualId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PrizeName,
                                RedeemPoint,
                                AnnualId,
                                Status = "A",
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

        #region //AddPoint -- 人員積分資料新增 -- Zoey 2023.02.20
        public string AddPoint(int AnnualId, string UserId, string PointType, int RedeemPrizeId, int PointQty, string Remark)
        {
            try
            {
                if (AnnualId <= 0) throw new SystemException("【年份】不能為空!");
                if (UserId.Length <= 0) throw new SystemException("【人員】不能為空!");
                if (PointType.Length <= 0) throw new SystemException("【類別】不能為空!");
                if (PointQty == 0) throw new SystemException("【點數】不能為０!");
                if (Remark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                if (Remark.Length <= 0) throw new SystemException("【備註】不能為空!");

                List<int> users = UserId.Split(',').Select(int.Parse).ToList();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        foreach (var user in users)
                        {
                            #region //判斷人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", user);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("人員資料錯誤!");
                            #endregion

                            switch (PointType)
                            {
                                case "Redeem":
                                    if (RedeemPrizeId <= 0) throw new SystemException("【獎勵】不能為空!");
                                    if (PointQty >= 0) throw new SystemException("兌換獎勵點數不可大於０!");
                                    break;

                                case "Other":
                                    break;

                                default:
                                    throw new SystemException("類別錯誤請重新選擇!");
                            }

                            if (PointQty < 0)
                            {
                                #region //判斷點數是否足夠兌換
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.Point
                                        WHERE UserId = @UserId
                                        AND AnnualId = @AnnualId
                                        HAVING SUM(PointQty) + @PointQty >= 0";
                                dynamicParameters.Add("UserId", user);
                                dynamicParameters.Add("AnnualId", AnnualId);
                                dynamicParameters.Add("PointQty", PointQty);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() <= 0) throw new SystemException("點數不足無法兌換!");
                                #endregion
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.Point (UserId, AnnualId, PointQty, PointType
                                    , ActivityId, ClubId, RedeemPrizeId, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PointId
                                    VALUES (@UserId, @AnnualId, @PointQty, @PointType
                                    , @ActivityId, @ClubId, @RedeemPrizeId, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId = user,
                                    AnnualId,
                                    PointQty,
                                    PointType,
                                    ActivityId = (int?)null,
                                    ClubId = (int?)null,
                                    RedeemPrizeId = RedeemPrizeId > 0 ? (int?)RedeemPrizeId : null,
                                    Remark,
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
        #endregion

        #region //Update
        #region //UpdateRedeemPrize -- 兌換獎勵資料更新 -- Zoey 2023.02.17
        public string UpdateRedeemPrize(int RedeemPrizeId, int AnnualId, string PrizeName, int RedeemPoint)
        {
            try
            {
                if (AnnualId <= 0) throw new SystemException("【年份】不能為空!");
                if (PrizeName.Length <= 0) throw new SystemException("【獎勵名稱】不能為空!");
                if (PrizeName.Length > 100) throw new SystemException("【獎勵名稱】長度錯誤!");
                if (RedeemPoint <= 0) throw new SystemException("兌換點數不可小於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷兌換獎勵資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RedeemPrize
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.Add("RedeemPrizeId", RedeemPrizeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("兌換獎勵資料錯誤!");
                        #endregion

                        #region //判斷獎勵名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RedeemPrize
                                WHERE PrizeName = @PrizeName
                                AND RedeemPrizeId != @RedeemPrizeId";
                        dynamicParameters.Add("PrizeName", PrizeName);
                        dynamicParameters.Add("RedeemPrizeId", RedeemPrizeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【獎勵名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.RedeemPrize SET
                                AnnualId = @AnnualId,
                                PrizeName = @PrizeName,
                                RedeemPoint = @RedeemPoint,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AnnualId,
                                PrizeName,
                                RedeemPoint,
                                LastModifiedDate,
                                LastModifiedBy,
                                RedeemPrizeId
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

        #region //UpdateRedeemPrizeStatus -- 兌換獎勵狀態更新 -- Zoey 2023.02.17
        public string UpdateRedeemPrizeStatus(int RedeemPrizeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷兌換獎勵資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.RedeemPrize
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.Add("RedeemPrizeId", RedeemPrizeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("兌換獎勵資料錯誤!");

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
                        sql = @"UPDATE EBP.RedeemPrize SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RedeemPrizeId
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
        #endregion

        #region //Delete
        #region //DeleteRedeemPrize -- 刪除兌換獎勵資料 -- Zoey 2023.02.17
        public string DeleteRedeemPrize(int RedeemPrizeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷兌換獎勵是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RedeemPrize
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.Add("RedeemPrizeId", RedeemPrizeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("兌換獎勵資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.RedeemPrize
                                WHERE RedeemPrizeId = @RedeemPrizeId";
                        dynamicParameters.Add("RedeemPrizeId", RedeemPrizeId);

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
        #endregion
    }
}
