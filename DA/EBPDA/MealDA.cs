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
    public class MealDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string ErpSysConnectionStrings = "";
        public string HrmConnectionStrings = "";

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

        public MealDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpSysConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

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
        #region //GetRestaurantMeal -- 取得餐點資料 -- Yi 2023.05.03
        public string GetRestaurantMeal(int MealId, int RestaurantId, string RestaurantName, string MealName, double MealPrice, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MealId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RestaurantId, a.MealName, a.MealPrice, a.MealImage, a.Status
                        , b.RestaurantName
                        , '【' + b.RestaurantName + '】' + a.MealName MealWithRestaurant";
                    sqlQuery.mainTables =
                        @"FROM EBP.RestaurantMeal a
                        INNER JOIN EBP.Restaurant b ON b.RestaurantId = a.RestaurantId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealId", @" AND a.MealId = @MealId", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND a.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantName", @" AND b.RestaurantName LIKE '%' + @RestaurantName + '%'", RestaurantName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealName", @" AND a.MealName LIKE '%' + @MealName + '%'", MealName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealStatus", @" AND a.Status IN @MealStatus", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MealId DESC";
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

        #region //GetRestaurantId -- 取得餐廳名稱(Cmb用) -- Yi 2023.05.03
        public string GetRestaurantId(int RestaurantId, string RestaurantNo, string RestaurantName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RestaurantId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.RestaurantNo
                        , CASE a.[Status] 
                            WHEN 'S' THEN a.RestaurantName + '(停用)'
                            ELSE a.RestaurantName
                        END RestaurantName";
                    sqlQuery.mainTables =
                        @"FROM EBP.Restaurant a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND a.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantNo", @" AND a.RestaurantNo LIKE '%' + @RestaurantNo + '%'", RestaurantNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantName", @" AND a.RestaurantName LIKE '%' + @RestaurantName + '%'", RestaurantName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RestaurantId";
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

        #region //GetMealCategory -- 取得餐點類別資料 -- Yi 2023.05.04
        public string GetMealCategory(int MealCgId, int MealId, int RestaurantId, string CategoryName, string CategoryType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MealCgId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MealId, b.RestaurantId, a.CategoryName
                        , CASE WHEN a.CategoryType = 'Y' THEN '是' WHEN a.CategoryType = 'N' THEN '否' END CategoryType";
                    sqlQuery.mainTables =
                        @"FROM EBP.MealCategory a
                        INNER JOIN EBP.RestaurantMeal b ON b.MealId = a.MealId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealCgId", @" AND a.MealCgId = @MealCgId", MealCgId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealId", @" AND a.MealId = @MealId", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND b.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CategoryName", @" AND a.CategoryName LIKE '%' + @CategoryName + '%'", CategoryName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CategoryType", @" AND a.CategoryType = @CategoryType", CategoryType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MealCgId";
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

        #region //GetMcDetail -- 取得餐點選項(細節)資料 -- Yi 2023.05.05
        public string GetMcDetail(int McDetailId, int MealCgId, int MealId, string CategoryName, string McDetailName
            , double McDetailPrice
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.McDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MealCgId, b.MealId, a.McDetailName, a.McDetailPrice
                        , b.CategoryName";
                    sqlQuery.mainTables =
                        @"FROM EBP.McDetail a
                        INNER JOIN EBP.MealCategory b ON b.MealCgId = a.MealCgId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McDetailId", @" AND a.McDetailId = @McDetailId", McDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealCgId", @" AND a.MealCgId = @MealCgId", MealCgId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealId", @" AND b.MealId = @MealId", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CategoryName", @" AND a.CategoryName LIKE '%' + @CategoryName + '%'", CategoryName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McDetailName", @" AND a.McDetailName = @McDetailName", McDetailName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.McDetailId";
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

        #region //GetMealCgId -- 取得類別名稱(Cmb用) -- Yi 2023.05.05
        public string GetMealCgId(int MealId, int MealCgId, string CategoryName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MealCgId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MealId, a.CategoryName";
                    sqlQuery.mainTables =
                        @"FROM EBP.MealCategory a
                        INNER JOIN EBP.RestaurantMeal b ON b.MealId = a.MealId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.MealId = @MealId";
                    dynamicParameters.Add("MealId", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealId", @" AND a.MealId = @MealId", MealId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MealCgId", @" AND a.MealCgId = @MealCgId", MealCgId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CategoryName", @" AND a.CategoryName LIKE '%' + @CategoryName + '%'", CategoryName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MealCgId";
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

        #region //GetDailyMeal -- 取得每日餐點資料 -- Yi 2023.05.08
        public string GetDailyMeal(int MealId, string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.CalendarId, FORMAT(a.CalendarDate, 'yyyy-MM-dd') CalendarDate
                            , SUBSTRING(b.MealInfo, 0, LEN(b.MealInfo)) MealInfo
                            FROM BAS.Calendar a
                            OUTER APPLY(
                                SELECT (
                                    SELECT CAST(ba.MealId as NVARCHAR)  + ','
                                    FROM EBP.DailyMeal ba
                                    WHERE ba.CalendarId = a.CalendarId
                                    FOR XML PATH('')
                                ) MealInfo
                            ) b
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CalendarDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CalendarDate <= @EndDate", EndDate);
                    sql += @"
                            ORDER BY a.CalendarDate";
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

        #region //GetLotDailyMeal -- 取得每日餐點資料(批次) -- Ben Ma 2023.11.03
        public string GetLotDailyMeal(string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.CalendarId, FORMAT(a.CalendarDate, 'yyyy-MM-dd') CalendarDate
                            , FORMAT(a.CalendarDate, 'yyyy-MM-dd') + ' (' +
                            CASE DATEPART(weekday, a.CalendarDate)
                                WHEN 1 THEN '日'
                                WHEN 2 THEN '一'
                                WHEN 3 THEN '二'
                                WHEN 4 THEN '三'
                                WHEN 5 THEN '四'
                                WHEN 6 THEN '五'
                                WHEN 7 THEN '六'
                            END + ')' CalendarWeekDate
                            , (
                                SELECT aa.MealId
                                , '【' + ac.RestaurantName + '】' + ab.MealName MealWithRestaurant
                                FROM EBP.DailyMeal aa
                                INNER JOIN EBP.RestaurantMeal ab ON aa.MealId = ab.MealId
                                INNER JOIN EBP.Restaurant ac ON ab.RestaurantId = ac.RestaurantId
                                WHERE aa.CalendarId = a.CalendarId
                                ORDER BY ac.RestaurantName, ab.MealName
                                FOR JSON PATH, ROOT('data')
                            ) DailyMeals
                            FROM BAS.Calendar a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CalendarDate >= @StartDate", StartDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CalendarDate <= @EndDate", EndDate);
                    sql += @" ORDER BY a.CalendarDate";

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

        #region //GetMealId -- 取得餐點資料(Cmb用) -- Yi 2023.05.09
        public string GetMealId()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.MealId, a.RestaurantId, b.RestaurantName + '-' + a.MealName MealName
                            , CASE a.[Status] 
                                WHEN 'S' THEN b.RestaurantName + '-' + a.MealName + '(停用)'
                                ELSE b.RestaurantName + '-' + a.MealName
                            END MealName
                            FROM EBP.RestaurantMeal a
                            INNER JOIN EBP.Restaurant b ON b.RestaurantId = a.RestaurantId
                            WHERE b.CompanyId = @CompanyId
                            ORDER BY a.MealId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
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

        #region //GetCalendarDate -- 取得訂餐開放時段資料 -- Yi 2023.11.13
        public string GetCalendarDate(string OrderType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.OrderType, FORMAT(a.StartTime, 'HH:mm') StartTime, FORMAT(a.EndTime, 'HH:mm') EndTime, b.TypeName OrderTypeName
                            FROM EBP.OrderSetting a
                            INNER JOIN BAS.[Type] b ON a.OrderType = b.TypeNo AND b.TypeSchema = 'OrderSetting.OrderType'
                            WHERE a.OrderType = @OrderType";
                    dynamicParameters.Add("OrderType", OrderType);
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
        #region //AddRestaurantMeal -- 餐點資料新增 -- Yi 2023.05.04
        public string AddRestaurantMeal(int RestaurantId, string MealName, double MealPrice, int MealImage)
        {
            try
            {
                #region //判斷餐點資料長度
                if (RestaurantId <= 0) throw new SystemException("【餐廳】不能為空!");
                if (MealName.Length <= 0) throw new SystemException("【餐點名稱】不能為空!");
                if (MealName.Length > 30) throw new SystemException("【餐點名稱】長度錯誤!");
                if (MealPrice <= 0) throw new SystemException("【餐點金額】不能為空!");
                //if (MealImage <= 0) throw new SystemException("【餐點圖片】不能為空!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.RestaurantMeal (RestaurantId, MealName, MealImage, MealPrice, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MealId
                                VALUES (@RestaurantId, @MealName, @MealImage, @MealPrice, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RestaurantId,
                                MealName,
                                MealImage = MealImage == -1 ? (int?)null : MealImage,
                                MealPrice,
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

        #region //AddMealCategory -- 新增餐點類別資料 -- Yi 2023.05.04
        public string AddMealCategory(int MealId, int RestaurantId, string CategoryName, string CategoryType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (CategoryName.Length <= 0) throw new SystemException("【客製化類別】不能為空!");
                        if (CategoryType.Length <= 0) throw new SystemException("【是否開放多選】不能為空!");

                        #region //確認餐點資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RestaurantId
                                FROM EBP.RestaurantMeal a
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐點資訊有誤!");

                        foreach (var item in result)
                        {
                            RestaurantId = item.RestaurantId;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.MealCategory (MealId, RestaurantId, CategoryName, CategoryType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MealCgId
                                VALUES (@MealId, @RestaurantId, @CategoryName, @CategoryType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MealId,
                                RestaurantId,
                                CategoryName,
                                CategoryType,
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

        #region //AddMcDetail -- 新增餐點選項(細節)資料 -- Yi 2023.05.05
        public string AddMcDetail(int MealCgId, string McDetailName, double McDetailPrice)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (McDetailName.Length <= 0) throw new SystemException("【客製化選項】不能為空!");
                        if (McDetailPrice < 0) throw new SystemException("【價格】不能為空!");

                        #region //確認類別資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.MealCategory a
                                INNER JOIN EBP.RestaurantMeal b ON b.MealId = a.MealId
                                WHERE MealCgId = @MealCgId";
                        dynamicParameters.Add("MealCgId", MealCgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客製化類別有誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.McDetail (MealCgId, McDetailName, McDetailPrice
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.McDetailId
                                VALUES (@MealCgId, @McDetailName, @McDetailPrice
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MealCgId,
                                McDetailName,
                                McDetailPrice,
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

        #region //AddDailyMeal -- 新增每日餐點資料 -- Yi 2023.05.10
        public string AddDailyMeal(int CalendarId, string Meals)
        {
            try
            {
                //if (Meals.Length <= 0) throw new SystemException("【餐點資料】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //判斷行事曆資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Calendar
                                WHERE CalendarId = @CalendarId";
                        dynamicParameters.Add("CalendarId", CalendarId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("行事曆資料錯誤!");
                        #endregion

                        #region //刪除每日餐點table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.DailyMeal
                                WHERE CalendarId = @CalendarId";
                        dynamicParameters.Add("CalendarId", CalendarId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (Meals.Length > 0)
                        {
                            string[] mealsList = Meals.Split(',');
                            #region //判斷每日餐點資料是否正確
                            if (mealsList.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT COUNT(1) TotalMeals
                                    FROM EBP.RestaurantMeal
                                    WHERE MealId IN @MealId";
                                dynamicParameters.Add("MealId", mealsList);

                                int totalMeals = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalMeals;
                                if (totalMeals != mealsList.Length) throw new SystemException("餐點資料錯誤!");
                            }
                            #endregion

                            foreach (var meal in mealsList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.DailyMeal (MealId, CalendarId
                                    , CreateDate, CreateBy)
                                    VALUES (@MealId, @CalendarId
                                    , @CreateDate, @CreateBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CalendarId,
                                        MealId = Convert.ToInt32(meal),
                                        CreateDate,
                                        CreateBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
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

        #region //AddLotDailyMeal -- 新增每日餐點資料(批次) -- Ben Ma 2023.11.03
        public string AddLotDailyMeal(string Calendars, string Meals)
        {
            try
            {
                if (Calendars.Length <= 0) throw new SystemException("【日期】不能為空!");
                if (Meals.Length <= 0) throw new SystemException("【餐點】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷日期資料是否正確
                        string[] calendarsList = Calendars.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalCalendars
                                FROM BAS.Calendar
                                WHERE CalendarId IN @CalendarId";
                        dynamicParameters.Add("CalendarId", calendarsList);

                        int totalCalendars = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalCalendars;
                        if (totalCalendars != calendarsList.Length) throw new SystemException("【日期】資料錯誤!");
                        #endregion

                        #region //判斷餐點資料是否正確
                        string[] mealsList = Meals.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalMeals
                                FROM EBP.RestaurantMeal
                                WHERE MealId IN @MealId";
                        dynamicParameters.Add("MealId", mealsList);

                        int totalMeals = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalMeals;
                        if (totalMeals != mealsList.Length) throw new SystemException("【餐點】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var calendar in calendarsList)
                        {
                            #region //刪除每日餐點table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE EBP.DailyMeal
                                    WHERE CalendarId = @CalendarId";
                            dynamicParameters.Add("CalendarId", Convert.ToInt32(calendar));
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            foreach (var meal in mealsList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.DailyMeal (MealId, CalendarId
                                        , CreateDate, CreateBy)
                                        VALUES (@MealId, @CalendarId
                                        , @CreateDate, @CreateBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CalendarId = Convert.ToInt32(calendar),
                                        MealId = Convert.ToInt32(meal),
                                        CreateDate,
                                        CreateBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
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

        #region //AddCopyRestaurantMeal-- 複製餐點資料 -- Yi 2023-05-23
        public string AddCopyRestaurantMeal(int CompanyId, int MealId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐點資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RestaurantId, a.MealName, a.MealImage, a.MealPrice, a.[Status]
                                FROM EBP.RestaurantMeal a
                                WHERE a.MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐點資料錯誤!");

                        int RestaurantId = 0;
                        string MealName = "";
                        int MealImage = 0;
                        double MealPrice = 0;
                        string Status = "";
                        foreach (var item in result)
                        {
                            RestaurantId = item.RestaurantId;
                            MealName = "Copy From: " + item.MealName;
                            MealImage = item.MealImage;
                            MealPrice = item.MealPrice;
                            Status = item.Status;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.RestaurantMeal (RestaurantId, MealName, MealImage, MealPrice, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MealId
                                VALUES (@RestaurantId, @MealName, @MealImage, @MealPrice, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RestaurantId,
                                MealName,
                                MealImage,
                                MealPrice,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int newMealId = 0;
                        foreach (var item in insertResult)
                        {
                            newMealId = item.MealId;
                        }

                        #region //複製餐點客製化類別
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MealId, a.RestaurantId, a.CategoryName, a.CategoryType
                                , (
                                    SELECT b.McDetailName, b.McDetailPrice
                                    FROM EBP.McDetail b
                                    INNER JOIN EBP.MealCategory ba ON ba.MealCgId = b.MealCgId
                                    WHERE b.MealCgId = a.MealCgId
                                    FOR JSON PATH, ROOT('data')
                                ) McDetail
                                FROM EBP.MealCategory a
                                WHERE a.MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            foreach (var item in result2)
                            {
                                RestaurantId = item.RestaurantId;
                                string CategoryName = item.CategoryName;
                                string CategoryType = item.CategoryType;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.MealCategory (MealId, RestaurantId, CategoryName, CategoryType
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.MealCgId
                                        VALUES (@MealId, @RestaurantId, @CategoryName, @CategoryType
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MealId = newMealId,
                                        RestaurantId,
                                        CategoryName,
                                        CategoryType,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int MealCgId = 1;
                                foreach (var item2 in insertResult)
                                {
                                    MealCgId = item2.MealCgId;
                                }

                                if (item.McDetail != null)
                                {
                                    var mcDetailJson = JObject.Parse(item.McDetail);

                                    foreach (var item3 in mcDetailJson.data)
                                    {
                                        #region //複製餐點客製化選項
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO EBP.McDetail (MealCgId, McDetailName, McDetailPrice
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.MealCgId, INSERTED.McDetailName
                                                VALUES (@MealCgId, @McDetailName, @McDetailPrice
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MealCgId,
                                                McDetailName = item3["McDetailName"].ToString(),
                                                McDetailPrice = item3["McDetailPrice"].ToString(),
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
                            }
                        }
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

        #region //Update
        #region //UpdateRestaurantMeal -- 餐點資料更新 -- Yi 2023.05.04
        public string UpdateRestaurantMeal(int MealId, int RestaurantId, string MealName, double MealPrice, int MealImage)
        {
            try
            {
                #region //判斷餐點資料長度
                if (RestaurantId <= 0) throw new SystemException("【餐廳】不能為空!");
                if (MealName.Length <= 0) throw new SystemException("【餐點名稱】不能為空!");
                if (MealName.Length > 30) throw new SystemException("【餐點名稱】長度錯誤!");
                if (MealPrice <= 0) throw new SystemException("【餐點金額】不能為空!");
                //if (MealImage <= 0) throw new SystemException("【餐點圖片】不能為空!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐點資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐點資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.RestaurantMeal SET
                                RestaurantId = @RestaurantId,
                                MealName = @MealName,
                                MealImage = @MealImage,
                                MealPrice = @MealPrice,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MealId = @MealId";
                        var parametersObject = new
                        {
                            RestaurantId,
                            MealName,
                            MealImage = MealImage == -1 ? (int?)null : MealImage,
                            MealPrice,
                            LastModifiedDate,
                            LastModifiedBy,
                            MealId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdateMealStatus -- 餐點狀態更新 -- Yi 2023.05.03
        public string UpdateMealStatus(int MealId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐點資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【餐點】資料錯誤!");

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
                        sql = @"UPDATE EBP.RestaurantMeal SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MealId = @MealId";
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            MealId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdateMealCategory -- 更新餐點類別資料 -- Yi 2023.05.04
        public string UpdateMealCategory(int MealCgId, int MealId, int RestaurantId, string CategoryName, string CategoryType)
        {
            try
            {
                if (CategoryName.Length <= 0) throw new SystemException("【客製化類別】不能為空!");
                if (CategoryType.Length <= 0) throw new SystemException("【是否開放多選】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認餐點資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RestaurantId
                                FROM EBP.RestaurantMeal a
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐點資訊有誤!");

                        foreach (var item in result)
                        {
                            RestaurantId = item.RestaurantId;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.MealCategory SET
                                MealId = @MealId,
                                RestaurantId = @RestaurantId,
                                CategoryName = @CategoryName,
                                CategoryType = @CategoryType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MealCgId = @MealCgId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MealId,
                                RestaurantId,
                                CategoryName,
                                CategoryType,
                                LastModifiedDate,
                                LastModifiedBy,
                                MealCgId
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

        #region //UpdateMcDetail -- 更新餐點選項(細節)資料 -- Yi 2023.05.05
        public string UpdateMcDetail(int McDetailId, int MealCgId, string McDetailName, double McDetailPrice)
        {
            try
            {
                if (McDetailName.Length <= 0) throw new SystemException("【客製化選項】不能為空!");
                if (McDetailPrice < 0) throw new SystemException("【價格】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認類別資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.MealCategory a
                                INNER JOIN EBP.RestaurantMeal b ON b.MealId = a.MealId
                                WHERE MealCgId = @MealCgId";
                        dynamicParameters.Add("MealCgId", MealCgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客製化類別有誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.McDetail SET
                                MealCgId = @MealCgId,
                                McDetailName = @McDetailName,
                                McDetailPrice = @McDetailPrice,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE McDetailId = @McDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MealCgId,
                                McDetailName,
                                McDetailPrice,
                                LastModifiedDate,
                                LastModifiedBy,
                                McDetailId
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

        #region //UpdateCalendarDate -- 訂餐開放時段資料更新 -- Yi 2023.11.13
        public string UpdateCalendarDate(string StartTime, string EndTime, string OrderType)
        {
            try
            {
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

                DateTime NewStartDate = Convert.ToDateTime(currentDate + " " + StartTime);
                DateTime NewEndDate = Convert.ToDateTime(currentDate + " " + EndTime);

                #region //判斷餐點資料長度
                if (StartTime.Length <= 4) throw new SystemException("【餐廳】長度錯誤!");
                if (EndTime.Length <= 4) throw new SystemException("【餐廳】長度錯誤!");
                if (OrderType.Length <= 0) throw new SystemException("【設定類型】不能為空!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂餐開放時段資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.OrderSetting
                                WHERE OrderType = @OrderType";
                        dynamicParameters.Add("OrderType", OrderType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("開放時段資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.OrderSetting SET
                                StartTime = @StartTime,
                                EndTime = @EndTime
                                WHERE OrderType = @OrderType";
                        var parametersObject = new
                        {
                            StartTime = NewStartDate,
                            EndTime = NewEndDate,
                            OrderType
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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
        #region //DeleteRestaurantMeal -- 刪除餐點資料 -- Yi 2023-05-24
        public string DeleteRestaurantMeal(int MealId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷餐點資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RestaurantMeal a
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐點資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.McDetail a
                                INNER JOIN EBP.MealCategory b ON a.MealCgId = b.MealCgId
                                WHERE b.MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM EBP.MealCategory
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

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

        #region //DeleteMealCategory -- 刪除餐點類別資料 -- Yi 2023.05.05
        public string DeleteMealCategory(int MealCgId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷餐點類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.MealCategory a
                                WHERE MealCgId = @MealCgId";
                        dynamicParameters.Add("MealCgId", MealCgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客製化類別資料錯誤!");
                        #endregion

                        #region //刪除類別table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.MealCategory
                                WHERE MealCgId = @MealCgId";
                        dynamicParameters.Add("MealCgId", MealCgId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteMcDetail -- 刪除餐點選項(細節)資料 -- Yi 2023.05.05
        public string DeleteMcDetail(int McDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷餐點選項(細節)資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.McDetail a
                                WHERE McDetailId = @McDetailId";
                        dynamicParameters.Add("McDetailId", McDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客製化選項資料錯誤!");
                        #endregion

                        #region //刪除細節table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.McDetail
                                WHERE McDetailId = @McDetailId";
                        dynamicParameters.Add("McDetailId", McDetailId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDailyMeal -- 刪除每日餐點資料 -- Ben Ma 2023.11.03
        public string DeleteDailyMeal(int CalendarId, int MealId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷日期資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Calendar
                                WHERE CalendarId = @CalendarId";
                        dynamicParameters.Add("CalendarId", CalendarId);

                        var resultCalendar = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCalendar.Count() <= 0) throw new SystemException("【日期】資料錯誤!");
                        #endregion

                        #region //判斷餐點資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);

                        var resultRestaurantMeal = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRestaurantMeal.Count() <= 0) throw new SystemException("【餐點】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.DailyMeal
                                WHERE CalendarId = @CalendarId
                                AND MealId = @MealId";
                        dynamicParameters.Add("CalendarId", CalendarId);
                        dynamicParameters.Add("MealId", MealId);

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
