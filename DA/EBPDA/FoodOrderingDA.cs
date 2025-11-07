using Dapper;
using Helpers;
using Newtonsoft.Json;
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
    public class FoodOrderingDA
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

        public string StartTime = "17:10:00";
        public string EndTime = "11:00:00";
        public string DiscountStartTime = "17:10:00";
        public string DiscountEndTime = "08:00:00";

        public FoodOrderingDA()
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

        #region //CheckisHoliday
        public bool CheckisHoliday()
        {
            string res = BaseHelper.GetWebRequest(string.Format("https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{0}.json", DateTime.Now.Year));

            List<CalendarData> calendarDatas = JsonConvert.DeserializeObject<List<CalendarData>>(res);

            return calendarDatas.Where(x => x.Date == DateTime.Now.ToString("yyyyMMdd")).Select(x => x.isHoliday).FirstOrDefault();
        }
        #endregion

        //#region //CheckWorkday
        //public string CheckWorkday()
        //{
        //    string res = BaseHelper.GetWebRequest(string.Format("https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{0}.json", DateTime.Now.DayOfYear));

        //    List<CalendarData> calendarDatas = JsonConvert.DeserializeObject<List<CalendarData>>(res);

        //    return calendarDatas.Where(x => x.Date == DateTime.Now.ToString("yyyyMMdd") && x.isHoliday == false).OrderBy(x => x.Date).Select(x => x.Date).FirstOrDefault();
        //}
        //#endregion

        #region //Get

        #region //GetMealInfo -- 取得餐點資訊 -- Yi 2023.05.24
        public string GetMealInfo(string MealDate, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    bool isHoliday = CheckisHoliday();

                    if (!isHoliday)
                    {
                        #region //查詢開放訂餐時間資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                        dynamicParameters.Add("OrderType", 1);
                        var resultOrderSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultOrderSetting.Count() <= 0) throw new SystemException("開放【點餐】時間資料錯誤!");

                        DateTime startTime = default(DateTime), endTime = default(DateTime);
                        foreach (var item in resultOrderSetting)
                        {
                            startTime = Convert.ToDateTime(item.StartTime);
                            endTime = Convert.ToDateTime(item.EndTime);
                        }
                        #endregion

                        //判斷日期，當日日期超過17:10就會跳下一天餐點
                        if (DateTime.Now.TimeOfDay >= endTime.TimeOfDay && DateTime.Now.TimeOfDay < startTime.TimeOfDay)
                        {
                            throw new SystemException("目前時間不開放訂購。");
                        }
                    }

                    sqlQuery.mainKey = "a.MealId,a.CalendarId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", b.RestaurantId, c.RestaurantName, b.MealName, ISNULL(b.MealImage, -1) MealImage, b.MealPrice, b.[Status]
                        , c.StartTime, c.EndTime, FORMAT(d.CalendarDate, 'yyyyMMdd') MealDate";
                    sqlQuery.mainTables =
                        @"FROM EBP.DailyMeal a
                        INNER JOIN EBP.RestaurantMeal b ON a.MealId = b.MealId
                        INNER JOIN EBP.Restaurant c ON b.RestaurantId = c.RestaurantId
                        INNER JOIN BAS.Calendar d ON a.CalendarId = d.CalendarId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND CONVERT(datetime, d.CalendarDate + c.EndTime) > GETDATE()
                                              AND d.CalendarDate = @MealDate";
                    dynamicParameters.Add("MealDate", MealDate);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "d.CalendarDate DESC, b.RestaurantId, b.MealName";
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

        #region //GetMcDetail -- 取得客製化項目資訊 -- Yi 2023.05.25
        public string GetMcDetail(int MealId, string MealDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    bool isHoliday = CheckisHoliday();

                    if (!isHoliday)
                    {
                        #region //查詢開放訂餐時間資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                        dynamicParameters.Add("OrderType", 1);
                        var resultOrderSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultOrderSetting.Count() <= 0) throw new SystemException("開放【點餐】時間資料錯誤!");

                        DateTime startTime = default(DateTime), endTime = default(DateTime);
                        foreach (var item in resultOrderSetting)
                        {
                            startTime = Convert.ToDateTime(item.StartTime);
                            endTime = Convert.ToDateTime(item.EndTime);
                        }
                        #endregion

                        //判斷日期，當日日期超過17:10就會跳下一天餐點
                        if (DateTime.Now.TimeOfDay >= endTime.TimeOfDay && DateTime.Now.TimeOfDay < startTime.TimeOfDay)
                        {
                            throw new SystemException("目前時間不開放訂購。");
                        }
                    }

                    sql = @"SELECT a.MealId, a.MealName, a.MealPrice, ISNULL(a.MealImage, -1) MealImage, b.RestaurantName, b.StartTime, b.EndTime
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
                                WHERE CONVERT(datetime, z.CalendarDate + b.EndTime) > GETDATE()
                                AND z.CalendarDate = @MealDate
                            )";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MealId", @" AND a.MealId = @MealId", MealId);
                    dynamicParameters.Add("MealDate", MealDate);
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

        #region //GetStaffUserId -- 取得事務維護人員名稱(Cmb用) -- Ellie 2023.05.30
        public string GetStaffUserId()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.StewardId, a.UserId, b.UserName, a.StaffUserId, c.UserNo + ' ' + c.UserName StaffUserName
                            FROM EBP.OrderMealSteward a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            INNER JOIN BAS.[User] c ON a.StaffUserId = c.UserId
                            WHERE a.UserId = @UserId
                            ORDER BY a.UserId";
                    dynamicParameters.Add("UserId", CurrentUser);
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

        #region //GetOrderHistory -- 取得點餐紀錄資訊 -- Yi 2023.05.30
        public string GetOrderHistory(int UserId, string UmoDate)
        {
            try
            {
                DateTime tempDate = default(DateTime);
                string avoidDate = "";

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (UserId <= 0) UserId = CurrentUser;
                    if (UmoDate.Length > 0) if (!DateTime.TryParse(UmoDate, out tempDate)) throw new SystemException("日期格式錯誤!");

                    if (UmoDate.Length > 0)
                    {
                        //判斷日期，當日日期超過17:10就會跳下一天餐點(目前餐點)
                        if (Convert.ToDateTime(UmoDate).TimeOfDay >= Convert.ToDateTime(StartTime).TimeOfDay && Convert.ToDateTime(UmoDate).TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                        {
                            UmoDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            UmoDate = DateTime.Now.ToString("yyyy-MM-dd");
                        }
                    }
                    else
                    {
                        //判斷日期，當日日期超過12:00就會跳下一天餐點(歷史餐點)
                        if (DateTime.Now.TimeOfDay >= Convert.ToDateTime("12:00:00").TimeOfDay && DateTime.Now.TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                        {
                            avoidDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            avoidDate = DateTime.Now.ToString("yyyy-MM-dd");
                        }
                    }

                    sql = @"SELECT a.UmoId, a.CompanyId, a.UserId, a.UmoNo
                            , b.UserNo + ' ' + b.UserName UserInfo
                            , FORMAT(a.UmoDate, 'yyyy-MM-dd') UmoDate, RIGHT(DATENAME(weekday, a.UmoDate), 1) AS Weekday
                            , a.UmoDiscount, a.UmoAmount
                            , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                            , (
                                SELECT aa.UmoDetailId, aa.MealId, aa.MealName, aa.UmoDetailRemark
                                , aa.UmoDetailQty, aa.UmoDetailPrice, aa.Pickup, ac.StartTime, ac.EndTime
                                , (
                                    SELECT aaa.McDetailId, aaa.McDetailName, aaa.UmoAdditionalPrice
                                    FROM EBP.UmoAdditional aaa
                                    WHERE aaa.UmoDetailId = aa.UmoDetailId
                                    FOR JSON PATH, ROOT ('data')
                                ) UmoAdditional
                                FROM EBP.UmoDetail aa
                                INNER JOIN EBP.RestaurantMeal ab ON ab.MealId = aa.MealId
                                INNER JOIN EBP.Restaurant ac ON ac.RestaurantId = ab.RestaurantId
                                WHERE aa.UmoId = a.UmoId
                                FOR JSON PATH, ROOT('data')
                            ) UmoDetail
                            , c.ItemQty, FORMAT(d.StartTime, 'yyyy-MM-dd HH:mm:ss') StartTime, FORMAT(d.EndTime, 'yyyy-MM-dd HH:mm:ss') EndTime
                            FROM EBP.UserMealOrder a
                            INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                            OUTER APPLY(
                                SELECT SUM(ca.UmoDetailQty) ItemQty
                                FROM EBP.UmoDetail ca
                                WHERE ca.UmoId = a.UmoId
                            ) c
                            OUTER APPLY (
                                SELECT MAX(dx.StartTime) StartTime, MIN(dx.EndTime) EndTime
                                FROM (
                                    SELECT CONVERT(datetime, 
                                    CASE WHEN DATEADD(day, 0, CONVERT(date, GETDATE())) >= a.UmoDate
                                        THEN DATEADD(day, -1, a.UmoDate) 
                                        ELSE DATEADD(day, 0, CONVERT(date, GETDATE())) END
                                    + db.StartTime) StartTime
                                    , CONVERT(datetime, a.UmoDate + db.EndTime) EndTime
                                    FROM EBP.RestaurantMeal da
                                    INNER JOIN EBP.Restaurant db ON da.RestaurantId = db.RestaurantId
                                    WHERE da.MealId IN (
                                        SELECT daa.MealId
                                        FROM EBP.UmoDetail daa
                                        WHERE daa.UmoId = a.UmoId
                                    )
                                ) dx
                            ) d
                            WHERE a.UserId = @UserId";
                    dynamicParameters.Add("UserId", UserId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UmoDate", @" AND FORMAT(a.UmoDate, 'yyyy-MM-dd') >= @UmoDate", UmoDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AvoidDate", @" AND FORMAT(a.UmoDate, 'yyyy-MM-dd') < @AvoidDate", avoidDate);
                    if (avoidDate.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "HistoryDate", @" AND FORMAT(a.UmoDate, 'yyyy-MM-dd') >= @HistoryDate", DateTime.Now.AddDays(-31).ToString("yyyy-MM-dd 00:00:00"));
                        sql += " ORDER BY a.UmoDate DESC, a.UmoId DESC";
                    }
                    else
                    {
                        sql += " ORDER BY a.UmoDate, a.UmoId";
                    }
                    

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

        #region //GetMealDate -- 取得每日餐點資料 -- Yi 2023.08.16
        public string GetMealDate()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT FORMAT(d.CalendarDate, 'yyyy-MM-dd') CalendarDate
                            , FORMAT(d.CalendarDate, 'yyyy-MM-dd') + '(' + RIGHT(DATENAME(weekday, d.CalendarDate), 1) + ')' AS Weekday
                            FROM EBP.DailyMeal a
                            INNER JOIN EBP.RestaurantMeal b ON a.MealId = b.MealId
                            INNER JOIN EBP.Restaurant c ON b.RestaurantId = c.RestaurantId
                            INNER JOIN BAS.Calendar d ON a.CalendarId = d.CalendarId
                            WHERE CONVERT(datetime, d.CalendarDate + c.EndTime) > GETDATE()
                            ORDER BY FORMAT(d.CalendarDate, 'yyyy-MM-dd')";
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

        #region //GetUserOrderHistory -- 取得個人餐點資料 -- Ben Ma 2023.08.25
        public string GetUserOrderHistory(string UserNo, string UmoDate)
        {
            try
            {
                if (UserNo.Length <= 0) throw new SystemException("【使用者工號】不能為空!");
                if (UmoDate.Length <= 0) throw new SystemException("【餐點日期】不能為空!");
                if (!DateTime.TryParse(UmoDate, out DateTime tempDate)) throw new SystemException("【餐點日期】格式錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 1
                            FROM EBP.UserMealOrder a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            WHERE b.UserNo = @UserNo
                            AND FORMAT(a.UmoDate, 'yyyy-MM-dd') = @UmoDate";
                    dynamicParameters.Add("UserNo", UserNo);
                    dynamicParameters.Add("UmoDate", UmoDate);

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

        #region //CheckWorkday -- 取得工作日資訊 -- Yi 2023.11.13
        public string CheckWorkday()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    string res = BaseHelper.GetWebRequest(string.Format("https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{0}.json", DateTime.Now.Year));

                    List<CalendarData> calendarDatas = JsonConvert.DeserializeObject<List<CalendarData>>(res);

                    string date = calendarDatas.Where(x => DateTime.ParseExact(x.Date, "yyyyMMdd", null) > DateTime.Now && x.isHoliday == false).OrderBy(x => x.Date).Select(x => x.Date).FirstOrDefault();
                    date = DateTime.ParseExact(date, "yyyyMMdd", null).ToString("yyyy-MM-dd");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = date
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

        #region //GetOrderTime -- 取得可點餐時間資料 -- Yi 2023.11.14
        public string GetOrderTime(string OrderType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT FORMAT(a.StartTime, 'HH:mm:ss') StartTime, FORMAT(a.EndTime, 'HH:mm:ss') EndTime
                            FROM EBP.OrderSetting a
                            WHERE a.OrderType = @OrderType";
                    dynamicParameters.Add("OrderType", 1);

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
        #region //AddMealCart -- 購物車資料新增 -- Yi 2023.05.29
        public string AddMealCart(string MealCart)
        {
            try
            {
                if (!MealCart.TryParseJson(out JObject tempJObject)) throw new SystemException("購物車資料格式錯誤");

                JObject cartJson = JObject.Parse(MealCart);

                #region //判斷訂餐時間
                string currentDate = cartJson["mealDate"].ToString();
                int discount = 0;

                if (DateTime.Now.Date > Convert.ToDateTime(currentDate).Date)
                {
                    throw new SystemException("購物車內餐點已過期。");
                }
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DateTime startTime = default(DateTime), endTime = default(DateTime);
                        bool isHoliday = CheckisHoliday();

                        if (!isHoliday)
                        {
                            #region //查詢開放訂餐時間資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                            dynamicParameters.Add("OrderType", 1);
                            var resultOrderSetting1 = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOrderSetting1.Count() <= 0) throw new SystemException("開放【點餐】時間資料錯誤!");

                            foreach (var item in resultOrderSetting1)
                            {
                                startTime = Convert.ToDateTime(item.StartTime);
                                endTime = Convert.ToDateTime(item.EndTime);
                            }
                            #endregion

                            //判斷日期，當日日期超過17:10就會跳下一天餐點
                            if (DateTime.Now.TimeOfDay >= endTime.TimeOfDay && DateTime.Now.TimeOfDay < startTime.TimeOfDay)
                            {
                                throw new SystemException("目前時間不開放訂購。");
                            }
                        }

                        #region //判斷使用者是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.UserId, b.CompanyId
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE a.UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", cartJson["userNo"].ToString());

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        int UserId = -1;
                        int companyId = -1;
                        foreach (var item in result)
                        {
                            UserId = item.UserId;
                            companyId = item.CompanyId;
                        }
                        #endregion

                        #region //判斷訂餐時間是否符合優惠

                        if (!isHoliday)
                        {
                            #region //查詢開放訂餐時間資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StartTime, a.EndTime
                                    FROM EBP.OrderSetting a
                                    WHERE a.OrderType = @OrderType";
                            dynamicParameters.Add("OrderType", 2);
                            var resultOrderSetting2 = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOrderSetting2.Count() <= 0) throw new SystemException("開放訂餐【優惠時段】時間資料錯誤!");

                            foreach (var item in resultOrderSetting2)
                            {
                                startTime = item.StartTime;
                                endTime = item.EndTime;
                            }
                            #endregion

                            if (DateTime.Now.TimeOfDay >= Convert.ToDateTime(startTime).TimeOfDay && DateTime.Now.TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                            {
                                discount = 10;
                            }

                            if (DateTime.Now.TimeOfDay >= Convert.ToDateTime("00:00:00").TimeOfDay && DateTime.Now.TimeOfDay < Convert.ToDateTime(endTime).TimeOfDay)
                            {
                                discount = 10;
                            }
                        }
                        else
                        {
                            discount = 10;
                        }
                        #endregion

                        #region //判斷餐點日期今天是否已有餐點
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.UserMealOrder a
                                WHERE a.UserId = @UserId
                                AND FORMAT(a.UmoDate, 'yyyy-MM-dd') = @UmoDate";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("UmoDate", currentDate);

                        var resultExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExist.Count() > 0) discount = 0;
                        #endregion

                        int totalPrice = 0;

                        for (int i = 0; i < cartJson["mealInfo"].Count(); i++)
                        {
                            #region //判斷餐點資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT d.CalendarDate, c.StartTime, c.EndTime, b.MealName, b.MealPrice
                                    , CONVERT(datetime, 
                                    CASE WHEN DATEADD(day, 0, CONVERT(date, GETDATE())) >= d.CalendarDate
                                         THEN DATEADD(day, -1, d.CalendarDate) 
                                         ELSE DATEADD(day, 0, CONVERT(date, GETDATE())) END
                                     + c.StartTime)
                                    , CONVERT(datetime, d.CalendarDate + c.EndTime)
                                    FROM EBP.DailyMeal a
                                    INNER JOIN EBP.RestaurantMeal b ON a.MealId = b.MealId
                                    INNER JOIN EBP.Restaurant c ON b.RestaurantId = c.RestaurantId
                                    INNER JOIN BAS.Calendar d ON a.CalendarId = d.CalendarId
                                    WHERE CONVERT(datetime, d.CalendarDate + c.EndTime) > DATEADD(day, 0, GETDATE())
                                    AND a.MealId = @MealId
                                    AND d.CalendarDate = @CurrentDate
                                    AND b.MealPrice = @MealPrice
                                    ORDER BY d.CalendarDate DESC";
                            dynamicParameters.Add("MealId", cartJson["mealInfo"][i]["mealId"].ToString());
                            dynamicParameters.Add("CurrentDate", currentDate);
                            dynamicParameters.Add("MealPrice", cartJson["mealInfo"][i]["price"].ToString());

                            var resultMeal = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMeal.Count() <= 0) throw new SystemException("餐點資料已過期或不存在!");
                            #endregion

                            //計算訂單總額
                            totalPrice += Convert.ToInt32(cartJson["mealInfo"][i]["price"]) * Convert.ToInt32(cartJson["mealInfo"][i]["qty"]);

                            for (int j = 0; j < cartJson["mealInfo"][i]["categoryInfo"].Count(); j++)
                            {
                                #region //判斷餐點客製化類別是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM EBP.MealCategory
                                    WHERE MealCgId = @MealCgId";
                                dynamicParameters.Add("MealCgId", cartJson["mealInfo"][i]["categoryInfo"][j]["mealCgId"].ToString());

                                var resultCategory = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCategory.Count() <= 0) throw new SystemException("餐點客製化【類別】資料錯誤!");
                                #endregion

                                for (int k = 0; k < cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"].Count(); k++)
                                {
                                    #region //判斷餐點客製化項目是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                    FROM EBP.McDetail
                                    WHERE McDetailId = @McDetailId
                                    AND McDetailPrice = @McDetailPrice";
                                    dynamicParameters.Add("McDetailId", cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailId"].ToString());
                                    dynamicParameters.Add("McDetailPrice", cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailPrice"].ToString());

                                    var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultDetail.Count() <= 0) throw new SystemException("餐點客製化【項目】資料錯誤!");
                                    #endregion

                                    totalPrice += Convert.ToInt32(cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailPrice"]) * Convert.ToInt32(cartJson["mealInfo"][i]["qty"]);
                                }
                            }
                        }

                        #region //訂單編號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(UmoNo), '000'), 3)) + 1 CurrentNum
                            FROM EBP.UserMealOrder
                            WHERE UmoNo LIKE @UmoNo";
                        dynamicParameters.Add("UmoNo", string.Format("{0}{1}___", "FD", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string umoNo = string.Format("{0}{1}{2}", "FD", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        int rowsAffected = 0;

                        #region //新增餐點訂單
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.UserMealOrder (CompanyId, UserId, UmoDate, UmoNo, UmoDiscount, UmoAmount
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.UmoId
                                    VALUES (@CompanyId, @UserId, @UmoDate, @UmoNo, @UmoDiscount, @UmoAmount
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = companyId,
                                UserId,
                                UmoDate = currentDate,
                                UmoNo = umoNo,
                                UmoDiscount = discount,
                                UmoAmount = totalPrice,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int umoId = -1;
                        foreach (var item in insertResult)
                        {
                            umoId = Convert.ToInt32(item.UmoId);
                        }
                        #endregion

                        for (int i = 0; i < cartJson["mealInfo"].Count(); i++)
                        {
                            int price = Convert.ToInt32(cartJson["mealInfo"][i]["price"]);
                            JToken meal = cartJson["mealInfo"][i];
                            int qty = Convert.ToInt32(cartJson["mealInfo"][i]["qty"]);

                            #region //新增餐點明細
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
                                    MealId = Convert.ToInt32(cartJson["mealInfo"][i]["mealId"]),
                                    MealName = cartJson["mealInfo"][i]["mealName"].ToString(),
                                    UmoDetailRemark = cartJson["mealInfo"][i]["remark"].ToString(),
                                    UmoDetailQty = Convert.ToInt32(cartJson["mealInfo"][i]["qty"]),
                                    UmoDetailPrice = Caculate(price, meal) * qty,
                                    Pickup = cartJson["mealInfo"][i]["pickup"].ToString(),
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

                            if (cartJson["mealInfo"][i]["categoryInfo"] != null)
                            {
                                for (int j = 0; j < cartJson["mealInfo"][i]["categoryInfo"].Count(); j++)
                                {
                                    for (int k = 0; k < cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"].Count(); k++)
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
                                                McDetailId = Convert.ToInt32(cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailId"]),
                                                McDetailName = cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailName"].ToString(),
                                                UmoAdditionalPrice = Convert.ToInt32(cartJson["mealInfo"][i]["categoryInfo"][j]["mcDetailInfo"][k]["mcDetailPrice"]),
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
                                UserNo = cartJson["userNo"].ToString(),
                                UmoNo = umoNo,
                                MotionDate = currentDate,
                                MotionRemark = "前台—【個人購物車】資料【新增】",
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

                    if (meal != null)
                    {
                        for (int x = 0; x < meal["categoryInfo"].Count(); x++)
                        {
                            for (int y = 0; y < meal["categoryInfo"][x]["mcDetailInfo"].Count(); y++)
                            {
                                totalprice += Convert.ToInt32(meal["categoryInfo"][x]["mcDetailInfo"][y]["mcDetailPrice"]);
                            }
                        }
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

        #region //AddLotCart -- 批量購物車資料新增 -- Ellie 2023.06.07
        public string AddLotCart(string LotCart)
        {
            try
            {
                if (!LotCart.TryParseJson(out JObject tempJObject)) throw new SystemException("批量購物車資料格式錯誤");

                JObject cartJson = JObject.Parse(LotCart);

                #region //判斷訂餐時間
                string currentDate = cartJson["mealDate"].ToString();
                int discount = 0;

                if (DateTime.Now.Date > Convert.ToDateTime(currentDate).Date)
                {
                    throw new SystemException("購物車內餐點已過期。");
                }
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        List<int> userList = new List<int>();

                        for (int i = 0; i < cartJson["userInfo"].Count(); i++)
                        {
                            DateTime startTime = default(DateTime), endTime = default(DateTime);
                            bool isHoliday = CheckisHoliday();

                            if (!isHoliday)
                            {
                                #region //查詢開放訂餐時間資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                                dynamicParameters.Add("OrderType", 1);
                                var resultOrderSetting1 = sqlConnection.Query(sql, dynamicParameters);
                                if (resultOrderSetting1.Count() <= 0) throw new SystemException("開放【點餐】時間資料錯誤!");

                                foreach (var item in resultOrderSetting1)
                                {
                                    startTime = Convert.ToDateTime(item.StartTime);
                                    endTime = Convert.ToDateTime(item.EndTime);
                                }
                                #endregion

                                //判斷日期，當日日期超過17:10就會跳下一天餐點
                                if (DateTime.Now.TimeOfDay >= endTime.TimeOfDay && DateTime.Now.TimeOfDay < startTime.TimeOfDay)
                                {
                                    throw new SystemException("目前時間不開放訂購。");
                                }
                            }

                            #region //判斷使用者是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                    INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                    WHERE a.UserId = @UserId
                                    AND c.CompanyId = @CompanyId";
                            dynamicParameters.Add("UserId", cartJson["userInfo"][i]["userId"].ToString());
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                            userList.Add(Convert.ToInt32(cartJson["userInfo"][i]["userId"]));
                            #endregion

                            #region //判斷訂餐時間是否符合優惠

                            if (!isHoliday)
                            {
                                #region //查詢開放訂餐時間資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.StartTime, a.EndTime
                                        FROM EBP.OrderSetting a
                                        WHERE a.OrderType = @OrderType";
                                dynamicParameters.Add("OrderType", 2);
                                var resultOrderSetting2 = sqlConnection.Query(sql, dynamicParameters);
                                if (resultOrderSetting2.Count() <= 0) throw new SystemException("開放訂餐【優惠時段】時間資料錯誤!");

                                foreach (var item in resultOrderSetting2)
                                {
                                    startTime = item.StartTime;
                                    endTime = item.EndTime;
                                }
                                #endregion

                                if (DateTime.Now.TimeOfDay >= Convert.ToDateTime(startTime).TimeOfDay && DateTime.Now.TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                                {
                                    discount = 10;
                                }

                                if (DateTime.Now.TimeOfDay >= Convert.ToDateTime("00:00:00").TimeOfDay && DateTime.Now.TimeOfDay < Convert.ToDateTime(endTime).TimeOfDay)
                                {
                                    discount = 10;
                                }
                            }
                            else
                            {
                                discount = 10;
                            }
                            #endregion

                            #region //判斷餐點日期今天是否已有餐點
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM EBP.UserMealOrder a
                                WHERE a.UserId = @UserId
                                AND FORMAT(a.UmoDate, 'yyyy-MM-dd') = @UmoDate";
                            dynamicParameters.Add("UserId", cartJson["userInfo"][i]["userId"].ToString());
                            dynamicParameters.Add("UmoDate", currentDate);

                            var resultExist = sqlConnection.Query(sql, dynamicParameters);
                            if (resultExist.Count() > 0) discount = 0;
                            #endregion
                        }

                        #region //判斷是否為該事務負責人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(a.StaffUserId) TotalStaffUser
                                FROM EBP.OrderMealSteward a
                                WHERE a.UserId = @UserId
                                AND a.StaffUserId IN @UserList";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("UserList", userList.ToArray());

                        var resultOrderMealSteward = sqlConnection.Query(sql, dynamicParameters);

                        int TotalStaffUser = 0;
                        foreach (var item in resultOrderMealSteward)
                        {
                            TotalStaffUser = Convert.ToInt32(item.TotalStaffUser);
                        }

                        if (TotalStaffUser != userList.Count) throw new SystemException("部分使用者不是該事務管轄!");
                        #endregion

                        int rowsAffected = 0;

                        for (int i = 0; i < cartJson["userInfo"].Count(); i++)
                        {
                            int totalPrice = 0;

                            for (int j = 0; j < cartJson["userInfo"][i]["mealInfo"].Count(); j++)
                            {
                                #region //判斷餐點資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT d.CalendarDate, c.StartTime, c.EndTime, b.MealName, b.MealPrice
                                        , CONVERT(datetime, 
                                        CASE WHEN DATEADD(day, 0, CONVERT(date, GETDATE())) >= d.CalendarDate
                                             THEN DATEADD(day, -1, d.CalendarDate) 
                                             ELSE DATEADD(day, 0, CONVERT(date, GETDATE())) END
                                         + c.StartTime)
                                        , CONVERT(datetime, d.CalendarDate + c.EndTime)
                                        FROM EBP.DailyMeal a
                                        INNER JOIN EBP.RestaurantMeal b ON a.MealId = b.MealId
                                        INNER JOIN EBP.Restaurant c ON b.RestaurantId = c.RestaurantId
                                        INNER JOIN BAS.Calendar d ON a.CalendarId = d.CalendarId
                                        WHERE CONVERT(datetime, d.CalendarDate + c.EndTime) > DATEADD(day, 0, GETDATE())
                                        AND a.MealId = @MealId
                                        AND d.CalendarDate = @CurrentDate
                                        AND b.MealPrice = @MealPrice
                                        ORDER BY d.CalendarDate DESC";
                                dynamicParameters.Add("MealId", cartJson["userInfo"][i]["mealInfo"][j]["mealId"].ToString());
                                dynamicParameters.Add("CurrentDate", currentDate);
                                dynamicParameters.Add("MealPrice", cartJson["userInfo"][i]["mealInfo"][j]["price"].ToString());

                                var resultMeal = sqlConnection.Query(sql, dynamicParameters);
                                if (resultMeal.Count() <= 0) throw new SystemException("餐點資料已過期或不存在!");
                                #endregion

                                //計算訂單總額
                                totalPrice += Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["price"]) * Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["qty"]);

                                for (int k = 0; k < cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"].Count(); k++)
                                {
                                    #region //判斷餐點客製化類別是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM EBP.MealCategory
                                            WHERE MealCgId = @MealCgId";
                                    dynamicParameters.Add("MealCgId", cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mealCgId"].ToString());

                                    var resultCategory = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultCategory.Count() <= 0) throw new SystemException("餐點客製化【類別】資料錯誤!");
                                    #endregion

                                    for (int l = 0; l < cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"].Count(); l++)
                                    {
                                        #region //判斷餐點客製化項目是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM EBP.McDetail
                                                WHERE McDetailId = @McDetailId
                                                AND McDetailPrice = @McDetailPrice";
                                        dynamicParameters.Add("McDetailId", cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailId"].ToString());
                                        dynamicParameters.Add("McDetailPrice", cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailPrice"].ToString());

                                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultDetail.Count() <= 0) throw new SystemException("餐點客製化【項目】資料錯誤!");
                                        #endregion

                                        totalPrice += Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailPrice"]) * Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["qty"]);
                                    }
                                }
                            }

                            #region //訂單編號自動取號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(UmoNo), '000'), 3)) + 1 CurrentNum
                                    FROM EBP.UserMealOrder
                                    WHERE UmoNo LIKE @UmoNo";
                            dynamicParameters.Add("UmoNo", string.Format("{0}{1}___", "FD", DateTime.Now.ToString("yyyyMMdd")));
                            int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                            string umoNo = string.Format("{0}{1}{2}", "FD", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                            #endregion                            

                            #region //新增餐點訂單
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.UserMealOrder (CompanyId, UserId, UmoDate, UmoNo, UmoDiscount, UmoAmount
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.UmoId
                                    VALUES (@CompanyId, @UserId, @UmoDate, @UmoNo, @UmoDiscount, @UmoAmount
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    UserId = Convert.ToInt32(cartJson["userInfo"][i]["userId"]),
                                    UmoDate = currentDate,
                                    UmoNo = umoNo,
                                    UmoDiscount = discount,
                                    UmoAmount = totalPrice,
                                    UmrDate = currentDate,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

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
                            dynamicParameters.Add("UserId", cartJson["userInfo"][i]["userId"].ToString());

                            var resultoUserNo = sqlConnection.Query(sql, dynamicParameters);

                            string userNo = "";
                            foreach (var item in resultoUserNo)
                            {
                                userNo = item.UserNo;
                            }
                            #endregion

                            for (int j = 0; j < cartJson["userInfo"][i]["mealInfo"].Count(); j++)
                            {
                                int price = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["price"]);
                                JToken meal = cartJson["userInfo"][i]["mealInfo"][j];
                                int qty = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["qty"]);

                                #region //新增餐點明細
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
                                        MealId = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["mealId"]),
                                        MealName = cartJson["userInfo"][i]["mealInfo"][j]["mealName"].ToString(),
                                        UmoDetailRemark = cartJson["userInfo"][i]["mealInfo"][j]["remark"].ToString(),
                                        UmoDetailQty = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["qty"]),
                                        UmoDetailPrice = Caculate(price, meal) * qty,
                                        Pickup = cartJson["userInfo"][i]["mealInfo"][j]["pickup"].ToString(),
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

                                if (cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"] != null)
                                {
                                    for (int k = 0; k < cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"].Count(); k++)
                                    {
                                        for (int l = 0; l < cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"].Count(); l++)
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
                                                    McDetailId = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailId"]),
                                                    McDetailName = cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailName"].ToString(),
                                                    UmoAdditionalPrice = Convert.ToInt32(cartJson["userInfo"][i]["mealInfo"][j]["categoryInfo"][k]["mcDetailInfo"][l]["mcDetailPrice"]),
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
                                    MotionRemark = "前台—【批量購物車】資料【新增】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
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

                #region //計算餐點費用總計
                int Caculate(int unitPrice, JToken meal)
                {
                    int totalprice = unitPrice;

                    if (meal != null)
                    {
                        for (int x = 0; x < meal["categoryInfo"].Count(); x++)
                        {
                            for (int y = 0; y < meal["categoryInfo"][x]["mcDetailInfo"].Count(); y++)
                            {
                                totalprice += Convert.ToInt32(meal["categoryInfo"][x]["mcDetailInfo"][y]["mcDetailPrice"]);
                            }
                        }
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
        #endregion

        #region //Delete
        #region //DeleteOrderHistory -- 單人點餐紀錄刪除 -- Yi 2023.06.02
        public string DeleteOrderHistory(int UmoId)
        {
            try
            {
                #region //判斷訂餐時間
                int discount = 0;
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        bool isHoliday = CheckisHoliday();

                        if (!isHoliday)
                        {
                            #region //查詢開放訂餐時間資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                            dynamicParameters.Add("OrderType", 1);
                            var resultOrderSetting = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOrderSetting.Count() <= 0) throw new SystemException("開放【點餐】時間資料錯誤!");

                            DateTime startTime = default(DateTime), endTime = default(DateTime);
                            foreach (var item in resultOrderSetting)
                            {
                                startTime = Convert.ToDateTime(item.StartTime);
                                endTime = Convert.ToDateTime(item.EndTime);
                            }
                            #endregion

                            //判斷日期，當日日期超過17:10就會跳下一天餐點
                            if (DateTime.Now.TimeOfDay >= endTime.TimeOfDay && DateTime.Now.TimeOfDay < startTime.TimeOfDay)
                            {
                                throw new SystemException("目前時間不開放取消訂餐。");
                            }
                        }

                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UmoId, UserId, UmoNo, FORMAT(UmoDate, 'yyyy-MM-dd') UmoDate
                                FROM EBP.UserMealOrder
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        int umoId = -1;
                        int userId = -1;
                        string umoNo = "";
                        string umoDate = "";
                        foreach (var item in result)
                        {
                            umoId = item.UmoId;
                            userId = item.UserId;
                            umoNo = item.UmoNo;
                            umoDate = item.UmoDate;
                        }
                        #endregion

                        #region //判斷是否為該事務負責人員
                        if (userId != CurrentUser)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EBP.OrderMealSteward a
                                    WHERE a.UserId = @UserId
                                    AND a.StaffUserId = @userId";
                            dynamicParameters.Add("UserId", CurrentUser);
                            dynamicParameters.Add("userId", userId);

                            var resultOrderMealSteward = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOrderMealSteward.Count() <= 0) throw new SystemException("無權限可取消該人員之訂餐紀錄!");
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

                        #region //判斷餐點日期是否已有餐點
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.UmoId, a.CreateDate
                                FROM EBP.UserMealOrder a
                                WHERE a.UserId = @UserId
                                AND FORMAT(a.UmoDate, 'yyyy-MM-dd') = @UmoDate
                                AND a.UmoId != @UmoId
                                ORDER BY a.UmoId";
                        dynamicParameters.Add("UserId", userId);
                        dynamicParameters.Add("UmoDate", umoDate);
                        dynamicParameters.Add("UmoId", umoId); //原本欲刪除

                        var resultExist = sqlConnection.Query(sql, dynamicParameters);

                        int updUmoId = -1;
                        DateTime createDate = default(DateTime);
                        foreach (var item in resultExist)
                        {
                            createDate = Convert.ToDateTime(item.CreateDate);
                            updUmoId = Convert.ToInt32(item.UmoId);
                        }
                        #endregion

                        #region //卡控當日餐點若有兩筆，且在折扣時間內點餐；即取消折扣訂單時，須更新第二筆訂單為符合折扣
                        if (resultExist.Count() > 0)
                        {
                            #region //判斷訂餐時間是否符合優惠

                            #region //查詢開放訂餐時間資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StartTime, a.EndTime
                                FROM EBP.OrderSetting a
                                WHERE a.OrderType = @OrderType";
                            dynamicParameters.Add("OrderType", 2);
                            var resultOrderSetting = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOrderSetting.Count() <= 0) throw new SystemException("開放訂餐【優惠時段】時間資料錯誤!");

                            string StartTime = "", EndTime = "";
                            foreach (var item in resultOrderSetting)
                            {
                                StartTime = Convert.ToDateTime(item.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                                EndTime = Convert.ToDateTime(item.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            #endregion

                            if (createDate.TimeOfDay >= Convert.ToDateTime(StartTime).TimeOfDay && createDate.TimeOfDay <= Convert.ToDateTime("23:59:59").TimeOfDay)
                            {
                                discount = 10;
                            }

                            if (createDate.TimeOfDay >= Convert.ToDateTime("00:00:00").TimeOfDay && createDate.TimeOfDay < Convert.ToDateTime(EndTime).TimeOfDay)
                            {
                                discount = 10;
                            }
                            #endregion

                            #region //更新訂單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EBP.UserMealOrder SET
                                    UmoDiscount = @UmoDiscount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UmoId = @UmoId";
                            var parametersObject = new
                            {
                                UmoDiscount = discount,
                                LastModifiedDate,
                                LastModifiedBy,
                                UmoId = updUmoId
                            };
                            dynamicParameters.AddDynamicParams(parametersObject);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //判斷該餐廳若超過結單時間，不予取消餐點
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM (
                                    SELECT CONVERT(datetime, 
                                        CASE WHEN DATEADD(day, 0, CONVERT(date, GETDATE())) >= c.UmoDate
                                            THEN DATEADD(day, -1, c.UmoDate) 
                                            ELSE DATEADD(day, 0, CONVERT(date, GETDATE())) END
                                        + b.StartTime) StartTime
                                        , CONVERT(datetime, c.UmoDate + b.EndTime) EndTime
                                    FROM EBP.RestaurantMeal a
                                    INNER JOIN EBP.Restaurant b ON b.RestaurantId = a.RestaurantId
                                    OUTER APPLY (
                                        SELECT ca.UmoId, ca.UmoDate
                                        FROM EBP.UserMealOrder ca
                                        WHERE ca.UmoId = @UmoId
                                    ) c
                                    WHERE a.MealId IN (
                                        SELECT aa.MealId
                                        FROM EBP.UmoDetail aa
                                        WHERE aa.UmoId = c.UmoId
                                    )
                                ) d
                                HAVING GETDATE() < MIN(d.EndTime)";
                        dynamicParameters.Add("UmoId", UmoId);

                        var resultMeal = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMeal.Count() <= 0) throw new SystemException("餐點已結單，無法取消餐點!");
                        #endregion

                        #region //刪除關聯table

                        #region //加購資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.UmoAdditional a
                                INNER JOIN EBP.UmoDetail b ON a.UmoDetailId = b.UmoDetailId
                                WHERE b.UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //餐點明細資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.UmoDetail
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
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
                                MotionDate = umoDate,
                                MotionRemark = "前台—【點餐紀錄】資料【刪除】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.UserMealOrder
                                WHERE UmoId = @UmoId";
                        dynamicParameters.Add("UmoId", UmoId);

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
