using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace EBPDA
{
    public class ActivityDA
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

        public ActivityDA()
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
        #region //GetActivity -- 取得活動資料 -- Ben Ma 2023.01.07
        public string GetActivity(int ActivityId, int AnnualId, string ActivityName, string StartDate, string EndDate, string Status
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ActivityId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ActivityName, a.AnnualId, FORMAT(a.ActivityStartDate, 'yyyy-MM-dd') ActivityStartDate
                        , FORMAT(a.ActivityEndDate, 'yyyy-MM-dd') ActivityEndDate, a.ActivityContent, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EBP.Activity a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityId", @" AND a.ActivityId = @ActivityId", ActivityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnualId", @" AND a.AnnualId = @AnnualId", AnnualId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityName", @" AND a.ActivityName LIKE '%' + @ActivityName + '%'", ActivityName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.ActivityStartDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.ActivityStartDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ActivityStartDate DESC, a.ActivityId";
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

        #region //GetAwards -- 取得活動獎項資料 -- Ben Ma 2023.01.07
        public string GetAwards(int AwardsId, int ActivityId, string AwardsName
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.AwardsId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ActivityId, a.AwardsName, ISNULL(a.AwardsPrice, 0) AwardsPrice, ISNULL(a.AwardsQuantity, 0) AwardsQuantity
                        , ISNULL(a.AwardsUnit, '') AwardsUnit, a.AwardsRemark, a.SortNumber, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EBP.Awards a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AwardsId", @" AND a.AwardsId = @AwardsId", AwardsId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityId", @" AND a.ActivityId = @ActivityId", ActivityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AwardsName", @" AND a.AwardsName LIKE '%' + @AwardsName + '%'", AwardsName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region //GetAwardsUser -- 取得獎項人員資料 -- Ben Ma 2023.01.07
        public string GetAwardsUser(int AwardsId, string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.GroupName
                            , (
                                SELECT aa.UserId, ISNULL(aa.Remark, '') Remark
                                , ab.UserNo, ab.UserName, ab.Gender
                                , ac.DepartmentNo, ac.DepartmentName
                                , ad.LogoIcon
                                FROM EBP.AwardsUser aa
                                INNER JOIN BAS.[User] ab ON aa.UserId = ab.UserId
                                INNER JOIN BAS.Department ac ON ab.DepartmentId = ac.DepartmentId
                                INNER JOIN BAS.Company ad ON ac.CompanyId = ad.CompanyId
                                WHERE aa.AwardsId = a.AwardsId
                                AND aa.GroupName = a.GroupName";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND (ab.UserNo LIKE '%' + @SearchKey + '%' OR ab.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sql += @"
                                FOR JSON PATH, ROOT('data')
                            ) AwardsUser
                            FROM EBP.AwardsUser a
                            WHERE a.AwardsId = @AwardsId
                            AND a.UserId IN (
                                SELECT aa.UserId
                                FROM EBP.AwardsUser aa
                                INNER JOIN BAS.[User] ab ON aa.UserId = ab.UserId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND (ab.UserNo LIKE '%' + @SearchKey + '%' OR ab.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sql += @"
                            )
                            ORDER BY a.GroupName";
                    dynamicParameters.Add("AwardsId", AwardsId);

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

        #region //GetAwardsKanban -- 取得取得獎項看板資料 -- Ben Ma 2023.01.09
        public string GetAwardsKanban(int ActivityId, int AwardsId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷活動資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM EBP.Activity
                            WHERE ActivityId = @ActivityId
                            AND Status = @Status";
                    dynamicParameters.Add("ActivityId", ActivityId);
                    dynamicParameters.Add("Status", "A");

                    var resultActivity = sqlConnection.Query(sql, dynamicParameters);
                    if (resultActivity.Count() <= 0) throw new SystemException("活動資料錯誤!");
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.AwardsId, a.AwardsName, ISNULL(a.AwardsPrice, 0) AwardsPrice, ISNULL(a.AwardsQuantity, 0) AwardsQuantity
                            , ISNULL(a.AwardsUnit, '') AwardsUnit, a.AwardsRemark
                            , (
                                SELECT DISTINCT aa.GroupName
                                , (
                                    SELECT aaa.UserId, ISNULL(aaa.Remark, '') Remark
                                    , aab.UserNo, aab.UserName, aab.Gender
                                    , aac.DepartmentNo, aac.DepartmentName
                                    , aad.LogoIcon
                                    FROM EBP.AwardsUser aaa
                                    INNER JOIN BAS.[User] aab ON aaa.UserId = aab.UserId
                                    INNER JOIN BAS.Department aac ON aab.DepartmentId = aac.DepartmentId
                                    INNER JOIN BAS.Company aad ON aac.CompanyId = aad.CompanyId
                                    WHERE aaa.AwardsId = aa.AwardsId
                                    AND aaa.GroupName = aa.GroupName
                                    FOR JSON PATH, ROOT('user')
                                ) AwardsUser
                                FROM EBP.AwardsUser aa
                                WHERE aa.AwardsId = a.AwardsId
                                ORDER BY aa.GroupName
                                FOR JSON PATH, ROOT('group')
                            ) AwardsGroup
                            FROM EBP.Awards a
                            INNER JOIN EBP.Activity b ON a.ActivityId = b.ActivityId
                            WHERE a.[Status] = @Status
                            AND b.ActivityId = @ActivityId
                            AND b.[Status] = @Status";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AwardsId", @" AND a.AwardsId = @AwardsId", AwardsId);
                    sql += @" 
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("ActivityId", ActivityId);

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

        #region //GetPoints -- 取得活動積分資料 -- Zoey 2023.02.07
        public string GetPoints(int ActivityPointId, int ActivityId
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ActivityPointId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ActivityId, a.PointType, a.PointQty, a.Remark, a.Status
                          , a.PointType + ' ' + a.Remark + ' (' + CONVERT(varchar, a.PointQty) + '點)' PointDetail";
                    sqlQuery.mainTables =
                        @"FROM EBP.ActivityPoint a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityPointId", @" AND a.ActivityPointId = @ActivityPointId", ActivityPointId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityId", @" AND a.ActivityId = @ActivityId", ActivityId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ActivityPointId";
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
        
        #region //GetParticipants -- 取得參與人員資料 -- Zoey 2023.02.09
        public string GetParticipants(int ActivityId, string SearchKey
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ParticipantId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserId, a.ActivityId, a.Status
                          , b.UserNo, b.UserName, b.Gender
                          , (
                              SELECT y.PointType, y.PointQty, y.Remark
                              FROM EBP.ParticipantPoint x
                              INNER JOIN EBP.ActivityPoint y ON y.ActivityPointId = x.ActivityPointId
                              WHERE x.ParticipantId = a.ParticipantId
                              FOR JSON PATH, ROOT('data')
                          ) PointDetail";
                    sqlQuery.mainTables =
                        @"FROM EBP.Participant a
                          INNER JOIN BAS.[User] b ON b.UserId = a.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ActivityId", @" AND a.ActivityId = @ActivityId", ActivityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ParticipantId";
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
        #region //AddActivity -- 活動資料新增 -- Ben Ma 2023.01.07
        public string AddActivity(int AnnualId, string ActivityName, string ActivityStartDate, string ActivityEndDate, string ActivityContent)
        {
            try
            {
                DateTime tempDate = default(DateTime);

                if (ActivityName.Length <= 0) throw new SystemException("【活動名稱】不能為空!");
                if (ActivityName.Length > 100) throw new SystemException("【活動名稱】長度錯誤!");
                if (ActivityStartDate.Length <= 0) throw new SystemException("【活動開始日期】不能為空!");
                if (!DateTime.TryParse(ActivityStartDate, out tempDate)) throw new SystemException("【活動開始日期】格式錯誤!");
                if (ActivityEndDate.Length <= 0) throw new SystemException("【活動開始日期】不能為空!");
                if (!DateTime.TryParse(ActivityEndDate, out tempDate)) throw new SystemException("【活動結束日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Activity (AnnualId, ActivityName, ActivityStartDate, ActivityEndDate, ActivityContent, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ActivityId
                                VALUES (@AnnualId, @ActivityName, @ActivityStartDate, @ActivityEndDate, @ActivityContent, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AnnualId,
                                ActivityName,
                                ActivityStartDate,
                                ActivityEndDate,
                                ActivityContent,
                                Status = "S", //停用
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

        #region //AddAwards -- 活動獎項資料新增 -- Ben Ma 2023.01.07
        public string AddAwards(int ActivityId, string AwardsName, int AwardsPrice, int AwardsQuantity
            , string AwardsUnit, string AwardsRemark)
        {
            try
            {
                if (AwardsName.Length <= 0) throw new SystemException("【獎項名稱】不能為空!");
                if (AwardsName.Length > 100) throw new SystemException("【獎項名稱】長度錯誤!");
                if (AwardsRemark.Length > 500) throw new SystemException("【獎項備註】長度錯誤!");
                if (AwardsUnit.Length > 50) throw new SystemException("【獎項數量單位】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        #region //判斷獎項名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Awards
                                WHERE AwardsName = @AwardsName
                                AND ActivityId = @ActivityId";
                        dynamicParameters.Add("AwardsName", AwardsName);
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【獎項名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM EBP.Awards";

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Awards (ActivityId, AwardsName, AwardsPrice, AwardsQuantity
                                , AwardsUnit, AwardsRemark, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ActivityId
                                VALUES (@ActivityId, @AwardsName, @AwardsPrice, @AwardsQuantity
                                , @AwardsUnit, @AwardsRemark, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ActivityId,
                                AwardsName,
                                AwardsPrice = AwardsPrice > 0 ? AwardsPrice : (int?)null,
                                AwardsQuantity = AwardsQuantity > 0 ? AwardsQuantity : (int?)null,
                                AwardsUnit,
                                AwardsRemark,
                                SortNumber = maxSort + 1,
                                Status = "S", //停用
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

        #region //AddAwardsUser -- 獎項人員資料新增 -- Ben Ma 2023.01.07
        public string AddAwardsUser(int AwardsId, string GroupName, string Users, string Remark)
        {
            try
            {
                if (GroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (Users.Length <= 0) throw new SystemException("【獎項人員】不能為空!");
                if (Remark.Length > 500) throw new SystemException("【備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動獎項資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動獎項資料錯誤!");
                        #endregion

                        #region //判斷獎項人員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("獎項人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            #region //判斷該獎項是否已經有該人員
                            #region //使用者資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserNo
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));

                            var resultUser = sqlConnection.Query(sql, dynamicParameters);

                            string UserNo = "";
                            foreach (var item in resultUser)
                            {
                                UserNo = item.UserNo;
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EBP.AwardsUser
                                    WHERE AwardsId = @AwardsId
                                    AND UserId = @UserId";
                            dynamicParameters.Add("AwardsId", AwardsId);
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));

                            var resultAwardsUser = sqlConnection.Query(sql, dynamicParameters);
                            if (resultAwardsUser.Count() > 0) throw new SystemException("【" + UserNo + "】已經重複領獎!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.AwardsUser (AwardsId, UserId, GroupName, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@AwardsId, @UserId, @GroupName, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AwardsId,
                                    UserId = Convert.ToInt32(user),
                                    GroupName,
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

        #region //AddPoints -- 活動積分資料新增 -- Zoey 2023.02.07
        public string AddPoints(int ActivityId, string PointType, int PointQty, string Remark)
        {
            try
            {
                if (PointType.Length <= 0) throw new SystemException("【積分類型】不能為空!");
                //if (PointQty <= 0) throw new SystemException("【積分數量】不能小於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.ActivityPoint (ActivityId, PointType, PointQty, Remark, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ActivityPointId
                                VALUES (@ActivityId, @PointType, @PointQty, @Remark, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ActivityId,
                                PointType,
                                PointQty,
                                Remark,
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

        #region //AddParticipants -- 參與人員資料新增 -- Zoey 2023.02.09
        public string AddParticipants(int ActivityId, string Users, string ActivityPoints)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【參與人員】不能為空!");
                if (ActivityPoints.Length <= 0) throw new SystemException("【積分類型】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
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

                        #region //判斷活動點數資料是否正確
                        string[] activityPointsList = ActivityPoints.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalActivityPoints
                                FROM EBP.ActivityPoint
                                WHERE ActivityPointId IN @ActivityPointId";
                        dynamicParameters.Add("ActivityPointId", activityPointsList);

                        int totalActivityPoints = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalActivityPoints;
                        if (totalActivityPoints != activityPointsList.Length) throw new SystemException("活動點數資料錯誤!");
                        #endregion

                        int rowsAffected = 0, participantId = -1;
                        foreach (var user in usersList)
                        {
                            #region //判斷人員是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ParticipantId
                                    FROM EBP.Participant
                                    WHERE UserId = @UserId
                                    AND ActivityId = @ActivityId";
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));
                            dynamicParameters.Add("ActivityId", ActivityId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            if (result2.Count() > 0)
                            {
                                foreach (var item in result2)
                                {
                                    participantId = item.ParticipantId;
                                }

                                #region //判斷人員點數是否已兌換
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.Participant
                                        WHERE Status = 'Y'
                                        AND ParticipantId =  @ParticipantId";
                                dynamicParameters.Add("ParticipantId", participantId);

                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0) throw new SystemException("人員點數已兌換，請至【點數異動】功能新增點數!");
                                #endregion
                            }
                            else
                            {
                                #region //新增EBP.Participant
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.Participant (UserId, ActivityId, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.ParticipantId
                                        VALUES (@UserId, @ActivityId, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        UserId = Convert.ToInt32(user),
                                        ActivityId,
                                        Status = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = insertResult.Count();

                                foreach (var item in insertResult)
                                {
                                    participantId = item.ParticipantId;
                                }
                                #endregion
                            }

                            foreach (var activityPoints in activityPointsList)
                            {
                                #region //判斷積分類型是否重複
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.ParticipantPoint
                                        WHERE ParticipantId = @ParticipantId
                                        AND ActivityPointId = @ActivityPointId";
                                dynamicParameters.Add("ParticipantId", participantId);
                                dynamicParameters.Add("ActivityPointId", activityPoints);

                                var result4 = sqlConnection.Query(sql, dynamicParameters);
                                if (result4.Count() > 0) throw new SystemException("【積分類型】重複，請重新輸入!");
                                #endregion

                                #region //新增EBP.ParticipantPoint
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.ParticipantPoint (ParticipantId, ActivityPointId, CreateDate, CreateBy)
                                        VALUES (@ParticipantId, @ActivityPointId, @CreateDate, @CreateBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ParticipantId = Convert.ToInt32(participantId),
                                        ActivityPointId = Convert.ToInt32(activityPoints),
                                        CreateDate,
                                        CreateBy
                                    });
                                var insertPointResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = insertPointResult.Count();
                                #endregion
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
        #endregion

        #region //Update
        #region //UpdateActivity -- 活動資料更新 -- Ben Ma 2023.01.07
        public string UpdateActivity(int ActivityId, int AnnualId, string ActivityName, string ActivityStartDate, string ActivityEndDate, string ActivityContent)
        {
            try
            {
                DateTime tempDate = default(DateTime);

                if (ActivityName.Length <= 0) throw new SystemException("【活動名稱】不能為空!");
                if (ActivityName.Length > 100) throw new SystemException("【活動名稱】長度錯誤!");
                if (ActivityStartDate.Length <= 0) throw new SystemException("【活動開始日期】不能為空!");
                if (!DateTime.TryParse(ActivityStartDate, out tempDate)) throw new SystemException("【活動開始日期】格式錯誤!");
                if (ActivityEndDate.Length <= 0) throw new SystemException("【活動開始日期】不能為空!");
                if (!DateTime.TryParse(ActivityEndDate, out tempDate)) throw new SystemException("【活動結束日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Activity SET
                                AnnualId = @AnnualId,
                                ActivityName = @ActivityName,
                                ActivityStartDate = @ActivityStartDate,
                                ActivityEndDate = @ActivityEndDate,
                                ActivityContent = @ActivityContent,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AnnualId, 
                                ActivityName,
                                ActivityStartDate,
                                ActivityEndDate,
                                ActivityContent,
                                LastModifiedDate,
                                LastModifiedBy,
                                ActivityId
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

        #region //UpdateActivityStatus -- 活動狀態更新 -- Ben Ma 2023.01.07
        public string UpdateActivityStatus(int ActivityId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");

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
                        sql = @"UPDATE EBP.Activity SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ActivityId
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

        #region //UpdateAwards -- 活動獎項資料更新 -- Ben Ma 2023.01.07
        public string UpdateAwards(int AwardsId, string AwardsName, int AwardsPrice, int AwardsQuantity
            , string AwardsUnit, string AwardsRemark)
        {
            try
            {
                if (AwardsName.Length <= 0) throw new SystemException("【獎項名稱】不能為空!");
                if (AwardsName.Length > 100) throw new SystemException("【獎項名稱】長度錯誤!");
                if (AwardsRemark.Length > 500) throw new SystemException("【獎項備註】長度錯誤!");
                if (AwardsUnit.Length > 50) throw new SystemException("【獎項數量單位】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動獎項資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ActivityId
                                FROM EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動獎項資料錯誤!");

                        int ActivityId = -1;
                        foreach (var item in result)
                        {
                            ActivityId = Convert.ToInt32(item.ActivityId);
                        }
                        #endregion

                        #region //判斷獎項名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Awards
                                WHERE AwardsName = @AwardsName
                                AND ActivityId = @ActivityId
                                AND AwardsId != @AwardsId";
                        dynamicParameters.Add("AwardsName", AwardsName);
                        dynamicParameters.Add("ActivityId", ActivityId);
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【獎項名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Awards SET
                                AwardsName = @AwardsName,
                                AwardsPrice = @AwardsPrice,
                                AwardsQuantity = @AwardsQuantity,
                                AwardsUnit = @AwardsUnit,
                                AwardsRemark = @AwardsRemark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AwardsName,
                                AwardsPrice = AwardsPrice > 0 ? AwardsPrice : (int?)null,
                                AwardsQuantity = AwardsQuantity > 0 ? AwardsQuantity : (int?)null,
                                AwardsUnit,
                                AwardsRemark,
                                LastModifiedDate,
                                LastModifiedBy,
                                AwardsId
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

        #region //UpdateAwardsStatus -- 活動獎項狀態更新 -- Ben Ma 2023.01.07
        public string UpdateAwardsStatus(int AwardsId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動獎項資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動獎項資料錯誤!");

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
                        sql = @"UPDATE EBP.Awards SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                AwardsId
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

        #region //UpdateAwardsSort -- 活動獎項順序調整 -- Ben Ma 2023.01.07
        public string UpdateAwardsSort(int ActivityId, string AwardsList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Awards SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ActivityId", ActivityId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] awardsSort = AwardsList.Split(',');

                        for (int i = 0; i < awardsSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EBP.Awards SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE AwardsId = @AwardsId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    AwardsId = Convert.ToInt32(awardsSort[i])
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

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

        #region //UpdatePoints -- 活動積分資料更新 -- Zoey 2023.02.07
        public string UpdatePoints(int ActivityPointId, string PointType, int PointQty, string Remark)
        {
            try
            {
                if (PointType.Length <= 0) throw new SystemException("【積分類型】不能為空!");
                //if (PointQty <= 0) throw new SystemException("【積分數量】不能小於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動積分資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ActivityPoint
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.Add("ActivityPointId", ActivityPointId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動積分資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.ActivityPoint SET
                                PointType = @PointType,
                                PointQty = @PointQty,
                                Remark = @Remark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PointType,
                                PointQty,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                ActivityPointId
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

        #region //UpdatePointStatus -- 活動積分狀態更新 -- Zoey 2023.02.07
        public string UpdatePointStatus(int ActivityPointId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動積分資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.ActivityPoint
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.Add("ActivityPointId", ActivityPointId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動積分資料錯誤!");

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
                        sql = @"UPDATE EBP.ActivityPoint SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ActivityPointId
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

        #region //UpdateRedeemPoint -- 參與人員狀態更新 -- Zoey 2023.02.09
        public string UpdateRedeemPoint(int ActivityId, int ParticipantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        if (ParticipantId > 0)
                        {
                            #region //判斷參與人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EBP.Participant
                                    WHERE ParticipantId = @ParticipantId";
                            dynamicParameters.Add("ParticipantId", ParticipantId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("參與人員資料錯誤!");
                            #endregion
                        }

                        int rowsAffected = 0;

                        #region //新增EBP.Point
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Point
                                OUTPUT INSERTED.PointId
                                SELECT b.UserId, d.AnnualId, c.PointQty, 'Activity' PointType, d.ActivityId
                                , NULL ClubId, NULL RedeemPrizeId, c.Remark, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                FROM EBP.ParticipantPoint a
                                INNER JOIN EBP.Participant b ON b.ParticipantId = a.ParticipantId
                                INNER JOIN EBP.ActivityPoint c ON c.ActivityPointId = a.ActivityPointId
                                INNER JOIN EBP.Activity d ON d.ActivityId = b.ActivityId AND d.ActivityId = c.ActivityId 
                                WHERE b.Status = 'N' ";

                        sql += ParticipantId > 0 ? @"AND a.ParticipantId = @Id" : @"AND d.ActivityId = @Id";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                Id = ParticipantId > 0 ? ParticipantId : ActivityId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //修改EBP.Participant
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Participant SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE Status = 'N' ";

                        sql += ParticipantId > 0 ? @"AND ParticipantId = @Id" : @"AND ActivityId = @Id";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                Id = ParticipantId > 0 ? ParticipantId : ActivityId
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
        #region //DeleteActivity -- 活動資料刪除 -- Ben Ma 2023.01.07
        public string DeleteActivity(int ActivityId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除獎項人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.AwardsUser a
                                INNER JOIN EBP.Awards b ON a.AwardsId = b.AwardsId
                                WHERE b.ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除活動獎項
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Awards
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Activity
                                WHERE ActivityId = @ActivityId";
                        dynamicParameters.Add("ActivityId", ActivityId);

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

        #region //DeleteAwards -- 活動獎項資料刪除 -- Ben Ma 2023.01.07
        public string DeleteAwards(int AwardsId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動獎項是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動獎項資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除獎項人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.AwardsUser
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

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

        #region //DeleteAwardsUser -- 刪除獎項人員資料 -- Ben Ma 2023.01.07
        public string DeleteAwardsUser(int AwardsId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動獎項資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Awards
                                WHERE AwardsId = @AwardsId";
                        dynamicParameters.Add("AwardsId", AwardsId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動獎項資料錯誤!");
                        #endregion

                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.AwardsUser
                                WHERE AwardsId = @AwardsId
                                AND UserId = @UserId";
                        dynamicParameters.Add("AwardsId", AwardsId);
                        dynamicParameters.Add("UserId", UserId);

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

        #region //DeletePoints -- 刪除活動積分資料 -- Zoey 2023.02.07
        public string DeletePoints(int ActivityPointId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷活動積分是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ActivityPoint
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.Add("ActivityPointId", ActivityPointId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("活動積分資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ActivityPoint
                                WHERE ActivityPointId = @ActivityPointId";
                        dynamicParameters.Add("ActivityPointId", ActivityPointId);

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

        #region //DeleteParticipants -- 刪除參與人員資料 -- Zoey 2023.02.09
        public string DeleteParticipants(int ParticipantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷參與人員是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Participant
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("參與人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除關聯table
                        #region //刪除EBP.ParticipantPoint
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ParticipantPoint
                                WHERE ParticipantId = @ParticipantId";
                        dynamicParameters.Add("ParticipantId", ParticipantId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Participant
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
        #endregion
    }
}