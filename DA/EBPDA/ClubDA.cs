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
    public class ClubDA
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

        public ClubDA()
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
        #region //GetClub -- 取得社團資料 -- Ben Ma 2023.02.16
        public string GetClub(int ClubId, string ClubName, string Status
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ClubId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ClubName, a.ClubApplicant, FORMAT(a.EstablishedDate, 'yyyy-MM-dd') EstablishedDate
                        , a.ClubProperty, a.ClubGoal, a.Appropriation, a.ActiveRegion, a.[Status]
                        , b.UserNo, b.UserName, b.Gender";
                    sqlQuery.mainTables =
                        @"FROM EBP.Club a
                        INNER JOIN BAS.[User] b ON a.ClubApplicant = b.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubId", @" AND a.ClubId = @ClubId", ClubId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubName", @" AND a.ClubName LIKE '%' + @ClubName + '%'", ClubName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.EstablishedDate DESC, a.ClubId";
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

        #region //GetClubMember -- 取得社團成員資料 -- Ben Ma 2023.02.16
        public string GetClubMember(int MemberId, int ClubId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ClubId, a.UserId, a.ClubJobId
                        , b.UserNo, b.UserName, b.Gender
                        , ISNULL(c.JobName, '') JobName";
                    sqlQuery.mainTables =
                        @"FROM EBP.ClubMember a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        LEFT JOIN EBP.ClubJob c ON a.ClubJobId = c.ClubJobId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubId", @" AND a.ClubId = @ClubId", ClubId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MemberId";
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
        #region //AddClub -- 社團資料新增 -- Ben Ma 2023.02.16
        public string AddClub(string ClubName, int ClubApplicant, string EstablishedDate, string ClubProperty
            , string ClubGoal, string Appropriation, string ActiveRegion)
        {
            try
            {
                if (ClubName.Length <= 0) throw new SystemException("【社團名稱】不能為空!");
                if (ClubName.Length > 100) throw new SystemException("【社團名稱】長度錯誤!");
                if (EstablishedDate.Length <= 0) throw new SystemException("【創立時間】不能為空!");
                if (!DateTime.TryParse(EstablishedDate, out DateTime tempDate)) throw new SystemException("【創立時間】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷申請人資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", ClubApplicant);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("申請人資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Club (ClubName, ClubApplicant, EstablishedDate, ClubProperty
                                , ClubGoal, Appropriation, ActiveRegion, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClubId
                                VALUES (@ClubName, @ClubApplicant, @EstablishedDate, @ClubProperty
                                , @ClubGoal, @Appropriation, @ActiveRegion, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClubName,
                                ClubApplicant,
                                EstablishedDate,
                                ClubProperty,
                                ClubGoal,
                                Appropriation,
                                ActiveRegion,
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

        #region //AddClubMember -- 社團成員資料新增 -- Ben Ma 2023.02.16
        public string AddClubMember(int ClubId, int ClubJobId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【專案成員】不能為空!");

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

                        #region //判斷社團成員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("社團成員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            #region //判斷是否重複參加
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserNo, b.UserName
                                    FROM EBP.ClubMember a
                                    INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                    WHERE a.ClubId = @ClubId
                                    AND a.UserId = @UserId";
                            dynamicParameters.Add("ClubId", ClubId);
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));

                            var resultJoin = sqlConnection.Query(sql, dynamicParameters);
                            if (resultJoin.Count() > 0)
                            {
                                foreach (var item in resultJoin)
                                {
                                    throw new SystemException(string.Format("{0} {1} 已是社團成員!", item.UserNo, item.UserName));
                                }
                            }
                            #endregion

                            #region //判斷社團成員是否參加超過3種社團
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(1) JoinClub, b.UserNo, b.UserName
                                    FROM EBP.ClubMember a
                                    INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                    WHERE a.UserId = @UserId
                                    GROUP BY b.UserNo, b.UserName";
                            dynamicParameters.Add("UserId", Convert.ToInt32(user));

                            var resultJoinMultiple = sqlConnection.Query(sql, dynamicParameters);
                            if (resultJoinMultiple.Count() > 0)
                            {
                                foreach (var item in resultJoinMultiple)
                                {
                                    if (Convert.ToInt32(item.JoinClub) >= 3) throw new SystemException(string.Format("{0} {1} 已參加超過三個社團!", item.UserNo, item.UserName));
                                }
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.ClubMember (ClubId, UserId, ClubJobId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MemberId
                                    VALUES (@ClubId, @UserId, @ClubJobId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ClubId,
                                    UserId = Convert.ToInt32(user),
                                    ClubJobId = ClubJobId <= 0 ? (int?)null : ClubJobId,
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
        #region //UpdateClub -- 社團資料更新 -- Ben Ma 2023.02.16
        public string UpdateClub(int ClubId, string ClubName, int ClubApplicant, string EstablishedDate
            , string ClubProperty, string ClubGoal, string Appropriation, string ActiveRegion)
        {
            try
            {
                if (ClubName.Length <= 0) throw new SystemException("【社團名稱】不能為空!");
                if (ClubName.Length > 100) throw new SystemException("【社團名稱】長度錯誤!");
                if (EstablishedDate.Length <= 0) throw new SystemException("【創立時間】不能為空!");
                if (!DateTime.TryParse(EstablishedDate, out DateTime tempDate)) throw new SystemException("【創立時間】格式錯誤!");

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

                        #region //判斷申請人資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", ClubApplicant);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("申請人資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Club SET
                                ClubName = @ClubName,
                                ClubApplicant = @ClubApplicant,
                                EstablishedDate = @EstablishedDate,
                                ClubProperty = @ClubProperty,
                                ClubGoal = @ClubGoal,
                                Appropriation = @Appropriation,
                                ActiveRegion = @ActiveRegion,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubId = @ClubId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClubName,
                                ClubApplicant,
                                EstablishedDate,
                                ClubProperty,
                                ClubGoal,
                                Appropriation,
                                ActiveRegion,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubId
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

        #region //UpdateClubStatus -- 社團狀態更新 -- Ben Ma 2023.02.16
        public string UpdateClubStatus(int ClubId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Club
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

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
                        sql = @"UPDATE EBP.Club SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubId = @ClubId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubId
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
        #region //DeleteClub -- 社團資料刪除 -- Ben Ma 2023.02.16
        public string DeleteClub(int ClubId)
        {
            try
            {
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

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除社團活動參與人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EBP.ClubParticipant a
                                INNER JOIN EBP.ClubActivity b ON a.ClubActivityId = b.ClubActivityId
                                WHERE b.ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除社團活動
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubActivity
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除社團成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubMember
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Club
                                WHERE ClubId = @ClubId";
                        dynamicParameters.Add("ClubId", ClubId);

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

        #region //DeleteClubMember -- 社團成員資料刪除 -- Ben Ma 2023.02.16
        public string DeleteClubMember(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubMember
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案成員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubMember
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

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
