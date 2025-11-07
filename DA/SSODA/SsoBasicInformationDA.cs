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

namespace SSODA
{
    public class SsoBasicInformationDA
    {
        public static string MainConnectionStrings = "";
        public static string ErpConnectionStrings = "";

        public static int CurrentCompany = -1;
        public static int CurrentUser = -1;
        public static int CreateBy = -1;
        public static int LastModifiedBy = -1;
        public static DateTime CreateDate = default(DateTime);
        public static DateTime LastModifiedDate = default(DateTime);

        public static string sql = "";
        public static JObject jsonResponse = new JObject();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static SqlQuery sqlQuery = new SqlQuery();

        public SsoBasicInformationDA()
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
        #region //GetDemandSource -- 取得需求來源資料 -- Ben Ma 2023.07.10
        public string GetDemandSource(int SourceId, string SearchKey, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SourceId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.SourceNo, a.SourceName, a.SourceDesc, a.[Status]
                        , a.SourceNo + ' ' + a.SourceName SourceWithNo";
                    sqlQuery.mainTables =
                        @"FROM SSO.DemandSource a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SourceId", @" AND a.SourceId = @SourceId", SourceId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.SourceNo LIKE '%' + @SearchKey + '%' OR a.SourceName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SourceNo, a.SourceId DESC";
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

        #region //GetDemandFlow -- 取得流程資料 -- Ben Ma 2023.07.10
        public string GetDemandFlow(int FlowId, string SearchKey, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.FlowId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FlowNo, a.FlowName, a.FlowImage, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM SSO.DemandFlow a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowId", @" AND a.FlowId = @FlowId", FlowId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.FlowNo LIKE '%' + @SearchKey + '%' OR a.FlowName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.FlowId DESC";
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

        #region //GetDemandRole -- 取得角色資料 -- Ben Ma 2023.07.10
        public string GetDemandRole(int RoleId, string SearchKey, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoleName, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM SSO.DemandRole a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.RoleName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoleId DESC";
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

        #region //GetRoleUser -- 取得角色使用者資料 -- Ben Ma 2023.07.10
        public string GetRoleUser(int RoleId, int CompanyId, int DepartmentId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoleId,a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM SSO.RoleUser a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND c.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "d.CompanyNo, c.DepartmentNo, b.UserNo";
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
        #region //AddDemandSource -- 需求來源資料新增 -- Ben Ma 2023.07.10
        public string AddDemandSource(string SourceNo, string SourceName, string SourceDesc)
        {
            try
            {
                if (SourceNo.Length <= 0) throw new SystemException("【來源代碼】不能為空!");
                if (SourceNo.Length > 10) throw new SystemException("【來源代碼】長度錯誤!");
                if (SourceName.Length <= 0) throw new SystemException("【來源名稱】不能為空!");
                if (SourceName.Length > 50) throw new SystemException("【來源名稱】長度錯誤!");
                if (SourceDesc.Length > 50) throw new SystemException("【來源描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷來源代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandSource
                                WHERE SourceNo = @SourceNo";
                        dynamicParameters.Add("SourceNo", SourceNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【來源代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SSO.DemandSource (CompanyId, SourceNo, SourceName, SourceDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SourceId
                                VALUES (@CompanyId, @SourceNo, @SourceName, @SourceDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                SourceNo,
                                SourceName,
                                SourceDesc,
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

        #region //AddDemandFlow -- 流程資料新增 -- Ben Ma 2023.07.10
        public string AddDemandFlow(string FlowNo, string FlowName, string FlowImage)
        {
            try
            {
                if (FlowNo.Length <= 0) throw new SystemException("【流程代號】不能為空!");
                if (FlowNo.Length > 50) throw new SystemException("【流程代號】長度錯誤!");
                if (FlowName.Length <= 0) throw new SystemException("【流程名稱】不能為空!");
                if (FlowName.Length > 50) throw new SystemException("【流程名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷流程代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowNo = @FlowNo";
                        dynamicParameters.Add("FlowNo", FlowNo);

                        var resultNo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultNo.Count() > 0) throw new SystemException("【流程代號】重複，請重新輸入!");
                        #endregion

                        #region //判斷流程名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowName = @FlowName";
                        dynamicParameters.Add("FlowName", FlowName);

                        var resultName = sqlConnection.Query(sql, dynamicParameters);
                        if (resultName.Count() > 0) throw new SystemException("【流程名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SSO.DemandFlow (FlowNo, FlowName, FlowImage, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FlowId
                                VALUES (@FlowNo, @FlowName, @FlowImage, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FlowNo,
                                FlowName,
                                FlowImage,
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

        #region //AddDemandRole -- 角色資料新增 -- Ben Ma 2023.07.10
        public string AddDemandRole(string RoleName)
        {
            try
            {
                if (RoleName.Length <= 0) throw new SystemException("【角色名稱】不能為空!");
                if (RoleName.Length > 50) throw new SystemException("【角色名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandRole
                                WHERE RoleName = @RoleName";
                        dynamicParameters.Add("RoleName", RoleName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SSO.DemandRole (CompanyId, RoleName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoleId
                                VALUES (@CompanyId, @RoleName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                RoleName,
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

        #region //AddRoleUser -- 角色使用者資料新增 -- Ben Ma 2023.07.10
        public string AddRoleUser(int RoleId, string Users)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        #region //判斷使用者資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SSO.RoleUser (RoleId, UserId
                                    , CreateDate, CreateBy)
                                    VALUES (@RoleId, @UserId
                                    , @CreateDate, @CreateBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoleId,
                                    UserId = Convert.ToInt32(user),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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
        #region //UpdateDemandSource -- 需求來源資料更新 -- Ben Ma 2023.07.10
        public string UpdateDemandSource(int SourceId, string SourceName, string SourceDesc)
        {
            try
            {
                if (SourceName.Length <= 0) throw new SystemException("【來源名稱】不能為空!");
                if (SourceName.Length > 50) throw new SystemException("【來源名稱】長度錯誤!");
                if (SourceDesc.Length > 50) throw new SystemException("【來源描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandSource
                                WHERE SourceId = @SourceId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SourceId", SourceId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.DemandSource SET
                                SourceName = @SourceName,
                                SourceDesc = @SourceDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SourceId = @SourceId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SourceName,
                                SourceDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                SourceId
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

        #region //UpdateDemandSourceStatus -- 需求來源狀態更新 -- Ben Ma 2023.07.10
        public string UpdateDemandSourceStatus(int SourceId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SSO.DemandSource
                                WHERE SourceId = @SourceId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SourceId", SourceId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源資料錯誤!");

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
                        sql = @"UPDATE SSO.DemandSource SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SourceId = @SourceId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SourceId
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

        #region //UpdateDemandFlow -- 流程資料更新 -- Ben Ma 2023.07.10
        public string UpdateDemandFlow(int FlowId, string FlowNo, string FlowName, string FlowImage)
        {
            try
            {
                if (FlowNo.Length <= 0) throw new SystemException("【流程代號】不能為空!");
                if (FlowNo.Length > 50) throw new SystemException("【流程代號】長度錯誤!");
                if (FlowName.Length <= 0) throw new SystemException("【流程名稱】不能為空!");
                if (FlowName.Length > 50) throw new SystemException("【流程名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷流程資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowId = @FlowId";
                        dynamicParameters.Add("FlowId", FlowId);

                        var resultExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExist.Count() <= 0) throw new SystemException("流程資料錯誤!");
                        #endregion

                        #region //判斷流程代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowNo = @FlowNo
                                AND FlowId != @FlowId";
                        dynamicParameters.Add("FlowNo", FlowNo);
                        dynamicParameters.Add("FlowId", FlowId);

                        var resultNo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultNo.Count() > 0) throw new SystemException("【流程代號】重複，請重新輸入!");
                        #endregion

                        #region //判斷流程名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowName = @FlowName
                                AND FlowId != @FlowId";
                        dynamicParameters.Add("FlowName", FlowName);
                        dynamicParameters.Add("FlowId", FlowId);

                        var resultName = sqlConnection.Query(sql, dynamicParameters);
                        if (resultName.Count() > 0) throw new SystemException("【流程名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.DemandFlow SET
                                FlowNo = @FlowNo,
                                FlowName = @FlowName,
                                FlowImage = @FlowImage,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FlowId = @FlowId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FlowNo,
                                FlowName,
                                FlowImage,
                                LastModifiedDate,
                                LastModifiedBy,
                                FlowId
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

        #region //UpdateDemandFlowStatus -- 流程狀態更新 -- Ben Ma 2023.07.10
        public string UpdateDemandFlowStatus(int FlowId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷流程資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SSO.DemandFlow
                                WHERE FlowId = @FlowId";
                        dynamicParameters.Add("FlowId", FlowId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("流程資料錯誤!");

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
                        sql = @"UPDATE SSO.DemandFlow SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FlowId = @FlowId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                FlowId
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

        #region //UpdateDemandRole -- 角色資料更新 -- Ben Ma 2023.07.10
        public string UpdateDemandRole(int RoleId, string RoleName)
        {
            try
            {
                if (RoleName.Length <= 0) throw new SystemException("【角色名稱】不能為空!");
                if (RoleName.Length > 50) throw new SystemException("【角色名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        #region //判斷角色名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandRole
                                WHERE RoleName = @RoleName
                                AND RoleId != @RoleId";
                        dynamicParameters.Add("RoleName", RoleName);
                        dynamicParameters.Add("RoleId", RoleId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.DemandRole SET
                                RoleName = @RoleName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoleId = @RoleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleName,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoleId
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

        #region //UpdateDemandRoleStatus -- 角色狀態更新 -- Ben Ma 2023.07.10
        public string UpdateDemandRoleStatus(int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SSO.DemandRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");

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
                        sql = @"UPDATE SSO.DemandRole SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoleId = @RoleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoleId
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
        #region //DeleteRoleUser -- 角色使用者資料刪除 -- Ben Ma 2023.07.10
        public string DeleteRoleUser(int RoleId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
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
                        sql = @"DELETE SSO.RoleUser
                                WHERE RoleId = @RoleId
                                AND UserId = @UserId";
                        dynamicParameters.Add("RoleId", RoleId);
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
        #endregion
    }
}
