using Dapper;
using Helpers;
using System;
using System.Web;
using System.Linq;
using System.Transactions;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace PMISDA
{
    public class TaskManagerDA
    {
        public static string MainConnectionStrings = "";
        public static string ErpConnectionStrings = "";
        public static string HrmConnectionStrings = "";

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
        public static MamoHelper mamoHelper = new MamoHelper();

        public TaskManagerDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
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
        #region //GetUser -- 取得使用者資料 -- Ben Ma 2022.03.31
        public string GetUser(int UserId, int DepartmentId, int CompanyId, string Departments
            , string UserNo, string UserName, string Gender, string Status, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender, ISNULL(a.Email, '') Email, ISNULL(a.Job, '') Job, ISNULL(a.JobType, '') JobType, a.Status
                        , b.DepartmentId, b.DepartmentNo, b.DepartmentName
                        , c.CompanyId, c.CompanyNo, c.CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon
                        , CASE a.Status 
                            WHEN 'S' THEN a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE a.UserNo + ' ' + a.UserName
                        END UserWithNo
                        , CASE a.Status 
                            WHEN 'S' THEN b.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE b.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName
                        END UserWithDepartment";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND b.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND b.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    if (Gender.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Gender", @" AND a.Gender IN @Gender", Gender.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.UserNo LIKE '%' + @SearchKey + '%' OR a.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.CompanyId, a.UserNo";
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

        #region //GetTaskType -- 取得任務類型 -- Ann 2025-02-24
        public string GetTaskType(int TaskTypeId, string TaskTypeName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.TaskTypeId, a.TaskTypeName
                            FROM PMIS.TaskType a 
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskTypeId", @" AND a.TaskTypeId = @TaskTypeId", TaskTypeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskTypeName", @" AND a.TaskTypeName LIKE '%' + @TaskTypeName + '%'", TaskTypeName);

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

        #region //GetDepartment -- 取得部門資料 -- Ann 2025-02-24
        public string GetDepartment(bool IsType, string CompanyNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.DepartmentId, a.DepartmentNo, a.DepartmentName
                            FROM BAS.Department a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo AND a.Status = 'A'";
                    dynamicParameters.Add("CompanyNo", CompanyNo);

                    if (IsType)
                    {
                        sql += @" AND a.DepartmentNo IN ('A751', 'A770', 'A780')";
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

        #region //GetCustomer -- 取得客戶資料 -- Ann 2025-02-24
        public string GetCustomer(string CompanyNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.CustomerId, a.CustomerNo, a.CustomerName
                            , a.CustomerNo + '-' + a.CustomerName CustomerFullNo
                            FROM SCM.Customer a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", CompanyNo);

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

        #region //GetTaskDistribution -- 取得任務(任務派發者) -- Shintokuro 2025.02.12
        public string GetTaskDistribution(string UserNo, int TdId, string TaskName, string StartDate, string ExpectedDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //確認任務發布者是否存在
                    int UserId = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = $@"SELECT a.UserId
                             FROM BAS.[User] a
                             WHERE UserNo = @UserNo";
                    dynamicParameters.Add("UserNo", UserNo);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");

                    foreach (var item in result)
                    {
                        UserId = item.UserId;
                    }
                    #endregion

                    sqlQuery.mainKey = "a.TdId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TaskTypeId, a.ProductionModel, a.CustomerId, a.DepartmentId
                          , a.TaskName, a.TaskContent
                          , a.AttendUser, a.ReturnFrequency,a.[Status]
                          , FORMAT(a.StartDate, 'yyyy-MM-dd') StartDate
                          , a.RepeatType, a.RepeatValue, a.RepeatEndType, a.RepeatEndValue,a.FileList,a.ImgList
                          , CASE WHEN a.ExpectedDate is NOT NULL  THEN FORMAT(a.ExpectedDate, 'yyyy-MM-dd')
                            ELSE '' END　ExpectedDate
                          , CASE 
                                WHEN GETDATE() >= DATEADD(DAY, -2, a.ExpectedDate) AND a.Status ='S'  THEN 'D'
                                ELSE a.Status
                            END AS TaskStatus
                          ";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TaskDistribution a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" AND a.CreateBy = @UserId";
                    dynamicParameters.Add("UserId", UserId);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TdId", @" AND a.TdId = @TdId", TdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TaskName", @" AND a.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.StartDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ExpectedDate", @" AND a.ExpectedDate <= @ExpectedDate ", ExpectedDate.Length > 0 ? Convert.ToDateTime(ExpectedDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0)
                    {
                        if (Status == "D") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND GETDATE() >= DATEADD(DAY, -2, a.ExpectedDate) AND a.Status = 'S'", Status.Split(','));
                        else
                        {
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                        }
                    }


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TdId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetAttendUserData -- 取得參與者 -- Shintokuro 2025.02.12
        public string GetAttendUserData(string AttendUser)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //確認參與者
                    dynamicParameters = new DynamicParameters();
                    sql = $@"SELECT STRING_AGG(a.UserNo + ' ' + a.UserName, ',') AS UserFull,STRING_AGG(CAST(a.UserId AS VARCHAR)  + '-' + a.UserName, ',') AS UseReportUser
                            FROM BAS.[User] a
                            WHERE UserId IN ({AttendUser})";
                    dynamicParameters.Add("AttendUser", AttendUser);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");
                    #endregion

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

        #region //GetReportContent -- 取得任務回報 -- Shintokuro 2025.02.12
        public string GetReportContent(int TdId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //確認參與者
                    dynamicParameters = new DynamicParameters();
                    sql = $@"SELECT a.MrId,a.ReportUserId,a.ReportContent, FORMAT(a.ReportDate, 'yyyy-MM-dd') ReportDate
                             ,b.UserNo + ' ' + b.UserName UserFull
                             FROM PMIS.MissionReport a
                             INNER JOIN BAS.[User] b on a.ReportUserId = b.UserId
                             WHERE a.TdId = @TdId
                             ORDER BY a.ReportUserId DESC, a.ReportDate DESC
                            ";
                    dynamicParameters.Add("TdId", TdId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #endregion

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

        #region //GetMissionReport -- 取得任務(參與者,任務回報清單) -- Shintokuro 2025.02.12
        public string GetMissionReport(int TdId, string TaskName, string StartDate, string ExpectedDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    

                    sqlQuery.mainKey = "a.MrId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Sort, a.Status ReportStatus 
                          , FORMAT(a.NextReportDate, 'yyyy-MM-dd') NextReportDate
                          , b.TaskName, b.TaskContent,b.AttendUser
                          , FORMAT(b.StartDate, 'yyyy-MM-dd') StartDate
                          , FORMAT(b.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, b.Status TaskStatus
                          , b.RepeatType, b.RepeatValue, b.RepeatEndType, b.RepeatEndValue
                          ";
                    sqlQuery.mainTables =
                        @"FROM PMIS.MissionReport a
                          INNER JOIN PMIS.TaskDistribution b on a.TdId = b.TdId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = $@"";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TdId", @" AND a.TdId = @TdId", TdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReportUserId", @" AND a.ReportUserId = @ReportUserId", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TaskName", @" AND a.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.StartDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ExpectedDate", @" AND b.ExpectedDate <= @ExpectedDate ", ExpectedDate.Length > 0 ? Convert.ToDateTime(ExpectedDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND b.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Sort ASC";
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

        #region //GetMrDetail -- 取得任務人員回報 -- Shintokuro 2025.03.04
        public string GetMrDetail(int MrDetailId, int TdId, int MrId, int ReportUserId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得任務Id
                    dynamicParameters = new DynamicParameters();
                    sql = $@"SELECT a.ReportUserId,a.TdId
                                 FROM PMIS.MissionReport a
                                 WHERE MrId = @MrId";
                    dynamicParameters.Add("MrId", MrId);
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in result)
                    {
                        TdId = item.TdId;
                    }
                    #endregion

                    sqlQuery.mainKey = "a.MrDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ReportContent,a.FileList,a.ImgList, FORMAT(a.ReportDate, 'yyyy-MM-dd HH:mm:ss') ReportDate
                          , 'B' + CAST(ROW_NUMBER() OVER (ORDER BY a.MrDetailId) AS VARCHAR) AS row
                          , c.UserNo , c.UserName 
                          , x.AttendUser
                          ";
                    sqlQuery.mainTables =
                        @"FROM PMIS.MrDetail a
                          INNER JOIN PMIS.MissionReport b on a.MrId = b.MrId
                          INNER JOIN BAS.[User] c on b.ReportUserId = c.UserId
                          OUTER APPLY(SELECT x1.AttendUser FROM PMIS.TaskDistribution x1 WHERE x1.TdId = b.TdId) x
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrDetailId", @" AND a.MrDetailId = @MrDetailId", MrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TdId", @" AND a.TdId = @TdId", TdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReportUserId", @" AND b.ReportUserId = @ReportUserId", ReportUserId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MrDetailId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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
        #region //AddTaskDistribution -- 任務新增 -- Shintokuro 2025.02.12
        public string AddTaskDistribution(string UserNo, string TaskName, string TaskContent, string ExpectedDate, string AttendUser
            , int ReturnFrequency, string RepeatType, int RepeatValue, string RepeatEndType, string RepeatEndValue, string FileList, string ImgList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (UserNo.Length <= 0) throw new SystemException("【任務發布者】不能為空!");
                        if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                        if (RepeatType.Length <= 0) throw new SystemException("【回報方式】不能為空!");

                        #region //確認任務發布者是否存在
                        int UserId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.UserId
                                 FROM BAS.[User] a
                                 WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");

                        foreach (var item in result)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.TaskDistribution (TaskName, TaskContent, StartDate, ExpectedDate, AttendUser, ReturnFrequency ,Status
                                , RepeatType, RepeatValue, RepeatEndType, RepeatEndValue, FileList, ImgList
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TdId
                                VALUES (@TaskName, @TaskContent, @StartDate, @ExpectedDate, @AttendUser, @ReturnFrequency, @Status
                                , @RepeatType, @RepeatValue, @RepeatEndType, @RepeatEndValue, @FileList, @ImgList
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TaskName,
                                TaskContent,
                                StartDate = CreateDate,
                                ExpectedDate = ExpectedDate != "" ? ExpectedDate : null,
                                AttendUser,
                                ReturnFrequency,
                                Status = "N", //啟用,
                                RepeatType,
                                RepeatValue,
                                RepeatEndType,
                                RepeatEndValue,
                                FileList,
                                ImgList,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //AddMrDetail -- 任務回報 -- Shintokuro 2025.02.13
        public string AddMrDetail(int MrId, string ReportContent)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (MrId <= 0) throw new SystemException("【任務】不能為空");
                        if (ReportContent.Length <= 0) throw new SystemException("【回報內容】不能為空!");
                        int rowsAffected = -1;
                        int TdId = -1;
                        int Sort = -1;
                        string Status = "";
                        #region //確認任務是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.ReportUserId,a.TdId,a.Status,ISNULL(x.Sort,1) Sort
                                 FROM PMIS.MissionReport a
                                 OUTER APPLY(
                                     SELECT TOP 1 x1.Sort+1 Sort
                                     FROM PMIS.MissionReport x1
                                     WHERE x1.Status = 'S' 
                                     AND x1.ReportUserId = @ReportUserId
                                     ORDER BY x1.Sort DESC
                                 ) x
                                 WHERE MrId = @MrId";
                        dynamicParameters.Add("MrId", MrId);
                        dynamicParameters.Add("ReportUserId", CurrentUser);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務不存在!");

                        foreach (var item in result)
                        {
                            if(item.ReportUserId != CurrentUser) throw new SystemException("非任務回報者!");
                            TdId = item.TdId;
                            Status = item.Status;
                            Sort = item.Sort;
                        }
                        #endregion

                        //if(Status == "N")
                        //{
                        //    #region //修改任務回報狀態
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = $@"UPDATE PMIS.MissionReport SET
                        //             Sort = @Sort,
                        //             Status = @Status,
                        //             StartDate = @StartDate,
                        //             LastModifiedDate = @LastModifiedDate,
                        //             LastModifiedBy = @LastModifiedBy
                        //             WHERE MrId = @MrId";
                        //    dynamicParameters.AddDynamicParams(
                        //        new
                        //        {
                        //            Sort,
                        //            StartDate = LastModifiedDate,
                        //            Status = "S",
                        //            LastModifiedDate,
                        //            LastModifiedBy,
                        //            MrId
                        //        });
                        //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //    #endregion
                        //}


                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.MrDetail (MrId, TdId, ReportContent, ReportDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MrDetailId
                                VALUES (@MrId, @TdId, @ReportContent, @ReportDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MrId,
                                TdId,
                                ReportContent,
                                ReportDate = CreateDate,
                                Status = "N", //啟用
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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
        #region //UpdateTaskDistribution -- 任務修改 -- Shintokuro 2025.02.12
        public string UpdateTaskDistribution(string UserNo, int TdId, int TaskTypeId, string TaskName, string ProductionModel, int CustomerId, int DepartmentId, string TaskContent, string ExpectedDate, string AttendUser
            , int ReturnFrequency, string RepeatType, int RepeatValue, string RepeatEndType, string RepeatEndValue, string FileList, string ImgList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (TdId <= 0) throw new SystemException("【任務】不能為空!");
                        if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                        if (RepeatType.Length <= 0) throw new SystemException("【重複類型】不能為空!");

                        #region //確認任務發布者是否存在
                        int UserId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.UserId
                                 FROM BAS.[User] a
                                 WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");

                        foreach (var item in result)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //確認任務否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT TOP 1 CreateBy
                                 FROM PMIS.TaskDistribution a
                                 WHERE TdId = @TdId";
                        dynamicParameters.Add("TdId", TdId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("任務不存在,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.CreateBy != UserId) throw new SystemException("當前使用者和任務發布者不一致,不能修改!");
                        }
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TaskDistribution SET
                                TaskName = @TaskName,
                                TaskContent = @TaskContent,
                                ExpectedDate = @ExpectedDate,
                                AttendUser = @AttendUser,
                                ReturnFrequency = @ReturnFrequency,
                                RepeatType = @RepeatType,
                                RepeatValue = @RepeatValue,
                                RepeatEndType = @RepeatEndType,
                                RepeatEndValue = @RepeatEndValue,
                                FileList = @FileList,
                                ImgList = @ImgList,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TdId = @TdId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TaskName,
                                TaskContent,
                                ExpectedDate = ExpectedDate != "" ? ExpectedDate : null,
                                AttendUser,
                                ReturnFrequency,
                                RepeatType,
                                RepeatValue,
                                RepeatEndType,
                                RepeatEndValue,
                                FileList,
                                ImgList,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                TdId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion


                        transactionScope.Complete();
                    }
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

        #region //UpdateTaskDistributionStatus -- 任務狀態修改 -- Shintokuro 2025.02.12
        public string UpdateTaskDistributionStatus(string UserNo, int TdId, string Status)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (TdId <= 0) throw new SystemException("【任務】不能為空!");
                        if (Status.Length <= 0) throw new SystemException("【狀態】不能為空!");
                        int rowsAffected = -1;
                        string TaskStatus = "";
                        string ActionMsg = "";
                        string UserInfo = "";
                        string UserName = "";
                        string TaskName = "";
                        string TaskContent = "";
                        string AttendUser = "";
                        string ReturnFrequency = "";
                        string StartDate = "";
                        string ExpectedDate = "";
                        string RepeatType = "";
                        int RepeatValue = 0;
                        string RepeatEndType = "";
                        string RepeatEndValue = "";
                        string TaskTypeName = "";
                        string CustomerNo = "";
                        string CustomerName = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ProductionModel = "";

                        #region //確認任務發布者是否存在
                        int UserId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.UserId,a.UserNo +' '+a.UserName UserInfo,a.UserName
                                 FROM BAS.[User] a
                                 WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");

                        foreach (var item in result)
                        {
                            UserId = item.UserId;
                            UserInfo = item.UserInfo;
                            UserName = item.UserName;
                        }
                        #endregion

                        #region //確認任務否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT TOP 1 a.CreateBy,a.Status,a.TaskName,a.TaskContent,a.AttendUser,a.ReturnFrequency
                                , a.RepeatType, a.RepeatValue, a.RepeatEndType, a.RepeatEndValue, a.ProductionModel
                                , FORMAT(a.StartDate, 'yyyy-MM-dd') StartDate
                                , FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate
                                 FROM PMIS.TaskDistribution a
                                 WHERE a.TdId = @TdId";
                        dynamicParameters.Add("TdId", TdId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("任務不存在,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.CreateBy != UserId) throw new SystemException("當前使用者和任務發布者不一致,不能修改!");
                            TaskStatus = item.Status;
                            TaskName = item.TaskName;
                            TaskContent = item.TaskContent;
                            AttendUser = item.AttendUser;
                            ReturnFrequency = item.ReturnFrequency.ToString();
                            StartDate = item.StartDate;
                            ExpectedDate = item.ExpectedDate;
                            RepeatType = item.RepeatType;
                            RepeatValue = item.RepeatValue;
                            RepeatEndType = item.RepeatEndType;
                            RepeatEndValue = item.RepeatEndValue;
                        }
                        #endregion

                        switch (Status)
                        {
                            case "N":
                                if (TaskStatus != "S") throw new SystemException("任務狀態須處於開始狀態才能反確認!");
                                ActionMsg = "修改任務成功!";
                                break;
                            case "S":
                                if (TaskStatus != "N") throw new SystemException("任務狀態須處於尚未開始狀態才能確認!");
                                ActionMsg = "發佈任務成功!";

                                DateTime? NextReportDate = default(DateTime);

                                switch (RepeatType)
                                {
                                    case "NONE":
                                        NextReportDate = null;
                                        break;
                                    case "DAILY":
                                        NextReportDate = CreateDate.AddDays(1);
                                        break;
                                    case "WEEKLY":
                                        int daysUntilFriday = ((int)DayOfWeek.Friday - (int)CreateDate.DayOfWeek + 7) % 7;
                                        NextReportDate = CreateDate.AddDays(daysUntilFriday == 0 ? 7 : daysUntilFriday);
                                        break;
                                    case "INTERVAL":
                                        NextReportDate = CreateDate.AddDays(3);
                                        break;
                                }



                                foreach (var User in AttendUser.Split(','))
                                {
                                    #region //確認任務回報者資料表是否建立
                                    int Sort = -1;
                                    dynamicParameters = new DynamicParameters();
                                    sql = $@"SELECT TOP 1 Sort+1 Sort
                                             FROM PMIS.MissionReport a
                                             WHERE ReportUserId = @ReportUserId
                                             AND Status = 'N'
                                            ORDER BY Sort DESC
                                            ";
                                    dynamicParameters.Add("TdId", TdId);
                                    dynamicParameters.Add("ReportUserId", User);

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() > 0)
                                    {
                                        foreach (var item in result)
                                        {
                                            Sort = item.Sort;
                                        }
                                    }
                                    else
                                    {
                                        Sort = 1;
                                    }
                                    #endregion


                                    #region //確認任務回報者資料表是否建立
                                    dynamicParameters = new DynamicParameters();
                                    sql = $@"SELECT TOP 1 1
                                             FROM PMIS.MissionReport a
                                             WHERE TdId = @TdId
                                             AND ReportUserId = @ReportUserId";
                                    dynamicParameters.Add("TdId", TdId);
                                    dynamicParameters.Add("ReportUserId", User);

                                    result = sqlConnection.Query(sql, dynamicParameters);

                                    if (result.Count() <= 0)
                                    {
                                        #region //新增SQL
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PMIS.MissionReport (TdId, Sort, ReportUserId, GetDate, StartDate, EndDate, NextReportDate, Status, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.MrId
                                                VALUES (@TdId, @Sort, @ReportUserId, @GetDate, @StartDate, @EndDate, @NextReportDate, @Status, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                TdId,
                                                Sort,
                                                ReportUserId = User,
                                                GetDate = CreateDate,
                                                StartDate = (string)null,
                                                EndDate = (string)null,
                                                NextReportDate,
                                                Status = "N", //啟用,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy = UserId,
                                                LastModifiedBy = UserId
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult1.Count();
                                        #endregion
                                    }
                                    #endregion
                                }


                                #region //MAMO通知
                                string Web = "https://bm.zy-tech.com.tw/";
                                string Url = Web + "TaskManager/MissionReport";

                                #region //確認公司資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.CompanyNo,a.UserNo + ' ' + a.UserName UserInfo
                                        FROM BAS.[User] a 
                                        INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                        INNER JOIN BAS.Company c on b.CompanyId = c.CompanyId
                                        WHERE a.UserNo = @UserNo";
                                dynamicParameters.Add("UserNo", UserNo);

                                var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                                if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                                string CompanyNo = "";
                                foreach (var item in CompanyResult)
                                {
                                    CompanyNo = item.CompanyNo;
                                    UserInfo = item.UserInfo;
                                }
                                #endregion

                                string Content = "";

                                string taskFullName = " 主旨:" + TaskName;

                                string repeatEndType = "";
                                switch (RepeatEndType)
                                {
                                    case "NEVER":
                                        repeatEndType = "永不結束";
                                        break;
                                    case "COUNT":
                                        repeatEndType = "重複" + RepeatEndValue + "次";
                                        break;
                                    case "UNTIL":
                                        repeatEndType = "重複到" + RepeatEndValue;
                                        break;
                                }

                                string taskRepeat = "";
                                switch (RepeatType)
                                {
                                    case "NONE":
                                        taskRepeat = "不重複";
                                        break;
                                    case "DAILY":
                                        taskRepeat = "每天，" + repeatEndType;
                                        break;
                                    case "WEEKLY":
                                        taskRepeat = "每週" + RepeatValue.ToString() + " " + repeatEndType;
                                        break;
                                    case "INTERVAL":
                                        taskRepeat = "每隔" + RepeatValue.ToString() + " " + repeatEndType;
                                        break;
                                }

                                Content = "### 【" + UserName + "】 發佈任務通知 \n" +
                                          "##### 任務名稱: " + taskFullName + "\n" +
                                          "##### 任務描述: " + TaskContent + "\n" +
                                          "##### 任務回報區: " + Url + "\n" +
                                          "##### 開始日期: " + StartDate + "\n" +
                                          "##### 預計結束日期: " + ExpectedDate + "\n" +
                                          "##### 回報頻率: " + taskRepeat + "\n" +
                                          "- 發信時間: " + CreateDate + "\n" +
                                          "- 發信人員: " + UserInfo + "\n";

                                #region //取得標記USER資料(原送測人員部門)
                                List<string> Tags = new List<string>();
                                foreach (var User in AttendUser.Split(','))
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.UserName, a.UserNo
                                    FROM BAS.[User] a
                                    WHERE a.UserNo = @UserNo OR TRY_CAST(a.UserId AS VARCHAR) = @UserNo";
                                    dynamicParameters.Add("UserNo", User);

                                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item in UserResult)
                                    {
                                        Tags.Add(item.UserNo);

                                        List<int> Files = new List<int>();

                                        #region //執行
                                        string MamoResult = mamoHelper.SendMessage(CompanyNo, UserId, "Personal", item.UserNo, Content, Tags, Files);

                                        JObject MamoResultJson = JObject.Parse(MamoResult);
                                        if (MamoResultJson["status"].ToString() != "success")
                                        {
                                            throw new SystemException(MamoResultJson["msg"].ToString());
                                        }
                                        #endregion
                                    }
                                }
                                #endregion
                                #endregion
                                break;
                            case "C":
                                if (TaskStatus == "C" || TaskStatus == "V") throw new SystemException("任務已完成或作廢");
                                ActionMsg = "結案成功!";
                                break;
                            case "V":
                                if (TaskStatus == "C" || TaskStatus == "V") throw new SystemException("任務已完成或作廢");
                                ActionMsg = "中止任務成功!";
                                break;
                        }

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TaskDistribution SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TdId = @TdId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                TdId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = ActionMsg
                        });
                        #endregion


                        transactionScope.Complete();
                    }
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

        #region //UpdateMissionReport -- 任務回報修改 -- Shintokuro 2025.02.13
        public string UpdateMissionReport(int MrId, string ReportContent)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (MrId <= 0) throw new SystemException("【任務回報】不能為空");
                        if (ReportContent.Length <= 0) throw new SystemException("【回報內容】不能為空!");

                        #region //確認任務是否存在
                        int ReportUserId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.ReportUserId
                                 FROM PMIS.MissionReport a
                                 WHERE MrId = @MrId";
                        dynamicParameters.Add("MrId", MrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務回報不存在!");

                        foreach (var item in result)
                        {
                            ReportUserId = item.ReportUserId;
                            if (ReportUserId != CurrentUser) throw new SystemException("當前使用者和任務回報者不一致,不能修改!");
                        }
                        #endregion


                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.MissionReport SET
                                ReportContent = @ReportContent,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MrId = @MrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ReportContent,
                                LastModifiedDate,
                                LastModifiedBy,
                                MrId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //UpdateMsStatusSort -- 任務回報排序修改 -- Shintokuro 2025.03.13
        public string UpdateMsStatusSort(int MrId, string TargetState, string TargetSort)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = -1;
                        string TaskStartDate = "";
                        string BaseStatus = "";
                        int BaseSort = -1;

                        if (MrId <= 0) throw new SystemException("【任務回報】不能為空");
                        if (TargetState.Length <= 0) throw new SystemException("【目標狀態】不能為空!");
                        if (TargetSort.Length <= 0) throw new SystemException("【目標排序】不能為空!");

                        #region //確認任務是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.ReportUserId,a.Status,a.Sort
                                 FROM PMIS.MissionReport a
                                 WHERE MrId = @MrId";
                        dynamicParameters.Add("MrId", MrId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("任務回報不存在!");
                        foreach(var item in result)
                        {
                            if(item.ReportUserId != CurrentUser) throw new SystemException("非任務回報者!");
                            if (item.Status == "N" && TargetState != "N")
                            {
                                #region //加入任務開始時間
                                TaskStartDate = $"StartDate = '{DateTime.Now:yyyy-MM-dd HH:mm:ss}',";
                                #endregion
                            }
                            BaseStatus = item.Status;
                            BaseSort = item.Sort;
                        }
                        #endregion

                        #region //修改任務回報狀態
                        dynamicParameters = new DynamicParameters();
                        sql = $@"UPDATE PMIS.MissionReport SET
                                Status = @Status,
                                {TaskStartDate}
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MrId = @MrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = TargetState,
                                LastModifiedDate,
                                LastModifiedBy,
                                MrId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion


                        #region //排序修改

                        int Sort = 1;
                        if (BaseStatus != TargetState) // 跨排
                        {
                            #region //原始位置排序
                            dynamicParameters = new DynamicParameters();
                            sql = $@"UPDATE PMIS.MissionReport SET
                                    Sort = Sort - 1
                                    WHERE Sort > @Sort
                                    AND Status = @BaseStatus";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Sort = BaseSort,
                                    BaseStatus
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //目標位置排序
                            foreach (var Id in TargetSort.Split(','))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = $@"UPDATE PMIS.MissionReport SET
                                         Sort = @Sort
                                         WHERE MrId = @MrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Sort,
                                        MrId = Id,
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                Sort++;
                            }
                            #endregion

                        }
                        else //同排
                        {
                            #region //目標位置排序
                            foreach (var Id in TargetSort.Split(','))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = $@"UPDATE PMIS.MissionReport SET
                                         Sort = @Sort
                                         WHERE MrId = @MrId
                                        ";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Sort,
                                        MrId = Id,
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                Sort++;
                            }
                            #endregion
                        }
                        #endregion



                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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
        #region //DeleteTaskDistribution -- 任務刪除 -- Shintokuro 2025.02.12
        public string DeleteTaskDistribution(string UserNo, int TdId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (TdId <= 0) throw new SystemException("【任務】不能為空!");

                        int rowsAffected = 0;

                        #region //確認任務發布者是否存在
                        int UserId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT a.UserId
                                 FROM BAS.[User] a
                                 WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("任務發布者不存在!");

                        foreach (var item in result)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //確認任務否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT TOP 1 a.CreateBy,a.Status
                                 FROM PMIS.TaskDistribution a
                                 WHERE TdId = @TdId";
                        dynamicParameters.Add("TdId", TdId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("任務不存在,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.CreateBy != UserId) throw new SystemException("當前使用者和任務發布者不一致,不能修改!");
                            if (item.Status != "N") throw new SystemException("當前任務狀態不可以刪除");
                        }
                        #endregion

                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TaskDistribution
                                WHERE TdId = @TdId";
                        dynamicParameters.Add("TdId", TdId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //DeleteMissionReport -- 回報刪除 -- Shintokuro 2025.02.12
        public string DeleteMissionReport(int MrId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (MrId <= 0) throw new SystemException("【回報】不能為空!");

                        int rowsAffected = 0;


                        #region //確認任務否存在
                        dynamicParameters = new DynamicParameters();
                        sql = $@"SELECT TOP 1 a.ReportUserId,a.Status
                                 FROM PMIS.MissionReport a
                                 WHERE MrId = @MrId";
                        dynamicParameters.Add("MrId", MrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("任務不存在,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.ReportUserId != CurrentUser) throw new SystemException("當前使用者和任務回報者不一致,不能修改!");
                            //if (item.Status != "N") throw new SystemException("當前任務狀態不可以刪除");
                        }
                        #endregion

                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.MissionReport
                                WHERE MrId = @MrId";
                        dynamicParameters.Add("MrId", MrId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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
