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

namespace PMISDA
{
    public class PmisBasicInformationDA
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

        public PmisBasicInformationDA()
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
        #region //GetAuthority -- 取得專案權限資料 -- Ben Ma 2022.06.22
        public string GetAuthority(int AuthorityId, string AuthorityType, string AuthorityName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.AuthorityId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.AuthorityType, a.AuthorityCode, a.AuthorityName, a.SortNumber, a.Status";
                    sqlQuery.mainTables =
                        @"FROM PMIS.Authority a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AuthorityId", @" AND a.AuthorityId = @AuthorityId", AuthorityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AuthorityType", @" AND a.AuthorityType = @AuthorityType", AuthorityType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AuthorityName", @" AND a.AuthorityName LIKE '%' + @AuthorityName + '%'", AuthorityName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AuthorityType, a.SortNumber";
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

        #region //GetFileLevel -- 取得檔案等級資料 -- Ben Ma 2022.06.28
        public string GetFileLevel(int LevelId, string LevelName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.LevelId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.LevelCode, a.LevelName, a.SortNumber, a.Status";
                    sqlQuery.mainTables =
                        @"FROM PMIS.FileLevel a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LevelId", @" AND a.LevelId = @LevelId", LevelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LevelName", @" AND a.LevelName LIKE '%' + @LevelName + '%'", LevelName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
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

        #region //GetTemplate -- 取得樣板資料 -- Ben Ma 2022.06.29
        public string GetTemplate(int TemplateId, string TemplateName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TemplateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TemplateName, a.TemplateDesc, a.WorkTimeStatus, a.WorkTimeInterval, a.TemplateOwner, a.Status
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM PMIS.Template a
                        INNER JOIN BAS.[User] b ON a.TemplateOwner = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @" AND (
                               a.TemplateOwner = @UserId
                               OR 
                               EXISTS (
                                   SELECT TOP 1 1
                                   FROM PMIS.TemplateCooperation aa
                                   WHERE aa.TemplateId = a.TemplateId
                                   AND aa.UserId = @UserId
                               )
                           )";
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TemplateId", @" AND a.TemplateId = @TemplateId", TemplateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TemplateName", @" AND a.TemplateName LIKE '%' + @TemplateName + '%'", TemplateName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TemplateId";
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

        #region //GetTemplateCooperation -- 取得樣板協作人員資料 -- Ben Ma 2022.08.17
        public string GetTemplateCooperation(int TemplateId, int CompanyId, int DepartmentId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TemplateCooperationId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TemplateId, a.UserId
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TemplateCooperation a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.TemplateId = @TemplateId";
                    dynamicParameters.Add("TemplateId", TemplateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND c.DepartmentId = @DepartmentId", DepartmentId);
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

        #region //GetTemplateTask -- 取得樣板任務資料 -- Ben Ma 2022.07.15
        public string GetTemplateTask(int TemplateTaskId, int TemplateId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"DECLARE @rowsAdded int

                            DECLARE @templateTask TABLE
                            ( 
                                TemplateTaskId int,
                                ParentTaskId int,
                                TaskLevel int,
                                TaskRoute nvarchar(MAX),
                                TaskSort int,
                                TaskName nvarchar(MAX),
                                TaskRouteName nvarchar(MAX),
                                processed int DEFAULT(0)
                            )

                            INSERT @templateTask
                                SELECT TemplateTaskId, ParentTaskId, 1 TaskLevel
                                , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) AS TaskRoute
                                , TaskSort
                                , TaskName
                                , '' AS TaskRouteName
                                , 0
                                FROM PMIS.TemplateTask
                                WHERE ParentTaskId IS NULL

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @templateTask SET processed = 1 WHERE processed = 0

                                INSERT @templateTask
                                    SELECT b.TemplateTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                    , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) AS TaskRoute
                                    , b.TaskSort
                                    , b.TaskName
                                    , CAST((CASE WHEN LEN(a.TaskRouteName) > 0 THEN a.TaskRouteName + ' > ' ELSE '' END) + a.TaskName AS nvarchar(MAX)) AS TaskRouteName
                                    , 0
                                    FROM @templateTask a
                                    INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.ParentTaskId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @templateTask SET processed = 2 WHERE processed = 1
                            END;

                            SELECT a.TemplateTaskId, a.TaskRouteName ParentTaskName, a.TaskSort, a.TaskName
                            , b.TemplateId, ISNULL(b.ParentTaskId, -1) ParentTaskId, b.Duration, b.ReplyFrequency
                            , c.SubTask
                            , SUBSTRING(d.TemplateTaskUser, 0, LEN(d.TemplateTaskUser)) TemplateTaskUser
                            FROM @templateTask a
                            INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.TemplateTask ca
                                WHERE ca.ParentTaskId = a.TemplateTaskId
                            ) c
                            OUTER APPLY (
                                SELECT (
                                    SELECT db.UserName + ','
                                    FROM PMIS.TemplateTaskUser da
                                    INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                    WHERE da.TemplateTaskId = b.TemplateTaskId
                                    ORDER BY da.AgentSort
                                    FOR XML PATH('')
                                ) TemplateTaskUser
                            ) d
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TemplateTaskId", @" AND b.TemplateTaskId = @TemplateTaskId", TemplateTaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TemplateId", @" AND b.TemplateId = @TemplateId", TemplateId);

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

        #region //GetTemplateTaskTree -- 取得樣板任務資料(樹狀結構) -- Ben Ma 2022.07.13
        public string GetTemplateTaskTree(int TemplateId, int ParentTaskId, int OpenedLevel)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //樣板資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TemplateName
                            FROM PMIS.Template
                            WHERE TemplateId = @TemplateId";
                    dynamicParameters.Add("TemplateId", TemplateId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");

                    string templateName = "";

                    foreach (var item in result)
                    {
                        templateName = item.TemplateName;
                    }
                    #endregion

                    #region //樣板任務資料
                    List<TemlpateTask> temlpateTasks = new List<TemlpateTask>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int

                            DECLARE @templateTask TABLE
                            ( 
                                TemplateTaskId int,
                                ParentTaskId int,
                                TaskLevel int,
                                TaskRoute nvarchar(MAX),
                                TaskSort int,
                                processed int DEFAULT(0)
                            )

                            INSERT @templateTask
                                SELECT TemplateTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                                , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) AS TaskRoute
                                , TaskSort, 0
                                FROM PMIS.TemplateTask
                                WHERE TemplateId = @TemplateId
                                AND ParentTaskId IS NULL

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @templateTask SET processed = 1 WHERE processed = 0

                                INSERT @templateTask
                                    SELECT b.TemplateTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                    , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) AS TaskRoute
                                    , b.TaskSort, 0
                                    FROM @templateTask a
                                    INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.ParentTaskId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @templateTask SET processed = 2 WHERE processed = 1
                            END;

                            SELECT a.TemplateTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort
                            , b.TaskName
                            FROM @templateTask a
                            INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                            WHERE 1=1
                            ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                    dynamicParameters.Add("TemplateId", TemplateId);

                    temlpateTasks = sqlConnection.Query<TemlpateTask>(sql, dynamicParameters).ToList();
                    #endregion

                    var data = new DHXTree
                    {
                        value = templateName,
                        id = "-1",
                        opened = true
                    };
                    
                    if (temlpateTasks.Count > 0)
                    {
                        data.items = temlpateTasks
                            .Where(x => x.TaskLevel == 1 && x.ParentTaskId == Convert.ToInt32(data.id))
                            .OrderBy(x => x.TaskLevel)
                            .ThenBy(x => x.TaskRoute)
                            .ThenBy(x => x.TaskSort)
                            .Select(x => new DHXTree
                            {
                                value = x.TaskName,
                                id = x.TemplateTaskId.ToString(),
                                opened = true,
                                level = (int)x.TaskLevel
                            })
                            .ToList();

                        if (data.items.Count > 0) Recursion(data.items);
                    }

                    void Recursion(List<DHXTree> taskTree)
                    {
                        if (taskTree.Count > 0)
                        {
                            for (int i = 0; i < taskTree.Count; i++)
                            {
                                taskTree[i].items = temlpateTasks
                                    .Where(x => x.TaskLevel == (taskTree[i].level + 1) && x.ParentTaskId == Convert.ToInt32(taskTree[i].id))
                                    .OrderBy(x => x.TaskLevel)
                                    .ThenBy(x => x.TaskRoute)
                                    .ThenBy(x => x.TaskSort)
                                    .Select(x => new DHXTree
                                    {
                                        value = x.TaskName,
                                        id = x.TemplateTaskId.ToString(),
                                        opened = taskTree[i].level < OpenedLevel ? true : false,
                                        level = (int)x.TaskLevel
                                    })
                                    .ToList();

                                if (taskTree[i].items.Count > 0) Recursion(taskTree[i].items);
                            }
                        }
                    }

                    List<DHXTree> trees = new List<DHXTree>
                    {
                        data
                    };

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = trees
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

        #region //GetTemplateTaskGantt -- 取得任務樣板資料(甘特圖) -- Ben Ma 2022.08.05
        public string GetTemplateTaskGantt(int TemplateId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //樣板資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TemplateName
                            FROM PMIS.Template
                            WHERE TemplateId = @TemplateId";
                    dynamicParameters.Add("TemplateId", TemplateId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");

                    string templateName = "";

                    foreach (var item in result)
                    {
                        templateName = item.TemplateName;
                    }
                    #endregion

                    #region //樣板任務資料
                    List<TemlpateTask> temlpateTasks = new List<TemlpateTask>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int

                            DECLARE @templateTask TABLE
                            ( 
                                TemplateTaskId int,
                                ParentTaskId int,
                                TaskLevel int,
                                TaskRoute nvarchar(MAX),
                                TaskSort int,
                                processed int DEFAULT(0)
                            )

                            INSERT @templateTask
                                SELECT TemplateTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                                , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) AS TaskRoute
                                , TaskSort, 0
                                FROM PMIS.TemplateTask
                                WHERE TemplateId = @TemplateId
                                AND ParentTaskId IS NULL

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @templateTask SET processed = 1 WHERE processed = 0

                                INSERT @templateTask
                                    SELECT b.TemplateTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                    , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) AS TaskRoute
                                    , b.TaskSort, 0
                                    FROM @templateTask a
                                    INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.ParentTaskId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @templateTask SET processed = 2 WHERE processed = 1
                            END;

                            SELECT a.TemplateTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort
                            , b.TaskName, b.StartPoint, b.Duration
                            , c.SubTask
                            FROM @templateTask a
                            INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.TemplateTask ca
                                WHERE ca.ParentTaskId = a.TemplateTaskId
                            ) c
                            WHERE 1=1
                            ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                    dynamicParameters.Add("TemplateId", TemplateId);

                    temlpateTasks = sqlConnection.Query<TemlpateTask>(sql, dynamicParameters).ToList();
                    #endregion

                    #region //樣板任務連結資料
                    List<TemlpateTaskLink> temlpateTaskLinks = new List<TemlpateTaskLink>();

                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.TemplateTaskLinkId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SourceTaskId, a.TargetTaskId, a.LinkType";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TemplateTaskLink a
                        INNER JOIN PMIS.TemplateTask b ON a.SourceTaskId = b.TemplateTaskId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TemplateId", @" AND b.TemplateId = @TemplateId", TemplateId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "a.TemplateTaskLinkId";

                    temlpateTaskLinks = BaseHelper.SqlQuery<TemlpateTaskLink>(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion

                    #region //組成任務json
                    DateTime startDate = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 08:00:00"));
                    var temp = temlpateTasks.Select(x => new { Total = (Convert.ToInt32(x.StartPoint) + Convert.ToInt32(x.Duration)) }).ToList();
                    int totalDuration = Convert.ToInt32(temp.Select(x => x.Total).Max());
                    List<DHXTask> dHXTasks = new List<DHXTask>
                    {
                        new DHXTask
                        {
                            id = "1.1",
                            text = templateName,
                            start_date = startDate.ToString("yyyy-MM-dd HH:mm"),
                            duration = totalDuration.ToString(),
                            progress = "0",
                            parent = "0",
                            type = "project",
                            order = 1
                        }
                    };

                    foreach (var task in temlpateTasks)
                    {
                        dHXTasks.Add(
                            new DHXTask
                            {
                                id = task.TemplateTaskId.ToString(),
                                text = task.TaskName,
                                start_date = startDate.AddMinutes(Convert.ToInt32(task.StartPoint)).ToString("yyyy-MM-dd HH:mm"),
                                duration = Convert.ToInt32(task.Duration).ToString(),
                                progress = "0",
                                parent = task.ParentTaskId > -1 ? task.ParentTaskId.ToString() : "1.1",
                                type = task.SubTask > 0 ? "project" : "task",
                                order = Convert.ToInt32(task.TaskSort)
                            });
                    }
                    #endregion

                    #region //組成任務連結json
                    List<DHXLink> dHXLinks = new List<DHXLink>();

                    foreach (var link in temlpateTaskLinks)
                    {
                        dHXLinks.Add(
                            new DHXLink
                            {
                                id = link.TemplateTaskLinkId.ToString(),
                                type = link.LinkType,
                                source = link.SourceTaskId.ToString(),
                                target = link.TargetTaskId.ToString()
                            });
                    }

                    DHXCollection dHXCollection = new DHXCollection
                    {
                        links = dHXLinks
                    };
                    #endregion

                    DHXGantt dHXGantt = new DHXGantt
                    {
                        data = dHXTasks,
                        collections = dHXCollection
                    };

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = dHXGantt
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

        #region //GetTemplateTaskUser -- 取得樣板任務成員資料 -- Ben Ma 2022.07.21
        public string GetTemplateTaskUser(int TemplateTaskId, int CompanyId, int DepartmentId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TemplateTaskUserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TemplateTaskId, a.UserId, a.LevelId, a.AgentSort
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon
                        , (
                            SELECT aa.AuthorityId, aa.AuthorityName, ab.AuthorityStatus
                            FROM PMIS.Authority aa
                            OUTER APPLY (
                                SELECT ISNULL((
                                    SELECT TOP 1 1
                                    FROM PMIS.TemplateTaskUserAuthority aaa
                                    WHERE aaa.TemplateTaskUserId = a.TemplateTaskUserId
                                    AND aaa.AuthorityId = aa.AuthorityId
                                ), 0) AuthorityStatus
                            ) ab
                            WHERE aa.AuthorityType = @AuthorityType
                            AND aa.[Status] = @AuthorityStatus
                            ORDER BY aa.SortNumber
                            FOR JSON PATH, ROOT('data')
                        ) Authority";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TemplateTaskUser a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.TemplateTaskId = @TemplateTaskId";
                    dynamicParameters.Add("AuthorityType", "T");
                    dynamicParameters.Add("AuthorityStatus", "A");
                    dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND c.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.UserNo LIKE '%' + @SearchKey + '%' OR b.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AgentSort";
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
        #region //AddAuthority -- 專案權限資料新增 -- Ben Ma 2022.06.22
        public string AddAuthority(string AuthorityType, string AuthorityCode, string AuthorityName)
        {
            try
            {
                if (AuthorityType.Length <= 0) throw new SystemException("【權限分類】不能為空!");
                if (AuthorityCode.Length <= 0) throw new SystemException("【權限代碼】不能為空!");
                if (AuthorityCode.Length > 50) throw new SystemException("【權限代碼】長度錯誤!");
                if (AuthorityName.Length <= 0) throw new SystemException("【權限名稱】不能為空!");
                if (AuthorityName.Length > 100) throw new SystemException("【權限名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷權限分類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "Authority.AuthorityType");
                        dynamicParameters.Add("TypeNo", AuthorityType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【權限分類】資料錯誤!");
                        #endregion

                        #region //判斷權限代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityCode = @AuthorityCode";
                        dynamicParameters.Add("AuthorityCode", AuthorityCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【權限代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷權限名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityName = @AuthorityName";
                        dynamicParameters.Add("AuthorityName", AuthorityName);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【權限名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM PMIS.Authority
                                WHERE AuthorityType = @AuthorityType";
                        dynamicParameters.Add("AuthorityType", AuthorityType);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.Authority (AuthorityType, AuthorityCode, AuthorityName, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.AuthorityId
                                VALUES (@AuthorityType, @AuthorityCode, @AuthorityName, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AuthorityType,
                                AuthorityCode,
                                AuthorityName,
                                SortNumber = maxSort + 1,
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

        #region //AddFileLevel -- 檔案等級資料新增 -- Ben Ma 2022.06.28
        public string AddFileLevel(string LevelCode, string LevelName)
        {
            try
            {
                if (LevelCode.Length <= 0) throw new SystemException("【權限代碼】不能為空!");
                if (LevelCode.Length > 50) throw new SystemException("【權限代碼】長度錯誤!");
                if (LevelName.Length <= 0) throw new SystemException("【權限名稱】不能為空!");
                if (LevelName.Length > 100) throw new SystemException("【權限名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷檔案等級代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelCode = @LevelCode";
                        dynamicParameters.Add("LevelCode", LevelCode);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【檔案等級代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷檔案等級名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelName = @LevelName";
                        dynamicParameters.Add("LevelName", LevelName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【檔案等級名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM PMIS.FileLevel";

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.FileLevel (LevelCode, LevelName, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LevelId
                                VALUES (@LevelCode, @LevelName, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LevelCode,
                                LevelName,
                                SortNumber = maxSort + 1,
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

        #region //AddTemplate -- 樣板資料新增 -- Ben Ma 2022.06.29
        public string AddTemplate(string TemplateName, string TemplateDesc, string WorkTimeStatus, string WorkTimeInterval)
        {
            try
            {
                if (TemplateName.Length <= 0) throw new SystemException("【樣板名稱】不能為空!");
                if (TemplateName.Length > 100) throw new SystemException("【樣板名稱】長度錯誤!");
                if (TemplateDesc.Length <= 0) throw new SystemException("【樣板描述】不能為空!");
                if (TemplateDesc.Length > 300) throw new SystemException("【樣板描述】長度錯誤!");
                if (WorkTimeStatus.Length <= 0) throw new SystemException("【工作天模式】不能為空!");
                if (WorkTimeStatus == "Y") if (WorkTimeInterval.Length <= 0) throw new SystemException("【工作時段】不能為空!");
                if (WorkTimeInterval.Length > 100) throw new SystemException("【工作時段】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.Template (TemplateName, TemplateDesc, WorkTimeStatus, WorkTimeInterval
                                , TemplateOwner, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TemplateId
                                VALUES (@TemplateName, @TemplateDesc, @WorkTimeStatus, @WorkTimeInterval
                                , @TemplateOwner, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TemplateName,
                                TemplateDesc,
                                WorkTimeStatus,
                                WorkTimeInterval,
                                TemplateOwner = CurrentUser,
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

        #region //AddTemplateCooperation -- 樣板協作人員資料新增 -- Ben Ma 2022.08.17
        public string AddTemplateCooperation(int TemplateId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【協作人員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");
                        #endregion

                        #region //判斷協作人員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("任務成員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PMIS.TemplateCooperation (TemplateId, UserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TemplateCooperationId
                                    VALUES (@TemplateId, @UserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TemplateId,
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

        #region //AddTemplateTask -- 樣板任務資料新增 -- Ben Ma 2022.07.15
        public string AddTemplateTask(int TemplateId, int ParentTaskId, string TaskName, int Duration, int ReplyFrequency)
        {
            try
            {
                if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                if (TaskName.Length > 50) throw new SystemException("【任務名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");
                        #endregion

                        #region //判斷父層樣板任務資料是否正確
                        if (ParentTaskId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PMIS.TemplateTask
                                    WHERE TemplateTaskId = @TemplateTaskId";
                            dynamicParameters.Add("TemplateTaskId", ParentTaskId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("父層樣板任務資料錯誤!");
                        }
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(TaskSort), 0) MaxSort
                                FROM PMIS.TemplateTask
                                WHERE TemplateId = @TemplateId
                                AND ISNULL(ParentTaskId, 0) = @ParentTaskId";
                        dynamicParameters.Add("TemplateId", TemplateId);
                        dynamicParameters.Add("ParentTaskId", ParentTaskId > 0 ? ParentTaskId : 0);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.TemplateTask (TemplateId, ParentTaskId, TaskSort, TaskName
                                , StartPoint, Duration, ReplyFrequency
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TemplateTaskId
                                VALUES (@TemplateId, @ParentTaskId, @TaskSort, @TaskName
                                , @StartPoint, @Duration, @ReplyFrequency
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TemplateId,
                                ParentTaskId = ParentTaskId > 0 ? ParentTaskId : (int?)null,
                                TaskSort = maxSort + 1,
                                TaskName,
                                StartPoint = 0,
                                Duration = Duration <= 0 ? (int?)null : Duration,
                                ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
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

        #region //AddTemplateTaskLink -- 樣板任務連結資料新增 -- Ben Ma 2022.08.15
        public string AddTemplateTaskLink(int SourceTaskId, int TargetTaskId, string LinkType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷來源樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ParentTaskId
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @SourceTaskId";
                        dynamicParameters.Add("SourceTaskId", SourceTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("來源樣板任務資料錯誤!");

                        int sourceParent = -1;

                        foreach (var item in result)
                        {
                            sourceParent = Convert.ToInt32(item.ParentTaskId);
                        }
                        #endregion

                        #region //判斷目標樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ParentTaskId
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TargetTaskId";
                        dynamicParameters.Add("TargetTaskId", TargetTaskId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("目標樣板任務資料錯誤!");

                        int targetParent = -1;

                        foreach (var item in result2)
                        {
                            targetParent = Convert.ToInt32(item.ParentTaskId);
                        }
                        #endregion

                        #region //判斷是否是相同階層的
                        if (sourceParent != targetParent) throw new SystemException("不同階層的任務無法連結連結!");
                        #endregion

                        #region //判斷連結是否存在循環參考
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM PMIS.TemplateTaskLink a
                                WHERE a.SourceTaskId = @TargetTaskId
                                AND a.TargetTaskId = @SourceTaskId";
                        dynamicParameters.Add("TargetTaskId", TargetTaskId);
                        dynamicParameters.Add("SourceTaskId", SourceTaskId);
                        var existResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (existResult != null) throw new SystemException("【專案任務】不得循環參考!");
                        #endregion

                        #region //判斷連結類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @TypeNo
                                AND TypeSchema = @TypeSchema";
                        dynamicParameters.Add("TypeNo", LinkType);
                        dynamicParameters.Add("TypeSchema", "LinkType");

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("連結類型資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PMIS.TemplateTaskLink (SourceTaskId, TargetTaskId, LinkType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TemplateTaskLinkId
                                VALUES (@SourceTaskId, @TargetTaskId, @LinkType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SourceTaskId,
                                TargetTaskId,
                                LinkType,
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

        #region //AddTemplateTaskUser -- 樣板任務成員新增 -- Ben Ma 2022.07.21
        public string AddTemplateTaskUser(int TemplateTaskId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【任務成員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");
                        #endregion

                        #region //判斷任務成員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("任務成員資料錯誤!");
                        #endregion

                        #region //抓取目前最小檔案等級
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                WHERE 1=1
                                ORDER BY SortNumber DESC";

                        int levelId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).LevelId;
                        #endregion

                        #region //抓取目前最大代理順序
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(AgentSort), 0) MaxAgentSort
                                FROM PMIS.TemplateTaskUser
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        int maxAgentSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxAgentSort;
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            maxAgentSort++;

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PMIS.TemplateTaskUser (TemplateTaskId, UserId, LevelId, AgentSort
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TemplateTaskUserId
                                    VALUES (@TemplateTaskId, @UserId, @LevelId, @AgentSort
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TemplateTaskId,
                                    UserId = Convert.ToInt32(user),
                                    LevelId = levelId,
                                    AgentSort = maxAgentSort,
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
        #region //UpdateAuthority -- 專案權限資料更新 -- Ben Ma 2022.06.23
        public string UpdateAuthority(int AuthorityId, string AuthorityName)
        {
            try
            {
                if (AuthorityName.Length <= 0) throw new SystemException("【權限名稱】不能為空!");
                if (AuthorityName.Length > 100) throw new SystemException("【權限名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案權限資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityId = @AuthorityId";
                        dynamicParameters.Add("AuthorityId", AuthorityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案權限資料錯誤!");
                        #endregion

                        #region //判斷權限名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityName = @AuthorityName
                                AND AuthorityId != @AuthorityId";
                        dynamicParameters.Add("AuthorityName", AuthorityName);
                        dynamicParameters.Add("AuthorityId", AuthorityId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【權限名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.Authority SET
                                AuthorityName = @AuthorityName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AuthorityId = @AuthorityId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                AuthorityName,
                                LastModifiedDate,
                                LastModifiedBy,
                                AuthorityId
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

        #region //UpdateAuthorityStatus -- 專案權限狀態更新 -- Ben Ma 2022.06.23
        public string UpdateAuthorityStatus(int AuthorityId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案權限資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM PMIS.Authority
                                WHERE AuthorityId = @AuthorityId";
                        dynamicParameters.Add("AuthorityId", AuthorityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案權限資料錯誤!");

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
                        sql = @"UPDATE PMIS.Authority SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AuthorityId = @AuthorityId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                AuthorityId
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

        #region //UpdateAuthoritySort -- 專案權限順序調整 -- Ben Ma 2022.06.23
        public string UpdateAuthoritySort(string AuthorityType, string AuthorityList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷權限分類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "Authority.AuthorityType");
                        dynamicParameters.Add("TypeNo", AuthorityType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【權限分類】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.Authority SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AuthorityType = @AuthorityType";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("AuthorityType", AuthorityType);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] authoritySort = AuthorityList.Split(',');

                        for (int i = 0; i < authoritySort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PMIS.Authority SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE AuthorityId = @AuthorityId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    AuthorityId = Convert.ToInt32(authoritySort[i])
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

        #region //UpdateFileLevel -- 檔案等級資料更新 -- Ben Ma 2022.06.28
        public string UpdateFileLevel(int LevelId, string LevelName)
        {
            try
            {
                if (LevelName.Length <= 0) throw new SystemException("【檔案等級名稱】不能為空!");
                if (LevelName.Length > 100) throw new SystemException("【檔案等級名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷檔案等級資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        dynamicParameters.Add("LevelId", LevelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        #region //判斷檔案等級名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelName = @LevelName
                                AND LevelId != @LevelId";
                        dynamicParameters.Add("LevelName", LevelName);
                        dynamicParameters.Add("LevelId", LevelId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【檔案等級名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.FileLevel SET
                                LevelName = @LevelName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LevelId = @LevelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LevelName,
                                LastModifiedDate,
                                LastModifiedBy,
                                LevelId
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

        #region //UpdateFileLevelStatus -- 檔案等級狀態更新 -- Ben Ma 2022.06.28
        public string UpdateFileLevelStatus(int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷檔案等級資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        dynamicParameters.Add("LevelId", LevelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案等級資料錯誤!");

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
                        sql = @"UPDATE PMIS.FileLevel SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LevelId = @LevelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                LevelId
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

        #region //UpdateFileLevelSort -- 檔案等級順序調整 -- Ben Ma 2022.06.28
        public string UpdateFileLevelSort(string FileLevelList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.FileLevel SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] fileLevelSort = FileLevelList.Split(',');

                        for (int i = 0; i < fileLevelSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PMIS.FileLevel SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE LevelId = @LevelId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    LevelId = Convert.ToInt32(fileLevelSort[i])
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTemplate -- 樣板資料更新 -- Ben Ma 2022.06.29
        public string UpdateTemplate(int TemplateId, string TemplateName, string TemplateDesc, string WorkTimeStatus, string WorkTimeInterval)
        {
            try
            {
                if (TemplateName.Length <= 0) throw new SystemException("【樣板名稱】不能為空!");
                if (TemplateName.Length > 100) throw new SystemException("【樣板名稱】長度錯誤!");
                if (TemplateDesc.Length <= 0) throw new SystemException("【樣板描述】不能為空!");
                if (TemplateDesc.Length > 300) throw new SystemException("【樣板描述】長度錯誤!");
                if (WorkTimeStatus.Length <= 0) throw new SystemException("【工作天模式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");
                        #endregion

                        #region //判斷擁有者才能更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template a
                                WHERE a.TemplateId = @TemplateId
                                AND (
                                    a.TemplateOwner = @UserId
                                    OR
                                    EXISTS (
                                        SELECT TOP 1 1
                                        FROM PMIS.TemplateCooperation aa
                                        WHERE aa.TemplateId = a.TemplateId
                                        AND aa.UserId = @UserId
                                    )
                                )";
                        dynamicParameters.Add("TemplateId", TemplateId);
                        dynamicParameters.Add("UserId", CurrentUser);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("樣板擁有者才有更新權限!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.Template SET
                                TemplateName = @TemplateName,
                                TemplateDesc = @TemplateDesc,
                                WorkTimeStatus = @WorkTimeStatus,
                                WorkTimeInterval = @WorkTimeInterval,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TemplateName,
                                TemplateDesc,
                                WorkTimeStatus,
                                WorkTimeInterval,
                                LastModifiedDate,
                                LastModifiedBy,
                                TemplateId
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

        #region //UpdateTemplateStatus -- 樣板資料狀態更新 -- Ben Ma 2022.06.29
        public string UpdateTemplateStatus(int TemplateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");

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
                        sql = @"UPDATE PMIS.Template SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                TemplateId
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

        #region //UpdateTemplateTask -- 樣板資料更新 -- Ben Ma 2022.06.29
        public string UpdateTemplateTask(int TemplateTaskId, string TaskName, int Duration, int ReplyFrequency)
        {
            try
            {
                if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                if (TaskName.Length > 50) throw new SystemException("【任務名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TemplateTask SET
                                TaskName = @TaskName,
                                Duration = @Duration,
                                ReplyFrequency = @ReplyFrequency,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TaskName,
                                Duration = Duration <= 0 ? (int?)null : Duration,
                                ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                LastModifiedDate,
                                LastModifiedBy,
                                TemplateTaskId
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

        #region //UpdateTemplateTaskSort -- 樣板任務順序調整 -- Ben Ma 2022.07.19
        public string UpdateTemplateTaskSort(int TemplateTaskId, string SortType)
        {
            try
            {
                if (!Regex.IsMatch(SortType, "^(top|bottom)$", RegexOptions.IgnoreCase)) throw new SystemException("【順序調整類型】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TemplateId, ParentTaskId, TaskSort
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");

                        int templateId = -1, parentTaskId = -1, taskSort = 0;

                        foreach (var item in result)
                        {
                            templateId = Convert.ToInt32(item.TemplateId);
                            parentTaskId = Convert.ToInt32(item.ParentTaskId);
                            taskSort = Convert.ToInt32(item.TaskSort);
                        }
                        #endregion

                        switch (SortType)
                        {
                            case "top":
                                if (taskSort > 1)
                                {
                                    #region //先將自己排序調整成負數
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = TaskSort * -1
                                            WHERE TemplateTaskId = @TemplateTaskId";
                                    dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //被調整的後移
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = TaskSort + 1,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE TemplateId = @TemplateId
                                            AND ISNULL(ParentTaskId, 0) = @ParentTaskId
                                            AND TaskSort = @TaskSort";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TemplateId = templateId,
                                            ParentTaskId = parentTaskId > 0 ? parentTaskId : 0,
                                            TaskSort = taskSort - 1
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將自己的排序前移
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = @TaskSort,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE TemplateTaskId = @TemplateTaskId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TaskSort = taskSort - 1,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TemplateTaskId
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("任務為第一順位，無法調整!");
                                }
                                break;
                            case "bottom":
                                #region //搜尋最大序號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MAX(TaskSort) MaxSort
                                        FROM PMIS.TemplateTask
                                        WHERE TemplateId = @TemplateId
                                        AND ISNULL(ParentTaskId, 0) = @ParentTaskId";
                                dynamicParameters.Add("TemplateId", templateId);
                                dynamicParameters.Add("ParentTaskId", parentTaskId > 0 ? parentTaskId : 0);

                                int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                                #endregion

                                if (taskSort < maxSort)
                                {
                                    #region //先將自己排序調整成負數
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = TaskSort * -1
                                            WHERE TemplateTaskId = @TemplateTaskId";
                                    dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //被調整的前移
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = TaskSort - 1,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE TemplateId = @TemplateId
                                            AND ISNULL(ParentTaskId, 0) = @ParentTaskId
                                            AND TaskSort = @TaskSort";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TemplateId = templateId,
                                            ParentTaskId = parentTaskId > 0 ? parentTaskId : 0,
                                            TaskSort = taskSort + 1
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將自己的排序後移
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PMIS.TemplateTask SET
                                            TaskSort = @TaskSort,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE TemplateTaskId = @TemplateTaskId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TaskSort = taskSort + 1,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TemplateTaskId
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("任務已為最後順位，無法調整!");
                                }
                                break;
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

        #region //UpdateTemplateTaskSchedule -- 樣板任務排程更新 -- Ben Ma 2022.08.15
        public string UpdateTemplateTaskSchedule(int TemplateTaskId, string StartDate, int Duration)
        {
            try
            {
                if (!DateTime.TryParse(StartDate, out DateTime tempDate)) throw new SystemException("任務【開始時間】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");
                        #endregion

                        DateTime benchmark = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 08:00:00"));
                        int StartPoint = Convert.ToInt32(Convert.ToDateTime(StartDate).Subtract(benchmark).TotalMinutes);

                        if (StartPoint <= 0) StartPoint = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TemplateTask SET
                                StartPoint = @StartPoint,
                                Duration = @Duration,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                StartPoint,
                                Duration,
                                LastModifiedDate,
                                LastModifiedBy,
                                TemplateTaskId
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

        #region //UpdateTemplateAllTaskSchedule -- 樣板任務排程更新(全部任務) -- Ben Ma 2022.10.07
        public string UpdateTemplateAllTaskSchedule(string TaskData)
        {
            try
            {
                if (!TaskData.TryParseJson(out JObject tempJObject)) throw new SystemException("專案任務資料格式錯誤");

                List<TemlpateTask> temlpateTasks = JsonConvert.DeserializeObject<List<TemlpateTask>>(JObject.Parse(TaskData)["data"].ToString());

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        foreach (var task in temlpateTasks)
                        {
                            #region //判斷樣板任務資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                            dynamicParameters.Add("TemplateTaskId", task.TemplateTaskId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");
                            #endregion

                            DateTime benchmark = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 08:00:00"));
                            int StartPoint = Convert.ToInt32(Convert.ToDateTime(task.StartDate).Subtract(benchmark).TotalMinutes);

                            if (StartPoint <= 0) StartPoint = 0;

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PMIS.TemplateTask SET
                                    StartPoint = @StartPoint,
                                    Duration = @Duration,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TemplateTaskId = @TemplateTaskId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    StartPoint,
                                    task.Duration,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    task.TemplateTaskId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTemplateTaskUserAgentSort -- 樣板任務成員代理順序更新 -- Ben Ma 2022.08.17
        public string UpdateTemplateTaskUserAgentSort(int TemplateTaskId, string TemplateTaskUserList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TemplateTaskUser SET
                                AgentSort = AgentSort * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] templateTaskUserSort = TemplateTaskUserList.Split(',');

                        for (int i = 0; i < templateTaskUserSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PMIS.TemplateTaskUser SET
                                    AgentSort = @AgentSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TemplateTaskId = @TemplateTaskId
                                    AND TemplateTaskUserId = @TemplateTaskUserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AgentSort = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TemplateTaskId,
                                    TemplateTaskUserId = Convert.ToInt32(templateTaskUserSort[i])
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

        #region //UpdateTemplateTaskUserAuthority -- 樣板任務成員權限更新 -- Ben Ma 2022.07.25
        public string UpdateTemplateTaskUserAuthority(int TemplateTaskUserId, int AuthorityId, bool Checked)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTaskUser
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務成員資料錯誤!");
                        #endregion

                        #region //判斷權限資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityId = @AuthorityId
                                AND AuthorityType = @AuthorityType";
                        dynamicParameters.Add("AuthorityId", AuthorityId);
                        dynamicParameters.Add("AuthorityType", "T");

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("權限資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        if (Checked)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PMIS.TemplateTaskUserAuthority (TemplateTaskUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TemplateTaskUserId, @AuthorityId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TemplateTaskUserId,
                                    AuthorityId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TemplateTaskUserAuthority
                                    WHERE TemplateTaskUserId = @TemplateTaskUserId
                                    AND AuthorityId = @AuthorityId";
                            dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);
                            dynamicParameters.Add("AuthorityId", AuthorityId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTemplateAllTaskUserAuthority
        public string UpdateTemplateAllTaskUserAuthority(int TemplateTaskUserId, bool Checked)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTaskUser
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案任務成員資料錯誤!");
                        #endregion

                        #region //判斷權限資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT AuthorityId, AuthorityType, AuthorityCode, AuthorityName, SortNumber, Status
                                FROM PMIS.Authority
                                WHERE AuthorityType = @AuthorityType";
                        dynamicParameters.Add("AuthorityType", "T");

                        var resultAuthority = sqlConnection.Query(sql, dynamicParameters);
                        if (resultAuthority.Count() <= 0) throw new SystemException("權限資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var item in resultAuthority)
                        {
                            if (Checked)
                            {
                                #region //新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PMIS.TemplateTaskUserAuthority (TemplateTaskUserId, AuthorityId
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        select @TemplateTaskUserId, a.AuthorityId, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy 
                                        from PMIS.Authority a 
                                        left join (select b.AuthorityId from PMIS.TemplateTaskUserAuthority b where b.TemplateTaskUserId = @TemplateTaskUserId) b on b.AuthorityId = a.AuthorityId
                                        where a.AuthorityType = @AuthorityType
                                        and b.AuthorityId is null
                                        order by a.SortNumber";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        AuthorityType = "T",
                                        TemplateTaskUserId,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PMIS.TemplateTaskUserAuthority
                                    WHERE TemplateTaskUserId = @TemplateTaskUserId";
                                dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTemplateTaskUserLevel -- 樣板任務成員閱覽檔案等級更新 -- Ben Ma 2022.07.25
        public string UpdateTemplateTaskUserLevel(int TemplateTaskUserId, int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTaskUser
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務成員資料錯誤!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        dynamicParameters.Add("LevelId", LevelId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.TemplateTaskUser SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy,
                                TemplateTaskUserId
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
        #region //DeleteTemplate -- 樣板資料刪除 -- Ben Ma 2022.07.15
        public string DeleteTemplate(int TemplateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");
                        #endregion

                        #region //判斷擁有者才能刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId
                                AND TemplateOwner = @TemplateOwner";
                        dynamicParameters.Add("TemplateId", TemplateId);
                        dynamicParameters.Add("TemplateOwner", CurrentUser);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("樣板擁有者才有刪除權限!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除樣板任務成員權限
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TemplateTaskUserAuthority a
                                INNER JOIN PMIS.TemplateTaskUser b ON a.TemplateTaskUserId = b.TemplateTaskUserId
                                INNER JOIN PMIS.TemplateTask c ON b.TemplateTaskId = c.TemplateTaskId
                                WHERE c.TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除樣板任務成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TemplateTaskUser a
                                INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                                WHERE b.TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除樣板任務連結資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TemplateTaskLink a
                                INNER JOIN PMIS.TemplateTask b ON a.SourceTaskId = b.TemplateTaskId
                                WHERE b.TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除樣板任務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTask
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除樣板協作人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateCooperation
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        dynamicParameters.Add("TemplateId", TemplateId);

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

        #region //DeleteTemplateCooperation -- 樣板協作人員資料刪除 -- Ben Ma 2022.08.17
        public string DeleteTemplateCooperation(int TemplateCooperationId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板協作人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateCooperation
                                WHERE TemplateCooperationId = @TemplateCooperationId";
                        dynamicParameters.Add("TemplateCooperationId", TemplateCooperationId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板協作人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateCooperation
                                WHERE TemplateCooperationId = @TemplateCooperationId";
                        dynamicParameters.Add("TemplateCooperationId", TemplateCooperationId);

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

        #region //DeleteTemplateTask -- 樣板任務資料刪除 -- Ben Ma 2022.07.15
        public string DeleteTemplateTask(int TemplateTaskId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //關聯任務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DECLARE @rowsAdded int

                                DECLARE @templateTask TABLE
                                ( 
                                    TemplateTaskId int,
                                    ParentTaskId int,
                                    TaskLevel int,
                                    TaskRoute nvarchar(MAX),
                                    TaskSort int,
                                    processed int DEFAULT(0)
                                )

                                INSERT @templateTask
                                    SELECT TemplateTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                                    , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) AS TaskRoute
                                    , TaskSort, 0
                                    FROM PMIS.TemplateTask
                                    WHERE ParentTaskId = @ParentTaskId

                                SET @rowsAdded=@@rowcount

                                WHILE @rowsAdded > 0
                                BEGIN
                                    UPDATE @templateTask SET processed = 1 WHERE processed = 0

                                    INSERT @templateTask
                                        SELECT b.TemplateTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                        , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) AS TaskRoute
                                        , b.TaskSort, 0
                                        FROM @templateTask a
                                        INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.ParentTaskId
                                        WHERE a.processed = 1

                                    SET @rowsAdded = @@rowcount

                                    UPDATE @templateTask SET processed = 2 WHERE processed = 1
                                END;

                                SELECT TemplateTaskId
                                FROM @templateTask
                                WHERE 1=1";
                        dynamicParameters.Add("ParentTaskId", TemplateTaskId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        if (result2.Count() > 0)
                        {
                            #region //刪除關聯任務連結
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TemplateTaskLink
                                    WHERE SourceTaskId IN @TemplateTask";
                            dynamicParameters.Add("TemplateTask", result2.Select(x => x.TemplateTaskId).ToArray());
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TemplateTaskLink
                                    WHERE TargetTaskId IN @TemplateTask";
                            dynamicParameters.Add("TemplateTask", result2.Select(x => x.TemplateTaskId).ToArray());
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除關聯任務成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM PMIS.TemplateTaskUserAuthority a
                                    INNER JOIN PMIS.TemplateTaskUser b ON a.TemplateTaskUserId = b.TemplateTaskUserId
                                    WHERE b.TemplateTaskId IN @TemplateTask";
                            dynamicParameters.Add("TemplateTask", result2.Select(x => x.TemplateTaskId).ToArray());
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TemplateTaskUser
                                    WHERE TemplateTaskId IN @TemplateTask";
                            dynamicParameters.Add("TemplateTask", result2.Select(x => x.TemplateTaskId).ToArray());
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除關聯任務
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TemplateTask
                                    WHERE TemplateTaskId IN @TemplateTask";
                            dynamicParameters.Add("TemplateTask", result2.Select(x => x.TemplateTaskId).ToArray());
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //刪除任務連結
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskLink
                                WHERE SourceTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskLink
                                WHERE TargetTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除任務成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TemplateTaskUserAuthority a
                                INNER JOIN PMIS.TemplateTaskUser b ON a.TemplateTaskUserId = b.TemplateTaskUserId
                                WHERE b.TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskUser
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTask
                                WHERE TemplateTaskId = @TemplateTaskId";
                        dynamicParameters.Add("TemplateTaskId", TemplateTaskId);

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

        #region //DeleteTemplateTaskLink -- 樣板任務連結資料刪除 -- Ben Ma 2022.08.15
        public string DeleteTemplateTaskLink(int TemplateTaskLinkId, int SourceTaskId, int TargetTaskId, string LinkType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務連結資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTaskLink
                                WHERE TemplateTaskLinkId = @TemplateTaskLinkId
                                AND SourceTaskId = @SourceTaskId
                                AND TargetTaskId = @TargetTaskId
                                AND LinkType = @LinkType";
                        dynamicParameters.Add("TemplateTaskLinkId", TemplateTaskLinkId);
                        dynamicParameters.Add("SourceTaskId", SourceTaskId);
                        dynamicParameters.Add("TargetTaskId", TargetTaskId);
                        dynamicParameters.Add("LinkType", LinkType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務連結資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskLink
                                WHERE TemplateTaskLinkId = @TemplateTaskLinkId";
                        dynamicParameters.Add("TemplateTaskLinkId", TemplateTaskLinkId);

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

        #region //DeleteTemplateTaskUser -- 樣板任務成員刪除 -- Ben Ma 2022.07.22
        public string DeleteTemplateTaskUser(int TemplateTaskUserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷樣板任務成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TemplateTaskUser
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("樣板任務成員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskUserAuthority
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TemplateTaskUser
                                WHERE TemplateTaskUserId = @TemplateTaskUserId";
                        dynamicParameters.Add("TemplateTaskUserId", TemplateTaskUserId);

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
