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
    public class TaskDA
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

        public TaskDA()
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
        #region //GetTaskChecklist -- 取得任務清單資料 -- Ben Ma 2022.11.01
        public string GetTaskChecklist(string Project, string TaskName, string TaskStatus)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int
                            DECLARE @projectTask TABLE
                            ( 
                                ProjectTaskId int,
                                ParentTaskId int,
                                TaskLevel int,
                                TaskRoute nvarchar(MAX),
                                TaskSort int,
                                TaskName nvarchar(MAX),
                                TaskRouteName nvarchar(MAX),
                                processed int DEFAULT(0)
                            )

                            INSERT @projectTask
                                SELECT ProjectTaskId, ParentTaskId, 1 TaskLevel
                                , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                                , TaskSort
                                , TaskName
                                , '' TaskRouteName
                                , 0
                                FROM PMIS.ProjectTask
                                WHERE ParentTaskId IS NULL

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @projectTask SET processed = 1 WHERE processed = 0

                                INSERT @projectTask
                                    SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                    , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                                    , b.TaskSort
                                    , b.TaskName
                                    , CAST((CASE WHEN LEN(a.TaskRouteName) > 0 THEN a.TaskRouteName + ' > ' ELSE '' END) + a.TaskName AS nvarchar(MAX)) TaskRouteName
                                    , 0
                                    FROM @projectTask a
                                    INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ParentTaskId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @projectTask SET processed = 2 WHERE processed = 1
                            END;";

                    sql += @"SELECT a.ProjectId, a.ProjectNo, a.ProjectName
                            , (
                                SELECT aa.ProjectTaskId, aa.TaskName, aa.TaskStatus
                                , ISNULL(FORMAT(aa.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                                , ISNULL(FORMAT(aa.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                                , ISNULL(FORMAT(aa.EstimateStart, 'yyyy-MM-dd HH:mm'), '') EstimateStart
                                , ISNULL(FORMAT(aa.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                                , ISNULL(FORMAT(aa.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                                , ISNULL(FORMAT(aa.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                                , ac.ParentTaskId, ac.TaskRouteName, ad.SubTask, ax.StatusName TaskStatusName
                                FROM PMIS.ProjectTask aa
                                INNER JOIN PMIS.TaskUser ab ON aa.ProjectTaskId = ab.TaskId
                                INNER JOIN @projectTask ac ON aa.ProjectTaskId = ac.ProjectTaskId
                                INNER JOIN BAS.[Status] ax ON ax.StatusNo = aa.TaskStatus AND ax.StatusSchema = 'ProjectTask.TaskStatus'
                                OUTER APPLY (
                                    SELECT COUNT(1) SubTask
                                    FROM PMIS.ProjectTask ada
                                    WHERE ada.ParentTaskId = aa.ProjectTaskId
                                ) ad
                                WHERE aa.ProjectId = a.ProjectId
                                AND ab.UserId = @UserId
                                AND ab.TaskType = 'Project'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND aa.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND aa.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    sql += @"   ORDER BY ISNULL(aa.ActualStart, ISNULL(aa.EstimateStart, aa.PlannedStart))
                                FOR JSON PATH, ROOT('data')
                            ) ProjectTask
                            FROM PMIS.Project a
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectTask aa
                                INNER JOIN PMIS.TaskUser ab ON aa.ProjectTaskId = ab.TaskId
                                WHERE aa.ProjectId = a.ProjectId
                                AND ab.UserId = @UserId
                                AND ab.TaskType = 'Project'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND aa.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND aa.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    sql += @")
                            AND a.ProjectStatus != 'P'";
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Project", @" AND (a.ProjectNo LIKE '%' + @Project + '%' OR a.ProjectName LIKE '%' + @Project + '%')", Project);

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

        #region //GetPersonalTaskChecklist --取得個人任務清單資料-New -- Chia Yuan 2023.12.13
        public string GetPersonalTaskChecklist(string TaskName, string TaskStatus)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PersonalTaskId, a.UserId, a.TaskName, a.TaskStatus
                            --, CASE WHEN @CreateDate between ISNULL(a.ActualStart, a.PlannedStart) AND ISNULL(a.ActualEnd, a.PlannedEnd) THEN 'I' ELSE a.TaskStatus END TaskStatus
                            , CASE WHEN a.UserId = @CurrentUser THEN 'Master' ELSE 'General' END UserRole
                            , ISNULL(FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                            , ISNULL(FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                            , ISNULL(FORMAT(ISNULL(a.ActualStart, a.PlannedStart), 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(ISNULL(a.ActualEnd, a.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , (
                                SELECT ub.UserName,ua.AgentSort
	                            FROM PMIS.TaskUser ua
	                            INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                            WHERE ua.TaskId = a.PersonalTaskId 
                                AND ua.TaskType = 'Personal'
                                ORDER BY ua.AgentSort
                                FOR JSON PATH, ROOT('data')
                            ) TaskUsers
                            , ISNULL(SUBSTRING(c.TaskUser, 0, LEN(c.TaskUser)), '') TaskUser
                            , d.StatusName TaskStatusName
                            FROM PMIS.PersonalTask a
                            INNER JOIN BAS.[Status] d ON d.StatusNo = a.TaskStatus AND d.StatusSchema = 'ProjectTask.TaskStatus'
                            OUTER APPLY (
                                SELECT (
                                    SELECT cb.UserName + ','
                                    FROM PMIS.TaskUser ca
                                    INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
                                    WHERE ca.TaskId = a.PersonalTaskId 
                                    AND ca.TaskType = 'Personal'
                                    ORDER BY ca.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) c
                            WHERE EXISTS (
	                            SELECT TOP 1 1
                                FROM PMIS.TaskUser ab
                                WHERE ab.TaskId = a.PersonalTaskId 
	                            AND (ab.UserId = @CurrentUser OR a.UserId = @CurrentUser)
                                AND ab.TaskType = 'Personal'
                            )";
                    dynamicParameters.Add("CreateDate", CreateDate);
                    dynamicParameters.Add("CurrentUser", CurrentUser);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND a.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND a.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
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

        #region //GetTaskStructure -- 取得任務結構資料 -- Ben Ma 2022.11.02
        public string GetTaskStructure(string Project, int ProjectTaskId, string TaskName, string TaskStatus, int OpenedLevel)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<TaskList> taskLists = new List<TaskList>();

                    #region //使用者任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ProjectId, a.ProjectNo, a.ProjectName
                            , (
                                SELECT aa.ProjectTaskId, aa.TaskName, aa.TaskStatus
                                FROM PMIS.ProjectTask aa
                                INNER JOIN PMIS.TaskUser ab ON aa.ProjectTaskId = ab.TaskId
                                WHERE aa.ProjectId = a.ProjectId
                                AND ab.TaskType = 'Project'
                                AND ab.UserId = @UserId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", @" AND aa.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND aa.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND aa.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    sql += @"   FOR JSON PATH
                            ) ProjectTask
                            FROM PMIS.Project a
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectTask aa
                                INNER JOIN PMIS.TaskUser ab ON aa.ProjectTaskId = ab.TaskId
                                WHERE aa.ProjectId = a.ProjectId
                                AND ab.TaskType = 'Project'
                                AND ab.UserId = @UserId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", @" AND aa.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND aa.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                    if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND aa.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    sql += @")
                            AND a.ProjectStatus != 'P'";
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Project", @" AND (a.ProjectNo LIKE '%' + @Project + '%' OR a.ProjectName LIKE '%' + @Project + '%')", Project);

                    taskLists = sqlConnection.Query<TaskList>(sql, dynamicParameters).ToList();
                    #endregion

                    List<DHXTree> trees = new List<DHXTree>();
                    foreach (var project in taskLists)
                    {
                        var userTask = JsonConvert.DeserializeObject<List<ProjectTask>>(project.ProjectTask);

                        List<ProjectTask> projectTasks = new List<ProjectTask>();

                        #region //專案任務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DECLARE @rowsAdded int

                                DECLARE @projectTask TABLE
                                ( 
                                    ProjectTaskId int,
                                    ParentTaskId int,
                                    TaskLevel int,
                                    TaskRoute nvarchar(MAX),
                                    TaskSort int,
                                    processed int DEFAULT(0)
                                )

                                INSERT @projectTask
                                    SELECT ProjectTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                                    , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                                    , TaskSort, 0
                                    FROM PMIS.ProjectTask
                                    WHERE ProjectId = @ProjectId
                                    AND ParentTaskId IS NULL

                                SET @rowsAdded=@@rowcount

                                WHILE @rowsAdded > 0
                                BEGIN
                                    UPDATE @projectTask SET processed = 1 WHERE processed = 0

                                    INSERT @projectTask
                                        SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                        , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                                        , b.TaskSort, 0
                                        FROM @projectTask a
                                        INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ParentTaskId
                                        WHERE a.processed = 1

                                    SET @rowsAdded = @@rowcount

                                    UPDATE @projectTask SET processed = 2 WHERE processed = 1
                                END;

                                SELECT a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort
                                , b.TaskName
                                , c.UserTask
                                FROM @projectTask a
                                INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM PMIS.TaskUser ca
                                        INNER JOIN PMIS.ProjectTask cb ON ca.TaskId = cb.ProjectTaskId
                                        WHERE ca.TaskId = b.ProjectTaskId
                                        AND ca.TaskType = 'Project'
                                        AND ca.UserId = @UserId";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", @" AND cb.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskName", @" AND cb.TaskName LIKE '%' + @TaskName + '%'", TaskName);
                        if (TaskStatus.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskStatus", @" AND cb.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                        sql += @"   ), 0) UserTask
                                ) c
                                ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                        dynamicParameters.Add("ProjectId", project.ProjectId);
                        dynamicParameters.Add("UserId", CurrentUser);

                        projectTasks = sqlConnection.Query<ProjectTask>(sql, dynamicParameters).ToList();

                        int maxLevel = (int)projectTasks.Max(x => x.TaskLevel);

                        while (maxLevel > 0)
                        {
                            var currentTask = projectTasks.Where(x => x.TaskLevel == maxLevel).OrderBy(x => x.TaskRoute);

                            foreach (var task in currentTask)
                            {
                                if (task.UserTask == 1)
                                {
                                    projectTasks
                                        .Where(x => x.ProjectTaskId == task.ProjectTaskId)
                                        .ToList()
                                        .ForEach(x =>
                                        {
                                            x.DisplayTask = 1;
                                        });
                                }
                                else
                                {
                                    if (projectTasks.Exists(x => x.ParentTaskId == Convert.ToInt32(task.ProjectTaskId) && x.DisplayTask == 1))
                                    {
                                        projectTasks
                                            .Where(x => x.ProjectTaskId == task.ProjectTaskId)
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.DisplayTask = 1;
                                            });
                                    }
                                }
                            }

                            maxLevel--;
                        }
                        #endregion

                        var data = new DHXTree
                        {
                            value = project.ProjectName,
                            id = "-1",
                            opened = true
                        };

                        if (projectTasks.Count > 0)
                        {
                            data.items = projectTasks
                                .Where(x => x.TaskLevel == 1 && x.ParentTaskId == Convert.ToInt32(data.id) && x.DisplayTask == 1)
                                .OrderBy(x => x.TaskLevel)
                                .ThenBy(x => x.TaskRoute)
                                .ThenBy(x => x.TaskSort)
                                .Select(x => new DHXTree
                                {
                                    value = x.TaskName,
                                    id = x.ProjectTaskId.ToString(),
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
                                    taskTree[i].items = projectTasks
                                        .Where(x => x.TaskLevel == (taskTree[i].level + 1) && x.ParentTaskId == Convert.ToInt32(taskTree[i].id) && x.DisplayTask == 1)
                                        .OrderBy(x => x.TaskLevel)
                                        .ThenBy(x => x.TaskRoute)
                                        .ThenBy(x => x.TaskSort)
                                        .Select(x => new DHXTree
                                        {
                                            value = x.TaskName,
                                            id = x.ProjectTaskId.ToString(),
                                            opened = taskTree[i].level < OpenedLevel ? true : false,
                                            level = (int)x.TaskLevel
                                        })
                                        .ToList();

                                    if (taskTree[i].items.Count > 0) Recursion(taskTree[i].items);
                                }
                            }
                        }

                        taskLists
                            .Where(x => x.ProjectId == project.ProjectId)
                            .ToList()
                            .ForEach(x =>
                            {
                                x.TaskTree = data;
                            });
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = taskLists
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

        #region //GetUserProject -- 取得使用者專案資料 -- Ben Ma 2022.11.04
        public string GetUserProject()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ProjectId, a.ProjectNo + ' ' + a.ProjectName ProjectWithNo
                            FROM PMIS.Project a
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectTask aa
                                INNER JOIN PMIS.TaskUser ab ON aa.ProjectTaskId = ab.TaskId
                                WHERE aa.ProjectId = a.ProjectId
                                AND ab.TaskType = 'Project'
                                AND ab.UserId = @CurrentUser
                            )
                            AND a.ProjectStatus != 'P'";
                    var result = sqlConnection.Query(sql, new { CurrentUser });

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

        #region //GetUserTask -- 取得使用者任務資料 -- Ben Ma 2022.11.07
        public string GetUserTask(int ProjectId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ProjectTaskId, a.TaskName
                            FROM PMIS.ProjectTask a
                            INNER JOIN PMIS.Project b ON a.ProjectId = b.ProjectId
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUser aa
                                WHERE aa.TaskId = a.ProjectTaskId
                                AND aa.UserId = @UserId
                                AND aa.TaskType = 'Project'
                            )
                            AND b.ProjectStatus != 'P'";
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);

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

        #region //GetUserPersonalTask -- 取得使用者個人任務資料-New -- Chia Yuan 2023.12.14
        public string GetUserPersonalTask(int PersonalTaskId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PersonalTaskId, a.TaskName
                            FROM PMIS.PersonalTask a
                            INNER JOIN PMIS.TaskUser b ON b.TaskId = a.PersonalTaskId
                            WHERE b.UserId = @UserId
                            AND b.TaskType = 'Personal'";
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

        #region GetPersonalTask -- 取得個人任務資料-New -- Chia Yuan 2023.12.05
        public string GetPersonalTask(int PersonalTaskId, int UserId, int CompanyId, int DepartmentId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得個人任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PersonalTaskId,a.TaskName, a.TaskDesc
                            , a.TaskStatus
                            , ISNULL(FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(a.PlannedDuration, 0) PlannedDuration
                            , ISNULL(FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                            , ISNULL(FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                            , ISNULL(a.ActualDuration, 0) ActualDuration
                            , ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                            , s.StatusName
                            , 0 SubTask
                            , 'Personal' TaskType
                            , (
                                SELECT ub.UserName,ua.AgentSort,
	                            (
		                            SELECT vb.AuthorityName FROM PMIS.TaskUserAuthority va
		                            INNER JOIN PMIS.Authority vb ON vb.AuthorityId = va.AuthorityId
		                            WHERE va.TaskUserId = ua.TaskUserId
		                            ORDER BY vb.SortNumber
		                            FOR JSON PATH, ROOT('data')
	                            ) Authoritys
	                            FROM PMIS.TaskUser ua
	                            INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                            WHERE ua.TaskId = a.PersonalTaskId 
                                AND ua.TaskType = 'Personal'
                                ORDER BY ua.AgentSort
                                FOR JSON PATH, ROOT('data')
                            ) TaskUsers
                            , SUBSTRING(u.TaskUser, 0, LEN(u.TaskUser)) TaskUser
                            FROM PMIS.PersonalTask a
                            INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                            INNER JOIN BAS.Department d ON d.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                            INNER JOIN BAS.[Status] s ON s.StatusNo = a.TaskStatus AND s.StatusSchema = 'ProjectTask.TaskStatus'
                            OUTER APPLY (
                                SELECT (
                                    SELECT cb.UserName + ','
                                    FROM PMIS.TaskUser ca
                                    INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
                                    WHERE ca.TaskId = a.PersonalTaskId 
                                    AND ca.TaskType = 'Personal'
                                    ORDER BY ca.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) u
                            WHERE a.PersonalTaskId = @PersonalTaskId
                            ORDER BY ISNULL(a.ActualStart, a.PlannedStart) DESC";
                    dynamicParameters.Add("PersonalTaskId", PersonalTaskId);
                    dynamicParameters.Add("dtNow", CreateDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", @" AND d.DepartmentId = @DepartmentId", DepartmentId);

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

        #region //GetUserTaskCalendar --取得行事曆任務資料-New -- Chia Yuan 2023.11.29
        public string GetUserTaskCalendar(int ProjectId, int TaskId, string TaskType)
        {
            try
            {
                if (!string.IsNullOrEmpty(TaskType) && !Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【任務類型】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int
                            DECLARE @projectTask TABLE
                            ( 
                                ProjectTaskId int,
                                ParentTaskId int,
                                TaskLevel int,
                                TaskRoute nvarchar(MAX),
                                TaskSort int,
                                TaskName nvarchar(MAX),
                                TaskRouteName nvarchar(MAX),
                                processed int DEFAULT(0)
                            )

                            INSERT @projectTask
                                SELECT a.ProjectTaskId, a.ParentTaskId, 1 TaskLevel
                                , CAST(ISNULL(a.ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                                , a.TaskSort
                                , a.TaskName
                                , '' TaskRouteName
                                , 0
                                FROM PMIS.ProjectTask a
	                            JOIN PMIS.TaskUser b ON b.TaskId = a.ProjectTaskId
	                            JOIN PMIS.Project c ON c.ProjectId = a.ProjectId
                                WHERE b.UserId = @CurrentUser
                                AND b.TaskType = 'Project'
	                            AND c.ProjectStatus <> 'P'

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @projectTask SET processed = 1 WHERE processed = 0

                                INSERT @projectTask
                                    SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                    , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                                    , b.TaskSort
                                    , b.TaskName
                                    , CAST((CASE WHEN LEN(a.TaskRouteName) > 0 THEN a.TaskRouteName + ' > ' ELSE '' END) + a.TaskName AS nvarchar(MAX)) TaskRouteName
                                    , 0
                                    FROM @projectTask a
                                    INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ParentTaskId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @projectTask SET processed = 2 WHERE processed = 1
                            END

                            SELECT x.TaskRouteName, a.ProjectTaskId, a.TaskStatus, a.TaskName, ISNULL(a.TaskDesc, '') TaskDesc
                            , ISNULL(FORMAT(ISNULL(a.ActualStart, a.EstimateStart), 'yyyy-MM-dd HH:mm'), '') EstimateStart 
                            , ISNULL(FORMAT(ISNULL(a.ActualEnd, a.EstimateEnd), 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                            , ISNULL(ISNULL(a.ActualDuration, a.EstimateDuration), 1) EstimateDuration
                            , ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                            , d.StatusName TaskStatusName
                            , e.SubTask
                            , 'Project' TaskType
                            , (
                                SELECT ub.UserName,ua.AgentSort,
	                            (
		                            SELECT vb.AuthorityName 
                                    FROM PMIS.TaskUserAuthority va
		                            INNER JOIN PMIS.Authority vb ON vb.AuthorityId = va.AuthorityId
		                            WHERE va.TaskUserId = ua.TaskUserId
		                            ORDER BY vb.SortNumber
		                            FOR JSON PATH, ROOT('data')
	                            ) Authoritys
	                            FROM PMIS.TaskUser ua
	                            INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                            WHERE ua.TaskId = a.ProjectTaskId 
                                AND ua.TaskType = 'Project'
                                ORDER BY ua.AgentSort
                                FOR JSON PATH, ROOT('data')
                            ) TaskUsers
                            , ISNULL(SUBSTRING(c.TaskUser, 0, LEN(c.TaskUser)), '') TaskUser
                            FROM @projectTask x
                            INNER JOIN PMIS.ProjectTask a ON a.ProjectTaskId = x.ProjectTaskId
                            INNER JOIN PMIS.Project b ON a.ProjectId = b.ProjectId
                            INNER JOIN BAS.[Status] d ON d.StatusNo = a.TaskStatus AND d.StatusSchema = 'ProjectTask.TaskStatus'
                            OUTER APPLY (
                                SELECT (
                                    SELECT cb.UserName + ','
                                    FROM PMIS.TaskUser ca
                                    INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
                                    WHERE ca.TaskId = a.ProjectTaskId 
                                    AND ca.TaskType = 'Project'
                                    ORDER BY ca.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) c
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ParentTaskId = a.ProjectTaskId
                            ) e
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUser aa
                                WHERE aa.TaskId = a.ProjectTaskId
                                AND aa.TaskType = 'Project'
                                AND aa.UserId = @CurrentUser
                            )
                            AND b.ProjectStatus <> 'P'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", @" AND a.ProjectTaskId = @ProjectTaskId", TaskId);

                    if (TaskType != "Project")
                    {
                        sql += @" UNION
                            SELECT null,a.PersonalTaskId
                            , a.TaskStatus
                            , a.TaskName, ISNULL(a.TaskDesc, '')
                            , ISNULL(FORMAT(ISNULL(a.ActualStart, a.PlannedStart),'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(ISNULL(a.ActualEnd, a.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(a.ActualDuration, a.PlannedDuration) PlannedDuration
                            , ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                            , d.StatusName
                            , 0 SubTask
                            , 'Personal'
                            , (
                                SELECT ub.UserName,ua.AgentSort,
	                            (
		                            SELECT vb.AuthorityName FROM PMIS.TaskUserAuthority va
		                            INNER JOIN PMIS.Authority vb ON vb.AuthorityId = va.AuthorityId
		                            WHERE va.TaskUserId = ua.TaskUserId
		                            ORDER BY vb.SortNumber
		                            FOR JSON PATH, ROOT('data')
	                            ) Authoritys
	                            FROM PMIS.TaskUser ua
	                            INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                            WHERE ua.TaskId = a.PersonalTaskId 
                                AND ua.TaskType = 'Personal'
                                ORDER BY ua.AgentSort
                                FOR JSON PATH, ROOT('data')
                            ) TaskUsers
                            , ISNULL(SUBSTRING(c.TaskUser, 0, LEN(c.TaskUser)), '') TaskUser
                            FROM PMIS.PersonalTask a
                            INNER JOIN BAS.[Status] d ON d.StatusNo = a.TaskStatus AND d.StatusSchema = 'ProjectTask.TaskStatus'
                            OUTER APPLY (
                                SELECT (
                                    SELECT cb.UserName + ','
                                    FROM PMIS.TaskUser ca
                                    INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
                                    WHERE ca.TaskId = a.PersonalTaskId 
                                    AND ca.TaskType = 'Personal'
                                    ORDER BY ca.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) c
                            WHERE EXISTS (
	                            SELECT TOP 1 1
                                FROM PMIS.TaskUser ab
                                WHERE ab.TaskId = a.PersonalTaskId 
	                            AND (ab.UserId = @CurrentUser OR a.UserId = @CurrentUser)
                                AND ab.TaskType = 'Personal'
                            )";
                    }

                    dynamicParameters.Add("CurrentUser", CurrentUser);
                    sql += " ORDER BY EstimateStart, EstimateDuration, a.TaskName";
                    var projectTasks = sqlConnection.Query(sql, dynamicParameters).ToList();

                    List<DHXProjectTask> tasks = new List<DHXProjectTask>();

                    foreach (var s in projectTasks)
                    {
                        tasks.Add(new DHXProjectTask
                        {
                            id = s.ProjectTaskId.ToString(),
                            type = s.TaskStatus,
                            text = s.TaskName,
                            details = s.TaskDesc,
                            start_date = s.EstimateStart,
                            end_date = s.EstimateEnd,
                            duration = s.EstimateDuration.ToString(),
                            ReplyFrequency = s.ReplyFrequency,
                            task_status_name = s.TaskStatusName,
                            TaskUserData = s.TaskUsers,
                            TaskUser = s.TaskUser,
                            SubTask = s.SubTask,
                            TaskRouteName = s.TaskRouteName,
                            TaskType = s.TaskType
                        });
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = tasks
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

        #region //GetLastReply -- 取得專案任務最後回報資料 -- Ben Ma 2022.11.16
        public string GetLastReply(int TaskId, string ReplyType)
        {
            try
            {
                if (!Regex.IsMatch(ReplyType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("回報類型錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    switch (ReplyType)
                    {
                        case "Project":
                            #region //取得專案任務資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DATEDIFF(MINUTE, ISNULL(b.ReplyDate, a.EstimateStart), GETDATE()) LastReplyInterval
                                    , a.EstimateStart, a.EstimateEnd, a.EstimateDuration, ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                                    FROM PMIS.ProjectTask a
                                    OUTER APPLY (
                                        SELECT TOP 1 ba.ReplyDate
                                        FROM PMIS.TaskReply ba
                                        WHERE ba.ReplyType = @ReplyType
                                        AND ba.TaskId = a.ProjectTaskId
                                        ORDER BY ba.ReplyDate
                                    ) b
                                    WHERE a.ProjectTaskId = @TaskId";
                            dynamicParameters.Add("ReplyType", ReplyType);
                            dynamicParameters.Add("TaskId", TaskId);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result1
                            });
                            #endregion
                            break;
                        case "Personal":
                            #region //取得個人任務資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DATEDIFF(MINUTE, ISNULL(b.ReplyDate, a.PlannedStart), GETDATE()) LastReplyInterval
                                    , a.ActualStart, a.ActualEnd, a.ActualDuration, ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                                    FROM PMIS.PersonalTask a
                                    OUTER APPLY (
                                        SELECT TOP 1 ba.ReplyDate
                                        FROM PMIS.TaskReply ba
                                        WHERE ba.ReplyType = @ReplyType
                                        AND ba.TaskId = a.PersonalTaskId
                                        ORDER BY ba.ReplyDate
                                    ) b
                                    WHERE a.PersonalTaskId = @PersonalTaskId";
                            dynamicParameters.Add("ReplyType", ReplyType);
                            dynamicParameters.Add("PersonalTaskId", TaskId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result2
                            });
                            #endregion
                            break;
                        default:
                            break;
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

        #region //GetTaskReply -- 取得任務回報內容資料 -- Ben Ma 2022.10.03
        public string GetTaskReply(int TaskId, string ReplyType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (!Regex.IsMatch(ReplyType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得任務回報資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ReplyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TaskId, a.ReplyType, a.ReplyContent, a.ReplyUser, a.ReplyStatus, a.DeferredDuration
                        , FORMAT(a.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                        , ISNULL(FORMAT(a.EventDate, 'yyyy-MM-dd HH:mm'), '') EventDate
                        , ISNULL(a.ReplyFile, '') ReplyFile, ISNULL(r.[FileName], '') [FileName], ISNULL(r.FileExtension, '') FileExtension, ISNULL(r.FileSize, 0) FileSize
                        , b.UserNo ReplyUserNo, b.UserName ReplyUserName
                        , d.LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TaskReply a
                        INNER JOIN BAS.[User] b ON a.ReplyUser = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                        OUTER APPLY (SELECT TOP 1 aa.[FileName], aa.FileExtension, aa.FileSize FROM BAS.[File] aa WHERE aa.FileId = a.ReplyFile) r";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @"AND a.TaskId = @TaskId
                        AND a.ReplyType = @ReplyType";
                    dynamicParameters.Add("TaskId", TaskId);
                    dynamicParameters.Add("ReplyType", ReplyType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ReplyId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
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

        #region //GetProjectTaskFile -- 取得使用者可存取檔案資料 -- Ben Ma 2022.11.07
        public string GetProjectTaskFile(int ProjectId, int TaskId, string FileName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得專案任務檔案資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FileId, a.FileType, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                            , b.[FileName], b.FileExtension, b.FileSize
                            , c.LevelName
                            , '' TaskName
                            , d.ProjectName
                            , a.ProjectId SourceId, 'P' SourceType
                            FROM PMIS.ProjectFile a
                            INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                            INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                            INNER JOIN PMIS.Project d ON a.ProjectId = d.ProjectId
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectUser aa
                                INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                                WHERE aa.ProjectId = a.ProjectId
                                AND aa.UserId = @UserId
                                AND ab.SortNumber <= c.SortNumber
                            )
                            AND d.ProjectStatus != 'P'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND d.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FileName", @" AND b.FileName LIKE '%' + @FileName + '%'", FileName);
                    sql += @" UNION ALL
                            SELECT a.FileId, a.FileType, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                            , b.[FileName], b.FileExtension, b.FileSize
                            , c.LevelName
                            , d.TaskName
                            , e.ProjectName
                            , a.ProjectTaskId SourceId, 'T' SourceType
                            FROM PMIS.ProjectTaskFile a
                            INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                            INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                            INNER JOIN PMIS.ProjectTask d ON a.ProjectTaskId = d.ProjectTaskId
                            INNER JOIN PMIS.Project e ON d.ProjectId = e.ProjectId
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUser aa
                                INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                                WHERE aa.TaskId = a.ProjectTaskId
                                AND aa.UserId = @UserId
                                AND aa.TaskType = 'Project'
                                AND ab.SortNumber <= c.SortNumber
                            )
                            AND EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUserAuthority aa
                                INNER JOIN PMIS.TaskUser ab ON aa.TaskUserId = ab.TaskUserId
                                INNER JOIN PMIS.Authority ac ON aa.AuthorityId = ac.AuthorityId
                                WHERE ab.TaskId = a.ProjectTaskId
                                AND ab.UserId = @UserId
                                AND ab.TaskType = 'Project'
                                AND ac.AuthorityCode = 'task-file-download'
                            )
                            AND e.ProjectStatus != 'P'";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND e.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaskId", @" AND d.ProjectTaskId = @TaskId", TaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FileName", @" AND b.FileName LIKE '%' + @FileName + '%'", FileName);
                    dynamicParameters.Add("UserId", CurrentUser);


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

        #region //GetPersonalTaskFile -- 取得使用者可存取檔案資料-New -- Chia Yuan 2023.12.14
        public string GetPersonalTaskFile(int PersonalTaskId, string FileName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得個人任務檔案資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PersonalTaskFileId, a.FileId, a.FileType, a.LevelId
                            , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                            , b.[FileName], b.FileExtension, b.FileSize
                            , c.LevelName
                            , d.TaskName
                            , a.PersonalTaskId SourceId, 'T' SourceType
                            , SUBSTRING(u.TaskUser, 0, LEN(u.TaskUser)) TaskUser
                            FROM PMIS.PersonalTaskFile a
                            INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                            INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                            INNER JOIN PMIS.PersonalTask d ON a.PersonalTaskId = d.PersonalTaskId
                            OUTER APPLY (
                                SELECT (
                                    SELECT cb.UserName + ','
                                    FROM PMIS.TaskUser ca
                                    INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
                                    WHERE ca.TaskId = a.PersonalTaskId 
                                    AND ca.TaskType = 'Personal'
                                    ORDER BY ca.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) u
                            WHERE EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUser aa
                                INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                                WHERE aa.TaskId = a.PersonalTaskId
                                AND aa.UserId = @UserId
                                AND ab.SortNumber <= c.SortNumber
                                AND aa.TaskType = 'Personal'
                            )
                            AND EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUserAuthority aa
                                INNER JOIN PMIS.TaskUser ab ON aa.TaskUserId = ab.TaskUserId
                                INNER JOIN PMIS.Authority ac ON aa.AuthorityId = ac.AuthorityId
                                WHERE ab.TaskId = a.PersonalTaskId
                                AND ab.UserId = @UserId
                                AND ac.AuthorityCode = 'task-file-download'
                                AND ab.TaskType = 'Personal'
                            )";

                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PersonalTaskId", @" AND d.PersonalTaskId = @PersonalTaskId", PersonalTaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FileName", @" AND b.FileName LIKE '%' + @FileName + '%'", FileName);

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

        #region //GetTaskGantt -- 取得任務資料(甘特圖) -- Ben Ma 2022.11.09
        public string GetTaskGantt(int TaskId, string TaskType)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    switch (TaskType)
                    {
                        case "Project":
                            List<ProjectTask> projectTasks = new List<ProjectTask>();
                            List<ProjectTaskLink> projectTaskLinks = new List<ProjectTaskLink>();

                            #region //專案資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ProjectId, ProjectName
                                    , ISNULL(FORMAT(EstimateStart, 'yyyy-MM-dd HH:mm'), '') EstimateStart
                                    , ISNULL(FORMAT(EstimateEnd, 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                                    , ISNULL(EstimateDuration, 1) EstimateDuration
                                    FROM PMIS.Project
                                    WHERE ProjectId = (SELECT ProjectId FROM PMIS.ProjectTask WHERE ProjectTaskId = @TaskId)";
                            dynamicParameters.Add("TaskId", TaskId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("專案資料錯誤!");

                            int projectId = -1;
                            string projectName = ""
                                , projectEstimateStart = "", projectEstimateEnd = "";

                            foreach (var item in result)
                            {
                                projectId = item.ProjectId;
                                projectName = item.ProjectName;
                                projectEstimateStart = item.EstimateStart;
                                projectEstimateEnd = item.EstimateEnd;
                            }
                            #endregion

                            #region //專案任務資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"DECLARE @rowsAdded int

                                    DECLARE @projectTask TABLE
                                    ( 
                                        ProjectTaskId int,
                                        ParentTaskId int,
                                        TaskLevel int,
                                        TaskRoute nvarchar(MAX),
                                        TaskSort int,
                                        processed int DEFAULT(0)
                                    )

                                    INSERT @projectTask
                                        SELECT ProjectTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                                        , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                                        , TaskSort, 0
                                        FROM PMIS.ProjectTask
                                        WHERE ProjectId = @ProjectId
                                        AND ParentTaskId IS NULL

                                    SET @rowsAdded=@@rowcount

                                    WHILE @rowsAdded > 0
                                    BEGIN
                                        UPDATE @projectTask SET processed = 1 WHERE processed = 0

                                        INSERT @projectTask
                                            SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                                            , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                                            , b.TaskSort, 0
                                            FROM @projectTask a
                                            INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ParentTaskId
                                            WHERE a.processed = 1

                                        SET @rowsAdded = @@rowcount

                                        UPDATE @projectTask SET processed = 2 WHERE processed = 1
                                    END;

                                    SELECT a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort
                                    , b.TaskName
                                    , ISNULL(b.ActualStart, b.EstimateStart) StartDate
                                    , ISNULL(b.ActualEnd, b.EstimateEnd) EndDate
                                    , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), 1) Duration
                                    , c.SubTask
                                    , d.DeferredStatus
                                    , b.TaskStatus
                                    , e.StatusName TaskStatusName
                                    FROM @projectTask a
                                    INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                                    OUTER APPLY (
                                        SELECT COUNT(1) SubTask
                                        FROM PMIS.ProjectTask ca
                                        WHERE ca.ParentTaskId = a.ProjectTaskId
                                    ) c
                                    OUTER APPLY (
                                        SELECT COUNT(1) DeferredStatus
                                        FROM PMIS.TaskReply da
                                        WHERE da.TaskId = a.ProjectTaskId
                                        AND da.ReplyType = 'Project'
                                        AND da.ReplyStatus = 'D'
                                    ) d
                                    OUTER APPLY (
                                        SELECT ea.StatusName FROM BAS.[Status] ea WHERE ea.StatusSchema = 'ProjectTask.TaskStatus' AND ea.StatusNo = b.TaskStatus
                                    ) e
                                    WHERE 1=1
                                    ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                            dynamicParameters.Add("ProjectId", projectId);

                            projectTasks = sqlConnection.Query<ProjectTask>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //專案任務連結資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectTaskLinkId, a.SourceTaskId, a.TargetTaskId, a.LinkType
                                    FROM PMIS.ProjectTaskLink a
                                    INNER JOIN PMIS.ProjectTask b ON a.SourceTaskId = b.ProjectTaskId
                                    WHERE b.ProjectId = @ProjectId
                                    ORDER BY a.ProjectTaskLinkId";
                            dynamicParameters.Add("ProjectId", projectId);

                            projectTaskLinks = sqlConnection.Query<ProjectTaskLink>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //組成任務json
                            List<DHXTask> dHXTasks = new List<DHXTask>
                            {
                                new DHXTask
                                {
                                    id = "1.1",
                                    text = projectName,
                                    start_date = projectEstimateStart,
                                    end_date = projectEstimateEnd,
                                    progress = "0",
                                    parent = "0",
                                    type = "project",
                                    order = 1
                                }
                            };

                            foreach (var task in projectTasks)
                            {
                                DateTime startDate = Convert.ToDateTime(task.StartDate);
                                int duration = Convert.ToInt32(task.Duration);
                                DateTime endDate = Convert.ToDateTime(task.EndDate);

                                dHXTasks.Add(
                                    new DHXTask
                                    {
                                        id = task.ProjectTaskId.ToString(),
                                        text = task.TaskName,
                                        start_date = startDate.ToString("yyyy-MM-dd HH:mm"),
                                        end_date = task.TaskStatus == "C" ? endDate.ToString("yyyy-MM-dd HH:mm") : null,
                                        duration = duration.ToString(),
                                        progress = "0",
                                        parent = task.ParentTaskId > -1 ? task.ParentTaskId.ToString() : "1.1",
                                        type = task.SubTask > 0 ? "project" : "task",
                                        order = Convert.ToInt32(task.TaskSort),
                                        sub_user = task.SubUser,
                                        task_status = task.TaskStatus,
                                        task_status_name = task.TaskStatusName
                                    });
                            }
                            #endregion

                            #region //組成任務連結json
                            List<DHXLink> dHXLinks = new List<DHXLink>();

                            foreach (var link in projectTaskLinks)
                            {
                                dHXLinks.Add(
                                    new DHXLink
                                    {
                                        id = link.ProjectTaskLinkId.ToString(),
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
                            break;
                        default:
                            break;
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

        #region //Add
        #region //AddTaskReply -- 任務回報內容新增 -- Ben Ma 2022.11.17
        public string AddTaskReply(int TaskId, string ReplyType, string ReplyStatus, string ActualStart, string ActualEnd
            , int DeferredDuration, string DurationUnit, int ReplyFile, string ReplyContent)
        {
            try
            {
                if (ReplyType.Length <= 0) throw new SystemException("【回報類型】不能為空!");
                if (ReplyStatus.Length <= 0) throw new SystemException("【回報狀態】不能為空!");
                if (!Regex.IsMatch(ReplyType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【回報類型】錯誤!");
                if (!Regex.IsMatch(ReplyStatus, "^(C|D|N|O|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【回報狀態】錯誤!");
                if (ReplyStatus == "D")
                {
                    if (DeferredDuration <= 0) throw new SystemException("【延後時間】必填且需大於0");
                    if (DurationUnit.Length <= 0) throw new SystemException("【時間單位】不能為空!");
                    if (ReplyContent.Length <= 0) throw new SystemException("【延宕原因】不能為空!");
                }
                if (ReplyStatus == "N")
                {
                    if (ReplyContent.Length <= 0) throw new SystemException("【尚未開始原因】不能為空!");
                }
                if (!string.IsNullOrWhiteSpace(ReplyContent)) ReplyContent = ReplyContent.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷回報類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusNo = @ReplyStatus
                                AND StatusSchema = 'TaskReply.ReplyStatus'";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ReplyStatus }) ?? throw new SystemException("【回報類型】資料錯誤!");
                        #endregion

                        bool canReply = false;
                        bool requiredFile = false;
                        int ProjectId = -1;
                        string taskStatus = string.Empty;
                        string estimateStart = string.Empty;
                        string estimateEnd = string.Empty;
                        string actualStart = string.Empty;
                        string projectName = string.Empty;
                        string taskName = string.Empty;
                        string projectAttribute = string.Empty;

                        switch (ReplyType)
                        {
                            case "Project":
                                ProjectDA projectDA = new ProjectDA();

                                #region //取得專案任務資料
                                sql = @"SELECT TOP 1 a.ProjectId
                                        , ISNULL(a.ParentTaskId, -1) ParentTaskId
                                        , a.TaskName, a.TaskStatus, a.RequiredFile
                                        , FORMAT(a.EstimateStart, 'yyyy-MM-dd HH:mm') EstimateStart
                                        , FORMAT(a.EstimateEnd, 'yyyy-MM-dd HH:mm') EstimateEnd
                                        , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                        , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
                                        , b.ProjectName, b.ProjectAttribute
                                        , c.UserId TaskUserId
                                        FROM PMIS.ProjectTask a
                                        INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                        INNER JOIN PMIS.TaskUser c ON c.TaskId = a.ProjectTaskId AND c.TaskType = 'Project'
                                        WHERE ProjectTaskId = @TaskId
                                        AND c.UserId = @CurrentUser";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { TaskId, CurrentUser }) ?? throw new SystemException("【專案】資料錯誤!");
                                #endregion

                                ProjectId = result?.ProjectId;
                                projectName = result?.ProjectName;
                                taskName = result?.TaskName;
                                taskStatus = result?.TaskStatus;
                                estimateStart = result?.EstimateStart;
                                estimateEnd = result?.EstimateEnd;
                                actualStart = result?.ActualStart;
                                projectAttribute = result?.ProjectAttribute;
                                requiredFile = Convert.ToBoolean(result?.RequiredFile);

                                if (requiredFile)
                                {
                                    switch (ReplyStatus)
                                    {
                                        case "O":
                                            if (ReplyFile <= 0) throw new SystemException("【任務回報】必須上傳附件!");
                                            break;
                                    }
                                }

                                #region //遞迴取得任務所有前置任務
                                dynamicParameters = new DynamicParameters();
                                sql = projectDA.GetSingleTaskRecursionSqlString(ref dynamicParameters, TaskId);
                                sql += @"
                                        SELECT a.ProjectTaskId
                                        , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                                        , a.TaskName, a.TaskStatus
                                        , ISNULL(FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                                        , ISNULL(FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                                        , ISNULL(a.PlannedDuration, 1) PlannedDuration
                                        , ISNULL(FORMAT(a.EstimateStart, 'yyyy-MM-dd HH:mm'), '') EstimateStart
                                        , ISNULL(FORMAT(a.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                                        , ISNULL(a.EstimateDuration, 1) EstimateDuration
                                        , ISNULL(FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                                        , ISNULL(FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                                        , ISNULL(a.ActualDuration, 1) ActualDuration
                                        , link.*
                                        FROM @projectTask x
                                        INNER JOIN PMIS.ProjectTask a on a.ProjectTaskId = x.ProjectTaskId
                                        OUTER APPLY (
	                                        SELECT ba.SourceTaskId, ba.TargetTaskId, bb.TaskName SourceTaskName, bb.TaskStatus SourceStatus
	                                        , FORMAT(bb.ActualStart, 'yyyy-MM-dd HH:mm') SourceActualStart
	                                        , FORMAT(bb.ActualEnd, 'yyyy-MM-dd HH:mm') SourceActualEnd
	                                        , ISNULL(bb.ActualDuration, 1) SourceActualDuration
	                                        FROM PMIS.ProjectTaskLink ba 
	                                        INNER JOIN PMIS.ProjectTask bb ON bb.ProjectTaskId = ba.SourceTaskId 
	                                        WHERE ba.TargetTaskId = a.ProjectTaskId) link
                                        OUTER APPLY (
	                                        SELECT ca.SourceTaskId, ca.TargetTaskId, cb.TaskName, cb.TaskStatus ParentSourceStatus
	                                        FROM PMIS.ProjectTaskLink ca 
	                                        INNER JOIN PMIS.ProjectTask cb ON cb.ProjectTaskId = ca.SourceTaskId 
	                                        WHERE ca.TargetTaskId = a.ParentTaskId) pLink
                                        ORDER BY x.TaskLevel";
                                var projectTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                                if (projectTasks.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");
                                #endregion

                                #region //Mail資料
                                sql = @"SELECT a.MailSubject, a.MailContent
							            , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                        , ISNULL(a.MailFrom, '') MailFrom
                                        , ISNULL(a.MailTo, '') MailTo
                                        , ISNULL(a.MailCc, '') MailCc
                                        , ISNULL(a.MailBcc, '') MailBcc
							            FROM BAS.Mail a
							            LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
							            WHERE EXISTS (
								            SELECT ca.MailId
								            FROM BAS.MailSendSetting ca
								            WHERE ca.MailId = a.MailId
								            AND ca.SettingSchema = 'ProjectTaskFlow'
								            AND ca.SettingNo = 'Y'
							            )";
                                var mailTemplate = sqlConnection.Query(sql);
                                if (mailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                                #endregion

                                var targetMailUsers = new List<dynamic>();

                                switch (taskStatus)
                                {
                                    case "B": //待處理(Backlog)
                                        bool inputDateStart = DateTime.TryParse(ActualStart, out DateTime resultActualStart);
                                        resultActualStart = inputDateStart ? resultActualStart : CreateDate;

                                        switch (ReplyStatus)
                                        {
                                            case "S": //開始(Start)
                                                canReply = true;

                                                #region //判斷前置任務是否完成
                                                var prevProjectTask = projectTasks.FirstOrDefault(f => f.SourceTaskId != null);
                                                if (prevProjectTask != null) //有前置任務
                                                {
                                                    if (prevProjectTask.SourceStatus != "C") throw new SystemException(string.Format("前置任務【{0}】尚未完成，無法回報開始!", prevProjectTask.SourceTaskName));
                                                    if (DateTime.TryParse(prevProjectTask.SourceActualEnd, out DateTime sourceActualEnd))
                                                    {
                                                        if (resultActualStart.CompareTo(sourceActualEnd) < 0)
                                                            throw new SystemException(string.Format("開始時間不得小於前置任務【{0}】的完成時間【{1}】", prevProjectTask.SourceTaskName, prevProjectTask.SourceActualEnd));
                                                    }
                                                }
                                                #endregion

                                                #region //取得子任務資料
                                                sql = @"SELECT a.TaskStatus
                                                        , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                                        , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ParentTaskId = @TaskId";
                                                var resultSubTask = sqlConnection.Query(sql, new { TaskId });
                                                if (resultSubTask.Any())
                                                {
                                                    #region //判斷子任務是否進行
                                                    if (!resultSubTask.Any(a => a.TaskStatus != "B"))
                                                    {
                                                        throw new SystemException("【子任務尚未進行】，無法回報開始!");
                                                    }
                                                    #endregion
                                                }
                                                #endregion

                                                #region //更新專案資訊(必須在任務狀態更新之前執行)
                                                sql = @"UPDATE a SET 
                                                        a.ActualStart = @ActualStart
                                                        FROM PMIS.Project a
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubTask
	                                                        FROM PMIS.ProjectTask ca
	                                                        WHERE ca.ProjectId = a.ProjectId
                                                        ) c
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubB
	                                                        FROM PMIS.ProjectTask da
	                                                        WHERE da.ProjectId = a.ProjectId
	                                                        AND da.TaskStatus = 'B'
                                                        ) d
                                                        WHERE a.ProjectId = @ProjectId
                                                        AND a.ProjectStatus <> 'P'
                                                        AND c.SubTask = d.SubB";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new {
                                                        ProjectId,
                                                        ActualStart = resultActualStart.ToString("yyyy-MM-dd HH:mm") //實際開始時間
                                                    });
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualStart = @ActualStart,
                                                        a.EstimateDuration = PlannedDuration,
                                                        a.TaskStatus = @TaskStatus,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId IN @ProjectTaskIds";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectTaskIds = projectTasks.Where(w => w.TaskStatus == "B").Select(s => s.ProjectTaskId).ToArray(),
                                                        ActualStart = resultActualStart.ToString("yyyy-MM-dd HH:mm"), //實際開始時間
                                                        TaskStatus = "I",
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion
                                                break;
                                            case "N": //尚未開始(Not started yet)
                                                canReply = true;
                                                if (CreateDate > Convert.ToDateTime(estimateStart)) throw new SystemException("已屆預估開始時間，請選擇開始或延宕!");
                                                break;
                                            case "D": //延宕(Delay)
                                                canReply = true;
                                                if (CreateDate <= Convert.ToDateTime(estimateStart)) throw new SystemException("未達預估開始時間，請選擇開始或尚未開始!");

                                                #region //計算遞延時間
                                                switch (DurationUnit)
                                                {
                                                    case "minute":
                                                        break;
                                                    case "hour":
                                                        DeferredDuration = DeferredDuration * 60;
                                                        break;
                                                    case "day":
                                                        DeferredDuration = DeferredDuration * 24 * 60;
                                                        break;
                                                    case "week":
                                                        DeferredDuration = DeferredDuration * 7 * 24 * 60;
                                                        break;
                                                    case "month":
                                                        DeferredDuration = DeferredDuration * 30 * 24 * 60;
                                                        break;
                                                }
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.EstimateStart = @EstimateStart,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Query(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        EstimateStart = Convert.ToDateTime(estimateStart).AddMinutes(DeferredDuration).ToString("yyyy-MM-dd HH:mm"), //預估開始時間
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    }).Count();
                                                #endregion

                                                #region //取得通知對象
                                                sql = @"SELECT a.MailUsers, ISNULL(SUBSTRING(a.TaskUsers, 0, LEN(a.TaskUsers)), '') TaskUsers
                                                        FROM (
                                                            SELECT (
                                                                SELECT ub.UserNo
                                                                , uc.CompanyId, uc.CompanyName, uc.CompanyNo
                                                                , ud.DepartmentName + '-' + ub.UserName + ':' + ub.Email MailTo
	                                                            FROM BAS.[User] ub
                                                                INNER JOIN BAS.Department ud ON ud.DepartmentId = ub.DepartmentId
                                                                INNER JOIN BAS.Company uc ON uc.CompanyId = ud.CompanyId
                                                                WHERE EXISTS(
                                                                    SELECT TOP 1 1 
                                                                    FROM PMIS.TaskUser ba 
                                                                    WHERE ba.TaskType = 'Project' 
                                                                    AND ba.TaskId IN @ProjectTaskIds
                                                                    AND ba.UserId = ub.UserId
                                                                )
                                                                FOR JSON PATH, ROOT('data')
                                                            ) MailUsers
                                                            , (
		                                                        SELECT cb.UserName + ','
		                                                        FROM BAS.[User] cb
                                                                WHERE EXISTS(
                                                                    SELECT TOP 1 1 
                                                                    FROM PMIS.TaskUser ba 
                                                                    WHERE ba.TaskType = 'Project' 
                                                                    AND ba.TaskId IN @ProjectTaskIds
                                                                    AND ba.UserId = cb.UserId
                                                                )
		                                                        FOR XML PATH('')
	                                                        ) TaskUsers
                                                        ) a";
                                                targetMailUsers = sqlConnection.Query(sql, 
                                                    new
                                                    {
                                                        ProjectTaskIds = projectTasks.Where(w => w.ParentTaskId > 0).Select(s => s.ParentTaskId).ToArray()
                                                    }).ToList();
                                                #endregion

                                                #region //延宕通知
                                                if (!targetMailUsers.Any(a => string.IsNullOrWhiteSpace(a.MailUsers)))
                                                {
                                                    foreach (var item in mailTemplate)
                                                    {
                                                        string mailSubject = item.MailSubject;
                                                        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                                                        string hyperLink = Regex.Replace(HttpContext.Current.Request.Url.AbsoluteUri, "api/.*|Task/.*", "Task/TaskManagement");
                                                        string pageUrl = hyperLink;
                                                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("查看可進行的任務"));

                                                        #region //Mail內容
                                                        foreach (var task in targetMailUsers)
                                                        {
                                                            mailSubject = mailSubject.Replace("[ProjectName]", projectName)
                                                                                     .Replace("[TaskName]", taskName);
                                                            mailContent = mailContent.Replace("[ProjectName]", projectName)
                                                                                     .Replace("[TaskName]", taskName)
                                                                                     .Replace("[TaskUser]", task.TaskUsers)
                                                                                     .Replace("[MailContent]", string.Format("專案任務【{0}】已延宕，請各成員持續追蹤任務情況。", taskName));

                                                            JObject mailUsersJson = JObject.Parse(task.MailUsers);
                                                            var mailTos = mailUsersJson["data"].Select(s => s["MailTo"]?.ToString()).ToArray();
                                                            var companyNos = mailUsersJson["data"].Select(s => s["CompanyNo"]?.ToString()).Distinct().ToList();

                                                            #region //MAMO個人訊息推播
                                                            //string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                                                            //mamoContent = mamoContent.Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                                                            //foreach (var companyNo in companyNos)
                                                            //{
                                                            //    var userNos = mailUsersJson["data"].Where(w => w["CompanyNo"]?.ToString() == companyNo).Select(s => s["UserNo"]?.ToString()).Distinct().ToList();
                                                            //    foreach (var userNo in userNos)
                                                            //    {
                                                            //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                                            //    }
                                                            //}
                                                            #endregion

                                                            #region //發送
                                                            //MailConfig mailConfig = new MailConfig
                                                            //{
                                                            //    Host = item.Host,
                                                            //    Port = Convert.ToInt32(item.Port),
                                                            //    SendMode = Convert.ToInt32(item.SendMode),
                                                            //    From = item.MailFrom,
                                                            //    Subject = mailSubject,
                                                            //    Account = item.Account,
                                                            //    Password = item.Password,
                                                            //    MailTo = string.Join(";", mailTos),
                                                            //    MailCc = item.MailCc,
                                                            //    MailBcc = item.MailBcc,
                                                            //    HtmlBody = mailContent.Replace("[hyperlink]", hyperLink),
                                                            //    TextBody = "-"
                                                            //};
                                                            //BaseHelper.MailSend(mailConfig);
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                #endregion
                                                break;
                                            default:
                                                throw new SystemException("不允許的【回報類型】!");
                                        }
                                        if (canReply)
                                        {
                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (TaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate, ReplyFile
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@TaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate, @ReplyFile
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    TaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateStart ? resultActualStart : CreateDate,
                                                    ReplyFile = ReplyFile > 0 ? ReplyFile : (int?)null,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            #endregion
                                        }
                                        break;
                                    case "I": //進行中
                                        bool inputDateEnd = DateTime.TryParse(ActualEnd, out DateTime resultActualEnd);
                                        resultActualEnd = inputDateEnd ? resultActualEnd : CreateDate;

                                        switch (ReplyStatus)
                                        { 
                                            case "O": //按計劃進行(On schedule)
                                                canReply = true;
                                                if (CreateDate > Convert.ToDateTime(estimateEnd)) throw new SystemException("已屆預估完成時間，請選擇完成或延宕!");
                                                break;
                                            case "D": //延宕(Delay)
                                                canReply = true;

                                                #region //計算遞延時間
                                                switch (DurationUnit)
                                                {
                                                    case "minute":
                                                        break;
                                                    case "hour":
                                                        DeferredDuration = DeferredDuration * 60;
                                                        break;
                                                    case "day":
                                                        DeferredDuration = DeferredDuration * 24 * 60;
                                                        break;
                                                    case "week":
                                                        DeferredDuration = DeferredDuration * 7 * 24 * 60;
                                                        break;
                                                    case "month":
                                                        DeferredDuration = DeferredDuration * 30 * 24 * 60;
                                                        break;
                                                }

                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.EstimateDuration = EstimateDuration + @DeferredDuration,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        DeferredDuration, //遞延時間(分鐘)
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion

                                                #region //取得通知對象
                                                sql = @"SELECT a.MailUsers, ISNULL(SUBSTRING(a.TaskUsers, 0, LEN(a.TaskUsers)), '') TaskUsers
                                                        FROM (
                                                            SELECT (
                                                                SELECT ub.UserNo
                                                                , uc.CompanyId, uc.CompanyName, uc.CompanyNo
                                                                , ud.DepartmentName + '-' + ub.UserName + ':' + ub.Email MailTo
	                                                            FROM BAS.[User] ub
                                                                INNER JOIN BAS.Department ud ON ud.DepartmentId = ub.DepartmentId
                                                                INNER JOIN BAS.Company uc ON uc.CompanyId = ud.CompanyId
                                                                WHERE EXISTS(
                                                                    SELECT TOP 1 1 
                                                                    FROM PMIS.TaskUser ba 
                                                                    WHERE ba.TaskType = 'Project' 
                                                                    AND ba.TaskId IN @ProjectTaskIds
                                                                    AND ba.UserId = ub.UserId
                                                                )
                                                                FOR JSON PATH, ROOT('data')
                                                            ) MailUsers
                                                            , (
		                                                        SELECT cb.UserName + ','
		                                                        FROM BAS.[User] cb
                                                                WHERE EXISTS(
                                                                    SELECT TOP 1 1 
                                                                    FROM PMIS.TaskUser ba 
                                                                    WHERE ba.TaskType = 'Project' 
                                                                    AND ba.TaskId IN @ProjectTaskIds
                                                                    AND ba.UserId = cb.UserId
                                                                )
		                                                        FOR XML PATH('')
	                                                        ) TaskUsers
                                                        ) a";
                                                targetMailUsers = sqlConnection.Query(sql, 
                                                    new
                                                    {
                                                        ProjectTaskIds = projectTasks.Where(w => w.ParentTaskId > 0).Select(s => s.ParentTaskId).ToArray()
                                                    }).ToList();
                                                #endregion

                                                #region //延宕通知
                                                if (!targetMailUsers.Any(a => string.IsNullOrWhiteSpace(a.MailUsers)))
                                                {
                                                    foreach (var item in mailTemplate)
                                                    {
                                                        string mailSubject = item.MailSubject;
                                                        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                                                        string hyperLink = Regex.Replace(HttpContext.Current.Request.Url.AbsoluteUri, "api/.*|Task/.*", "Task/TaskManagement");
                                                        string pageUrl = hyperLink;
                                                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("查看可進行的任務"));

                                                        #region //Mail內容
                                                        foreach (var task in targetMailUsers)
                                                        {
                                                            mailSubject = mailSubject.Replace("[ProjectName]", projectName)
                                                                                     .Replace("[TaskName]", taskName);
                                                            mailContent = mailContent.Replace("[ProjectName]", projectName)
                                                                                     .Replace("[TaskName]", taskName)
                                                                                     .Replace("[TaskUser]", task.TaskUsers)
                                                                                     .Replace("[MailContent]", string.Format("專案任務【{0}】已延宕，請各成員持續追蹤任務情況。", taskName));

                                                            JObject mailUsersJson = JObject.Parse(task.MailUsers);
                                                            var mailTos = mailUsersJson["data"].Select(s => s["MailTo"]?.ToString()).ToArray();
                                                            var companyNos = mailUsersJson["data"].Select(s => s["CompanyNo"]?.ToString()).Distinct().ToList();

                                                            #region //MAMO個人訊息推播
                                                            //string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                                                            //mamoContent = mamoContent.Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                                                            //foreach (var companyNo in companyNos)
                                                            //{
                                                            //    var userNos = mailUsersJson["data"].Where(w => w["CompanyNo"]?.ToString() == companyNo).Select(s => s["UserNo"]?.ToString()).Distinct().ToList();
                                                            //    foreach (var userNo in userNos)
                                                            //    {
                                                            //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                                            //    }
                                                            //}
                                                            #endregion

                                                            #region //發送
                                                            //MailConfig mailConfig = new MailConfig
                                                            //{
                                                            //    Host = item.Host,
                                                            //    Port = Convert.ToInt32(item.Port),
                                                            //    SendMode = Convert.ToInt32(item.SendMode),
                                                            //    From = item.MailFrom,
                                                            //    Subject = mailSubject,
                                                            //    Account = item.Account,
                                                            //    Password = item.Password,
                                                            //    MailTo = string.Join(";", mailTos),
                                                            //    MailCc = item.MailCc,
                                                            //    MailBcc = item.MailBcc,
                                                            //    HtmlBody = mailContent.Replace("[hyperlink]", hyperLink),
                                                            //    TextBody = "-"
                                                            //};
                                                            //BaseHelper.MailSend(mailConfig);
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                #endregion
                                                break;
                                            case "C": //完成(Complete)
                                                canReply = true;
                                                if (resultActualEnd.CompareTo(Convert.ToDateTime(actualStart)) <= 0) throw new SystemException("【完成時間】須大於【開始時間】");

                                                #region //取得子任務資料
                                                sql = @"SELECT a.TaskStatus
                                                        , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                                        , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ParentTaskId = @TaskId";
                                                var resultSubTask = sqlConnection.Query(sql, new { TaskId });
                                                if (resultSubTask.Any())
                                                {
                                                    #region //判斷子任務是否完成
                                                    if (resultSubTask.Any(a => a.TaskStatus != "C"))
                                                    {
                                                        throw new SystemException("【子任務尚未完成】，無法回報完成!");
                                                    }
                                                    #endregion

                                                    #region //判斷回報時間是否小於子任務最後完成時間
                                                    if (DateTime.TryParse(resultSubTask.Max(m => m.ActualEnd), out DateTime subMaxActualEnd))
                                                    {
                                                        if (resultActualEnd.CompareTo(subMaxActualEnd) <= 0) throw new SystemException(string.Format("【完成時間】不得小於子任務【最後完成時間】{0}", resultSubTask.Max(m => m.ActualEnd)));
                                                    }
                                                    #endregion
                                                }
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualEnd = @ActualEnd,
                                                        a.TaskStatus = 'C',
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        ActualEnd = resultActualEnd.ToString("yyyy-MM-dd HH:mm"),
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion

                                                foreach (var TempParentTaskId in projectTasks.Where(w => w.ParentTaskId > 0).Select(s => s.ParentTaskId))
                                                {
                                                    #region //取得該層所有任務
                                                    sql = @"SELECT a.ProjectTaskId, a.TaskStatus
                                                            , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                                            , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
												            , a.ActualDuration
                                                            FROM PMIS.ProjectTask a
                                                            WHERE a.ParentTaskId = @TempParentTaskId";
                                                    resultSubTask = sqlConnection.Query(sql, new { TempParentTaskId });
                                                    #endregion

                                                    if (resultSubTask.Any())
                                                    {
                                                        bool hasMaxActualEnd = DateTime.TryParse(resultSubTask.Max(m => m.ActualEnd), out DateTime maxActualEnd);

                                                        #region //判斷子任務是否皆已完成
                                                        int subTaskCount = resultSubTask.Where(w => w.TaskStatus == "C").Count();
                                                        if (subTaskCount == resultSubTask.Count())
                                                        {
                                                            #region //自動完成上層任務
                                                            sql = @"UPDATE a SET
                                                                    a.ActualEnd = @ActualEnd,
                                                                    a.TaskStatus = @TaskStatus,
                                                                    a.LastModifiedDate = @LastModifiedDate,
                                                                    a.LastModifiedBy = @LastModifiedBy
                                                                    FROM PMIS.ProjectTask a
                                                                    WHERE a.ProjectTaskId = @TempParentTaskId
                                                                    AND a.TaskStatus <> @TaskStatus";
                                                            rowsAffected += sqlConnection.Execute(sql,
                                                                new
                                                                {
                                                                    TempParentTaskId,
                                                                    ActualEnd = hasMaxActualEnd ? maxActualEnd.ToString("yyyy-MM-dd HH:mm") : resultActualEnd.ToString("yyyy-MM-dd HH:mm"),
                                                                    TaskStatus = "C",
                                                                    LastModifiedDate,
                                                                    LastModifiedBy
                                                                });
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                }

                                                #region //更新專案資訊(必須在任務狀態更新之後執行)
                                                sql = @"UPDATE a SET 
                                                        a.ProjectStatus = 'C',
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.Project a
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubTask
	                                                        FROM PMIS.ProjectTask ca
	                                                        WHERE ca.ProjectId = a.ProjectId
                                                        ) c
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubTaskC
	                                                        FROM PMIS.ProjectTask da
	                                                        WHERE da.ProjectId = a.ProjectId
	                                                        AND da.TaskStatus = 'C'
                                                        ) d
                                                        WHERE c.SubTask = d.SubTaskC
                                                        AND a.ProjectStatus <> 'P'
                                                        AND a.ProjectId = @ProjectId";
                                                dynamicParameters.AddDynamicParams(
                                                    new
                                                    {
                                                        ProjectId,
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion

                                                #region //取得通知對象 (停用)
                                                //dynamicParameters = new DynamicParameters();
                                                //sql = projectDA.GetTaskRecursionSqlString(ref dynamicParameters, new List<dynamic> { projectId });
                                                //sql += @"
                                                //    SELECT
                                                //    a.ProjectTaskId, a.TaskStatus
                                                //    , b.ProjectNo, b.ProjectName, b.ProjectDesc
                                                //    , a.TaskName, a.TaskDesc
                                                //    , a.TaskStatus, s2.StatusName TaskStatusName
                                                //    , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                                                //    , ISNULL(e.DepartmentNo, '') DepartmentNo, ISNULL(e.DepartmentName, '') DepartmentName
                                                //    , ISNULL(e.CompanyNo, '') CompanyNo, ISNULL(e.CompanyName, '') CompanyName, ISNULL(e.LogoIcon, -1) LogoIcon
                                                //    , ISNULL(SUBSTRING(d.TaskUsers, 0, LEN(d.TaskUsers)), '') TaskUsers
                                                //    , u.UserName ProjectMasterName, u.UserNo ProjectMasterNo, u.Gender ProjectMasterGender, u.UserName + '(' + u.UserNo + ')' ProjectMaster
                                                //    , ISNULL(ISNULL(link.TaskName, pLink.TaskName), '-') PrevTaskName
                                                //    ,  (
                                                //        SELECT ub.UserNo
                                                //        , uc.CompanyId, uc.CompanyName, uc.CompanyNo
                                                //        , ud.DepartmentName + '-' + ub.UserName + ':' + ub.Email MailTo
	                                               //     FROM BAS.[User] ub
                                                //        INNER JOIN BAS.Department ud ON ud.DepartmentId = ub.DepartmentId
                                                //        INNER JOIN BAS.Company uc ON uc.CompanyId = ud.CompanyId
                                                //        WHERE EXISTS(
                                                //            SELECT TOP 1 1 
                                                //            FROM PMIS.TaskUser ba 
                                                //            WHERE ba.TaskType = 'Project' 
                                                //            AND ba.TaskId = a.ProjectTaskId
                                                //            AND ba.UserId = ub.UserId
                                                //        )
                                                //        FOR JSON PATH, ROOT('data')
                                                //    ) MailUsers
                                                //    FROM @projectTask x
                                                //    INNER JOIN PMIS.ProjectTask a ON a.ProjectTaskId = x.ProjectTaskId
                                                //    INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                                //    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                                                //    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                                                //    INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                                                //    INNER JOIN BAS.[Type] t1 ON t1.TypeNo = b.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                                                //    INNER JOIN BAS.[Type] t2 ON t2.TypeNo = b.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                                                //    OUTER APPLY (
	                                               //     SELECT ea.DepartmentNo, ea.DepartmentName
	                                               //     , eb.CompanyNo, eb.CompanyName , eb.LogoIcon
	                                               //     FROM BAS.[Department] ea
	                                               //     INNER JOIN BAS.[Company] eb ON eb.CompanyId = ea.CompanyId
	                                               //     WHERE ea.DepartmentId = b.DepartmentId
                                                //    ) e
                                                //    OUTER APPLY (
                                                //        SELECT (
                                                //            SELECT db.UserName + ','
                                                //            FROM PMIS.TaskUser da
                                                //            INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                                //            WHERE da.TaskId = a.ProjectTaskId
                                                //            AND da.TaskType = 'Project'
                                                //            ORDER BY da.AgentSort
                                                //            FOR XML PATH('')
                                                //        ) TaskUsers
                                                //    ) d
                                                //    OUTER APPLY (
	                                               //     SELECT TOP 1 ub.UserId, ub.UserNo, ub.UserName, ub.Gender
	                                               //     FROM PMIS.ProjectUser ua 
	                                               //     INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                                               //     WHERE ua.ProjectId = a.ProjectId 
	                                               //     AND ua.AgentSort = 1
                                                //    ) u
                                                //    OUTER APPLY (
                                                //        SELECT ba.SourceTaskId, ba.TargetTaskId, bb.TaskName, bb.TaskStatus SourceStatus 
                                                //        FROM PMIS.ProjectTaskLink ba 
                                                //        INNER JOIN PMIS.ProjectTask bb ON bb.ProjectTaskId = ba.SourceTaskId 
                                                //        WHERE ba.TargetTaskId = a.ProjectTaskId
                                                //    ) link
                                                //    OUTER APPLY (
                                                //        SELECT ca.SourceTaskId, ca.TargetTaskId, cb.TaskName, cb.TaskStatus ParentSourceStatus
                                                //        FROM PMIS.ProjectTaskLink ca
                                                //        INNER JOIN PMIS.ProjectTask cb ON cb.ProjectTaskId = ca.SourceTaskId
                                                //        WHERE ca.TargetTaskId = a.ParentTaskId
                                                //    ) pLink
                                                //    WHERE NOT EXISTS (SELECT TOP 1 1 FROM @projectTask aa WHERE aa.ParentTaskId = a.ProjectTaskId) --排除父層任務
                                                //    AND ISNULL(link.SourceStatus, '') IN ('', 'C') --前置任務需完成
                                                //    AND ISNULL(pLink.ParentSourceStatus, '') IN ('', 'C') --父層前置任務需完成
                                                //    AND a.TaskStatus IN ('B') --待開始
                                                //    AND (EXISTS (
		                                              //      SELECT TOP 1 ISNULL(da.ProjectTaskLinkId, 1)
		                                              //      FROM PMIS.ProjectTaskLink da 
		                                              //      LEFT JOIN PMIS.ProjectTask db ON db.ProjectTaskId = da.TargetTaskId
		                                              //      LEFT JOIN PMIS.ProjectTask dc ON dc.ProjectTaskId = da.SourceTaskId
		                                              //      WHERE da.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')) AND ISNULL(dc.TaskStatus, '') IN ('', 'C')
	                                               //     ) 
	                                               //     OR NOT EXISTS (SELECT TOP 1 1 FROM PMIS.ProjectTaskLink ea WHERE ea.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')))
                                                //    )
                                                //    ORDER BY ISNULL(ISNULL(a.ActualStart, a.EstimateStart), a.PlannedStart)";
                                                //targetMailUsers = sqlConnection.Query(sql, dynamicParameters).ToList();
                                                #endregion

                                                #region //完成通知
                                                //if (!targetMailUsers.Any(a => string.IsNullOrWhiteSpace(a.MailUsers)))
                                                //{
                                                //    foreach (var item in mailTemplate)
                                                //    {
                                                //        string mailSubject = item.MailSubject;
                                                //        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                                                //        string hyperLink = Regex.Replace(HttpContext.Current.Request.Url.AbsoluteUri, "api/.*|Task/.*", "Task/TaskManagement");
                                                //        string pageUrl = hyperLink;
                                                //        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("查看可進行的任務"));

                                                        #region //Mail內容
                                                        //foreach (var task in targetMailUsers)
                                                        //{
                                                        //    mailSubject = mailSubject.Replace("[ProjectName]", task.ProjectName)
                                                        //                             .Replace("[TaskName]", task.TaskName);
                                                        //    mailContent = mailContent.Replace("[ProjectName]", task.ProjectName)
                                                        //                             .Replace("[PrevTaskName]", task.PrevTaskName)
                                                        //                             .Replace("[TaskName]", task.TaskName)
                                                        //                             .Replace("[TaskUser]", task.TaskUsers)
                                                        //                             .Replace("[MailContent]", string.Format("專案任務【{0}】已可進行，請安排時間開始進行。", task.TaskName));

                                                        //    JObject mailUsersJson = JObject.Parse(task.MailUsers);
                                                        //    var mailTos = mailUsersJson["data"].Select(s => s["MailTo"]?.ToString()).ToArray();
                                                        //    var companyNos = mailUsersJson["data"].Select(s => s["CompanyNo"]?.ToString()).Distinct().ToList();

                                                            #region //MAMO個人訊息推播
                                                            //string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                                                            //mamoContent = mamoContent.Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                                                            //foreach (var companyNo in companyNos)
                                                            //{
                                                            //    var userNos = mailUsersJson["data"].Where(w => w["CompanyNo"]?.ToString() == companyNo).Select(s => s["UserNo"]?.ToString()).Distinct().ToList();
                                                            //    foreach (var userNo in userNos)
                                                            //    {
                                                            //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                                            //    }
                                                            //}
                                                            #endregion

                                                            #region //發送
                                                            //MailConfig mailConfig = new MailConfig
                                                            //{
                                                            //    Host = item.Host,
                                                            //    Port = Convert.ToInt32(item.Port),
                                                            //    SendMode = Convert.ToInt32(item.SendMode),
                                                            //    From = item.MailFrom,
                                                            //    Subject = mailSubject,
                                                            //    Account = item.Account,
                                                            //    Password = item.Password,
                                                            //    MailTo = string.Join(";", mailTos),
                                                            //    MailCc = item.MailCc,
                                                            //    MailBcc = item.MailBcc,
                                                            //    HtmlBody = mailContent.Replace("[hyperlink]", hyperLink),
                                                            //    TextBody = "-"
                                                            //};
                                                            //BaseHelper.MailSend(mailConfig);
                                                            #endregion
                                                        //}
                                                        #endregion
                                                    //}
                                                //}
                                                #endregion
                                                break;
                                            default:
                                                throw new SystemException("不允許的【回報類型】!");
                                        }
                                        if (canReply)
                                        {
                                            #region //取得回報紀錄
                                            sql = @"SELECT TOP 1 a.ReplyFile
                                                    FROM PMIS.TaskReply a
                                                    WHERE EXISTS (
                                                        SELECT TOP 1 1
                                                        FROM PMIS.TaskUser aa
                                                        WHERE aa.TaskId = a.TaskId
                                                        AND aa.UserId = @CurrentUser 
                                                        AND aa.TaskType = 'Project'
                                                    ) 
                                                    AND a.TaskId = @TaskId
                                                    AND ISNULL(a.ReplyFile, -1) <> -1";
                                            var resultReplyFile = sqlConnection.QueryFirstOrDefault(sql, new { TaskId, CurrentUser });
                                            if (requiredFile && resultReplyFile == null && ReplyFile == -1) throw new SystemException("請選擇並上傳回報附件!");
                                            #endregion

                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (TaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate, ReplyFile
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@TaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate, @ReplyFile
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    TaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateEnd ? resultActualEnd : CreateDate,
                                                    ReplyFile = ReplyFile > 0 ? ReplyFile : (int?)null,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            #endregion
                                        }
                                        break;
                                    case "T":
                                        throw new SystemException("不允許的【回報類型】!");
                                    case "C":
                                        throw new SystemException("不允許的【回報類型】!");
                                }
                                break;
                            case "Personal":
                                #region //取得專案任務資料
                                sql = @"SELECT TOP 1 a.TaskName
                                        , a.TaskStatus
                                        , FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm') PlannedStart
                                        , FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm') PlannedEnd
                                        , ISNULL(a.PlannedDuration, 0) PlannedDuration
                                        , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                        , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
                                        , ISNULL(a.ActualDuration, 0) ActualDuration
                                        FROM PMIS.PersonalTask a
                                        WHERE a.PersonalTaskId = @TaskId";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { TaskId }) ?? throw new SystemException("【個人任務】資料錯誤!");

                                taskName = result?.TaskName;
                                taskStatus = result?.TaskStatus;
                                estimateStart = result?.PlannedStart;
                                estimateEnd = result?.PlannedEnd;
                                actualStart = result?.ActualStart;
                                ActualEnd = result?.ActualEnd ?? string.Empty;

                                switch (taskStatus)
                                {
                                    case "B": //待處理(Backlog)
                                        bool inputDateStart = DateTime.TryParse(ActualStart, out DateTime resultActualStart);
                                        resultActualStart = inputDateStart ? resultActualStart : CreateDate;

                                        switch (ReplyStatus)
                                        {
                                            case "S": //開始(Start)
                                            case "O": //按計劃進行(On schedule)
                                                canReply = true;

                                                #region //更新任務資料
                                                sql = @"UPDATE PMIS.PersonalTask SET
                                                        ActualStart = @ActualStart,
                                                        TaskStatus = @TaskStatus,
                                                        LastModifiedDate = @LastModifiedDate,
                                                        LastModifiedBy = @LastModifiedBy
                                                        WHERE PersonalTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        ActualStart = resultActualStart.ToString("yyyy-MM-dd HH:mm"), //實際開始時間
                                                        TaskStatus = "I",
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion
                                                break;
                                            case "N": //尚未開始(Not started yet)
                                                canReply = true;
                                                if (CreateDate > Convert.ToDateTime(estimateStart)) throw new SystemException("已屆預估開始時間，請選擇開始或延宕!");
                                                break;
                                            case "D": //延宕(Delay)
                                                canReply = true;
                                                if (CreateDate <= Convert.ToDateTime(estimateStart)) throw new SystemException("未達預估開始時間，請選擇開始或尚未開始!");

                                                #region //計算遞延時間
                                                switch (DurationUnit)
                                                {
                                                    case "minute":
                                                        break;
                                                    case "hour":
                                                        DeferredDuration = DeferredDuration * 60;
                                                        break;
                                                    case "day":
                                                        DeferredDuration = DeferredDuration * 24 * 60;
                                                        break;
                                                    case "week":
                                                        DeferredDuration = DeferredDuration * 7 * 24 * 60;
                                                        break;
                                                    case "month":
                                                        DeferredDuration = DeferredDuration * 30 * 24 * 60;
                                                        break;
                                                }
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualStart = @ActualStart,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.PersonalTask a
                                                        WHERE a.PersonalTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        ActualStart = Convert.ToDateTime(estimateStart).AddMinutes(DeferredDuration), //預估開始時間
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion
                                                break;
                                        }
                                        if (canReply)
                                        {
                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (TaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate, ReplyFile
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@TaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate, @ReplyFile
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    TaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateStart ? resultActualStart : CreateDate,
                                                    ReplyFile = ReplyFile > 0 ? ReplyFile : (int?)null,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            #endregion
                                        }
                                        break;
                                    case "I": //進行中
                                        bool inputDateEnd = DateTime.TryParse(ActualEnd, out DateTime resultActualEnd);
                                        resultActualEnd = inputDateEnd ? resultActualEnd : CreateDate;

                                        switch (ReplyStatus)
                                        {
                                            case "O": //按計劃進行(On schedule)
                                                canReply = true;
                                                if (CreateDate > Convert.ToDateTime(estimateEnd)) throw new SystemException("已屆預估完成時間，請選擇完成或延宕!");
                                                break;
                                            case "D": //延宕(Delay)
                                                canReply = true;

                                                #region //計算遞延時間
                                                switch (DurationUnit)
                                                {
                                                    case "minute":
                                                        break;
                                                    case "hour":
                                                        DeferredDuration = DeferredDuration * 60;
                                                        break;
                                                    case "day":
                                                        DeferredDuration = DeferredDuration * 24 * 60;
                                                        break;
                                                    case "week":
                                                        DeferredDuration = DeferredDuration * 7 * 24 * 60;
                                                        break;
                                                    case "month":
                                                        DeferredDuration = DeferredDuration * 30 * 24 * 60;
                                                        break;
                                                }

                                                #endregion

                                                #region //更新任務資料
                                                resultActualEnd = Convert.ToDateTime(string.IsNullOrEmpty(ActualEnd) ? estimateEnd : ActualEnd); //取得現在任務最後完成時間
                                                sql = @"UPDATE a SET
                                                        a.ActualEnd = @ActualEnd,
                                                        a.ActualDuration = ISNULL(ActualDuration, PlannedDuration) + @DeferredDuration,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.PersonalTask a
                                                        WHERE a.PersonalTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        ActualEnd = resultActualEnd.AddMinutes(DeferredDuration).ToString("yyyy-MM-dd HH:mm"),
                                                        DeferredDuration, //遞延時間(分鐘)
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion
                                                break;
                                            case "C": //完成(Complete)
                                                canReply = true;
                                                if (resultActualEnd.CompareTo(Convert.ToDateTime(actualStart)) <= 0) throw new SystemException("【完成時間】須大於【開始時間】!");

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualEnd = @ActualEnd,
                                                        a.ActualDuration = @ActualDuration,
                                                        a.TaskStatus = @TaskStatus,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.PersonalTask a
                                                        WHERE a.PersonalTaskId = @TaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        TaskId,
                                                        TaskStatus = "C",
                                                        ActualEnd = resultActualEnd.ToString("yyyy-MM-dd HH:mm"),
                                                        ActualDuration = (resultActualEnd - Convert.ToDateTime(actualStart)).TotalMinutes,
                                                        LastModifiedDate,
                                                        LastModifiedBy
                                                    });
                                                #endregion
                                                break;
                                            default:
                                                throw new SystemException("不允許的【回報類型】!");
                                        }
                                        if (canReply)
                                        {
                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (TaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate, ReplyFile
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@TaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate, @ReplyFile
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    TaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateEnd ? resultActualEnd : CreateDate,
                                                    ReplyFile = ReplyFile > 0 ? ReplyFile : (int?)null,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            #endregion
                                        }
                                        break;
                                    case "T":
                                        throw new SystemException("不允許的【回報類型】!");
                                    case "C":
                                        throw new SystemException("不允許的【回報類型】!");
                                }

                                #endregion
                                break;
                            default:
                                break;
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //AddPersonalTask -- 個人任務新增-New -- Chia Yuan 2023.12.06
        public string AddPersonalTask(string TaskName, string TaskDesc, string PlannedStart, int PlannedDuration, int ReplyFrequency)
        {
            try
            {
                if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                if (TaskName.Length > 50) throw new SystemException("【任務名稱】長度錯誤!");
                if (TaskDesc.Length > 300) throw new SystemException("【任務描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        bool inputPlannedStart = DateTime.TryParse(PlannedStart, out DateTime plannedStart);
                        if (!inputPlannedStart) plannedStart = CreateDate.Date; //開始時間
                        PlannedDuration = PlannedDuration <= 0 ? 1440 : PlannedDuration; //預設為1天
                        DateTime plannedEnd = CreateDate.Date;
                        plannedEnd = plannedStart.AddMinutes(PlannedDuration); //完成時間

                        #region //判斷使用者資料是否正確
                        sql = @"SELECT TOP 1 c.CompanyNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department d ON d.DepartmentId = a.DepartmentId
                                INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                WHERE a.UserId = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("【使用者】資料錯誤!");
                        string CompanyNo = result?.CompanyNo ?? string.Empty;
                        #endregion

                        #region //抓取目前最大檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                ORDER BY SortNumber";
                        int LevelId = sqlConnection.QueryFirstOrDefault(sql).LevelId;
                        #endregion

                        #region //任務新增
                        sql = @"INSERT INTO PMIS.PersonalTask (UserId, TaskName, TaskDesc
                                , PlannedStart, PlannedEnd, PlannedDuration, ReplyFrequency
                                , TaskStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PersonalTaskId
                                VALUES (@CurrentUser, @TaskName, @TaskDesc
                                , @PlannedStart, @PlannedEnd, @PlannedDuration, @ReplyFrequency
                                , @TaskStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                CurrentUser,
                                TaskName,
                                TaskDesc,
                                PlannedStart = plannedStart,
                                PlannedEnd = plannedEnd,
                                PlannedDuration,
                                ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                TaskStatus = "B",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected++;
                        #endregion

                        int TaskId = insertResult?.PersonalTaskId ?? -1;

                        #region //成員新增
                        sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TaskUserId
                                VALUES (@TaskId, 'Personal', @CurrentUser, @LevelId, @AgentSort
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                TaskId,
                                CurrentUser,
                                LevelId,
                                AgentSort = 1,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected++;
                        #endregion

                        int TaskUserId = insertResult?.QueryFirstOrDefault ?? -1;

                        #region //成員權限新增
                        sql = @"INSERT INTO PMIS.TaskUserAuthority (TaskUserId, AuthorityId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                SELECT @TaskUserId, a.AuthorityId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy 
                                FROM PMIS.Authority a 
                                LEFT JOIN (SELECT b.AuthorityId FROM PMIS.TaskUserAuthority b WHERE b.TaskUserId = @TaskUserId) b on b.AuthorityId = a.AuthorityId
                                WHERE a.AuthorityType = 'T'
                                AND a.[Status] = 'A'
                                AND b.AuthorityId is null
                                ORDER BY a.SortNumber";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                TaskUserId,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        #endregion

                        #region //取得任務成員資料
                        sql = @"SELECT u.UserId,u.UserNo,u.UserName
                                FROM PMIS.TaskUser a
                                INNER JOIN BAS.[User] u ON u.UserId = a.UserId
                                WHERE a.TaskId = @TaskId
                                AND a.TaskType = 'Personal'";
                        var resultUsers = sqlConnection.Query(sql, new { TaskId });
                        #endregion

                        if (!string.IsNullOrWhiteSpace(CompanyNo))
                        {
                            #region //MAMO團隊新增 (停用)
                            //int teamId = -1;
                            //var mamoResult = mamoHelper.CreateTeams(companyNo, CurrentUser, string.Format("PP-{0}-{1}", TaskName, CreateDate.ToString("yyyyMMddHHmm")), "Personal");
                            //JObject mamoResultJson = JObject.Parse(mamoResult);

                            //if (mamoResultJson["status"].ToString() == "success")
                            //{
                            //    foreach (var item in mamoResultJson["result"])
                            //    {
                            //        teamId = Convert.ToInt32(item["TeamId"].ToString());
                            //    }

                            //    mamoHelper.AddTeamMembers(companyNo, CurrentUser, teamId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                            //}
                            #endregion

                            #region //TeamId更新 (停用)
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"UPDATE a SET a.TeamId = @TeamId
                            //        FROM PMIS.PersonalTask a
                            //        WHERE a.PersonalTaskId = @PersonalTaskId";
                            //dynamicParameters.AddDynamicParams(
                            //    new
                            //    {
                            //        TeamId = teamId,
                            //        PersonalTaskId = taskId
                            //    });
                            //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

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

        #region //AddPersonalTaskByCopy -- 個人任務複製-New -- Chia Yuan 2023.12.07
        public string AddPersonalTaskByCopy(int PersonalTaskId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        #region //取得來源任務資料
                        sql = @"SELECT a.UserId, a.TaskName, a.TaskDesc, a.PlannedStart, a.PlannedEnd, a.PlannedDuration, ISNULL(a.ReplyFrequency , 0) ReplyFrequency
                                FROM PMIS.PersonalTask a
                                WHERE a.PersonalTaskId = @PersonalTaskId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskId }) ?? throw new SystemException("【個人任務】資料錯誤!");
                        #endregion

                        #region //取得來源任務成員資料
                        sql = @"SELECT a.TaskUserId, a.TaskId, a.UserId, a.LevelId, a.AgentSort
                                , (
                                    SELECT aa.AuthorityId, aa.AuthorityName, ab.AuthorityStatus
                                    FROM PMIS.Authority aa
                                    OUTER APPLY (
                                        SELECT ISNULL((
                                            SELECT TOP 1 1
                                            FROM PMIS.TaskUserAuthority va
                                            WHERE va.TaskUserId = a.TaskUserId
                                            AND va.AuthorityId = aa.AuthorityId
                                        ), 0) AuthorityStatus
                                    ) ab
                                    WHERE aa.AuthorityType = 'T'
                                    AND aa.[Status] = 'A'
                                    ORDER BY aa.SortNumber
                                    FOR JSON PATH, ROOT('data')
                                ) Authority
                                FROM PMIS.TaskUser a
                                INNER JOIN PMIS.PersonalTask b ON a.TaskId = b.PersonalTaskId
                                WHERE b.PersonalTaskId = @PersonalTaskId
                                AND a.TaskType = 'Personal'";
                        var resultUsers = sqlConnection.Query(sql, new { PersonalTaskId });
                        #endregion

                        #region //任務資料新增
                        sql = @"INSERT INTO PMIS.PersonalTask (UserId,TaskName,TaskDesc
                                ,PlannedStart,PlannedEnd,PlannedDuration,ReplyFrequency,TaskStatus
                                ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.PersonalTaskId
                                VALUES (@CurrentUser, @TaskName, @TaskDesc
                                , @PlannedStart, @PlannedEnd, @PlannedDuration, @ReplyFrequency,'B'
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CurrentUser,
                                TaskName = result.TaskName + "(複製)",
                                result.TaskDesc,
                                result.PlannedStart,
                                result.PlannedEnd,
                                result.PlannedDuration,
                                result.ReplyFrequency,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        int TaskId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters)?.PersonalTaskId; //取得新產生的任務id
                        rowsAffected++;
                        #endregion

                        #region 任務成員新增
                        foreach (var user in resultUsers)
                        {
                            sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TaskUserId
                                    VALUES (@TaskId, 'Personal', @UserId, @LevelId, @AgentSort
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TaskId,
                                    user.UserId,
                                    user.LevelId,
                                    user.AgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            int TaskUserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters)?.TaskUserId; //取得新產生的成員id
                            rowsAffected++;

                            #region //權限新增
                            JObject Authority = JObject.Parse(user.Authority);
                            for (int i = 0; i < Authority["data"].Count(); i++)
                            {
                                if (Convert.ToInt32(Authority["data"][i]["AuthorityStatus"]) == 1)
                                {
                                    sql = @"INSERT INTO PMIS.TaskUserAuthority (TaskUserId, AuthorityId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@TaskUserId, @AuthorityId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            TaskUserId,
                                            AuthorityId = Convert.ToInt32(Authority["data"][i]["AuthorityId"]),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Update
        #region //UpdateProjectAllTaskEstimate -- 專案任務排程預估更新(全部任務) -- Ben Ma 2022.11.25
        public string UpdateProjectAllTaskEstimate(int ProjectId, int ProjectTaskId, string TaskData)
        {
            try
            {
                if (!TaskData.TryParseJson(out JObject tempJObject)) throw new SystemException("【專案任務】格式錯誤!");

                List<ProjectTask> projectTasks = JsonConvert.DeserializeObject<List<ProjectTask>>(JObject.Parse(TaskData)["data"].ToString());

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ProjectStatus
                                , FORMAT(PlannedStart, 'yyyy-MM-dd HH:mm') PlannedStart
                                , FORMAT(PlannedEnd, 'yyyy-MM-dd HH:mm') PlannedEnd
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!!");
                        string projectStatus = string.Empty;
                        DateTime projectPlannedStart = DateTime.Now;
                        DateTime projectPlannedEnd = DateTime.Now.AddMinutes(1);
                        projectStatus = result.ProjectStatus;
                        projectPlannedStart = Convert.ToDateTime(result.PlannedStart);
                        projectPlannedEnd = Convert.ToDateTime(result.PlannedEnd);
                        #endregion

                        foreach (var task in projectTasks)
                        {
                            if (task.ProjectTaskId.ToString() != "1.1")
                            {
                                #region //判斷專案任務資料是否正確
                                sql = @"SELECT TOP 1 1
                                        FROM PMIS.ProjectTask
                                        WHERE ProjectTaskId = @ProjectTaskId";
                                dynamicParameters.Add("ProjectTaskId", task.ProjectTaskId);
                                result = sqlConnection.QueryFirstOrDefault(sql, new { task.ProjectTaskId }) ?? throw new SystemException("【專案任務】資料錯誤!");
                                #endregion

                                switch (task.TaskStatus)
                                {
                                    case "B":
                                    case "I":
                                        #region //更新任務資料
                                        sql = @"UPDATE a SET
                                                a.EstimateStart = @EstimateStart,
                                                a.EstimateEnd = @EstimateEnd,
                                                a.EstimateDuration = @EstimateDuration,
                                                a.LastModifiedDate = @LastModifiedDate,
                                                a.LastModifiedBy = @LastModifiedBy
                                                FROM PMIS.ProjectTask a
                                                WHERE a.ProjectTaskId = @ProjectTaskId";
                                        rowsAffected += sqlConnection.Execute(sql,
                                            new
                                            {
                                                task.ProjectTaskId,
                                                EstimateStart = Convert.ToDateTime(task.EstimateStart),
                                                EstimateEnd = Convert.ToDateTime(task.EstimateEnd),
                                                task.EstimateDuration,
                                                LastModifiedDate,
                                                LastModifiedBy
                                            });
                                        #endregion
                                        break;
                                    case "C":
                                        #region //更新任務資料
                                        sql = @"UPDATE a SET
                                                a.ActualDuration = @EstimateDuration,
                                                a.LastModifiedDate = @LastModifiedDate,
                                                a.LastModifiedBy = @LastModifiedBy
                                                FROM PMIS.ProjectTask a
                                                WHERE a.ProjectTaskId = @ProjectTaskId";
                                        rowsAffected += sqlConnection.Execute(sql,
                                            new
                                            {
                                                task.ProjectTaskId,
                                                task.EstimateDuration,
                                                LastModifiedDate,
                                                LastModifiedBy
                                            });
                                        #endregion
                                        break;
                                }
                            }
                            else
                            {
                                switch (projectStatus)
                                {
                                    case "B":
                                    case "I":
                                        #region //更新任務資料
                                        sql = @"UPDATE a SET
                                                a.EstimateStart = @EstimateStart,
                                                a.EstimateEnd = @EstimateEnd,
                                                a.EstimateDuration = @EstimateDuration,
                                                a.LastModifiedDate = @LastModifiedDate,
                                                a.LastModifiedBy = @LastModifiedBy
                                                FROM PMIS.Project a
                                                WHERE a.ProjectId = @ProjectId";
                                        rowsAffected += sqlConnection.Execute(sql,
                                            new
                                            {
                                                ProjectId,
                                                EstimateStart = Convert.ToDateTime(task.EstimateStart),
                                                EstimateEnd = Convert.ToDateTime(task.EstimateEnd),
                                                task.EstimateDuration,
                                                LastModifiedDate,
                                                LastModifiedBy
                                            });
                                        #endregion
                                        break;
                                    case "C":
                                        #region //更新任務資料
                                        sql = @"UPDATE a SET
                                                a.ActualEnd = @EstimateEnd,
                                                a.ActualDuration = @EstimateDuration,
                                                a.LastModifiedDate = @LastModifiedDate,
                                                a.LastModifiedBy = @LastModifiedBy
                                                FROM PMIS.Project a
                                                WHERE a.ProjectId = @ProjectId";
                                        rowsAffected += sqlConnection.Execute(sql,
                                            new
                                            {
                                                ProjectId,
                                                task.EstimateEnd,
                                                task.EstimateDuration,
                                                LastModifiedDate,
                                                LastModifiedBy
                                            });
                                        #endregion
                                        break;
                                }
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

        #region //UpdatePersonalTask -- 個人任務更新-New -- Chia Yuan 2023.12.07
        public string UpdatePersonalTask(int PersonalTaskId, string TaskName , string TaskDesc
            , string PlannedStart , int PlannedDuration, int ReplyFrequency)
        {
            try
            {
                if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                if (TaskName.Length > 50) throw new SystemException("【任務名稱】長度錯誤!");
                if (TaskDesc.Length > 300) throw new SystemException("【任務描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        #region //判斷任務是否正確
                        sql = @"SELECT TOP 1 TaskStatus
                                FROM PMIS.PersonalTask
                                WHERE PersonalTaskId = @PersonalTaskId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskId }) ?? throw new SystemException("【個人任務】資料錯誤!");
                        string TaskStatus = result?.TaskStatus ?? string.Empty;
                        if (TaskStatus == "C") throw new SystemException("【個人任務】已完成，無法更新!");
                        #endregion

                        switch (TaskStatus)
                        {
                            case "B": //待處理
                            case "I": //進行中
                            case "T": //待檢驗
                                if (string.IsNullOrEmpty(PlannedStart)) throw new SystemException("【計畫開始時間】不能為空!");
                                if (!DateTime.TryParse(PlannedStart, out DateTime plannedStart)) throw new SystemException("【計畫開始時間】格式有誤!");

                                #region //個人任務更新
                                sql = @"UPDATE PMIS.PersonalTask SET
                                        TaskName = @TaskName,
                                        TaskDesc = @TaskDesc,
                                        PlannedStart = @PlannedStart,
                                        PlannedEnd = @PlannedEnd,
                                        PlannedDuration = @PlannedDuration,
                                        ReplyFrequency = @ReplyFrequency,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PersonalTaskId = @PersonalTaskId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        PersonalTaskId,
                                        TaskName,
                                        TaskDesc,
                                        PlannedStart = plannedStart,
                                        PlannedEnd = plannedStart.AddMinutes(PlannedDuration),
                                        PlannedDuration,
                                        ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                        LastModifiedDate,
                                        LastModifiedBy
                                    });
                                #endregion
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
        #endregion

        #region //Delete
        #region //DeleteReplyFile -- 刪除回報附件 -- Chia Yuan 2024.12.07
        public string DeleteReplyFile(int key)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        sql = @"DELETE a
                                FROM BAS.[File] a
                                WHERE a.FileId = @key";
                        rowsAffected += sqlConnection.Execute(sql, new { key });

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
