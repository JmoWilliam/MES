using Dapper;
using Helpers;
using NLog;
using System.Linq;
using System;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PMISDA
{
    public class ProjectTrackDA
    {
        public static string MainConnectionStrings = string.Empty;

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

        public ProjectTrackDA()
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
        #region //ProjectPageList
        private List<dynamic> ProjectPageList(int ProjectId, string ProjectUserId, string TaskUserId, string ProjectMaster, int CompanyId, int DepartmentId, string Project, string TaskName
            , string ProjectStatus, string TaskStatus, string ReplyStatus, string TaskStartDate, string TaskEndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
            {
                #region //篩選資料
                sqlQuery.mainKey = "a.ProjectId";
                sqlQuery.auxKey = "";
                sqlQuery.columns = ", a.ProjectName";
                sqlQuery.mainTables =
                    @" FROM PMIS.Project a";
                sqlQuery.auxTables = "";
                sqlQuery.distinct = true;
                sql = "";

                if (ProjectId > 0)
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                }

                if (!string.IsNullOrWhiteSpace(ProjectUserId))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "ProjectUserIds"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.Project ba
		                    INNER JOIN PMIS.ProjectUser bb ON bb.ProjectId = ba.ProjectId
		                    WHERE ba.ProjectId = a.ProjectId
		                    AND bb.UserId IN @ProjectUserIds
	                    )", ProjectUserId.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(TaskUserId))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "TaskUserIds"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.TaskUser ca
                            INNER JOIN PMIS.ProjectTask cb ON cb.ProjectTaskId = ca.TaskId AND ca.TaskType = 'Project'
		                    WHERE cb.ProjectId = a.ProjectId
		                    AND ca.UserId IN @TaskUserIds
	                    )", TaskUserId.Split(','));
                }

                if (CompanyId > 0 || DepartmentId > 0)
                {
                    sql += @" AND (EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.Project da
                            INNER JOIN PMIS.ProjectUser db ON db.ProjectId = da.ProjectId AND db.AgentSort = 1
                            INNER JOIN BAS.[User] dc ON dc.UserId = db.UserId
                            INNER JOIN BAS.Department dd ON dd.DepartmentId = dc.DepartmentId
		                    WHERE da.ProjectId = a.ProjectId";
                    if (CompanyId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", " AND dd.CompanyId = @CompanyId", CompanyId);
                    }
                    if (DepartmentId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", " AND dd.DepartmentId = @DepartmentId", DepartmentId);
                    }
                    sql += @") OR 
                            EXISTS (
                            SELECT TOP 1 1 
                            FROM BAS.Department de 
                            WHERE de.DepartmentId = a.DepartmentId";
                    if (CompanyId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", " AND de.CompanyId = @CompanyId", CompanyId);
                    }
                    if (DepartmentId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", " AND de.DepartmentId = @DepartmentId", DepartmentId);
                    }

                    sql += @") OR 
                            EXISTS (
                            SELECT TOP 1 1 
                            FROM PMIS.ProjectUser du
                            INNER JOIN PMIS.ProjectUserAuthority da ON da.ProjectUserId = du.ProjectUserId
                            INNER JOIN PMIS.Authority db ON db.AuthorityId = da.AuthorityId AND db.AuthorityType = 'P'
                            WHERE du.ProjectId = a.ProjectId
                            AND du.UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

                    sql += "))";
                }

                if (!string.IsNullOrWhiteSpace(Project))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "Project"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.Project ea
		                    WHERE ea.ProjectId = a.ProjectId
		                    AND (ea.ProjectNo LIKE N'%' + @Project + '%' OR ea.ProjectName LIKE N'%' + @Project + '%')
	                    )", Project.Trim());
                }

                if (!string.IsNullOrWhiteSpace(TaskName))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "TaskName"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.ProjectTask fa
		                    WHERE fa.ProjectId = a.ProjectId
		                    AND fa.TaskName LIKE N'%' + @TaskName + '%'
	                    )", TaskName.Trim());
                }

                if (!string.IsNullOrWhiteSpace(ProjectMaster))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "ProjectMasters"
                        , @" AND EXISTS (
	                        SELECT TOP 1 1
	                        FROM PMIS.Project pa
	                        INNER JOIN PMIS.ProjectUser pb ON pb.ProjectId = pa.ProjectId
	                        WHERE pa.ProjectId = a.ProjectId
	                        AND pb.AgentSort = 1
	                        AND pb.UserId IN @ProjectMasters
                        )", ProjectMaster.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(ProjectStatus))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "ProjectStatus"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.Project ga
		                    WHERE ga.ProjectId = a.ProjectId
		                    AND ga.ProjectStatus IN @ProjectStatus
	                    )", ProjectStatus.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(TaskStatus))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "TaskStatus"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.ProjectTask ha
		                    WHERE ha.ProjectId = a.ProjectId
		                    AND ha.TaskStatus IN @TaskStatus
	                    )", TaskStatus.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(ReplyStatus))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "ReplyStatus"
                        , @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.TaskReply ia
                            INNER JOIN PMIS.ProjectTask ib ON ib.ProjectTaskId = ia.TaskId AND ia.ReplyType = 'Project'
		                    WHERE ib.ProjectId = a.ProjectId
		                    AND ia.ReplyStatus IN @ReplyStatus
	                    )", ReplyStatus.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(TaskStartDate) || !string.IsNullOrWhiteSpace(TaskEndDate))
                {
                    sql += @" AND EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.ProjectTask ja
                            WHERE ja.ProjectId = a.ProjectId";
                    if (DateTime.TryParse(TaskStartDate, out DateTime startDate))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", " AND ISNULL(ISNULL(ja.ActualStart, ja.EstimateStart), ja.PlannedStart) >= @StartDate", startDate.ToString("yyyy-MM-dd"));
                    }
                    if (DateTime.TryParse(TaskEndDate, out DateTime endDate))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", " AND ISNULL(ISNULL(ja.ActualStart, ja.EstimateStart), ja.PlannedStart) <= @EndDate", endDate.AddDays(1).ToString("yyyy-MM-dd"));
                    }
                    sql += ")";
                }

                sqlQuery.conditions = sql;
                sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectId";
                sqlQuery.pageIndex = PageIndex;
                sqlQuery.pageSize = PageSize;

                return BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
                #endregion
            }
        }
        #endregion

        #region //GetPersonalTaskFromProject
        public string GetPersonalTaskFromProject(List<int> ProjectIds)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PersonalTaskId, a.TaskName, a.TaskDesc, a.TaskStatus, s1.StatusName TaskStatusName
                        , ISNULL(FORMAT(ISNULL(a.ActualStart,a.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
	                    , ISNULL(FORMAT(ISNULL(a.ActualEnd, a.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
	                    , ISNULL(a.ActualDuration, a.PlannedDuration) TaskDuration
                        , ISNULL(d.DepartmentNo, '') DepartmentNo
                        , ISNULL(d.DepartmentName, '') DepartmentName
                        , ISNULL(e.CompanyName, '') CompanyName
                        , ISNULL(e.CompanyNo, '') CompanyNo
                        , ISNULL(e.LogoIcon, -1) LogoIcon
                        , u.UserId, u.UserNo, u.UserName, u.Gender
                        , u.UserName + '(' + u.UserNo + ')' PersonalMaster
                        , 'Personal' TaskType
                        , (SELECT cb.UserId, cb.UserNo, cb.UserName
	                        FROM PMIS.TaskUser da
	                        INNER JOIN BAS.[User] cb ON cb.UserId = da.UserId
	                        WHERE da.TaskType = 'Personal'
	                        AND da.TaskId = a.PersonalTaskId
	                        AND cb.[Status] = 'A'
	                        ORDER BY da.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) TaskUser
                        , (SELECT fb.UserNo, fb.UserName
                            , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                            , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                            , fa.ReplyStatus, fa.ReplyContent, fa.DeferredDuration, fc.StatusName
	                        FROM PMIS.TaskReply fa
		                    INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
		                    INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                        WHERE fa.ReplyType = 'Personal'
	                        AND fa.TaskId = a.PersonalTaskId
		                    AND fb.[Status] = 'A'
	                        ORDER BY fa.ReplyDate
	                        FOR JSON PATH, ROOT('data')
                        ) TaskReply
	                    FROM PMIS.PersonalTask a
	                    INNER JOIN BAS.[User] u ON u.UserId = a.UserId
	                    INNER JOIN BAS.Department d ON d.DepartmentId = u.DepartmentId
	                    INNER JOIN BAS.Company e ON e.CompanyId = d.CompanyId
	                    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.TaskStatus AND s1.StatusSchema = 'PersonalTask.TaskStatus'
	                    WHERE EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.Project ba
		                    INNER JOIN PMIS.ProjectUser bb ON bb.ProjectId = ba.ProjectId
		                    INNER JOIN BAS.[User] bc ON bc.UserId = bb.UserId
		                    WHERE bb.UserId = a.UserId
		                    AND bc.[Status] = 'A'
		                    AND (ISNULL(a.ActualStart, a.PlannedStart) BETWEEN ISNULL(ISNULL(ba.ActualStart, ba.EstimateStart), ba.PlannedStart) AND ISNULL(ISNULL(ba.ActualEnd, ba.EstimateEnd), ba.PlannedEnd) AND a.TaskStatus <> 'C')
		                    AND ba.ProjectId IN @ProjectIds
	                    )";
                    dynamicParameters.Add("ProjectIds", ProjectIds);

                    var result = sqlConnection.Query(sql, dynamicParameters).ToList();

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

        #region //GetTreeBaseSqlString
        private string GetProjectTaskRecursionString(ref DynamicParameters dynamicParameters, List<dynamic> ProjectIds)
        {
            var recursionString = "";
            recursionString = @"
                    DECLARE @rowsAdded int
                    DECLARE @projectTask TABLE
                    (
                        ProjectTaskId int,
                        ParentTaskId int,
                        TaskLevel int,
                        TaskRoute NVARCHAR(128),
                        TaskSort int,
                        TaskName NVARCHAR(128),
                        TaskRouteName NVARCHAR(512),
                        processed int DEFAULT(0)
                    )
                    INSERT @projectTask
                    SELECT a.ProjectTaskId
                    , ISNULL(a.ParentTaskId, -1) ParentTaskId
                    , 1 TaskLevel
                    , CAST(ISNULL(a.ParentTaskId, -1) AS NVARCHAR(128)) TaskRoute
                    , TaskSort
                    , a.TaskName
                    , '' TaskRouteName
                    , 0
                    FROM PMIS.ProjectTask a
                    WHERE a.ParentTaskId IS NULL
                    AND a.ProjectId IN @ProjectIds";

            dynamicParameters.Add("ProjectIds", ProjectIds);

            recursionString += @"
                    SET @rowsAdded=@@rowcount
                    WHILE @rowsAdded > 0
                    BEGIN
                        UPDATE @projectTask SET processed = 1 WHERE processed = 0
                        INSERT @projectTask
                        SELECT b.ProjectTaskId
                        , b.ParentTaskId
                        , (a.TaskLevel + 1) TaskLevel
                        , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS NVARCHAR(128)) AS NVARCHAR(128)) TaskRoute
                        , b.TaskSort
                        , b.TaskName
                        , CAST((CASE WHEN LEN(a.TaskRouteName) > 0 THEN a.TaskRouteName + ' > ' ELSE '' END) + a.TaskName AS NVARCHAR(128)) TaskRouteName
                        , 0
                        FROM @projectTask a
                        INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ParentTaskId
                        WHERE a.processed = 1
                        SET @rowsAdded = @@rowcount
                        UPDATE @projectTask SET processed = 2 WHERE processed = 1
                    END
            ";

            return recursionString;
        }
        #endregion

        #region //GetProjectTaskTree --取得專案任務資料(樹狀結構) -- Chia Yuan 2024.05.03
        public string GetProjectTaskTree(int ProjectId, string ProjectUserId, string TaskUserId, string ProjectMaster, int CompanyId, int DepartmentId, string Project, string TaskName
            , string ProjectStatus, string TaskStatus, string ReplyStatus, string TaskStartDate, string TaskEndDate
            , int OpenedLevel
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                var projects = ProjectPageList(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , OrderBy, PageIndex, PageSize);

                var projectIds = projects.Select(s => s.ProjectId).ToList();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetProjectTaskRecursionString(ref dynamicParameters, projectIds);
                    sql += @"SELECT c.ProjectId
	                    , ISNULL(FORMAT(ISNULL(ISNULL(c.ActualStart, c.EstimateStart), c.PlannedStart), 'yyyy-MM-dd HH:mm'), '') ProjectStart
	                    , ISNULL(FORMAT(ISNULL(ISNULL(c.ActualEnd, c.EstimateEnd), c.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') ProjectEnd
	                    , ISNULL(ISNULL(c.ActualDuration, c.EstimateDuration), c.PlannedDuration) ProjectDuration
                        , c.ProjectNo, c.ProjectName, c.ProjectDesc, c.CustomerRemark, c.ProjectRemark
                        , c.ProjectStatus, s1.StatusName ProjectStatusName
                        , c.ProjectType, t1.TypeName ProjectTypeName
                        , c.ProjectAttribute, t2.TypeName ProjectAttributeName
                        , c.WorkTimeStatus, c.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                        , a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort, a.TaskRouteName
                        , b.TaskName, b.TaskDesc
                        , b.TaskStatus, s2.StatusName TaskStatusName
	                    , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
	                    , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
	                    , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration) TaskDuration
                        , ISNULL(d.DepartmentNo, m.MasterDepartmentNo) DepartmentNo
                        , ISNULL(d.DepartmentName, m.MasterDepartmentName) DepartmentName
                        , ISNULL(e.CompanyName, m.MasterCompanyName) CompanyName
                        , ISNULL(e.CompanyNo, m.MasterCompanyNo) CompanyNo
                        , ISNULL(e.LogoIcon, -1) LogoIcon
                        , 'Project' TaskType
                        , (SELECT cb.UserId, cb.UserNo, cb.UserName, cb.Gender
	                        FROM PMIS.ProjectUser ca
	                        INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
	                        WHERE ca.ProjectId = c.ProjectId
	                        AND cb.[Status] = 'A'
	                        ORDER BY ca.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) ProjectUser
                        , (SELECT db.UserId, db.UserNo, db.UserName, db.Gender
	                        FROM PMIS.TaskUser da
	                        INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                        WHERE da.TaskType = 'Project'
	                        AND da.TaskId = a.ProjectTaskId
	                        AND db.[Status] = 'A'
	                        ORDER BY da.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) TaskUser
                        , (SELECT fb.UserNo, fb.UserName
                            , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                            , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                            , fa.ReplyContent, fa.DeferredDuration, fa.ReplyStatus, fc.StatusName ReplyStatusName
	                        FROM PMIS.TaskReply fa
		                    INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
		                    INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                        WHERE fa.ReplyType = 'Project'
	                        AND fa.TaskId = a.ProjectTaskId
		                    AND fb.[Status] = 'A'
	                        ORDER BY fa.ReplyDate
	                        FOR JSON PATH, ROOT('data')
                        ) TaskReply
                        , ISNULL((
	                        SELECT COUNT(1)
	                        FROM PMIS.TaskUser ea
                            INNER JOIN BAS.[User] eb ON eb.UserId = ea.UserId
	                        WHERE ea.TaskId = a.ProjectTaskId
	                        AND ea.TaskType = 'Project'
                            AND eb.[Status] = 'A'), 0
                        ) TaskUserCount
                        , ISNULL((
                            SELECT COUNT(1)
                            FROM PMIS.TaskReply ra
                            WHERE ra.TaskId = a.ProjectTaskId
                            AND ra.ReplyType = 'Project'
                            AND ra.ReplyStatus = 'D'), 0
                        ) DeferredCount
                        , m.*
                        FROM @projectTask a
                        INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                        INNER JOIN PMIS.Project c ON c.ProjectId = b.ProjectId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = c.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = c.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = c.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                        INNER JOIN BAS.[Type] t2 ON t2.TypeNo = c.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                        LEFT JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        LEFT JOIN BAS.Company e ON e.CompanyId = d.CompanyId
                        OUTER APPLY (
	                        SELECT ROW_NUMBER() OVER (PARTITION BY mb.ProjectId ORDER BY mb.AgentSort) SortPM
                            , ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                            , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
                            , mc.CompanyNo MasterCompanyNo, mc.CompanyName MasterCompanyName
                            , md.DepartmentNo MasterDepartmentNo, md.DepartmentName MasterDepartmentName
	                        FROM BAS.[User] ma 
	                        INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId
                            INNER JOIN BAS.Department md ON md.DepartmentId = ma.DepartmentId
                            INNER JOIN BAS.Company mc ON mc.CompanyId = md.CompanyId
	                        WHERE mb.ProjectId = c.ProjectId
                        ) m
                        WHERE m.SortPM = 1";
                    //if (!string.IsNullOrWhiteSpace(TaskStatus))
                    //{
                    //    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                    //        , "TaskStatus"
                    //        , @" AND b.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    //}
                    //if (!string.IsNullOrWhiteSpace(ReplyStatus))
                    //{
                    //    sql += " AND r.DeferredCount > 0";
                    //}
                    sql += @" ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                    var projectTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    List<ProjectList> taskLists = projectTasks
                        .GroupBy(g => g.ProjectId)
                        .Select(s => new ProjectList
                        {
                            ProjectId = Convert.ToString(s.Key),
                            ProjectNo = s.First()?.ProjectNo,
                            ProjectName = s.First()?.ProjectName,
                            StartDate = s.First()?.ProjectStart ?? "",
                            EndDate = s.First()?.ProjectEnd ?? "",
                            Duration = s.First()?.ProjectDuration ?? 1,
                            ProjectDesc = s.First()?.ProjectDesc,
                            CompanyName = s.First()?.CompanyName,
                            CompanyNo = s.First()?.CompanyNo,
                            LogoIcon = s.First()?.LogoIcon,
                            DepartmentNo = s.First()?.DepartmentNo,
                            DepartmentName = s.First()?.DepartmentName,
                            ProjectMaster = s.First()?.ProjectMaster,
                            MasterGender = s.First()?.Gender,
                            ProjectUser = s.First()?.ProjectUser,
                            ProjectStatus = s.First()?.ProjectStatus,
                            ProjectStatusName = s.First()?.ProjectStatusName,
                            ProjectType = s.First()?.ProjectType,
                            ProjectTypeName = s.First()?.ProjectTypeName,
                            ProjectAttribute = s.First()?.ProjectAttribute,
                            ProjectAttributeName = s.First()?.ProjectAttributeName,
                            WorkTimeStatus = s.First()?.WorkTimeStatus,
                            WorkTimeStatusName = s.First()?.WorkTimeStatusName,
                            TotalCount = projects.FirstOrDefault(f => f.ProjectId == s.Key)?.TotalCount,
                            ProjectTask = JsonConvert.SerializeObject(projectTasks.Where(w => w.ProjectId == s.Key)) //Task data
                        }).Distinct().ToList();

                    foreach (var p in taskLists) //projectTasks.Select(s => s.ProjectId).Distinct().ToList()
                    {
                        int.TryParse(p.ProjectId, out int projectId);
                        var tempProjectTasks = projectTasks.Where(w => w.ProjectId == projectId).ToList();

                        #region //停用
                        //int maxLevel = (int)tempProjectTasks.Max(x => x.TaskLevel);
                        //while (maxLevel > 0)
                        //{
                        //    var currentTask = tempProjectTasks.Where(x => x.TaskLevel == maxLevel).OrderBy(x => x.TaskRoute);
                        //    foreach (var task in currentTask)
                        //    {
                        //        if (task.UserTask == 1)
                        //        {
                        //            tempProjectTasks
                        //                .Where(x => x.ProjectTaskId == task.ProjectTaskId)
                        //                .ToList()
                        //                .ForEach(x =>
                        //                {
                        //                    x.DisplayTask = 1;
                        //                });
                        //        }
                        //        else
                        //        {
                        //            if (tempProjectTasks.Exists(x => x.ParentTaskId == task.ProjectTaskId && x.DisplayTask == 1))
                        //            {
                        //                tempProjectTasks
                        //                    .Where(x => x.ProjectTaskId == task.ProjectTaskId)
                        //                    .ToList()
                        //                    .ForEach(x =>
                        //                    {
                        //                        x.DisplayTask = 1;
                        //                    });
                        //            }
                        //        }
                        //    }
                        //    maxLevel--;
                        //}
                        #endregion

                        var data = new DHXTree
                        {
                            value = p.ProjectName,
                            id = "-1",
                            opened = 0 < OpenedLevel ? true : false,
                            level = 0,
                            status = p.ProjectStatus,
                            //start_date = p.StartDate ?? "",
                            //end_date = p.EndDate ?? "",
                            //duration = p.Duration,
                            //deferred_status = "",
                            //status_name = p.ProjectStatusName,
                            //task_status = "",
                            //task_status_name = ""
                        };

                        if (tempProjectTasks.Count > 0)
                        {
                            data.items = tempProjectTasks
                                .Where(x => x.TaskLevel == 1 && x.ParentTaskId == Convert.ToInt32(data.id)) // && x.DisplayTask == 1
                                .OrderBy(x => x.TaskLevel)
                                .ThenBy(x => x.TaskRoute)
                                .ThenBy(x => x.TaskSort)
                                .Select(x => new DHXTree
                                {
                                    value = x.TaskName,
                                    id = x.ProjectTaskId.ToString(),
                                    opened = data.level < OpenedLevel ? true : (ReplyStatus.Length > 0 ? true : false),
                                    level = (int)x.TaskLevel,
                                    status = x.ProjectStatus,
                                    //start_date = x.TaskStart ?? "",
                                    //end_date = x.TaskEnd ?? "",
                                    //duration = x.TaskDuration ?? 1,
                                    //deferred_status = x.DeferredCount > 0 ? "Y" : "N",
                                    //status_name = x.ProjectStatusName,
                                    //task_status = x.TaskStatus,
                                    //task_status_name = x.TaskStatusName
                                })
                                .ToList();

                            if (data.items.Count > 0) Recursion(data.items);
                        }

                        void Recursion(List<DHXTree> taskTree)
                        {
                            taskTree.ForEach(f =>
                            {
                                f.items = tempProjectTasks
                                            .Where(x => x.TaskLevel == (f.level + 1) && x.ParentTaskId == Convert.ToInt32(f.id)) // && x.DisplayTask == 1
                                            .OrderBy(x => x.TaskLevel)
                                            .ThenBy(x => x.TaskRoute)
                                            .ThenBy(x => x.TaskSort)
                                            .Select(x => new DHXTree
                                            {
                                                value = x.TaskName,
                                                id = x.ProjectTaskId.ToString(),
                                                opened = f.level < OpenedLevel ? true : (ReplyStatus.Length > 0 ? true : false),
                                                level = (int)x.TaskLevel,
                                                status = x.ProjectStatus,
                                                //start_date = x.TaskStart ?? "",
                                                //end_date = x.TaskEnd ?? "",
                                                //duration = x.TaskDuration ?? 1,
                                                //deferred_status = x.DeferredCount > 0 ? "Y" : "N",
                                                //status_name = x.ProjectStatusName,
                                                //task_status = x.TaskStatus,
                                                //task_status_name = x.TaskStatusName
                                            })
                                            .ToList();
                                if (f.items.Count > 0) Recursion(f.items);
                            });
                        }

                        taskLists
                            .Where(x => x.ProjectId == p.ProjectId)
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

        #region //GetProjectTaskGantt -- 取得專案任務資料(甘特圖) -- Chia Yuan 2024.05.16
        public string GetProjectTaskGantt(int ProjectId, string ProjectUserId, string TaskUserId, string ProjectMaster, int CompanyId, int DepartmentId, string Project, string TaskName
            , string ProjectStatus, string TaskStatus, string ReplyStatus, string TaskStartDate, string TaskEndDate
            , string TreeGroupBy)
        {
            try
            {
                var projects = ProjectPageList(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , "", -1, -1);

                var projectIds = projects.Select(s => s.ProjectId).ToList();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetProjectTaskRecursionString(ref dynamicParameters, projectIds);
                    sql += @"SELECT c.ProjectId
	                    , ISNULL(FORMAT(ISNULL(ISNULL(c.ActualStart, c.EstimateStart), c.PlannedStart), 'yyyy-MM-dd HH:mm'), '') ProjectStart
	                    , ISNULL(FORMAT(ISNULL(ISNULL(c.ActualEnd, c.EstimateEnd), c.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') ProjectEnd
	                    , ISNULL(ISNULL(ISNULL(c.ActualDuration, c.EstimateDuration), c.PlannedDuration), 1) ProjectDuration
	                    , FORMAT(c.PlannedStart, 'yyyy-MM-dd HH:mm') ProjectPlannedStart
	                    , FORMAT(c.PlannedEnd, 'yyyy-MM-dd HH:mm') ProjectPlannedEnd
	                    , ISNULL(c.PlannedDuration, 1) ProjectPlannedDuration
	                    , FORMAT(ISNULL(c.ActualStart, c.EstimateStart), 'yyyy-MM-dd HH:mm') ProjectEstimateStart
	                    , FORMAT(ISNULL(c.ActualEnd, c.EstimateEnd), 'yyyy-MM-dd HH:mm') ProjectEstimateEnd
	                    , ISNULL(ISNULL(c.ActualDuration, c.EstimateDuration), 1) ProjectEstimateDuration
                        , c.ProjectNo, c.ProjectName, c.ProjectDesc, c.CustomerRemark, c.ProjectRemark
                        , c.ProjectStatus, s1.StatusName ProjectStatusName
                        , c.ProjectType, t1.TypeName ProjectTypeName
                        , c.ProjectAttribute, t2.TypeName ProjectAttributeName
                        , c.WorkTimeStatus, c.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                        , a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort, a.TaskRouteName
                        , b.TaskName, b.TaskDesc
                        , b.TaskStatus, s2.StatusName TaskStatusName
	                    , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
	                    , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
	                    , ISNULL(ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration), 1) TaskDuration
	                    , FORMAT(b.PlannedStart, 'yyyy-MM-dd HH:mm') TaskPlannedStart
	                    , FORMAT(b.PlannedEnd, 'yyyy-MM-dd HH:mm') TaskPlannedEnd
	                    , ISNULL(b.PlannedDuration, 1) TaskPlannedDuration
	                    , FORMAT(ISNULL(b.ActualStart, b.EstimateStart), 'yyyy-MM-dd HH:mm') TaskEstimateStart
	                    , FORMAT(ISNULL(b.ActualEnd, b.EstimateEnd), 'yyyy-MM-dd HH:mm') TaskEstimateEnd
	                    , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), 1) TaskEstimateDuration
                        , ISNULL(d.DepartmentNo, m.MasterDepartmentNo) DepartmentNo
                        , ISNULL(d.DepartmentName, m.MasterDepartmentName) DepartmentName
                        , ISNULL(e.CompanyName, m.MasterCompanyName) CompanyName
                        , ISNULL(e.CompanyNo, m.MasterCompanyNo) CompanyNo
                        , ISNULL(e.LogoIcon, -1) LogoIcon
                        , 'Project' TaskType
                        , (SELECT cb.UserId, cb.UserNo, cb.UserName, cb.Gender
	                        FROM PMIS.ProjectUser ca
	                        INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
	                        WHERE ca.ProjectId = c.ProjectId
	                        AND cb.[Status] = 'A'
	                        ORDER BY ca.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) ProjectUser
                        , (SELECT db.UserId, db.UserNo, db.UserName, db.Gender
	                        FROM PMIS.TaskUser da
	                        INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                        WHERE da.TaskType = 'Project'
	                        AND da.TaskId = a.ProjectTaskId
	                        AND db.[Status] = 'A'
	                        ORDER BY da.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) TaskUser
                        , (SELECT fb.UserNo, fb.UserName
                            , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                            , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                            , fa.ReplyContent, fa.DeferredDuration, fa.ReplyStatus, fc.StatusName ReplyStatusName
	                        FROM PMIS.TaskReply fa
		                    INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
		                    INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                        WHERE fa.ReplyType = 'Project'
	                        AND fa.TaskId = a.ProjectTaskId
		                    AND fb.[Status] = 'A'
	                        ORDER BY fa.ReplyDate
	                        FOR JSON PATH, ROOT('data')
                        ) TaskReply
                        , ISNULL((
	                        SELECT COUNT(1)
	                        FROM PMIS.TaskUser ea
                            INNER JOIN BAS.[User] eb ON eb.UserId = ea.UserId
	                        WHERE ea.TaskId = a.ProjectTaskId
	                        AND ea.TaskType = 'Project'
                            AND eb.[Status] = 'A'), 0
                        ) TaskUserCount
                        , ISNULL((
	                        SELECT COUNT(1)
	                        FROM PMIS.ProjectUser ha
                            INNER JOIN BAS.[User] hb ON hb.UserId = ha.UserId
	                        WHERE ha.ProjectId = c.ProjectId
                            AND hb.[Status] = 'A'), 0
                        ) ProjectUserCount
                        , ISNULL((
                            SELECT COUNT(1)
                            FROM PMIS.TaskReply ra
                            WHERE ra.TaskId = a.ProjectTaskId
                            AND ra.ReplyType = 'Project'
                            AND ra.ReplyStatus = 'D'), 0
                        ) DeferredCount
                        , ISNULL((
                            SELECT COUNT(1)
                            FROM PMIS.ProjectTask ga
                            WHERE ga.ParentTaskId = a.ProjectTaskId), 0
                        ) SubTaskCount
                        , m.*
                        FROM @projectTask a
                        INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                        INNER JOIN PMIS.Project c ON c.ProjectId = b.ProjectId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = c.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = c.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = c.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                        INNER JOIN BAS.[Type] t2 ON t2.TypeNo = c.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                        LEFT JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        LEFT JOIN BAS.Company e ON e.CompanyId = d.CompanyId
                        OUTER APPLY (
	                        SELECT ROW_NUMBER() OVER (PARTITION BY mb.ProjectId ORDER BY mb.AgentSort) SortPM
                            , ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                            , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
                            , mc.CompanyNo MasterCompanyNo, mc.CompanyName MasterCompanyName
                            , md.DepartmentNo MasterDepartmentNo, md.DepartmentName MasterDepartmentName
	                        FROM BAS.[User] ma 
	                        INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId
                            INNER JOIN BAS.Department md ON md.DepartmentId = ma.DepartmentId
                            INNER JOIN BAS.Company mc ON mc.CompanyId = md.CompanyId
	                        WHERE mb.ProjectId = c.ProjectId
                        ) m";
                    //if (!string.IsNullOrWhiteSpace(TaskStatus))
                    //{
                    //    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                    //        , "TaskStatus"
                    //        , @" AND b.TaskStatus IN @TaskStatus", TaskStatus.Split(','));
                    //}
                    //if (!string.IsNullOrWhiteSpace(ReplyStatus))
                    //{
                    //    sql += " AND r.DeferredCount > 0";
                    //}
                    sql += @" ORDER BY ISNULL(e.CompanyId, -1), ISNULL(d.DepartmentId, -1), c.ProjectId, a.TaskLevel, a.TaskRoute, a.TaskSort";
                    var projectTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //專案任務連結資料
                    List<ProjectTaskLink> projectTaskLinks = new List<ProjectTaskLink>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ProjectTaskLinkId, a.SourceTaskId, a.TargetTaskId, a.LinkType
                            FROM PMIS.ProjectTaskLink a
                            INNER JOIN PMIS.ProjectTask b ON a.SourceTaskId = b.ProjectTaskId
                            WHERE b.ProjectId IN @ProjectIds
                            ORDER BY b.ProjectId, a.ProjectTaskLinkId";
                    dynamicParameters.Add("ProjectIds", projectTasks.Select(s => Convert.ToInt32(s.ProjectId)).Distinct().ToArray());
                    projectTaskLinks = sqlConnection.Query<ProjectTaskLink>(sql, dynamicParameters).ToList();
                    #endregion

                    List<ProjectList> projLists = projectTasks
                        .GroupBy(g => g.ProjectId)
                        .Select(s => new ProjectList
                        {
                            ProjectId = Convert.ToString(s.Key),
                            ProjectNo = s.First()?.ProjectNo,
                            ProjectName = s.First()?.ProjectName,
                            StartDate = s.First()?.ProjectStart ?? "",
                            EndDate = s.First()?.ProjectEnd ?? "",
                            Duration = s.First()?.ProjectDuration,
                            ProjectPlannedStart = s.First()?.ProjectPlannedStart,
                            ProjectPlannedEnd= s.First()?.ProjectPlannedEnd,
                            ProjectPlannedDuration = s.First()?.ProjectPlannedDuration,
                            ProjectDesc = s.First()?.ProjectDesc,
                            CompanyName = s.First()?.CompanyName,
                            CompanyNo = s.First()?.CompanyNo,
                            LogoIcon = s.First()?.LogoIcon,
                            DepartmentNo = s.First()?.DepartmentNo,
                            DepartmentName = s.First()?.DepartmentName,
                            ProjectMaster = s.First()?.ProjectMaster,
                            MasterUserNo = s.First()?.UserNo,
                            MasterGender = s.First()?.Gender,
                            ProjectUser = s.First()?.ProjectUser,
                            ProjectUserCount = s.First()?.ProjectUserCount,
                            ProjectStatus = s.First()?.ProjectStatus,
                            ProjectStatusName = s.First()?.ProjectStatusName,
                            ProjectType = s.First()?.ProjectType,
                            ProjectTypeName = s.First()?.ProjectTypeName,
                            ProjectAttribute = s.First()?.ProjectAttribute,
                            ProjectAttributeName = s.First()?.ProjectAttributeName,
                            WorkTimeStatus = s.First()?.WorkTimeStatus,
                            WorkTimeStatusName = s.First()?.WorkTimeStatusName,
                            TotalCount = projects.FirstOrDefault(f => f.ProjectId == s.Key)?.TotalCount,
                            ProjectTask = JsonConvert.SerializeObject(s.ToList()) //Task data
                        }).Distinct().ToList();

                    List<DHXTask> dHXTasks = new List<DHXTask>();
                    List<DHXTask> dHXMaster = new List<DHXTask>();
                    switch (TreeGroupBy)
                    {
                        case "0.1":
                            dHXMaster.AddRange(projLists.GroupBy(g => g.CompanyNo).Select(s => new DHXTask
                            {
                                id = string.Format("{0},{1}", TreeGroupBy, s.Key),
                                text = s.First()?.CompanyName,
                                start_date = s.Min(m => m.StartDate),
                                duration = s.Sum(s1 => s1.Duration).ToString(),
                                progress = "0",
                                parent = "0",
                                type = "project",
                                order = 1,
                                open = 0,
                                planned_start = s.Min(m => m.ProjectPlannedStart),
                                planned_end = s.Max(m => m.ProjectPlannedEnd),
                                planned_duration = Convert.ToDateTime(s.Max(m => m.ProjectPlannedEnd)).Subtract(Convert.ToDateTime(s.Max(m => m.ProjectPlannedStart))).TotalMinutes.ToString(),
                            }));
                            break;
                        case "0.2":
                            dHXMaster.AddRange(projLists.GroupBy(g => g.DepartmentNo).Select(s => new DHXTask
                            {
                                id = string.Format("{0},{1}", TreeGroupBy, s.Key),
                                text = s.First()?.DepartmentName,
                                start_date = s.Min(m => m.StartDate),
                                duration = s.Sum(s1 => s1.Duration).ToString(),
                                progress = "0",
                                parent = "0",
                                type = "project",
                                order = 1,
                                open = 0,
                                planned_start = s.Min(m => m.ProjectPlannedStart),
                                planned_end = s.Max(m => m.ProjectPlannedEnd),
                                planned_duration = Convert.ToDateTime(s.Max(m => m.ProjectPlannedEnd)).Subtract(Convert.ToDateTime(s.Max(m => m.ProjectPlannedStart))).TotalMinutes.ToString(),
                            }));
                            break;
                        case "0.3":
                            dHXMaster.AddRange(projLists.GroupBy(g => g.MasterUserNo).Select(s => new DHXTask
                            {
                                id = string.Format("{0},{1}", TreeGroupBy, s.Key),
                                text = s.First()?.ProjectMaster,
                                start_date = s.Min(m => m.StartDate),
                                duration = s.Sum(s1 => s1.Duration).ToString(),
                                progress = "0",
                                parent = "0",
                                type = "project",
                                order = 1,
                                open = 0,
                                planned_start = s.Min(m => m.ProjectPlannedStart),
                                planned_end = s.Max(m => m.ProjectPlannedEnd),
                                planned_duration = Convert.ToDateTime(s.Max(m => m.ProjectPlannedEnd)).Subtract(Convert.ToDateTime(s.Max(m => m.ProjectPlannedStart))).TotalMinutes.ToString(),
                            }));
                            break;
                        case "0.0":
                            dHXMaster.Add(new DHXTask
                            {
                                id = TreeGroupBy,
                            });
                            break;
                    }

                    #region //甘特圖結構
                    dHXMaster.ForEach(master =>
                    {
                        var tempProjLists = new List<ProjectList>();
                        var treeGroupBy = master.id.Split(',');
                        if (treeGroupBy.Length > 1)
                        {
                            dHXTasks.Add(master);

                            switch (treeGroupBy[0])
                            {
                                case "0.1":
                                    tempProjLists = projLists.Where(w => w.CompanyNo == treeGroupBy[1]).ToList();
                                    break;
                                case "0.2":
                                    tempProjLists = projLists.Where(w => w.DepartmentNo == treeGroupBy[1]).ToList();
                                    break;
                                case "0.3":
                                    tempProjLists = projLists.Where(w => w.MasterUserNo == treeGroupBy[1]).ToList();
                                    break;
                            }
                        }
                        else
                        {
                            tempProjLists = projLists;
                        }

                        tempProjLists.ForEach(proj =>
                        {
                            dHXTasks.Add(
                                new DHXTask
                                {
                                    id = string.Format("1.1,{0}", proj.ProjectId),
                                    text = treeGroupBy.Length > 1 ? string.Format("{0}-{1}", proj.ProjectMaster, proj.ProjectName) : proj.ProjectName,
                                    start_date = proj.StartDate,
                                    duration = proj.Duration.ToString(),
                                    progress = "0",
                                    parent = treeGroupBy.Length > 1 ? string.Format("{0},{1}", treeGroupBy[0], treeGroupBy[1]) : "0",
                                    type = "project",
                                    order = 1,
                                    open = 0,
                                    planned_start = proj.ProjectPlannedStart,
                                    planned_end = proj.ProjectPlannedEnd,
                                    planned_duration = proj.ProjectPlannedDuration.ToString(),
                                    sub_user = proj.ProjectUserCount
                                });

                            foreach (var task in projectTasks.Where(w => Convert.ToString(w.ProjectId) == proj.ProjectId))
                            {
                                DateTime startDate = task.ActualStart != null ? Convert.ToDateTime(task.ActualStart) : Convert.ToDateTime(task.EstimateStart);
                                DateTime endDate = task.ActualEnd != null ? Convert.ToDateTime(task.ActualEnd) : Convert.ToDateTime(task.EstimateEnd);
                                int duration = task.ActualDuration != null ? Convert.ToInt32(task.ActualDuration) : Convert.ToInt32(task.EstimateDuration);

                                dHXTasks.Add(
                                    new DHXTask
                                    {
                                        id = task.ProjectTaskId.ToString(),
                                        text = task.TaskName,
                                        start_date = task.TaskStart,
                                        duration = task.TaskDuration.ToString(),
                                        progress = "0",
                                        parent = task.ParentTaskId > -1 ? task.ParentTaskId.ToString() : string.Format("1.1,{0}", proj.ProjectId),
                                        type = task.SubTaskCount > 0 ? "project" : "task",
                                        order = Convert.ToInt32(task.TaskSort),
                                        open = 0,
                                        planned_start = task.TaskPlannedStart,
                                        planned_end = task.TaskPlannedEnd,
                                        planned_duration = task.TaskPlannedDuration.ToString(),
                                        deferred_status = task.DeferredCount > 0 ? "Y" : "N",
                                        sub_user = task.TaskUserCount,
                                        task_status = task.TaskStatus,
                                        task_status_name = task.TaskStatusName
                                    });
                            }
                        });
                    });
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

        #region //GetProjectTaskTimeline -- 取得專案任務資料(時間軸圖) -- Chia Yuan 2024.07.15
        public string GetProjectTaskTimeline(int ProjectId, string ProjectUserId, string TaskUserId, string ProjectMaster, int CompanyId, int DepartmentId, string Project, string TaskName
            , string ProjectStatus, string TaskStatus, string ReplyStatus, string TaskStartDate, string TaskEndDate
            , string DurationSplit)
        {
            try
            {
                if (!Regex.IsMatch(DurationSplit, "^(yyyy-MM|yyyy-MM-dd|yyyy-MM-dd HH:00|yyyy-MM-dd HH:mm)$", RegexOptions.IgnoreCase)) throw new SystemException("【節點間隔】單位錯誤!");

                var projects = ProjectPageList(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , "", -1, -1);

                var projectIds = projects.Select(s => s.ProjectId).ToList();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //專案資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetProjectTaskRecursionString(ref dynamicParameters, projectIds);
                    sql += @"SELECT x.*
                            , ROUND(x.SubTaskC / x.SubTask, 2) CompletionRate
                            FROM (
	                            SELECT b.ProjectNo, b.ProjectName, b.ProjectDesc, b.ProjectStart, b.ProjectEnd, b.ProjectDuration
                                , b.ProjectType, b.ProjectTypeName
                                , b.ProjectAttribute, b.ProjectAttributeName
                                , b.ProjectStatus, b.ProjectStatusName
                                , b.WorkTimeStatus, b.WorkTimeStatusName
                                , b.DepartmentNo, b.DepartmentName, b.CompanyName, b.CompanyNo, b.LogoIcon
                                , b.SubTask, b.SubTaskC, b.UserId, b.UserNo, b.UserName, b.Gender, b.ProjectMaster
	                            , a.*
	                            , DATEDIFF(MINUTE, a.TaskStart, a.TaskEnd) diffTaskDuration
	                            , DATEDIFF(MINUTE, b.ProjectStart, b.ProjectEnd) diffProjectDuration
	                            , LAG(a.ProjectTaskId) OVER (PARTITION BY a.ProjectId ORDER BY a.TaskEnd, a.TaskLevel DESC, a.TaskSort) PreTaskId
	                            , LEAD(a.ProjectTaskId) OVER (PARTITION BY a.ProjectId ORDER BY a.TaskEnd, a.TaskLevel DESC, a.TaskSort) NextTaskId
	                            , LAG(a.TaskEnd) OVER (PARTITION BY a.ProjectId ORDER BY a.TaskEnd, a.TaskLevel DESC, a.TaskSort) PreTaskEnd
	                            , LEAD(a.TaskEnd) OVER (PARTITION BY a.ProjectId ORDER BY a.TaskEnd, a.TaskLevel DESC, a.TaskSort) NextTaskEnd
	                            , ROW_NUMBER() OVER (PARTITION BY a.ProjectId ORDER BY a.TaskStart, a.TaskLevel DESC, a.TaskSort) SortDuration
	                            , ROW_NUMBER() OVER (PARTITION BY a.ProjectId ORDER BY a.TaskEnd, a.TaskLevel DESC, a.TaskSort) SortDuration2
	                            FROM (
		                            SELECT aa.ProjectTaskId, aa.ProjectId, aa.TaskName, ISNULL(aa.TaskDesc, '') TaskDesc
                                    , CASE ab.ReplyStatus WHEN 'D' THEN 'D' ELSE aa.TaskStatus END TaskStatus
                                    , s2.StatusName + CASE ab.ReplyStatus WHEN 'D' THEN '(' + ab.ReplyStatusName + ')' ELSE '' END TaskStatusName
		                            , FORMAT(ISNULL(ISNULL(aa.ActualStart, aa.EstimateStart), aa.PlannedStart), @DurationSplit) TaskStart
		                            , FORMAT(ISNULL(ISNULL(aa.ActualEnd, aa.EstimateEnd), aa.PlannedEnd), @DurationSplit) TaskEnd
		                            , ISNULL(ISNULL(ISNULL(aa.ActualDuration, aa.EstimateDuration), aa.PlannedDuration), 1) TaskDuration
		                            , aa.TaskSort
                                    , ISNULL (
                                        STUFF((
                                            SELECT ',' + bu.UserName
                                            FROM PMIS.TaskUser ba
                                            INNER JOIN BAS.[User] bu ON bu.UserId = ba.UserId
                                            WHERE ba.TaskId = aa.ProjectTaskId 
			                                AND ba.TaskType = 'Project'
                                            ORDER BY bu.UserNo
                                            FOR XML PATH('')
                                        ), 1, 1, '')
                                    , '') TaskUser
                                    , ax.TaskLevel
		                            FROM PMIS.ProjectTask aa
		                            OUTER APPLY (
			                            SELECT TOP 1 ba.ReplyStatus, bb.StatusName ReplyStatusName
			                            FROM PMIS.TaskReply ba
                                        INNER JOIN BAS.[Status] bb ON bb.StatusNo = ba.ReplyStatus AND bb.StatusSchema = 'TaskReply.ReplyStatus'
			                            WHERE ba.TaskId = aa.ProjectTaskId
			                            AND ba.ReplyType = 'Project'
			                            ORDER BY ba.ReplyId DESC
		                            ) ab
                                    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = aa.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                                    INNER JOIN @projectTask ax ON ax.ProjectTaskId = aa.ProjectTaskId
	                            ) a
	                            CROSS APPLY (
		                            SELECT TOP 1 aa.ProjectId, aa.ProjectNo, aa.ProjectName, aa.ProjectDesc
                                    , aa.ProjectType, t1.TypeName ProjectTypeName
                                    , aa.ProjectAttribute, t2.TypeName ProjectAttributeName
                                    , aa.ProjectStatus, s1.StatusName ProjectStatusName
                                    , aa.WorkTimeStatus, s3.StatusName WorkTimeStatusName
		                            , CONVERT(FLOAT, ta.SubTask) SubTask
		                            , CONVERT(FLOAT, ISNULL(tc.SubTaskC, 1)) SubTaskC
		                            , FORMAT(ISNULL(ISNULL(aa.ActualStart, aa.EstimateStart), aa.PlannedStart), @DurationSplit) ProjectStart
		                            , FORMAT(ISNULL(ISNULL(aa.ActualEnd, aa.EstimateEnd), aa.PlannedEnd), @DurationSplit) ProjectEnd
                                    , ISNULL(ISNULL(ISNULL(aa.ActualDuration, aa.EstimateDuration), aa.PlannedDuration), 1) ProjectDuration
                                    , ISNULL(ad.DepartmentNo, '') DepartmentNo
                                    , ISNULL(ad.DepartmentName, '') DepartmentName
                                    , ISNULL(ac.CompanyName, '') CompanyName
                                    , ISNULL(ac.CompanyNo, '') CompanyNo
                                    , ISNULL(ac.LogoIcon, -1) LogoIcon
                                    , am.*
		                            FROM PMIS.Project aa
                                    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = aa.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                                    INNER JOIN BAS.[Status] s3 ON s3.StatusNo = aa.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                                    INNER JOIN BAS.[Type] t1 ON t1.TypeNo = aa.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                                    INNER JOIN BAS.[Type] t2 ON t2.TypeNo = aa.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                                    LEFT JOIN BAS.Department ad ON ad.DepartmentId = aa.DepartmentId
                                    LEFT JOIN BAS.Company ac ON ac.CompanyId = ad.CompanyId
                                    OUTER APPLY (
                                        SELECT COUNT(1) SubTask
                                        FROM PMIS.ProjectTask ba
                                        WHERE ba.ProjectId= aa.ProjectId 
                                    ) ta
                                    OUTER APPLY (
                                        SELECT COUNT(1) SubTaskC
                                        FROM PMIS.ProjectTask ca
                                        WHERE ca.ProjectId= aa.ProjectId 
			                            AND ca.TaskStatus = 'C'
                                    ) tc
                                    OUTER APPLY (
	                                    SELECT ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                                        , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
	                                    FROM BAS.[User] ma 
	                                    INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId 
	                                    WHERE mb.ProjectId = aa.ProjectId
                                        AND mb.AgentSort = 1
                                    ) am
		                            WHERE aa.ProjectId = a.ProjectId
	                            ) b
                            ) x
                            ORDER BY x.ProjectId, x.SortDuration2";
                    dynamicParameters.Add("DurationSplit", DurationSplit);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");
                    #endregion

                    #region //example
                    //{
                    //    "id": "u1692878476101",
                    //    "type": "circle",
                    //    "width": 20,
                    //    "height": 20,
                    //    "fill": "#11C169",
                    //    "stroke": "#11C169",
                    //    "fontStyle": "normal",
                    //    "fontColor": "#FFFFFF",
                    //    "x": 320,
                    //    "y": -50,
                    //},
                    //{
                    //    "id": "u1692946984831",
                    //    "type": "dash",
                    //    "text": "123",
                    //    "textAlign": "center",
                    //    "connectType": "straight",
                    //    "from": "u1692878359567",
                    //    "to": "u1692878476100",
                    //    "fromSide": "right",
                    //    "toSide": "left"
                    //},
                    //{
                    //    "id": 1,
                    //    "type": "task",
                    //    "x": -40,
                    //    "y": -125,
                    //    "TaskStatus": "https://snippet.dhtmlx.com/codebase/data/diagram/05/img/desktop.svg",
                    //    "Name": "-kick off meeting.",
                    //    "Text": "",
                    //    "TaskUser": "胡嘉原"
                    //},
                    //{
                    //    "id": "duration-1",
                    //    "type": "duration",
                    //    "x": -40,
                    //    "y": -50,
                    //    "Text": "2024-07-15 10:00 ~ 2024-07-15 12:00"
                    //},
                    #endregion

                    projectIds = result.Select(s => { return Convert.ToInt32(s.ProjectId); }).Distinct().ToList();
                    string[] colorArray = new string[] { "#FF003E", "#FF8000", "#FFCC00", "#11AA00", "#0AC4A2", "#2D30EA", "#9700FF" };
                    List<DHXCircle> dHXCircle = new List<DHXCircle>();
                    List<TimeLineShapes> dHXShapes = new List<TimeLineShapes>();
                    List<DHXDiagramLine> dHXDiagramLines = new List<DHXDiagramLine>();
                    List<GroupDiagram> groupDiagram = new List<GroupDiagram>();

                    projectIds.ForEach(projectId =>
                    {
                        int x1 = 20, x2 = -30;
                        int y1 = -50, y2 = 0, y3 = -40;
                        int d1 = 160, w_node = 45, w_info = 150, w_date = 140;
                        int i = 0;

                        dHXCircle = new List<DHXCircle>();
                        dHXShapes = new List<TimeLineShapes>();
                        dHXDiagramLines = new List<DHXDiagramLine>();
                        GroupDiagram tempGroupDiagram = new GroupDiagram();

                        var tempResult = result.Where(w => w.ProjectId == projectId);

                        #region //Task Info
                        foreach (var item in tempResult)
                        {
                            if (item.TaskEnd == item.PreTaskEnd) y2 += 125; //同一時間軸，調整y軸上的位子

                            //節點
                            if (i == 0)
                            {
                                double rate = item.CompletionRate * 100;
                                string rateColor = rate < 5 ? colorArray[0] : (rate >= 5 && rate < 25 ? colorArray[1] : (rate >= 25 && rate < 75 ? colorArray[2] : (rate >= 75 && rate <= 99 ? colorArray[4] : colorArray[3])));
                                dHXCircle.Add(new DHXCircle
                                {
                                    id = string.Format(@"P_{0}", item.ProjectTaskId),
                                    type = "circle",
                                    x = x2 - (w_node * 2),
                                    y = y2,
                                    width = w_node,
                                    height = w_node,
                                    text = string.Format(@"{0}%", rate),
                                    fill = rateColor,
                                    stroke = rateColor,
                                    fontStyle = "normal",
                                    fontColor = "#FFFFFF",
                                    //fill = colorArray[i % colorArray.Length],
                                    //stroke = colorArray[i % colorArray.Length],
                                    //x = x1,
                                    //y = y1,
                                });
                            }
                            //x1 += d1;

                            //狀態資訊
                            dHXShapes.Add(new TimeLineShapes
                            {
                                id = string.Format(@"{0}", item.ProjectTaskId),
                                type = "status",
                                x = x2,
                                y = y2,
                                width = w_node,
                                height = w_node,
                                text = item.TaskStatusName,
                                TaskStatus = item.TaskStatus,
                                IconPath = IconPath(item.TaskStatus)
                                //TaskStatus = i == 0 ? "/Content/images/pmis/icon-progress.png" : IconPath(item.TaskStatus)
                                //"https://snippet.dhtmlx.com/codebase/data/diagram/05/img/desktop.svg"
                            });

                            //事件資訊
                            dHXShapes.Add(new TimeLineShapes
                            {
                                id = string.Format(@"T_{0}", item.ProjectTaskId),
                                type = "task",
                                x = x2 - ((w_info - w_node) / 2),
                                y = y2 + w_node,
                                width = w_info,
                                height = 60,
                                Name = Regex.Replace(item.TaskName, @"[^\w\、,，.。()（）] ", ""),
                                text = Regex.Replace(item.TaskDesc, @"[^\w\、,，.。()（）] ", ""),
                                TaskUser = item.TaskUser
                            });

                            if (item.TaskEnd != item.NextTaskEnd)
                            {
                                //日期資訊
                                dHXShapes.Add(new TimeLineShapes
                                {
                                    id = string.Format(@"D_{0}", item.ProjectTaskId),
                                    type = "duration",
                                    x = x2 - ((w_date - w_node) / 2),
                                    y = y3,
                                    width = w_date,
                                    height = 45,
                                    DateLine = i == 0 ? string.Format("{0}\r\n~{1}", item.TaskStart, item.TaskEnd) : item.TaskEnd
                                });

                                y2 = 0;
                                x2 += d1;
                            }
                            i++;
                        }
                        #endregion

                        #region //Line Json組成
                        foreach (var item in tempResult.Where(w => w.PreTaskId != null))
                        {
                            if (item.TaskEnd != item.PreTaskEnd)
                            {
                                dHXDiagramLines.Add(new DHXDiagramLine
                                {
                                    id = string.Format(@"L_{0}_{1}", item.PreTaskId, item.ProjectTaskId),
                                    type = "line",
                                    from = string.Format(@"{0}", item.PreTaskId),
                                    to = string.Format(@"{0}", item.ProjectTaskId),
                                    connectType = "elbow",
                                    forwardArrow = ""
                                });
                            }
                            else
                            {
                                dHXDiagramLines.Add(new DHXDiagramLine
                                {
                                    id = string.Format(@"L_{0}_{1}", item.PreTaskId, item.ProjectTaskId),
                                    type = "dash",
                                    from = string.Format(@"{0}", item.PreTaskId),
                                    fromSide = "bottom",
                                    to = string.Format(@"{0}", item.ProjectTaskId),
                                    toSide = "top",
                                    connectType = "straight",
                                    forwardArrow = ""
                                });
                            }
                        }
                        #endregion

                        var project = result.FirstOrDefault(f => f.ProjectId == projectId);
                        if (project != null)
                        {
                            tempGroupDiagram.ProjectId = projectId.ToString();
                            tempGroupDiagram.ProjectNo = project?.ProjectNo;
                            tempGroupDiagram.ProjectName = project?.ProjectName;
                            tempGroupDiagram.StartDate = project?.ProjectStart;
                            tempGroupDiagram.EndDate = project?.ProjectEnd;
                            tempGroupDiagram.ProjectStatus = project?.ProjectStatus;
                            tempGroupDiagram.ProjectStatusName = project?.ProjectStatusName;
                            tempGroupDiagram.LogoIcon = project?.LogoIcon;
                            tempGroupDiagram.CompanyName = project?.CompanyName;
                            tempGroupDiagram.DepartmentName = project?.DepartmentName;
                            tempGroupDiagram.MasterGender = project?.Gender;
                            tempGroupDiagram.ProjectMaster = project?.ProjectMaster;
                            tempGroupDiagram.WorkTimeStatusName = project?.WorkTimeStatusName;
                            tempGroupDiagram.ProjectTypeName = project?.ProjectTypeName;
                            tempGroupDiagram.ProjectDesc = project?.ProjectDesc;
                            tempGroupDiagram.CompletionRate = project?.CompletionRate * 100;
                        }

                        tempGroupDiagram.dHXCircle = dHXCircle;
                        tempGroupDiagram.dHXShapes = dHXShapes;
                        tempGroupDiagram.dHXDiagramLines = dHXDiagramLines;
                        groupDiagram.Add(tempGroupDiagram);
                    });

                    string IconPath(string status)
                    {
                        string path = string.Empty;
                        switch (status)
                        {
                            case "B":
                                path = "/Content/images/pmis/icon-not-start.png";
                                break;
                            case "C":
                                path = "/Content/images/pmis/icon-finish.png";
                                break;
                            case "I":
                                path = "/Content/images/pmis/icon-progress.png";
                                break;
                            case "D":
                                path = "/Content/images/pmis/icon-pending.png";
                                break;
                            case "T":
                                path = "/Content/images/pmis/icon-pending.png";
                                break;
                        }
                        return path;
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        dataGroupDiagram = groupDiagram
                        //dataNodes = dHXCircle,
                        //dataLines = dHXDiagramLines,
                        //dataShapes = dHXShapes,
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

        #region //GetProject -- 取得專案資料 -- Chia Yuan 2024.05.16
        public string GetProject(int ProjectId, int ProjectTaskId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //專案資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "c.ProjectId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", ISNULL(FORMAT(ISNULL(ISNULL(c.ActualStart, c.EstimateStart), c.PlannedStart), 'yyyy-MM-dd HH:mm'), '') ProjectStart
                        , ISNULL(FORMAT(ISNULL(ISNULL(c.ActualEnd, c.EstimateEnd), c.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') ProjectEnd
                        , ISNULL(ISNULL(ISNULL(c.ActualDuration, c.EstimateDuration), c.PlannedDuration), 1) ProjectDuration
                        , ISNULL(FORMAT(c.PlannedStart, 'yyyy-MM-dd HH:mm'), '') ProjectPlannedStart
                        , ISNULL(FORMAT(c.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') ProjectPlannedEnd
                        , ISNULL(c.PlannedDuration, 1) ProjectPlannedDuration
                        , ISNULL(FORMAT(c.EstimateStart, 'yyyy-MM-dd HH:mm'), '') ProjectEstimateStart
                        , ISNULL(FORMAT(c.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') ProjectEstimateEnd
                        , ISNULL(c.EstimateDuration, 1) ProjectEstimateDuration
                        , ISNULL(FORMAT(c.ActualStart, 'yyyy-MM-dd HH:mm'), '') ProjectActualStart
                        , ISNULL(FORMAT(c.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ProjectActualEnd
                        , ISNULL(c.ActualDuration, 1) ProjectActualDuration
                        , c.ProjectNo, c.ProjectName, c.ProjectDesc, c.CustomerRemark, c.ProjectRemark
                        , c.ProjectStatus, s1.StatusName ProjectStatusName
                        , c.ProjectType, t1.TypeName ProjectTypeName
                        , c.ProjectAttribute, t2.TypeName ProjectAttributeName
                        , c.WorkTimeStatus, c.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                        , ISNULL(d.DepartmentNo, '') DepartmentNo
                        , ISNULL(d.DepartmentName, '') DepartmentName
                        , ISNULL(e.CompanyName, '') CompanyName
                        , ISNULL(e.CompanyNo, '') CompanyNo
                        , ISNULL(e.LogoIcon, -1) LogoIcon
                        , 'Project' TaskType
                        , (SELECT DISTINCT cb.UserId, cb.UserNo, cb.UserName, cb.Gender, ca.AgentSort
	                        FROM PMIS.ProjectUser ca
	                        INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
	                        WHERE ca.ProjectId = c.ProjectId
	                        AND cb.[Status] = 'A'
	                        ORDER BY ca.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) ProjectUser
                        , (SELECT fb.UserNo, fb.UserName
                            , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                            , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                            , fa.ReplyContent, fa.DeferredDuration
                            , fa.ReplyStatus, fc.StatusName ReplyStatusName
                            , ISNULL(fd.ReplyFrequency, 0) ReplyFrequency
	                        FROM PMIS.TaskReply fa
	                        INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
	                        INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                        INNER JOIN PMIS.ProjectTask fd ON fd.ProjectTaskId = fa.TaskId
	                        WHERE fa.ReplyType = 'Project'
	                        AND fb.[Status] = 'A'
	                        AND fd.ProjectId = c.ProjectId
	                        ORDER BY fa.ReplyDate
	                        FOR JSON PATH, ROOT('data')
                        ) TaskReply
                        , ISNULL((
                            SELECT COUNT(1)
                            FROM PMIS.ProjectTask ga
                            WHERE ga.ProjectId = c.ProjectId), 0
                        ) SubTaskCount
                        , m.*";
                    sqlQuery.mainTables =
                        @" FROM PMIS.Project c
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = c.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = c.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = c.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                        INNER JOIN BAS.[Type] t2 ON t2.TypeNo = c.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                        LEFT JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        LEFT JOIN BAS.Company e ON e.CompanyId = d.CompanyId
                        OUTER APPLY (
	                        SELECT ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                            , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
	                        FROM BAS.[User] ma 
	                        INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId 
	                        WHERE mb.ProjectId = c.ProjectId
                            AND mb.AgentSort = 1
                        ) m";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND c.ProjectId = @ProjectId";
                    dynamicParameters.Add("ProjectId", ProjectId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SortNumber", @" AND m.SortPM = @SortNumber", 1);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters
                        , "ProjectTaskId"
                        , @" AND EXISTS (
                                SELECT TOP 1 1 
                                FROM PMIS.ProjectTask aa 
                                WHERE aa.ProjectId = c.ProjectId 
                                AND aa.ProjectTaskId = @ProjectTaskId)", ProjectTaskId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.ProjectId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
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

        #region //GetProjectTask -- 取得任務資料 -- Chia Yuan 2024.05.16
        public string GetProjectTask(int TaskId, int ProjectId, string TaskType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【任務類型】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得任務資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProjectTaskId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProjectId";
                    sqlQuery.mainTables =
                        @" FROM PMIS.ProjectTask a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectTaskId", @" AND a.ProjectTaskId = @ProjectTaskId", TaskId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectTaskId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion

                    #region //任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetProjectTaskRecursionString(ref dynamicParameters, result.Select(s => s.ProjectId).ToList());
                    sql += @"SELECT a.ProjectTaskId
	                    , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
                        , ISNULL(ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration), 1) TaskDuration
                        , ISNULL(FORMAT(b.PlannedStart, 'yyyy-MM-dd HH:mm'), '') TaskPlannedStart
                        , ISNULL(FORMAT(b.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') TaskPlannedEnd
                        , ISNULL(b.PlannedDuration, 1) TaskPlannedDuration
                        , ISNULL(FORMAT(b.EstimateStart, 'yyyy-MM-dd HH:mm'), '') TaskEstimateStart
                        , ISNULL(FORMAT(b.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') TaskEstimateEnd
                        , ISNULL(b.EstimateDuration, 1) TaskEstimateDuration
                        , ISNULL(FORMAT(b.ActualStart, 'yyyy-MM-dd HH:mm'), '') TaskActualStart
                        , ISNULL(FORMAT(b.ActualEnd, 'yyyy-MM-dd HH:mm'), '') TaskActualEnd
                        , ISNULL(b.ActualDuration, 1) TaskActualDuration
                        , a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort, a.TaskRouteName
                        , b.TaskName, b.TaskDesc
                        , b.TaskStatus, s2.StatusName TaskStatusName
                        , c.ProjectNo, c.ProjectName, c.ProjectDesc, c.CustomerRemark, c.ProjectRemark
                        , c.ProjectStatus, s1.StatusName ProjectStatusName
                        , c.ProjectType, t1.TypeName ProjectTypeName
                        , c.ProjectAttribute, t2.TypeName ProjectAttributeName
                        , c.WorkTimeStatus, c.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                        , ISNULL(d.DepartmentNo, '') DepartmentNo
                        , ISNULL(d.DepartmentName, '') DepartmentName
                        , ISNULL(e.CompanyName, '') CompanyName
                        , ISNULL(e.CompanyNo, '') CompanyNo
                        , ISNULL(e.LogoIcon, -1) LogoIcon
                        , 'Project' TaskType
                        , (SELECT db.UserId, db.UserNo, db.UserName, db.Gender
	                        FROM PMIS.TaskUser da
	                        INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                        WHERE da.TaskType = 'Project'
	                        AND da.TaskId = a.ProjectTaskId
	                        AND db.[Status] = 'A'
	                        ORDER BY da.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) TaskUser
                        , (SELECT fb.UserNo, fb.UserName
                            , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                            , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                            , fa.ReplyContent, fa.DeferredDuration
                            , fa.ReplyStatus, fc.StatusName ReplyStatusName
                            , ISNULL(fd.ReplyFrequency, 0) ReplyFrequency
	                        FROM PMIS.TaskReply fa
	                        INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
	                        INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                        INNER JOIN PMIS.ProjectTask fd ON fd.ProjectTaskId = fa.TaskId
	                        WHERE fa.ReplyType = 'Project'
	                        AND fb.[Status] = 'A'
	                        AND fd.ProjectTaskId = a.ProjectTaskId
	                        ORDER BY fa.ReplyDate
	                        FOR JSON PATH, ROOT('data')
                        ) TaskReply
                        , ISNULL((
                            SELECT COUNT(1)
                            FROM PMIS.ProjectTask ga
                            WHERE ga.ProjectId = c.ProjectId), 0
                        ) SubTaskCount
                        , m.*
                        FROM @projectTask a
                        INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                        INNER JOIN PMIS.Project c ON c.ProjectId = b.ProjectId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = c.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = c.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = c.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                        INNER JOIN BAS.[Type] t2 ON t2.TypeNo = c.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                        LEFT JOIN BAS.Department d ON d.DepartmentId = c.DepartmentId
                        LEFT JOIN BAS.Company e ON e.CompanyId = d.CompanyId
                        OUTER APPLY (
	                        SELECT ROW_NUMBER() OVER (PARTITION BY mb.ProjectId ORDER BY mb.AgentSort) SortPM
                            , ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                            , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
	                        FROM BAS.[User] ma 
	                        INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId 
	                        WHERE mb.ProjectId = c.ProjectId
                        ) m
                        WHERE m.SortPM = 1
                        AND a.ProjectTaskId IN @ProjectTaskIds";

                    dynamicParameters.Add("ProjectTaskIds", result.Select(s => s.ProjectTaskId).ToArray());

                    sql += @" ORDER BY c.ProjectId, a.ProjectTaskId";

                    var projectTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = projectTasks
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

        #region //GetProjectReply -- 取得任務回報內容資料 -- Chia Yuan 2024.5.16
        public string GetProjectReply(int ProjectId, int TaskId, string ReplyType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (!Regex.IsMatch(ReplyType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【任務類型】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //任務回報資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ReplyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TaskId, a.ReplyType, a.ReplyContent, a.ReplyUser
                        , FORMAT(a.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                        , ISNULL(FORMAT(a.EventDate, 'yyyy-MM-dd HH:mm'), '') EventDate
                        , a.ReplyStatus, s3.StatusName ReplyStatusName, a.DeferredDuration
                        , b.UserNo ReplyUserNo
                        , b.UserName ReplyUserName
                        , d.LogoIcon
                        , e.TaskName
                        , e.TaskStatus, s2.StatusName TaskStatusName";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TaskReply a
                        INNER JOIN BAS.[User] b ON a.ReplyUser = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                        INNER JOIN PMIS.ProjectTask e ON e.ProjectTaskId = a.TaskId
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = e.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = a.ReplyStatus AND s3.StatusSchema = 'TaskReply.ReplyStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @" AND a.ReplyType = @ReplyType";
                    dynamicParameters.Add("ReplyType", ReplyType);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters
                           , "ProjectId"
                           , @" AND EXISTS (
                                SELECT TOP 1 1 
                                FROM PMIS.ProjectTask aa 
                                WHERE aa.ProjectTaskId = a.TaskId
                                AND aa.ProjectId = @ProjectId)", ProjectId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters
                           , "ProjectTaskId"
                           , @" AND EXISTS (
                                SELECT TOP 1 1 
                                FROM PMIS.ProjectTask aa 
                                WHERE aa.ProjectTaskId = a.TaskId
                                AND aa.ProjectTaskId = @ProjectTaskId)", TaskId);

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

        #region //GetProjectFile -- 取得專案檔案資料 -- Chia Yuan 2024.5.16
        public string GetProjectFile(int ProjectId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //專案檔案
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProjectFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FileId, a.FileType, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                        , b.[FileName], b.FileExtension, b.FileSize
                        , c.LevelName
                        , d.ProjectId, d.ProjectNo, d.ProjectName
                        , t1.TypeName FileTypeName
                        , 'P' SourceType";
                    sqlQuery.mainTables =
                        @"FROM PMIS.ProjectFile a
                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                        INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                        INNER JOIN PMIS.Project d ON d.ProjectId = a.ProjectId
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = a.FileType AND t1.TypeSchema = 'ProjectFile.FileType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND d.ProjectStatus != 'P'";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectFileId";
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

        #region //GetProjectTaskFile -- 取得專案檔案資料 -- Chia Yuan 2024.5.16
        public string GetProjectTaskFile(int ProjectTaskId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //任務檔案
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProjectTaskFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FileId, a.FileType, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                        , b.[FileName], b.FileExtension, b.FileSize
                        , c.LevelId, c.LevelName
                        , e.ProjectTaskId
                        , d.ProjectId, d.ProjectNo, d.ProjectName
                        , t1.TypeName FileTypeName
                        , 'T' SourceType";
                    sqlQuery.mainTables =
                        @"FROM PMIS.ProjectTaskFile a
                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                        INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                        INNER JOIN PMIS.ProjectTask e ON e.ProjectTaskId = a.ProjectTaskId
                        INNER JOIN PMIS.Project d ON d.ProjectId = e.ProjectId
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = a.FileType AND t1.TypeSchema = 'ProjectFile.FileType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND d.ProjectStatus != 'P'";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectTaskId", @" AND a.ProjectTaskId = @ProjectTaskId", ProjectTaskId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectTaskFileId";
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

        #region //GetTrackFilterDefault -- 取得過濾器設定值 -- Chia Yuan 2024.5.16
        public string GetTrackFilterDefault()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //任務回報資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.UserId, a.FilterName,a.FilterValue
                            FROM PMIS.TrackFilterDefault a
                            WHERE a.UserId = @UserId";
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
        #endregion

        #region //Add

        #endregion

        #region //Update
        #region //UpdateTrackFilterDefault -- 使用者篩選預設值更新 -- Chia Yuan 2024.5.20
        public string UpdateTrackFilterDefault(string FilterJsonData)
        {
            try
            {
                int rowsAffected = 0;
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    JArray jsonData = JArray.Parse(FilterJsonData);

                    #region //取得
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(a.FilterName)) FilterName
                            FROM PMIS.TrackFilterDefault a
                            WHERE a.FilterName IN @FilterNames
                            AND a.UserId = @UserId";
                    dynamicParameters.Add("FilterNames", jsonData.Select(s => s["FilterName"]?.ToString()).ToArray());
                    dynamicParameters.Add("UserId", CurrentUser);
                    var result = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //刪除資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"DELETE a
                            FROM PMIS.TrackFilterDefault a
                            WHERE NOT(a.FilterName IN @FilterNames)
                            AND a.UserId = @UserId";
                    dynamicParameters.Add("FilterNames", jsonData.Select(s => s["FilterName"]?.ToString()).ToArray());
                    dynamicParameters.Add("UserId", CurrentUser);
                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                    #endregion

                    foreach (var item in jsonData)
                    {
                        if (result.FirstOrDefault(f => f.FilterName == item["FilterName"]?.ToString()) == null)
                        {
                            #region //新增資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PMIS.TrackFilterDefault (UserId, FilterName, FilterValue
                                    , LastModifiedDate, LastModifiedBy)
                                    VALUES (@UserId, @FilterName, @FilterValue
                                    , @LastModifiedDate, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId = CurrentUser,
                                    FilterName = item["FilterName"]?.ToString(),
                                    FilterValue = item["FilterValue"]?.ToString(),
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //更新資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    a.FilterValue = @FilterValue,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM PMIS.TrackFilterDefault a
                                    WHERE a.FilterName = @FilterName
                                    AND a.UserId = @UserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId = CurrentUser,
                                    FilterName = item["FilterName"]?.ToString(),
                                    FilterValue = item["FilterValue"]?.ToString(),
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
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

        #endregion

        #region //Api
        #region //SendMessage -- MAMO訊息推播 -- Chia Yuan 2024.02.07
        public string SendMessage(int TaskId, string TaskType, string Content)
        {
            try
            {
                if (!string.IsNullOrEmpty(TaskType) && !Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【任務類型】錯誤!");

                if (Content.Length > 300) throw new SystemException("【訊息】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷專案任務資料是否正確
                    sql = @"SELECT TOP 1 ProjectId
                            FROM PMIS.ProjectTask
                            WHERE ProjectTaskId = @TaskId";
                    var result = sqlConnection.QueryFirstOrDefault(sql, new { TaskId }) ?? throw new SystemException("【專案任務】資料錯誤!");
                    #endregion

                    #region //任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetProjectTaskRecursionString(ref dynamicParameters, new List<dynamic> { result.ProjectId });
                    sql += @"SELECT a.ProjectTaskId
	                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
                            , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
                            , ISNULL(ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration), 1) TaskDuration

                            , a.ProjectTaskId, a.ParentTaskId, a.TaskLevel, a.TaskRoute, a.TaskSort, a.TaskRouteName
                            , b.TaskName, b.TaskDesc
                            , b.TaskStatus, s2.StatusName TaskStatusName

                            , c.ProjectNo, c.ProjectName, c.ProjectDesc, c.CustomerRemark, c.ProjectRemark
                            , c.ProjectStatus, s1.StatusName ProjectStatusName
                            , c.ProjectType, t1.TypeName ProjectTypeName
                            , c.ProjectAttribute, t2.TypeName ProjectAttributeName
                            , c.WorkTimeStatus, c.WorkTimeInterval, s3.StatusName WorkTimeStatusName

                            , (SELECT ub.UserNo
                                , uc.CompanyId, uc.CompanyName, uc.CompanyNo
                                , ud.DepartmentName + '-' + ub.UserName + ':' + ub.Email MailTo
	                            FROM BAS.[User] ub
                                INNER JOIN BAS.Department ud ON ud.DepartmentId = ub.DepartmentId
                                INNER JOIN BAS.Company uc ON uc.CompanyId = ud.CompanyId
                                WHERE EXISTS (
                                    SELECT TOP 1 1 
                                    FROM PMIS.TaskUser ba 
                                    WHERE ba.TaskType = @TaskType
                                    AND ba.TaskId = a.ProjectTaskId
                                    AND ba.UserId = ub.UserId
                                )
                                FOR JSON PATH, ROOT('data')
                            ) MailUsers
                            , (SELECT db.UserId, db.UserNo, db.UserName, db.Gender
	                            FROM PMIS.TaskUser da
	                            INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                            WHERE da.TaskType = @TaskType
	                            AND da.TaskId = a.ProjectTaskId
	                            AND db.[Status] = 'A'
	                            ORDER BY da.AgentSort
	                            FOR JSON PATH, ROOT('data')
                            ) TaskUser
                            , ISNULL (
                                STUFF((
                                    SELECT ',' + bu.UserName
                                    FROM PMIS.TaskUser ba
                                    INNER JOIN BAS.[User] bu ON bu.UserId = ba.UserId
                                    WHERE ba.TaskId = a.ProjectTaskId 
			                        AND ba.TaskType = @TaskType
                                    ORDER BY bu.UserNo
                                    FOR XML PATH('')
                                ), 1, 1, '')
                            , '') TaskUsers
                            , (SELECT fb.UserNo, fb.UserName
                                , FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm') ReplyDate
                                , FORMAT(fa.EventDate, 'yyyy-MM-dd HH:mm') EventDate
                                , fa.ReplyContent, fa.DeferredDuration
                                , fa.ReplyStatus, fc.StatusName ReplyStatusName
                                , ISNULL(fd.ReplyFrequency, 0) ReplyFrequency
	                            FROM PMIS.TaskReply fa
	                            INNER JOIN BAS.[User] fb ON fb.UserId = fa.ReplyUser
	                            INNER JOIN BAS.[Status] fc ON fc.StatusNo = fa.ReplyStatus AND fc.StatusSchema = 'TaskReply.ReplyStatus'
	                            INNER JOIN PMIS.ProjectTask fd ON fd.ProjectTaskId = fa.TaskId
	                            WHERE fa.ReplyType = @TaskType
	                            AND fb.[Status] = 'A'
	                            AND fd.ProjectTaskId = a.ProjectTaskId
	                            ORDER BY fa.ReplyDate
	                            FOR JSON PATH, ROOT('data')
                            ) TaskReply
                            , ISNULL((
                                SELECT COUNT(1)
                                FROM PMIS.ProjectTask ga
                                WHERE ga.ProjectId = c.ProjectId), 0
                            ) SubTaskCount
                            , m.*
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                            INNER JOIN PMIS.Project c ON c.ProjectId = b.ProjectId
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = c.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                            INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                            INNER JOIN BAS.[Status] s3 ON s3.StatusNo = c.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                            INNER JOIN BAS.[Type] t1 ON t1.TypeNo = c.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                            INNER JOIN BAS.[Type] t2 ON t2.TypeNo = c.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                            OUTER APPLY (
	                            SELECT ROW_NUMBER() OVER (PARTITION BY mb.ProjectId ORDER BY mb.AgentSort) SortPM
                                , ma.UserId, ma.UserNo, ma.UserName, ma.Gender
                                , ma.UserName + '(' + ma.UserNo + ')' ProjectMaster
	                            FROM BAS.[User] ma 
	                            INNER JOIN PMIS.ProjectUser mb ON mb.UserId = ma.UserId 
	                            WHERE mb.ProjectId = c.ProjectId
                            ) m
                            WHERE m.SortPM = 1
                            AND a.ProjectTaskId = @TaskId";

                    dynamicParameters.Add("TaskId", TaskId);
                    dynamicParameters.Add("TaskType", TaskType);

                    var Task = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                    #endregion

                    if (Task.MailUsers == null) throw new SystemException("沒有可通知對象");

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
                    var mailTemplate = sqlConnection.Query(sql, dynamicParameters).ToList();
                    if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                    #endregion

                    string tableTemplate = @"<table><thead>
                                                <tr><th>專案名稱</th><th>任務名稱</th><th>參與人員</th><th>任務描述</th><th>起始時間</th><th>結束時間</th><th>任務狀態</th></tr>
                                            </thead>
                                            <tbody>
                                                <tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>
                                            </tbody></table>";
                    string mamoTemplate = tableTemplate;
                    FormatHtmlTagForMamo(ref mamoTemplate);

                    string mamoTableContent = "\r\n\r\n" + string.Format(mamoTemplate, Task?.ProjectName, Task?.TaskName, Task?.TaskUsers, Task?.TaskDesc, Task?.TaskStart, Task?.TaskEnd, Task?.TaskStatusName)
                        + "\r\n\r\n" + string.Format("通知訊息: {0}", Content);
                    JObject mailUsersJson = JObject.Parse(Task.MailUsers);
                    var companyNos = mailUsersJson["data"].Select(s => s["CompanyNo"]?.ToString()).Distinct().ToList();
                    var subUsers = mailUsersJson["data"].Select(s => s["UserNo"]?.ToString()).Distinct();

                    MamoHelper mamoHelper = new MamoHelper();
                    foreach (var item in mailTemplate)
                    {
                        string mailSubject = item.MailSubject;
                        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                        string hyperLink = Regex.Replace(HttpContext.Current.Request.Url.AbsoluteUri, "api/.*|ProjectTrack/.*", "Task/TaskManagement");
                        string pageUrl = hyperLink;
                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("查看可進行的任務"));

                        mailSubject = mailSubject.Replace("[ProjectName]", Task?.ProjectName)
                                                 .Replace("[TaskName]", Task?.TaskName);
                        mailContent = mailContent.Replace("[ProjectName]", Task?.ProjectName)
                                                 .Replace("[TaskName]", Task?.TaskName)
                                                 .Replace("[TaskUser]", Task?.TaskUsers)
                                                 .Replace("[MailContent]", mamoTableContent);

                        #region //MAMO個人訊息推播
                        string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                        mamoContent = mamoContent.Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                        foreach (var companyNo in companyNos)
                        {
                            var userNos = mailUsersJson["data"].Where(w => w["CompanyNo"]?.ToString() == companyNo).Select(s => s["UserNo"]?.ToString()).Distinct().ToList();
                            foreach (var userNo in userNos)
                            {
                                mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                            }
                        }
                        #endregion
                    }
                    
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + subUsers.Count() + " rows affected)"
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

        public void FormatHtmlTagForMamo(ref string mamoTemplate)
        {
            if (string.IsNullOrWhiteSpace(mamoTemplate)) mamoTemplate = "";
            else
            {
                mamoTemplate = Regex.Replace(mamoTemplate, "\r\n", "");
                mamoTemplate = Regex.Replace(mamoTemplate, @"\s+", "");
                mamoTemplate = Regex.Replace(mamoTemplate, "<table.*?><thead.*?>", "");
                mamoTemplate = Regex.Replace(mamoTemplate, "</thead>", "");
                mamoTemplate = Regex.Replace(mamoTemplate, "<tr.*?><th.*?>", "|"); //起始
                mamoTemplate = Regex.Replace(mamoTemplate, "</th><th.*?>", "|");  //中間
                mamoTemplate = Regex.Replace(mamoTemplate, "</th></tr>", "|\n");  // 結束 換行
                mamoTemplate = Regex.Replace(mamoTemplate, "<tr.*?><td.*?>", "|"); //起始
                mamoTemplate = Regex.Replace(mamoTemplate, "</td><td.*?>", "|");  //中間
                mamoTemplate = Regex.Replace(mamoTemplate, "</td></tr>", "|\n");  // 結束 換行
                mamoTemplate = Regex.Replace(mamoTemplate, "<tbody.*?>", "");
                mamoTemplate = Regex.Replace(mamoTemplate, "</tbody></table>", "");

                var rowList = mamoTemplate.Remove(mamoTemplate.Length - 1).Split('\n');
                var signList = rowList.Select((s, x) => s).ToList();
                var splitList = rowList.First().Split('|').Where(w => !string.IsNullOrWhiteSpace(w)).Select(s => "---");

                var splitLine = string.Format("|{0}|", string.Join("|", splitList));
                signList.Insert(1, splitLine);
                mamoTemplate = string.Join("\r\n", signList);
            }
        }
        #endregion
    }
}
