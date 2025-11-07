using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Web;
using System.Linq;
using System.Transactions;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using static Dapper.SqlMapper;

namespace PMISDA
{
    public class ProjectDA
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

        public ProjectDA()
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
        #region //GetSubTaskStrucSQL -- 取得子任務結構資料 -- Chia Yuan 2024.12.3
        /// <summary>
        /// 取得子任務結構資料
        /// </summary>
        /// <returns></returns>
        public string GetSubTaskStrucSQL()
        {
            return @"DECLARE @rowsAdded int
                    DECLARE @projectTask TABLE
                    ( 
                        ProjectTaskId int,
                        ParentTaskId int,
                        TaskLevel int,
                        TaskRoute nvarchar(MAX),
                        TaskSort int,
	                    TaskStatus varchar(2),
                        processed int DEFAULT(0)
                    )

                    INSERT @projectTask
                        SELECT ProjectTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                        , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                        , TaskSort, TaskStatus, 0
                        FROM PMIS.ProjectTask
                        WHERE ProjectTaskId = @ProjectTaskId

                    SET @rowsAdded=@@rowcount

                    WHILE @rowsAdded > 0
                    BEGIN
                        UPDATE @projectTask SET processed = 1 WHERE processed = 0
                        INSERT @projectTask
                            SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                            , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                            , b.TaskSort, b.TaskStatus, 0
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON b.ParentTaskId= a.ProjectTaskId 
                            WHERE a.processed = 1
                        SET @rowsAdded = @@rowcount
                        UPDATE @projectTask SET processed = 2 WHERE processed = 1
                    END

                    SELECT b.TaskStatus, b.TaskName, b.TaskDesc, b.TaskSort
                    , a.TaskLevel, a.ProjectTaskId, ISNULL(b.ParentTaskId, -1) ParentTaskId, ISNULL(b.ChannelId, -1) ChannelId
                    , ISNULL(b.PlannedDuration, 1) PlannedDuration, ISNULL(b.EstimateDuration, 1) EstimateDuration, ISNULL(b.ActualDuration, 1) ActualDuration
                    , ISNULL(b.PlannedStart, '') PlannedStart
                    , ISNULL(b.PlannedEnd, '') PlannedEnd
                    , ISNULL(b.EstimateStart, '') EstimateStart
                    , ISNULL(b.EstimateEnd, '') EstimateEnd
                    , ISNULL(b.ActualStart, '') ActualStart
                    , ISNULL(b.ActualEnd, '') ActualEnd
                    FROM @projectTask a
                    INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                    ORDER BY a.TaskLevel";
        }
        #endregion

        #region //GetParentTaskStrucSQL -- 取得父任務結構資料 -- Chia Yuan 2024.12.3
        /// <summary>
        /// 取得父任務結構資料
        /// </summary>
        /// <returns></returns>
        public string GetParentTaskStrucSQL()
        {
            return @"DECLARE @rowsAdded int
                    DECLARE @projectTask TABLE
                    ( 
                        ProjectTaskId int,
                        ParentTaskId int,
                        TaskLevel int,
                        TaskRoute nvarchar(MAX),
                        TaskSort int,
	                    TaskStatus varchar(2),
                        processed int DEFAULT(0)
                    )

                    INSERT @projectTask
                        SELECT ProjectTaskId, ISNULL(ParentTaskId, -1) ParentTaskId, 1 TaskLevel
                        , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
                        , TaskSort, TaskStatus, 0
                        FROM PMIS.ProjectTask
                        WHERE ProjectTaskId = @ProjectTaskId

                    SET @rowsAdded=@@rowcount

                    WHILE @rowsAdded > 0
                    BEGIN
                        UPDATE @projectTask SET processed = 1 WHERE processed = 0
                        INSERT @projectTask
                            SELECT b.ProjectTaskId, b.ParentTaskId, ( a.TaskLevel + 1 ) TaskLevel
                            , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                            , b.TaskSort, b.TaskStatus, 0
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON a.ParentTaskId= b.ProjectTaskId 
                            WHERE a.processed = 1
                        SET @rowsAdded = @@rowcount
                        UPDATE @projectTask SET processed = 2 WHERE processed = 1
                    END

                    SELECT b.TaskStatus, b.TaskName, b.TaskDesc, b.TaskSort
                    , a.TaskLevel, a.ProjectTaskId, ISNULL(b.ParentTaskId, -1) ParentTaskId, ISNULL(b.ChannelId, -1) ChannelId
                    , ISNULL(b.PlannedDuration, 1) PlannedDuration, ISNULL(b.EstimateDuration, 1) EstimateDuration, ISNULL(b.ActualDuration, 1) ActualDuration
                    , ISNULL(b.PlannedStart, '') PlannedStart
                    , ISNULL(b.PlannedEnd, '') PlannedEnd
                    , ISNULL(b.EstimateStart, '') EstimateStart
                    , ISNULL(b.EstimateEnd, '') EstimateEnd
                    , ISNULL(b.ActualStart, '') ActualStart
                    , ISNULL(b.ActualEnd, '') ActualEnd
                    FROM @projectTask a
                    INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.ProjectTaskId
                    ORDER BY a.TaskLevel";
        }
        #endregion

        #region //GetPrevTaskSQL --取得前置任務 -- Chia Yuan 2024.12.3
        /// <summary>
        /// 取得前置任務
        /// </summary>
        /// <returns></returns>
        public string GetPrevTaskSQL()
        {
            return @"SELECT b.TaskStatus, b.TaskName
                    , b.ProjectTaskId, ISNULL(b.ParentTaskId, -1) ParentTaskId, ISNULL(b.ChannelId, -1) ChannelId
                    , ISNULL(b.PlannedDuration, 1) PlannedDuration, ISNULL(b.EstimateDuration, 1) EstimateDuration, ISNULL(b.ActualDuration, 1) ActualDuration
                    , ISNULL(b.PlannedStart, '') PlannedStart
                    , ISNULL(b.PlannedEnd, '') PlannedEnd
                    , ISNULL(b.EstimateStart, '') EstimateStart
                    , ISNULL(b.EstimateEnd, '') EstimateEnd
                    , ISNULL(b.ActualStart, '') ActualStart
                    , ISNULL(b.ActualEnd, '') ActualEnd
                    FROM PMIS.ProjectTaskLink a
                    LEFT JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.SourceTaskId
                    WHERE a.TargetTaskId = @ProjectTaskId";
        }
        #endregion

        #region //GetProjectNo -- 取得專案代碼 -- Ben Ma 2022.08.18
        public string GetProjectNo(string CompanyType, string ProjectType)
        {
            try
            {
                if (CompanyType.Length <= 0) throw new SystemException("【公司別】不能為空!");
                if (ProjectType.Length <= 0) throw new SystemException("【專案分類】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司別資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM BAS.[Type]
                            WHERE TypeSchema = @TypeSchema
                            AND TypeNo = @TypeNo";
                    dynamicParameters.Add("TypeSchema", "Project.CompanyType");
                    dynamicParameters.Add("TypeNo", CompanyType);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    #region //判斷專案分類資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM BAS.[Type]
                            WHERE TypeSchema = @TypeSchema
                            AND TypeNo = @TypeNo";
                    dynamicParameters.Add("TypeSchema", "Project.ProjectType");
                    dynamicParameters.Add("TypeNo", ProjectType);

                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                    if (result2.Count() <= 0) throw new SystemException("【專案分類】資料錯誤!");
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(ProjectNo), '000'), 3)) + 1 CurrentNum
                            FROM PMIS.Project
                            WHERE ProjectNo LIKE @ProjectNo";
                    dynamicParameters.Add("ProjectNo", string.Format("{0}{1}{2}___", CompanyType, ProjectType, DateTime.Now.ToString("yyMM")));
                    int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                    string projectNo = string.Format("{0}{1}{2}{3}", CompanyType, ProjectType, DateTime.Now.ToString("yyMM"), string.Format("{0:000}", currentNum));

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = projectNo
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

        #region //GetProject -- 取得專案資料 -- Ben Ma 2022.08.19
        public string GetProject(int ProjectId, int DepartmentId, string ProjectNo, string ProjectName
            , string ProjectAttribute, string ProjectStatus, int ProjectTaskId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProjectId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", ISNULL(a.DepartmentId, -1) DepartmentId, ISNULL(a.TeamId, -1) TeamId, a.CompanyType, a.ProjectType, a.ProjectNo
                        , a.ProjectName, a.ProjectDesc, a.ProjectAttribute, a.WorkTimeStatus, a.WorkTimeInterval, a.CustomerRemark
                        , a.ProjectRemark, a.ProjectStatus
                        , ISNULL(b.DepartmentNo, '') DepartmentNo, ISNULL(b.DepartmentName, '') DepartmentName
                        , ISNULL(c.LogoIcon, -1) LogoIcon
                        , d.AuthUpdate, d.AuthDelete, d.AuthTask
                        , e.UserName, e.UserNo, e.Gender
                        , c1.SubTask, c2.SubChannel";
                    sqlQuery.mainTables =
                        @"FROM PMIS.Project a
                        LEFT JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        LEFT JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                        OUTER APPLY (
                            SELECT COUNT(1) SubTask
                            FROM PMIS.ProjectTask aa
	                        INNER JOIN MAMO.Channels ab ON ab.ChannelId = aa.ChannelId
                            WHERE aa.ProjectId = a.ProjectId
	                        AND ab.[Status] = 'A'
                        ) c1
                        OUTER APPLY (
	                        SELECT COUNT(1) SubChannel
	                        FROM PMIS.ProjectTask aa
	                        WHERE aa.ProjectId = a.ProjectId AND NOT(aa.ChannelId IS NULL)
                        ) c2
                        OUTER APPLY (
                            SELECT ISNULL((
                                SELECT TOP 1 1
                                FROM PMIS.ProjectUser aa
                                INNER JOIN PMIS.ProjectUserAuthority ab ON aa.ProjectUserId = ab.ProjectUserId
                                INNER JOIN PMIS.Authority ac ON ab.AuthorityId = ac.AuthorityId
                                WHERE aa.ProjectId = a.ProjectId
                                AND aa.UserId = @UserId
                                AND ac.AuthorityCode = 'project-update'
                            ), 0) AuthUpdate
                            , ISNULL((
                                SELECT TOP 1 1
                                FROM PMIS.ProjectUser aa
                                INNER JOIN PMIS.ProjectUserAuthority ab ON aa.ProjectUserId = ab.ProjectUserId
                                INNER JOIN PMIS.Authority ac ON ab.AuthorityId = ac.AuthorityId
                                WHERE aa.ProjectId = a.ProjectId
                                AND aa.UserId = @UserId
                                AND ac.AuthorityCode = 'project-delete'
                            ), 0) AuthDelete
                            , ISNULL((
                                SELECT TOP 1 1
                                FROM PMIS.ProjectUser aa
                                INNER JOIN PMIS.ProjectUserAuthority ab ON aa.ProjectUserId = ab.ProjectUserId
                                INNER JOIN PMIS.Authority ac ON ab.AuthorityId = ac.AuthorityId
                                WHERE aa.ProjectId = a.ProjectId
                                AND aa.UserId = @UserId
                                AND ac.AuthorityCode = 'task-manage'
                            ), 0) AuthTask
                        ) d
                        OUTER APPLY (
	                        SELECT TOP 1 eu.UserName, eu.UserNo, eu.Gender
	                        FROM PMIS.ProjectUser ea 
	                        INNER JOIN BAS.[User] eu ON eu.UserId = ea.UserId
	                        WHERE ea.ProjectId = a.ProjectId 
	                        AND ea.AgentSort = 1
                        ) e";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("UserId", CurrentUser);
                    if (ProjectTaskId <= 0) {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "IncludeUser",
                            @" AND (
                                EXISTS (
                                    SELECT TOP 1 1
                                    FROM PMIS.ProjectUser aa
                                    WHERE aa.ProjectId = a.ProjectId
                                    AND aa.UserId = @IncludeUser
                                )
                                OR
                                EXISTS (
                                    SELECT TOP 1 1
                                    FROM PMIS.TaskUser aa
                                    INNER JOIN PMIS.ProjectTask ab ON aa.TaskId = ab.ProjectTaskId
                                    WHERE ab.ProjectId = a.ProjectId
                                    AND aa.UserId = @IncludeUser
                                    AND aa.TaskType = 'Project'
                                )
                            )",
                            CurrentUser);
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectNo", @" AND a.ProjectNo LIKE '%' + @ProjectNo + '%'", ProjectNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectName", @" AND a.ProjectName LIKE '%' + @ProjectName + '%'", ProjectName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectAttribute", @" AND a.ProjectAttribute = @ProjectAttribute", ProjectAttribute);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectStatus", @" AND a.ProjectStatus = @ProjectStatus", ProjectStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectTaskId", 
                        @" AND EXISTS (
                            SELECT TOP 1 1 FROM PMIS.ProjectTask aa 
                            WHERE aa.ProjectId = a.ProjectId 
                            AND aa.ProjectTaskId = @ProjectTaskId)", 
                        ProjectTaskId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectId";
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

        #region //GetProjectUser -- 取得專案成員資料 -- Ben Ma 2022.08.22
        public string GetProjectUser(int ProjectId, int CompanyId, int DepartmentId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProjectUserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProjectId, a.UserId, a.LevelId, a.AgentSort
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon
                        , e.LevelName
                        , (
                            SELECT aa.AuthorityId, aa.AuthorityCode, aa.AuthorityName, ab.AuthorityStatus
                            FROM PMIS.Authority aa
                            OUTER APPLY (
                                SELECT ISNULL((
                                    SELECT TOP 1 1
                                    FROM PMIS.ProjectUserAuthority aaa
                                    WHERE aaa.ProjectUserId = a.ProjectUserId
                                    AND aaa.AuthorityId = aa.AuthorityId
                                ), 0) AuthorityStatus
                            ) ab
                            WHERE aa.AuthorityType = @AuthorityType
                            AND aa.[Status] = @AuthorityStatus
                            ORDER BY aa.SortNumber
                            FOR JSON PATH, ROOT('data')
                        ) Authority
                        , case when a.UserId = @CurrentUser and a.AgentSort = 1 then 'Admin' else 'General' end as UserRole";
                    sqlQuery.mainTables =
                        @"FROM PMIS.ProjectUser a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                        INNER JOIN PMIS.FileLevel e ON a.LevelId = e.LevelId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.ProjectId = @ProjectId";
                    dynamicParameters.Add("CurrentUser", CurrentUser);
                    dynamicParameters.Add("AuthorityType", "P");
                    dynamicParameters.Add("AuthorityStatus", "A");
                    dynamicParameters.Add("ProjectId", ProjectId);
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

        #region //GetProjectUserAuthority -- 取得專案成員權限資料 -- Ben Ma 2022.08.23
        public string GetProjectUserAuthority(int ProjectId, string UserNo, string AuthorityCode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT result.AuthorityCode, result.Authority
                            FROM (
                                SELECT a.AuthorityCode, b.Authority
                                FROM PMIS.Authority a
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM PMIS.ProjectUserAuthority ba
                                        INNER JOIN PMIS.ProjectUser bb ON ba.ProjectUserId = bb.ProjectUserId
                                        INNER JOIN BAS.[User] bc ON bb.UserId = bc.UserId
                                        WHERE ba.AuthorityId = a.AuthorityId
                                        AND bb.ProjectId = @ProjectId
                                        AND bc.UserNo = @UserNo
                                    ), 0) Authority
                                ) b
                                WHERE 1=1
                                AND a.AuthorityType = @AuthorityType
                                AND a.Status = @Status
                                UNION ALL
                                SELECT 'change' AuthorityCode, CASE a.AgentSort WHEN 1 THEN 1 ELSE 0 END Authority
                                FROM PMIS.ProjectUser a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ProjectId = @ProjectId
                                AND b.UserNo = @UserNo
                            ) result
                            WHERE 1=1";
                    dynamicParameters.Add("ProjectId", ProjectId);
                    dynamicParameters.Add("UserNo", UserNo);
                    dynamicParameters.Add("AuthorityType", "P");
                    dynamicParameters.Add("Status", "A");
                    if (AuthorityCode.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AuthorityCode", @" AND result.AuthorityCode IN @AuthorityCode AND result.Authority > 0", AuthorityCode.Split(','));

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

        #region //GetProjectFile -- 取得專案檔案資料 -- Ben Ma 2022.08.30
        public string GetProjectFile(int ProjectFileId, int ProjectId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProjectFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProjectId, a.FileId, a.LevelId, a.FileType
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                        , b.[FileName], b.FileContent, b.FileExtension, b.FileSize
                        , c.LevelName";
                    sqlQuery.mainTables =
                        @"FROM PMIS.ProjectFile a
                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                        INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM PMIS.ProjectUser aa
                            INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                            WHERE aa.ProjectId = a.ProjectId
                            AND aa.UserId = @UserId
                            AND ab.SortNumber <= c.SortNumber
                        )";
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectFileId", @" AND a.ProjectFileId = @ProjectFileId", ProjectFileId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectFileId";
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

        #region //GetProjectTask -- 取得專案任務資料 -- Ben Ma 2022.09.16
        public string GetProjectTask(int ProjectTaskId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得專案資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.ProjectId
                            FROM PMIS.ProjectTask a
                            WHERE a.ProjectTaskId = @ProjectTaskId";
                    dynamicParameters.Add("ProjectTaskId", ProjectTaskId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    int projectId = -1;
                    foreach (var item in result)
                    {
                        projectId = item.ProjectId;
                    }
                    #endregion

                    #region //取得專案任務資料
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
                                WHERE ProjectId = @ProjectId
                                AND ParentTaskId IS NULL

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
                            END;

                            SELECT a.ProjectTaskId, a.TaskRouteName ParentTaskName, a.TaskSort, a.TaskName
                            , b.ProjectId, b.TaskDesc, b.TaskStatus, b.RequiredFile
                            , e.ProjectStatus, c.SubTask
                            , ISNULL(FORMAT(b.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(b.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(b.PlannedDuration, 0) PlannedDuration
                            , ISNULL(FORMAT(b.EstimateStart, 'yyyy-MM-dd HH:mm'), '') EstimateStart
                            , ISNULL(FORMAT(b.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                            , ISNULL(b.EstimateDuration, 0) EstimateDuration
                            , ISNULL(FORMAT(b.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                            , ISNULL(FORMAT(b.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                            , ISNULL(b.ActualDuration, 0) ActualDuration
                            , ISNULL(b.ParentTaskId, -1) ParentTaskId, ISNULL(b.ReplyFrequency, 0) ReplyFrequency
                            , FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm') LastEnd
                            , ISNULL(SUBSTRING(d.TaskUser, 0, LEN(d.TaskUser)), '') TaskUser
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                            INNER JOIN PMIS.Project e ON e.ProjectId = b.ProjectId
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ParentTaskId = a.ProjectTaskId
                            ) c
                            OUTER APPLY (
                                SELECT (
                                    SELECT db.UserName + ','
                                    FROM PMIS.TaskUser da
                                    INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                    WHERE da.TaskId = b.ProjectTaskId
                                    AND da.TaskType = 'Project'
                                    ORDER BY da.AgentSort
                                    FOR XML PATH('')
                                ) TaskUser
                            ) d
                            WHERE 1=1";
                    dynamicParameters.Add("ProjectId", projectId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", @" AND b.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", @" AND b.ProjectId = @ProjectId", projectId);

                    result = sqlConnection.Query(sql, dynamicParameters);
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

        #region //GetProjectTaskTree -- 取得專案任務資料(樹狀結構) -- Ben Ma 2022.08.23
        public string GetProjectTaskTree(int ProjectId, int ParentTaskId, int OpenedLevel)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string projectName = string.Empty, projectStatus = string.Empty;

                    #region //專案資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ProjectName,ProjectId,ProjectStatus
                            FROM PMIS.Project
                            WHERE ProjectId = @ProjectId";
                    dynamicParameters.Add("ProjectId", ProjectId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");

                    foreach (var item in result)
                    {
                        projectName = item.ProjectName;
                        projectStatus = item.ProjectStatus;
                    }
                    #endregion

                    #region //任務資料
                    List<ProjectTask> projectTasks = new List<ProjectTask>();

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
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                            ORDER BY a.TaskLevel, a.TaskRoute, a.TaskSort";
                    dynamicParameters.Add("ProjectId", ProjectId);

                    projectTasks = sqlConnection.Query<ProjectTask>(sql, dynamicParameters).ToList();
                    #endregion

                    var data = new DHXTree
                    {
                        status = projectStatus,
                        value = projectName,
                        id = "-1",
                        opened = true
                    };

                    if (projectTasks.Count > 0)
                    {
                        data.items = projectTasks
                            .Where(x => x.TaskLevel == 1 && x.ParentTaskId == Convert.ToInt32(data.id))
                            .OrderBy(x => x.TaskLevel)
                            .ThenBy(x => x.TaskRoute)
                            .ThenBy(x => x.TaskSort)
                            .Select(x => new DHXTree
                            {
                                status = projectStatus,
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
                                    .Where(x => x.TaskLevel == (taskTree[i].level + 1) && x.ParentTaskId == Convert.ToInt32(taskTree[i].id))
                                    .OrderBy(x => x.TaskLevel)
                                    .ThenBy(x => x.TaskRoute)
                                    .ThenBy(x => x.TaskSort)
                                    .Select(x => new DHXTree
                                    {
                                        status = projectStatus,
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

        #region //GetProjectTaskGantt -- 取得專案任務資料(甘特圖) -- Ben Ma 2022.09.26
        public string GetProjectTaskGantt(int ProjectId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //專案資料
                    sql = @"SELECT ProjectId, ProjectName, ProjectStatus
                            , ISNULL(FORMAT(PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(PlannedDuration, 1) PlannedDuration
                            , FORMAT(ISNULL(ActualStart, EstimateStart), 'yyyy-MM-dd HH:mm') StartDate
                            , FORMAT(ISNULL(ActualEnd, EstimateEnd), 'yyyy-MM-dd HH:mm') EndDate
                            , ISNULL(ISNULL(ActualDuration, EstimateDuration), 1) Duration
                            FROM PMIS.Project
                            WHERE ProjectId = @ProjectId";
                    var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");

                    string projectName =result.ProjectName,
                            projectStatus = result.ProjectStatus, 
                            projectPlannedStart = result.PlannedStart, 
                            projectPlannedEnd = result.PlannedEnd,
                            projectStartDate = result.StartDate,
                            projectEndDate = result.EndDate;
                        //, projectEstimateStart = "", projectEstimateEnd = "";
                        //, projectActualStart = "", projectActualEnd = "";

                    int projectPlannedDuration = result.PlannedDuration, projectDuration = result.Duration; //, projectEstimateDuration = result.EstimateDuration, projectActualDuration = esult.ActualDuration;
                    //projectEstimateStart = item.EstimateStart;
                    //projectEstimateEnd = item.EstimateEnd;
                    //projectActualStart = item.ActualStart;
                    //projectActualEnd = item.ActualEnd;
                    #endregion

                    #region //專案任務資料
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
                            , b.PlannedStart, b.PlannedEnd, b.PlannedDuration
                            , ISNULL(b.ActualStart, b.EstimateStart) StartDate
                            , ISNULL(b.ActualEnd, b.EstimateEnd) EndDate
                            , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), 1) Duration
                            , c.SubTask
                            , d.DeferredStatus
                            , b.TaskStatus
                            , e.StatusName TaskStatusName
                            , f.SubUser
                            FROM @projectTask a
                            INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ParentTaskId = a.ProjectTaskId
                            ) c
                            OUTER APPLY (
	                            SELECT COUNT(1) SubUser
	                            FROM PMIS.TaskUser fa
	                            WHERE fa.TaskId = a.ProjectTaskId
                                AND fa.TaskType = 'Project'
                            ) f
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
                    List<ProjectTask> projectTasks = sqlConnection.Query<ProjectTask>(sql, new { ProjectId }).ToList();
                    #endregion

                    #region //專案任務連結資料
                    sql = @"SELECT a.ProjectTaskLinkId, a.SourceTaskId, a.TargetTaskId, a.LinkType
                            FROM PMIS.ProjectTaskLink a
                            INNER JOIN PMIS.ProjectTask b ON a.SourceTaskId = b.ProjectTaskId
                            WHERE b.ProjectId = @ProjectId
                            ORDER BY a.ProjectTaskLinkId";
                    List<ProjectTaskLink> projectTaskLinks = sqlConnection.Query<ProjectTaskLink>(sql, new { ProjectId }).ToList();
                    #endregion

                    #region //組成任務json
                    List<DHXTask> dHXTasks = new List<DHXTask>();

                    switch (projectStatus)
                    {
                        case "P":
                            #region //專案規劃
                            dHXTasks.Add(
                                new DHXTask
                                {
                                    id = "1.1",
                                    text = projectName,
                                    start_date = projectPlannedStart,
                                    end_date = projectPlannedEnd,
                                    progress = "0",
                                    parent = "0",
                                    type = "project",
                                    order = 1
                                });

                            foreach (var task in projectTasks)
                            {
                                DateTime plannedStart = task.PlannedStart != null ? task.PlannedStart.Value : DateTime.Now;
                                int plannedDuration = task.PlannedDuration != null ? task.PlannedDuration.Value : 1;
                                DateTime plannedEnd = task.PlannedEnd != null ? task.PlannedEnd.Value : plannedStart.AddMinutes(plannedDuration);

                                dHXTasks.Add(
                                    new DHXTask
                                    {
                                        id = task.ProjectTaskId.ToString(),
                                        text = task.TaskName,
                                        start_date = plannedStart.ToString("yyyy-MM-dd HH:mm"),
                                        duration = plannedDuration.ToString(),
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
                            break;
                        default:
                            #region //專案進行中&專案完成
                            dHXTasks.Add(
                                new DHXTask
                                {
                                    id = "1.1",
                                    text = projectName,
                                    start_date = projectStartDate,
                                    duration = projectDuration.ToString(),
                                    progress = "0",
                                    parent = "0",
                                    type = "project",
                                    order = 1,
                                    planned_start = projectPlannedStart,
                                    planned_end = projectPlannedEnd,
                                    planned_duration = projectPlannedDuration.ToString()
                                });

                            foreach (var task in projectTasks)
                            {
                                DateTime startDate = task.StartDate.Value;
                                int duration = task.Duration != null ? task.Duration.Value : 1;
                                DateTime endDate = task.EndDate != null ? task.EndDate.Value: startDate.AddMinutes(duration);

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
                                        planned_start = Convert.ToDateTime(task.PlannedStart).ToString("yyyy-MM-dd HH:mm"),
                                        planned_end = Convert.ToDateTime(task.PlannedEnd).ToString("yyyy-MM-dd HH:mm"),
                                        planned_duration = task.PlannedDuration.ToString(),
                                        deferred_status = task.DeferredStatus > 0 ? "Y" : "N",
                                        sub_user = task.SubUser,
                                        task_status = task.TaskStatus,
                                        task_status_name = task.TaskStatusName
                                    });
                            }
                            #endregion
                            break;
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

        #region //GetTaskUser -- 取得任務成員資料-New (需傳入任務類型) -- Chia Yuan 2023.12.06
        public string GetTaskUser(int TaskId, int CompanyId, int DepartmentId, string SearchKey, string TaskType
            , string OrderBy, int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TaskUserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TaskId, a.UserId, a.TaskType, a.LevelId, a.AgentSort
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentNo, c.DepartmentName
                        , d.LogoIcon
                        , e.LevelName
                        , (
                            SELECT aa.AuthorityId, aa.AuthorityName, ab.AuthorityStatus
                            FROM PMIS.Authority aa
                            OUTER APPLY (
                                SELECT ISNULL((
                                    SELECT TOP 1 1
                                    FROM PMIS.TaskUserAuthority aaa
                                    WHERE aaa.TaskUserId = a.TaskUserId
                                    AND aaa.AuthorityId = aa.AuthorityId
                                ), 0) AuthorityStatus
                            ) ab
                            WHERE aa.AuthorityType = 'T'
                            AND aa.[Status] = 'A'
                            ORDER BY aa.SortNumber
                            FOR JSON PATH, ROOT('data')
                        ) Authority
                        , case when a.UserId = @CurrentUser and a.AgentSort = 1 then 'Admin' else 'General' end as UserRole";
                    sqlQuery.mainTables =
                        @"FROM PMIS.TaskUser a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                        INNER JOIN PMIS.FileLevel e ON a.LevelId = e.LevelId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.TaskId = @TaskId AND a.TaskType = @TaskType";
                    dynamicParameters.Add("CurrentUser", CurrentUser);
                    dynamicParameters.Add("TaskId", TaskId);
                    dynamicParameters.Add("TaskType", TaskType);
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

        #region //GetTaskUserAuthority -- 取得任務成員權限資料-New  (需傳入任務類型) -- Chia Yuan 2023-12.08
        public string GetTaskUserAuthority(int TaskId, string TaskType, string UserNo, string AuthorityCode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT result.AuthorityCode, result.Authority
                            FROM (
                                SELECT a.AuthorityCode, b.Authority
                                FROM PMIS.Authority a
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM PMIS.TaskUserAuthority ba
                                        INNER JOIN PMIS.TaskUser bb ON ba.TaskUserId = bb.TaskUserId
                                        INNER JOIN BAS.[User] bc ON bb.UserId = bc.UserId
                                        WHERE ba.AuthorityId = a.AuthorityId
                                        AND bb.TaskId = @TaskId
                                        AND bb.TaskType = @TaskType
                                        AND bc.UserNo = @UserNo
                                    ), 0) Authority
                                ) b
                                WHERE 1=1
                                AND a.AuthorityType = 'T'
                                AND a.Status = 'A'
                                UNION ALL
                                SELECT 'change' AuthorityCode, CASE a.AgentSort WHEN 1 THEN 1 ELSE 0 END Authority
                                FROM PMIS.TaskUser a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.TaskId = @TaskId
                                AND a.TaskType = @TaskType
                                AND b.UserNo = @UserNo
                            ) result
                            WHERE 1=1";
                    if (AuthorityCode.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AuthorityCode", @" AND result.AuthorityCode IN @AuthorityCode AND result.Authority > 0", AuthorityCode.Split(','));

                    dynamicParameters.Add("TaskId", TaskId);
                    dynamicParameters.Add("TaskType", TaskType);
                    dynamicParameters.Add("UserNo", UserNo);

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

        #region //GetProjectTaskFile -- 取得專案任務檔案 -- Ben Ma 2022.09.23
        public string GetProjectTaskFile(int ProjectTaskFileId, int ProjectTaskId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProjectTaskFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProjectTaskId, a.FileId, a.LevelId, a.FileType
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                        , b.[FileName], b.FileContent, b.FileExtension, b.FileSize
                        , c.LevelName";
                    sqlQuery.mainTables =
                        @"FROM PMIS.ProjectTaskFile a
                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                        INNER JOIN PMIS.FileLevel c ON a.LevelId = c.LevelId
                        INNER JOIN PMIS.ProjectTask d ON a.ProjectTaskId = d.ProjectTaskId";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @" AND a.ProjectTaskId = @TaskId
                        AND (
                            EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.TaskUser aa
                                INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                                WHERE aa.TaskId = a.ProjectTaskId
                                AND aa.TaskType = 'Project'
                                AND aa.UserId = @UserId
                                AND ab.SortNumber <= c.SortNumber
                            ) OR
                            EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectUser aa
                                INNER JOIN PMIS.FileLevel ab ON aa.LevelId = ab.LevelId
                                WHERE aa.ProjectId = d.ProjectId
                                AND aa.UserId = @UserId
                                AND ab.SortNumber <= c.SortNumber
                            )
                        )";
                    dynamicParameters.Add("TaskId", ProjectTaskId);
                    dynamicParameters.Add("UserId", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectTaskFileId", @" AND a.ProjectTaskFileId = @ProjectTaskFileId", ProjectTaskFileId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectTaskFileId";
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
        #region //AddProject -- 專案資料新增 -- Ben Ma 2022.08.19
        public string AddProject(int DepartmentId, string CompanyType, string ProjectType, string ProjectNo
            , string ProjectName, string ProjectDesc, string ProjectAttribute, string WorkTimeStatus
            , string WorkTimeInterval, string CustomerRemark, string ProjectRemark)
        {
            try
            {
                if (CompanyType.Length <= 0) throw new SystemException("【公司別】不能為空!");
                if (ProjectType.Length <= 0) throw new SystemException("【專案分類】不能為空!");
                if (ProjectNo.Length <= 0) throw new SystemException("【專案代碼】不能為空!");
                if (ProjectNo.Length != 10) throw new SystemException("【專案代碼】長度錯誤!");
                if (ProjectName.Length <= 0) throw new SystemException("【專案名稱】不能為空!");
                if (ProjectName.Length > 100) throw new SystemException("【專案名稱】長度錯誤!");
                if (ProjectDesc.Length <= 0) throw new SystemException("【專案描述】不能為空!");
                if (ProjectDesc.Length > 300) throw new SystemException("【專案描述】長度錯誤!");
                if (ProjectAttribute.Length <= 0) throw new SystemException("【專案屬性】不能為空!");
                if (WorkTimeStatus.Length <= 0) throw new SystemException("【工作天模式】不能為空!");
                if (WorkTimeStatus == "Y") if (WorkTimeInterval.Length <= 0) throw new SystemException("【工作時段】不能為空!");
                if (WorkTimeInterval.Length > 100) throw new SystemException("【工作時段】長度錯誤!");
                if (CustomerRemark.Length > 500) throw new SystemException("【客戶備註】長度錯誤!");
                if (ProjectRemark.Length > 500) throw new SystemException("【專案備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確
                        if (DepartmentId > 0)
                        {
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE DepartmentId = @DepartmentId";
                            var result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("部門資料錯誤!");
                        }
                        #endregion

                        #region //判斷公司別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        var result2 = sqlConnection.QueryFirstOrDefault(sql, new { TypeSchema = "Project.CompanyType", TypeNo = CompanyType }) ?? throw new SystemException("【公司別】資料錯誤!");
                        #endregion

                        #region //判斷專案分類資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        var result3 = sqlConnection.QueryFirstOrDefault(sql, new { TypeSchema = "Project.ProjectType", TypeNo = ProjectType }) ?? throw new SystemException("【專案分類】資料錯誤!");
                        #endregion

                        #region //專案新增
                        sql = @"INSERT INTO PMIS.Project (DepartmentId, CompanyType, ProjectType, ProjectNo
                                , ProjectName, ProjectDesc, ProjectAttribute, WorkTimeStatus
                                , WorkTimeInterval, CustomerRemark, ProjectRemark, ProjectStatus
                                , PlannedStart
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectId
                                VALUES (@DepartmentId, @CompanyType, @ProjectType, @ProjectNo
                                , @ProjectName, @ProjectDesc, @ProjectAttribute, @WorkTimeStatus
                                , @WorkTimeInterval, @CustomerRemark, @ProjectRemark, 'P'
                                , @PlannedStart
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                DepartmentId = DepartmentId <= 0 ? (int?)null : DepartmentId,
                                CompanyType,
                                ProjectType,
                                ProjectNo,
                                ProjectName,
                                ProjectDesc,
                                ProjectAttribute,
                                WorkTimeStatus,
                                WorkTimeInterval,
                                CustomerRemark,
                                ProjectRemark,
                                PlannedStart = CreateDate.ToString("yyyy-MM-dd"),
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected++;
                        int ProjectId = Convert.ToInt32(insertResult?.ProjectId ?? -1);
                        #endregion

                        #region //取得目前最高檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                WHERE Status = 'A'
                                ORDER BY SortNumber";
                        var result4 = sqlConnection.QueryFirstOrDefault(sql) ?? throw new SystemException("檔案等級資料不存在，請確認是否有新增!");
                        int maxLevel = Convert.ToInt32(result4?.LevelId ?? -1);
                        #endregion

                        #region //專案成員新增
                        sql = @"INSERT INTO PMIS.ProjectUser (ProjectId, UserId, LevelId, AgentSort
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectUserId
                                VALUES (@ProjectId, @UserId, @LevelId, @AgentSort
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var resultUsers = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                ProjectId,
                                UserId = CurrentUser,
                                LevelId = maxLevel,
                                AgentSort = 1,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected++;
                        int ProjectUserId = Convert.ToInt32(resultUsers?.ProjectUserId ?? -1);
                        #endregion

                        #region //抓取所有的專案權限
                        sql = @"SELECT AuthorityId
                                FROM PMIS.Authority
                                WHERE AuthorityType = 'P'
                                AND Status = 'A'";
                        var result6 = sqlConnection.Query(sql);
                        if (!result6.Any()) throw new SystemException("專案權限資料不存在，請確認是否有新增!");
                        #endregion

                        #region //專案成員權限新增
                        foreach (var item in result6)
                        {
                            sql = @"INSERT INTO PMIS.ProjectUserAuthority (ProjectUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@ProjectUserId, @AuthorityId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProjectUserId,
                                    AuthorityId = Convert.ToInt32(item.AuthorityId),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        #endregion

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

        #region //AddProjectUser -- 專案成員資料新增 -- Ben Ma 2022.08.22
        public string AddProjectUser(int ProjectId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【專案成員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int teamId = -1;
                        string companyNo = string.Empty;

                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ISNULL(a.TeamId, -1) TeamId, c.CompanyNo
                                FROM PMIS.Project a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE a.ProjectId = @ProjectId";
                        var result = sqlConnection.Query(sql, new { ProjectId });
                        if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");
                        foreach (var item in result)
                        {
                            teamId = item.TeamId;
                            companyNo = item.CompanyNo;
                        }
                        #endregion

                        string[] usersList = Users.Split(',').Distinct().ToArray();

                        #region //判斷專案成員資料是否正確
                        sql = @"SELECT UserId, UserNo
                                FROM BAS.[User]
                                WHERE UserId IN @UserIds";
                        var resultUsers = sqlConnection.Query(sql, new { UserIds  = usersList });
                        if (resultUsers.Count() != usersList.Length) throw new SystemException("【專案成員】資料錯誤!");
                        #endregion

                        #region //抓取目前最小檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                ORDER BY SortNumber DESC";
                        int LevelId = sqlConnection.QueryFirstOrDefault(sql).LevelId;
                        #endregion

                        #region //抓取目前最大代理順序
                        sql = @"SELECT ISNULL(MAX(AgentSort), 0) MaxAgentSort
                                FROM PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId";
                        int maxAgentSort = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }).MaxAgentSort;
                        #endregion

                        #region //專案成員新增
                        foreach (var item in resultUsers)
                        {
                            maxAgentSort++;
                            sql = @"INSERT INTO PMIS.ProjectUser (ProjectId, UserId, LevelId, AgentSort
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectUserId
                                    VALUES (@ProjectId, @UserId, @LevelId, @AgentSort
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProjectId,
                                    UserId = Convert.ToInt32(item.UserId),
                                    LevelId,
                                    AgentSort = maxAgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        #endregion

                        #region //MAMO團隊成員新增 (停用)
                        //if (!string.IsNullOrWhiteSpace(companyNo) && teamId > 0)
                        //{
                        //    var mamoResult = mamoHelper.AddTeamMembers(companyNo, CurrentUser, teamId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                        //    JObject mamoResultJson = JObject.Parse(mamoResult);
                        //}
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

        #region //AddProjectFile -- 專案檔案資料新增 -- Ben Ma 2022.08.30
        public string AddProjectFile(int ProjectId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        #endregion

                        #region //判斷檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileId }) ?? throw new SystemException("檔案資料錯誤!");
                        #endregion

                        #region //判斷是否是專案成員
                        sql = @"SELECT LevelId
                                FROM PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId
                                AND UserId = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, CurrentUser }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        int LevelId = result?.LevelId ?? -1;
                        #endregion

                        #region //新增專案檔案
                        sql = @"INSERT INTO PMIS.ProjectFile (ProjectId, FileId, LevelId, FileType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectFileId
                                VALUES (@ProjectId, @FileId, @LevelId, @FileType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                FileId,
                                LevelId,
                                FileType = "99",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
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

        #region //AddProjectTask -- 專案任務資料新增 -- Ben Ma 2022.09.15
        public string AddProjectTask(int ProjectId, int ParentTaskId, string TaskName, string TaskDesc
            , string PlannedStart, int PlannedDuration, int ReplyFrequency, bool RequiredFile)
        {
            try
            {
                if (TaskName.Length <= 0) throw new SystemException("【任務名稱】不能為空!");
                if (TaskName.Length > 50) throw new SystemException("【任務名稱】長度錯誤!");
                if (TaskDesc.Length > 300) throw new SystemException("【任務描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    var rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int teamId = -1;
                        string companyNo = string.Empty;
                        string projectStatus = string.Empty;

                        PlannedDuration = PlannedDuration <= 0 ? 1440 : PlannedDuration; //預設為1天
                        DateTime plannedStart = PlannedStart.Length <= 0 ? CreateDate : Convert.ToDateTime(PlannedStart);

                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ISNULL(a.TeamId, -1) TeamId, a.ProjectStatus, c.CompanyNo
                                , FORMAT(a.PlannedStart,'yyyy-MM-dd HH:mm') PlannedStart
                                FROM PMIS.Project a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE a.ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (projectStatus == "P") DateTime.TryParse(result.PlannedStart, out plannedStart);
                        teamId = result.TeamId;
                        companyNo = result.CompanyNo;
                        projectStatus = result.ProjectStatus;
                        #endregion

                        #region //判斷父層專案任務資料是否正確
                        if (ParentTaskId > 0)
                        {
                            sql = @"SELECT TOP 1 FORMAT(ISNULL(ActualStart, ISNULL(EstimateStart, PlannedStart)),'yyyy-MM-dd HH:mm') StartDate
                                    , FORMAT(ISNULL(ActualEnd, ISNULL(EstimateEnd, PlannedEnd)),'yyyy-MM-dd HH:mm') EndDate
                                    , FORMAT(PlannedStart,'yyyy-MM-dd HH:mm') PlannedStart
                                    , FORMAT(PlannedEnd,'yyyy-MM-dd HH:mm') PlannedEnd
                                    , FORMAT(EstimateStart,'yyyy-MM-dd HH:mm') EstimateStart
                                    , FORMAT(EstimateEnd,'yyyy-MM-dd HH:mm') EstimateEnd
                                    , FORMAT(ActualStart,'yyyy-MM-dd HH:mm') ActualStart
                                    , FORMAT(ActualEnd,'yyyy-MM-dd HH:mm') ActualEnd
                                    , TaskStatus
                                    FROM PMIS.ProjectTask
                                    WHERE ProjectTaskId = @ParentTaskId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { ParentTaskId }) ?? throw new SystemException("【上層專案任務】資料錯誤!");
                            DateTime.TryParse(result.StartDate, out plannedStart);
                        }
                        #endregion

                        #region //抓取目前最大序號
                        sql = @"SELECT ISNULL(MAX(TaskSort), 0) MaxSort
                                FROM PMIS.ProjectTask
                                WHERE ProjectId = @ProjectId
                                AND ISNULL(ParentTaskId, -1) = @ParentTaskId";
                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, ParentTaskId })?.MaxSort ?? 0;
                        #endregion

                        int projectTaskId = -1;
                        dynamic insertResult;

                        #region //專案任務新增
                        if (projectStatus == "P")
                        {
                            sql = @"INSERT INTO PMIS.ProjectTask (ProjectId, ParentTaskId, TaskSort, TaskName
                                    , TaskDesc, ReplyFrequency, PlannedStart, PlannedEnd, PlannedDuration
                                    , RequiredFile, TaskStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectTaskId
                                    VALUES (@ProjectId, @ParentTaskId, @TaskSort, @TaskName
                                    , @TaskDesc, @ReplyFrequency, @PlannedStart, @PlannedEnd, @PlannedDuration
                                    , @RequiredFile, 'B'
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    ProjectId,
                                    ParentTaskId = ParentTaskId > 0 ? ParentTaskId : (int?)null,
                                    TaskSort = maxSort + 1,
                                    TaskName,
                                    TaskDesc,
                                    ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                    PlannedStart = plannedStart,
                                    PlannedEnd = plannedStart.AddMinutes(PlannedDuration),
                                    PlannedDuration,
                                    RequiredFile,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected++;
                            projectTaskId = insertResult.ProjectTaskId;
                        }
                        else
                        {
                            sql = @"INSERT INTO PMIS.ProjectTask (ProjectId, ParentTaskId, TaskSort, TaskName
                                    , TaskDesc, ReplyFrequency, PlannedStart, PlannedEnd, PlannedDuration
                                    , EstimateStart, EstimateEnd, EstimateDuration
                                    , RequiredFile, TaskStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectTaskId
                                    VALUES (@ProjectId, @ParentTaskId, @TaskSort, @TaskName
                                    , @TaskDesc, @ReplyFrequency, @PlannedStart, @PlannedEnd, @PlannedDuration
                                    , @EstimateStart, @EstimateEnd, @EstimateDuration
                                    , @RequiredFile, 'B'
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    ProjectId,
                                    ParentTaskId = ParentTaskId > 0 ? ParentTaskId : (int?)null,
                                    TaskSort = maxSort + 1,
                                    TaskName,
                                    TaskDesc,
                                    ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                    PlannedStart = plannedStart,
                                    PlannedEnd = plannedStart.AddMinutes(PlannedDuration),
                                    PlannedDuration,
                                    EstimateStart = plannedStart,
                                    EstimateEnd = plannedStart.AddMinutes(PlannedDuration),
                                    EstimateDuration = PlannedDuration,
                                    RequiredFile,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected++;
                            projectTaskId = insertResult.ProjectTaskId;
                        }
                        #endregion

                        #region //變更上層狀態
                        if (projectStatus == "I")
                        {
                            #region //取得父任務結構
                            sql = GetParentTaskStrucSQL();
                            var resultParentTasks = sqlConnection.Query(sql, new { ProjectTaskId = ParentTaskId });
                            var projectTaskIds = resultParentTasks.Select(s => s.ProjectTaskId).ToArray();
                            #endregion

                            #region //更新主任務狀態
                            sql = @"UPDATE a SET 
                                    a.ActualEnd = NULL,
                                    a.ActualDuration = NULL,
                                    a.TaskStatus = CASE a.TaskStatus WHEN 'C' THEN 'I' ELSE a.TaskStatus END
                                    FROM PMIS.ProjectTask a
                                    WHERE a.ProjectTaskId IN @ProjectTaskIds";
                            rowsAffected += sqlConnection.Execute(sql, new { ProjectTaskIds = projectTaskIds });
                            #endregion

                            #region //MAMO (停用)
                            //if (!string.IsNullOrWhiteSpace(companyNo) && teamId > 0)
                            //{
                            //    #region //MAMO頻道新增
                            //    int channelId = -1;
                            //    var mamoResult = mamoHelper.CreateChannels(companyNo, CurrentUser, teamId, TaskName, "描述：" + TaskDesc);
                            //    JObject mamoResultJson = JObject.Parse(mamoResult);

                            //    if (mamoResultJson["status"].ToString() == "success")
                            //    {
                            //        foreach (var item in mamoResultJson["result"])
                            //        {
                            //            channelId = Convert.ToInt32(item["ChannelId"].ToString());
                            //        }
                            //    }
                            //    #endregion

                            //    #region //TeamId更新
                            //    sql = @"UPDATE a SET a.ChannelId = @ChannelId
                            //FROM PMIS.ProjectTask a
                            //WHERE a.ProjectTaskId = @ProjectTaskId";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            ProjectTaskId = projectTaskId,
                            //            ChannelId = channelId
                            //        });
                            //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //    #endregion
                            //}
                            #endregion
                        }
                        #endregion

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

        #region //AddProjectTaskByCopy -- 新增複製任務 -- Chia Yuan 2023.11.09
        public string AddProjectTaskByCopy(int ProjectTaskId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string projectStatus = string.Empty, companyNo = string.Empty;
                        int ParentTaskId = -1, ProjectId = -1;
                        //int teamId = -1, parentTaskId = -1;
                        //string companyNo = string.Empty;
                        //string projectStatus = string.Empty;
                        //string parentTaskStatus = string.Empty;

                        #region //判斷狀態資料是否正確
                        sql = @"SELECT TOP 1 a.ProjectId, a.ProjectName, a.ProjectStatus
                                , c.CompanyNo, ISNULL(t.ParentTaskId, -1) ParentTaskId, t.TaskStatus, t.TaskName
                                FROM PMIS.Project a
                                INNER JOIN PMIS.ProjectTask t ON t.ProjectId = a.ProjectId
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE t.ProjectTaskId = @ProjectTaskId";
                        var result = sqlConnection.Query(sql, new { ProjectTaskId });
                        if (result.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.ProjectStatus == "C") throw new SystemException(string.Format("【{0}】專案已完成，不得操作!", item.ProjectName));
                            ProjectId = item.ProjectId;
                            projectStatus = item.ProjectStatus;
                            companyNo = item.CompanyNo;
                            ParentTaskId = item.ParentTaskId;
                        }
                        #endregion

                        #region //抓取目前最大序號
                        sql = @"SELECT ISNULL(MAX(TaskSort), 0) MaxSort
                                FROM PMIS.ProjectTask
                                WHERE ProjectId = @ProjectId
                                AND ISNULL(ParentTaskId, -1) = @ParentTaskId";
                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, ParentTaskId })?.MaxSort ?? 0;
                        #endregion

                        #region //取得子任務結構
                        sql = GetSubTaskStrucSQL();
                        var resultSubTasks = sqlConnection.Query<ProjectTask>(sql, new { ProjectTaskId }).ToList();
                        var ProjectTaskIds = resultSubTasks.Select(s => s.ProjectTaskId).ToArray();
                        #endregion

                        #region //取得來源任務連結資料
                        sql = @"SELECT a.ProjectTaskLinkId, a.SourceTaskId, a.TargetTaskId, a.LinkType
                                FROM PMIS.ProjectTaskLink a
                                WHERE a.SourceTaskId IN @ProjectTaskIds 
                                    OR a.TargetTaskId IN @ProjectTaskIds
                                ORDER BY a.ProjectTaskLinkId";
                        List<ProjectTaskLink> resultTaskLink = sqlConnection.Query<ProjectTaskLink>(sql, 
                            new 
                            {
                                ProjectTaskIds = resultSubTasks.Select(s => s.ProjectTaskId).Distinct().ToArray()
                            }).ToList();
                        #endregion

                        #region //取得來源任務成員資料
                        sql = @"SELECT a.TaskUserId, a.TaskId, a.UserId, u.UserNo, a.LevelId, a.AgentSort
                                , (
                                    SELECT aa.AuthorityId, aa.AuthorityName, ab.AuthorityStatus
                                    FROM PMIS.Authority aa
                                    OUTER APPLY (
                                        SELECT ISNULL((
                                            SELECT TOP 1 1
                                            FROM PMIS.TaskUserAuthority ba
                                            WHERE ba.TaskUserId = a.TaskUserId
                                            AND ba.AuthorityId = aa.AuthorityId
                                        ), 0) AuthorityStatus
                                    ) ab
                                    WHERE aa.AuthorityType = 'T'
                                    AND aa.[Status] = 'A'
                                    ORDER BY aa.SortNumber
                                    FOR JSON PATH, ROOT('data')
                                ) Authority
                                FROM PMIS.TaskUser a
                                INNER JOIN BAS.[User] u ON u.UserId = a.UserId
                                INNER JOIN PMIS.ProjectTask b ON a.TaskId = b.ProjectTaskId
                                WHERE a.TaskType = 'Project' 
                                AND b.ProjectTaskId IN @ProjectTaskIds";
                        List<ProjectTaskUser> resultTaskUser = sqlConnection.Query<ProjectTaskUser>(sql, new { ProjectTaskIds }).ToList();
                        #endregion

                        List<ProjectTaskLink> projectTaskLinks = resultTaskLink; //暫存新複製的任務連結
                        List<ProjectTaskUser> projectTaskUsers = new List<ProjectTaskUser>(); //暫存新複製的任務成員

                        resultSubTasks.ForEach(task =>
                        {
                            #region //任務結構新增
                            sql = @"INSERT INTO PMIS.ProjectTask (ProjectId, TaskSort, TaskName
                                    , TaskDesc, ReplyFrequency
                                    , PlannedStart, PlannedEnd, PlannedDuration
                                    , EstimateStart, EstimateEnd, EstimateDuration, TaskStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectTaskId
                                    VALUES (@ProjectId, @TaskSort, @TaskName
                                    , @TaskDesc, @ReplyFrequency, @PlannedStart, @PlannedEnd, @PlannedDuration
                                    , @EstimateStart, @EstimateEnd, @EstimateDuration, 'B'
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected++;
                            int projectTaskId = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    ProjectId,
                                    TaskSort = task.ProjectTaskId == ProjectTaskId ? maxSort + 1 : Convert.ToInt32(task.TaskSort), //第一筆 取得所屬層最大排序+1
                                    TaskName = task.ProjectTaskId == ProjectTaskId ? task.TaskName + "(複製)" : task.TaskName,
                                    task.TaskDesc,
                                    task.ReplyFrequency,
                                    task.PlannedStart,
                                    task.PlannedEnd,
                                    task.PlannedDuration,
                                    task.EstimateStart,
                                    task.EstimateEnd,
                                    task.EstimateDuration,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                })?.ProjectTaskId; //取得新產生的任務id
                            #endregion

                            #region //新複製的任務連結
                            foreach (var link in projectTaskLinks.Where(w => w.TargetTaskId == task.ProjectTaskId))
                            {
                                link.TargetTaskId = projectTaskId;
                            }
                            foreach (var link in projectTaskLinks.Where(w => w.SourceTaskId == task.ProjectTaskId))
                            {
                                link.SourceTaskId = projectTaskId;
                            }
                            #endregion

                            #region //新複製的任務成員
                            foreach (var user in resultTaskUser.Where(w => w.TaskId == task.ProjectTaskId))
                            {
                                projectTaskUsers.Add(
                                    new ProjectTaskUser
                                    {
                                        TaskId = projectTaskId,
                                        UserId = user.UserId,
                                        UserNo = user.UserNo,
                                        LevelId = user.LevelId,
                                        AgentSort = user.AgentSort,
                                        Authority = user.Authority
                                    });
                            }
                            #endregion

                            resultSubTasks
                                .Where(w => w.ParentTaskId == task.ProjectTaskId) //以舊的關聯篩選出資料
                                .ToList()
                                .ForEach(item =>
                                {
                                    item.ParentTaskId = projectTaskId; //重新指派子任務的 ParentTaskId 為新的 ProjectTaskId
                                });
                            task.ProjectTaskId = projectTaskId; //指派新的 ProjectTaskId
                        });

                        resultSubTasks.ForEach(task =>
                        {
                            #region //新複製的任務ParentTaskId更新
                            sql = @"UPDATE a SET
                                    a.ParentTaskId = @ParentTaskId,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM PMIS.ProjectTask a
                                    WHERE ProjectTaskId = @ProjectTaskId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    task.ProjectTaskId,
                                    ParentTaskId = task.ParentTaskId < 0 ? null : task.ParentTaskId,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            #endregion
                        });

                        #region //寫入任務連結資料
                        projectTaskLinks.ForEach(link =>
                        {
                            sql = @"INSERT INTO PMIS.ProjectTaskLink (SourceTaskId, TargetTaskId, LinkType
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectTaskLinkId
                                    VALUES (@SourceTaskId, @TargetTaskId, @LinkType
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    link.SourceTaskId,
                                    link.TargetTaskId,
                                    link.LinkType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        });
                        #endregion

                        #region //寫入任務成員資料 
                        projectTaskUsers.ForEach(user =>
                        {
                            sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TaskUserId
                                    VALUES (@TaskId, 'Project', @UserId, @LevelId, @AgentSort
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    user.TaskId,
                                    user.UserId,
                                    user.LevelId,
                                    user.AgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected ++;

                            int TaskUserId = insertResult?.TaskUserId ?? -1;

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
                        });
                        #endregion

                        if (projectStatus == "I")
                        {
                            #region //取得父任務結構
                            sql = GetParentTaskStrucSQL();
                            var resultParentTasks = sqlConnection.Query(sql, new { ProjectTaskId });
                            var ParentTaskIds = resultParentTasks.Where(w => w.ParentTaskId > 0).Select(s => s.ParentTaskId).ToArray();
                            #endregion

                            #region //更新主任務狀態
                            sql = @"UPDATE a SET 
                                    a.ActualEnd = NULL,
                                    a.ActualDuration = NULL,
                                    a.TaskStatus = CASE a.TaskStatus WHEN 'C' THEN 'I' ELSE a.TaskStatus END
                                    FROM PMIS.ProjectTask a
                                    WHERE a.ProjectTaskId IN @ParentTaskIds";
                            rowsAffected += sqlConnection.Execute(sql, new { ParentTaskIds });
                            #endregion

                            #region MAMO (停用)
                            //if (!string.IsNullOrWhiteSpace(companyNo) && teamId > 0)
                            //{
                            //    foreach (var task in projectTasks)
                            //    {
                            //        #region //MAMO頻道新增
                            //        int channelId = -1;
                            //        var mamoResult = mamoHelper.CreateChannels(companyNo, CurrentUser, teamId, task.TaskName, "描述：" + task.TaskDesc);
                            //        JObject mamoResultJson = JObject.Parse(mamoResult);

                            //        if (mamoResultJson["status"].ToString() == "success")
                            //        {
                            //            foreach (var item in mamoResultJson["result"])
                            //            {
                            //                channelId = Convert.ToInt32(item["ChannelId"].ToString());
                            //            }
                            //        }
                            //        #endregion

                            //        #region //MAMO頻道成員新增
                            //        if (channelId > 0)
                            //        {
                            //            mamoResult = mamoHelper.AddChannelMembers(companyNo, CurrentUser, channelId, projectTaskUsers.Where(w => w.ProjectTaskId.ToString() == task.ProjectTaskId).Select(s => { return s.UserNo; }).Distinct().ToList());
                            //            mamoResultJson = JObject.Parse(mamoResult);
                            //        }
                            //        #endregion

                            //        #region //TeamId更新
                            //        sql = @"UPDATE a SET a.ChannelId = @ChannelId
                            //    FROM PMIS.ProjectTask a
                            //    WHERE a.ProjectTaskId = @ProjectTaskId";
                            //        dynamicParameters.AddDynamicParams(
                            //            new
                            //            {
                            //                task.ProjectTaskId,
                            //                ChannelId = channelId
                            //            });
                            //        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //        #endregion
                            //    }
                            //}
                            #endregion
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

        #region //AddProjectTaskByTemplate -- 專案任務資料新增 -- Ben Ma 2022.10.05
        public string AddProjectTaskByTemplate(int ProjectId, int ParentTaskId, int TemplateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", new { ProjectId });
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");
                        #endregion

                        #region //判斷父層專案任務資料是否正確
                        if (ParentTaskId > 0)
                        {
                            sql = @"SELECT TOP 1 1
                                    FROM PMIS.ProjectTask
                                    WHERE ProjectTaskId = @ParentTaskId";
                            result = sqlConnection.Query(sql, new { ParentTaskId });
                            if (result.Count() <= 0) throw new SystemException("父層專案任務資料錯誤!");
                        }
                        #endregion

                        #region //判斷樣板資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Template
                                WHERE TemplateId = @TemplateId";
                        result = sqlConnection.Query(sql, new { TemplateId });
                        if (result.Count() <= 0) throw new SystemException("樣板資料錯誤!");
                        #endregion

                        #region //樣板任務資料
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
                                    , CAST(ISNULL(ParentTaskId, -1) AS nvarchar(MAX)) TaskRoute
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
                                        , CAST(a.TaskRoute + ',' + CAST(b.ParentTaskId AS nvarchar(MAX)) AS nvarchar(MAX)) TaskRoute
                                        , b.TaskSort
                                        , 0
                                        FROM @templateTask a
                                        INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.ParentTaskId
                                        WHERE a.processed = 1

                                    SET @rowsAdded = @@rowcount

                                    UPDATE @templateTask SET processed = 2 WHERE processed = 1
                                END;

                                SELECT a.TemplateTaskId
                                , b.ParentTaskId, b.TaskSort, b.TaskName, ISNULL(b.ReplyFrequency, 0) ReplyFrequency, b.Duration
                                FROM @templateTask a
                                INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                                WHERE 1=1
                                ORDER BY a.TaskLevel, a.TaskRoute, b.TaskSort";
                        List<TemlpateTask> temlpateTasks = sqlConnection.Query<TemlpateTask>(sql, new { TemplateId }).ToList();
                        #endregion

                        #region //樣板任務連結資料
                        sql = @"SELECT a.TemplateTaskLinkId, a.SourceTaskId, a.TargetTaskId, a.LinkType
                                FROM PMIS.TemplateTaskLink a
                                INNER JOIN PMIS.TemplateTask b ON a.SourceTaskId = b.TemplateTaskId
                                WHERE b.TemplateId = @TemplateId
                                ORDER BY a.TemplateTaskLinkId";
                        List<TemlpateTaskLink> temlpateTaskLinks = sqlConnection.Query<TemlpateTaskLink>(sql, new { TemplateId }).ToList();
                        #endregion

                        #region //樣板任務成員資料
                        sql = @"SELECT a.TemplateTaskUserId, a.TemplateTaskId, a.UserId, a.LevelId, a.AgentSort
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
                                    WHERE aa.AuthorityType = 'T'
                                    AND aa.[Status] = 'A'
                                    ORDER BY aa.SortNumber
                                    FOR JSON PATH, ROOT('data')
                                ) Authority
                                FROM PMIS.TemplateTaskUser a
                                INNER JOIN PMIS.TemplateTask b ON a.TemplateTaskId = b.TemplateTaskId
                                WHERE b.TemplateId = @TemplateId";
                        List<TemplateTaskUser> templateTaskUsers = sqlConnection.Query<TemplateTaskUser>(sql, new { TemplateId }).ToList();
                        #endregion

                        int index = 0;
                        foreach (var task in temlpateTasks)
                        {
                            sql = @"INSERT INTO PMIS.ProjectTask (ProjectId, TaskSort, TaskName
                                    , PlannedDuration, ReplyFrequency, TaskStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectTaskId
                                    VALUES (@ProjectId, @TaskSort, @TaskName
                                    , @PlannedDuration, @ReplyFrequency, 'B'
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            int ProjectTaskId = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    ProjectId,
                                    TaskSort = Convert.ToInt32(task.TaskSort),
                                    task.TaskName,
                                    PlannedDuration = task.Duration,
                                    task.ReplyFrequency,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                }).ProjectTaskId;

                            temlpateTasks
                                .Where(x => x.TemplateTaskId == task.TemplateTaskId)
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.ProjectTaskId = ProjectTaskId;
                                });

                            rowsAffected++;
                            index++;
                        }

                        foreach (var task in temlpateTasks)
                        {
                            if (task.ParentTaskId > 0)
                            {
                                int? parentTaskId = temlpateTasks.Where(x => x.TemplateTaskId == task.ParentTaskId).Select(x => x.ProjectTaskId).FirstOrDefault();

                                temlpateTasks
                                    .Where(x => x.TemplateTaskId == task.TemplateTaskId)
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.ParentProjectTaskId = parentTaskId;
                                    });
                            }
                        }

                        foreach (var task in temlpateTasks)
                        {
                            sql = @"UPDATE PMIS.ProjectTask SET
                                    ParentTaskId = @ParentTaskId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectTaskId = @ProjectTaskId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    task.ProjectTaskId,
                                    ParentTaskId = task.ParentProjectTaskId > 0 ? task.ParentProjectTaskId : (ParentTaskId > 0 ? ParentTaskId : (int?)null),
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                        }

                        foreach (var task in temlpateTasks)
                        {
                            #region //抓取目前最大序號
                            sql = @"SELECT ISNULL(MAX(TaskSort), 0) MaxSort, COUNT(1) CountTask
                                    FROM PMIS.ProjectTask
                                    WHERE ProjectId = @ProjectId
                                    AND ISNULL(ParentTaskId, -1) = @ParentTaskId";
                            var resultSort = sqlConnection.Query(sql,
                                new 
                                {
                                    ProjectId,
                                    ParentTaskId = task.ParentProjectTaskId > 0 ? task.ParentProjectTaskId : ParentTaskId
                                });

                            int maxSort = 0, countTask = 0;
                            foreach (var item in resultSort)
                            {
                                maxSort = item.MaxSort;
                                countTask = item.CountTask;
                            }
                            #endregion

                            sql = @"UPDATE PMIS.ProjectTask SET
                                    TaskSort = @TaskSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectTaskId = @ProjectTaskId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    task.ProjectTaskId,
                                    TaskSort = countTask == maxSort ? task.TaskSort : (maxSort + 1),
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                        }

                        #region //複製任務連結
                        List<ProjectTaskLink> projectTaskLinks = new List<ProjectTaskLink>();
                        foreach (var link in temlpateTaskLinks)
                        {
                            projectTaskLinks.Add(
                                new ProjectTaskLink
                                {
                                    SourceTaskId = (int)temlpateTasks.Where(x => x.TemplateTaskId == link.SourceTaskId).Select(x => x.ProjectTaskId).FirstOrDefault(),
                                    TargetTaskId = (int)temlpateTasks.Where(x => x.TemplateTaskId == link.TargetTaskId).Select(x => x.ProjectTaskId).FirstOrDefault(),
                                    LinkType = link.LinkType
                                });
                        }

                        projectTaskLinks
                            .ToList()
                            .ForEach(x =>
                            {
                                x.CreateDate = CreateDate;
                                x.LastModifiedDate = LastModifiedDate;
                                x.CreateBy = CreateBy;
                                x.LastModifiedBy = LastModifiedBy;
                            });

                        sql = @"INSERT INTO PMIS.ProjectTaskLink (SourceTaskId, TargetTaskId, LinkType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectTaskLinkId
                                VALUES (@SourceTaskId, @TargetTaskId, @LinkType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, projectTaskLinks);
                        #endregion

                        #region //複製任務成員
                        List<ProjectTaskUser> projectTaskUsers = new List<ProjectTaskUser>();
                        foreach (var user in templateTaskUsers)
                        {
                            projectTaskUsers.Add(
                                new ProjectTaskUser
                                {
                                    TaskId = (int)temlpateTasks.Where(x => x.TemplateTaskId == user.TemplateTaskId).Select(x => x.ProjectTaskId).FirstOrDefault(),
                                    UserId = user.UserId,
                                    LevelId = user.LevelId,
                                    AgentSort = user.AgentSort,
                                    Authority = user.Authority
                                });
                        }

                        foreach (var user in projectTaskUsers)
                        {
                            sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TaskUserId
                                    VALUES (@TaskId, 'Project', @UserId, @LevelId, @AgentSort
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    user.TaskId,
                                    user.UserId,
                                    user.LevelId,
                                    user.AgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected ++;

                            int TaskUserId = insertResult?.TaskUserId ?? -1;

                            #region //權限複製
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

        #region //AddProjectTaskLink -- 專案任務連結資料新增 -- Ben Ma 2022.09.27
        public string AddProjectTaskLink(int SourceTaskId, int TargetTaskId, string LinkType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷來源樣板任務資料是否正確
                        sql = @"SELECT TOP 1 ParentTaskId
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @SourceTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { SourceTaskId }) ?? throw new SystemException("【來源專案任務】資料錯誤!");
                        int SourceParent = result?.ParentTaskId ?? -1;
                        #endregion

                        #region //判斷目標樣板任務資料是否正確
                        sql = @"SELECT TOP 1 ParentTaskId, TaskStatus
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @TargetTaskId";
                        dynamicParameters.Add("TargetTaskId", TargetTaskId);
                        result = sqlConnection.Query(sql, dynamicParameters) ?? throw new SystemException("【目標專案任務】資料錯誤!");
                        int TargetParent = result?.ParentTaskId ?? -1;
                        string TargetTaskStatus = result?.TaskStatus ?? string.Empty;
                        #endregion

                        if (SourceParent != TargetParent) throw new SystemException("不同階層的任務無法連結連結!");
                        if (TargetTaskStatus != "B") throw new SystemException("【目標專案任務】必須是待處理狀態!");

                        #region //判斷連結是否存在循環參考
                        sql = @"SELECT TOP 1 1 
                                FROM PMIS.ProjectTaskLink a
                                WHERE a.SourceTaskId = @TargetTaskId
                                AND a.TargetTaskId = @SourceTaskId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { TargetTaskId, SourceTaskId }) ?? throw new SystemException("【專案任務】不得循環參考!");
                        #endregion

                        #region //判斷連結類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @LinkType
                                AND TypeSchema = 'LinkType'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { LinkType }) ?? throw new SystemException("【連結類型】資料錯誤!");
                        #endregion

                        #region //新增專案任務連結
                        sql = @"INSERT INTO PMIS.ProjectTaskLink (SourceTaskId, TargetTaskId, LinkType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectTaskLinkId
                                VALUES (@SourceTaskId, @TargetTaskId, @LinkType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
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
                        rowsAffected++;
                        #endregion

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

        #region //AddProjectTaskUser -- 專案任務成員新增(停用) -- Ben Ma 2022.09.19
        public string AddProjectTaskUser(int ProjectTaskId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【任務成員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTask a
                                INNER JOIN 
                                WHERE ProjectTaskId = @ProjectTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId }) ?? throw new SystemException("專案任務】資料錯誤!");
                        #endregion

                        string[] usersList = Users.Split(',').Distinct().ToArray();

                        #region //判斷任務成員資料是否正確
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserIds";
                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, new { UserIds = usersList }).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("【任務成員】資料錯誤!");
                        #endregion

                        #region //抓取目前最小檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                ORDER BY SortNumber DESC";
                        int LevelId = sqlConnection.QueryFirstOrDefault(sql).LevelId;
                        #endregion

                        #region //抓取目前最大代理順序
                        sql = @"SELECT ISNULL(MAX(AgentSort), 0) MaxAgentSort
                                FROM PMIS.TaskUser
                                WHERE TaskId = @ProjectTaskId
                                AND TaskType = 'Project'";
                        int maxAgentSort = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId }).MaxAgentSort;
                        #endregion

                        #region //專案任務成員新增
                        List<object> ParametersList = new List<object>();
                        foreach (var UserId in usersList)
                        {
                            maxAgentSort++;
                            ParametersList.Add(
                                new
                                {
                                    TaskId = ProjectTaskId,
                                    UserId,
                                    LevelId,
                                    AgentSort = maxAgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TaskUserId
                                VALUES (@TaskId, 'Project', @UserId, @LevelId, @AgentSort
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, ParametersList);
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

        #region //AddTaskUser 任務成員新增-New (需傳入任務類型) -- Chia Yuan 2023.12.06
        public string AddTaskUser(int TaskId, string TaskType, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【任務成員】不能為空!");
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int channelId = -1;
                        string companyNo = string.Empty;

                        #region //取得專案任務資料
                        switch (TaskType)
                        {
                            case "Project":
                                #region //判斷專案任務資料是否正確
                                sql = @"SELECT TOP 1 c.CompanyNo, ISNULL(a.ChannelId, -1) ChannelId
                                        FROM PMIS.ProjectTask a
                                        INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                        LEFT JOIN BAS.Department d ON d.DepartmentId = b.DepartmentId
                                        LEFT JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                        WHERE a.ProjectTaskId = @TaskId";
                                var result1 = sqlConnection.Query(sql, new { TaskId });
                                if (result1.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");
                                foreach (var item in result1)
                                {
                                    channelId = item.ChannelId;
                                    companyNo = item.CompanyNo;
                                }
                                #endregion
                                break;
                            case "Personal":
                                #region //判斷個人任務資料是否正確
                                sql = @"SELECT TOP 1 1
                                        FROM PMIS.PersonalTask
                                        WHERE PersonalTaskId = @TaskId";
                                var result2 = sqlConnection.Query(sql, new { TaskId });
                                if (result2.Count() <= 0) throw new SystemException("【個人任務】資料錯誤!");
                                #endregion
                                break;
                            default:
                                break;
                        }
                        #endregion

                        #region //判斷專案成員資料是否正確
                        string[] usersList = Users.Split(',').Distinct().ToArray();
                        sql = @"SELECT UserId, UserNo
                                FROM BAS.[User]
                                WHERE UserId IN @UserIds";
                        var resultUsers = sqlConnection.Query(sql, new { UserIds = usersList });
                        if (resultUsers.Count() != usersList.Length) throw new SystemException("【任務成員】資料錯誤!");
                        #endregion

                        #region //抓取目前最小檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                ORDER BY SortNumber DESC";
                        int LevelId = sqlConnection.QueryFirstOrDefault(sql).LevelId;
                        #endregion

                        #region //抓取目前最大代理順序
                        sql = @"SELECT ISNULL(MAX(AgentSort), 0) MaxAgentSort
                                FROM PMIS.TaskUser
                                WHERE TaskId = @TaskId
                                AND TaskType = @TaskType";
                        int maxAgentSort = sqlConnection.QueryFirstOrDefault(sql, new { TaskId, TaskType }).MaxAgentSort;
                        #endregion

                        #region //專案任務成員新增
                        List<object> ParametersList = new List<object>();
                        foreach (var UserId in usersList)
                        {
                            maxAgentSort++;
                            ParametersList.Add(
                                new
                                {
                                    TaskId,
                                    TaskType,
                                    UserId,
                                    LevelId,
                                    AgentSort = maxAgentSort,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        sql = @"INSERT INTO PMIS.TaskUser (TaskId, TaskType, UserId, LevelId, AgentSort
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TaskUserId
                                VALUES (@TaskId, @TaskType, @UserId, @LevelId, @AgentSort
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, ParametersList);
                        #endregion

                        #region //MAMO頻道成員新增 (停用)
                        //if (!string.IsNullOrWhiteSpace(companyNo) && TaskType == "Project" && channelId > 0)
                        //{
                        //    var mamoResult = mamoHelper.AddChannelMembers(companyNo, CurrentUser, channelId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                        //    JObject mamoResultJson = JObject.Parse(mamoResult);
                        //}
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

        #region //AddProjectTaskFile -- 專案任務檔案資料新增 -- Ben Ma 2022.09.22
        public string AddProjectTaskFile(int ProjectTaskId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @ProjectTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId }) ?? throw new SystemException("專案任務資料錯誤!");
                        #endregion

                        #region //判斷檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileId }) ?? throw new SystemException("檔案資料錯誤!");
                        #endregion

                        #region //判斷是否是專案成員
                        sql = @"SELECT LevelId
                                FROM PMIS.ProjectUser a
                                INNER JOIN PMIS.ProjectTask b ON a.ProjectId = b.ProjectId
                                WHERE b.ProjectTaskId = @ProjectTaskId
                                AND UserId = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId, CurrentUser }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        int LevelId = result?.LevelId ?? -1;
                        #endregion

                        #region //新增專案任務檔案
                        sql = @"INSERT INTO PMIS.ProjectTaskFile (ProjectTaskId, FileId, LevelId, FileType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectTaskFileId
                                VALUES (@ProjectTaskId, @FileId, @LevelId, @FileType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectTaskId,
                                FileId,
                                LevelId,
                                FileType = "99",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
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

        #region //AddPersonalTaskFile 個人任務檔案資料新增-New -- Chia Yuan 2023.12.14
        public string AddPersonalTaskFile(int PersonalTaskId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.PersonalTask
                                WHERE PersonalTaskId = @PersonalTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskId }) ?? throw new SystemException("【專案任務】資料錯誤!");
                        #endregion

                        #region //判斷檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileId }) ?? throw new SystemException("【檔案】資料錯誤!");
                        #endregion

                        #region //判斷是否是任務成員
                        sql = @"SELECT LevelId, c.AuthorityName, c.AuthorityCode, c.[Status]
                                FROM PMIS.TaskUser a
                                INNER JOIN PMIS.TaskUserAuthority b ON b.TaskUserId = a.TaskUserId
                                INNER JOIN PMIS.Authority c ON c.AuthorityId = b.AuthorityId
                                WHERE a.TaskId = @PersonalTaskId
                                AND a.TaskType = 'Personal'
                                AND a.UserId = @CurrentUser
                                AND c.AuthorityCode = 'task-file-upload'
                                AND c.[Status] = 'A'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskId, CurrentUser }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        int LevelId = result?.LevelId ?? -1;
                        #endregion

                        #region //新增個人任務檔案
                        sql = @"INSERT INTO PMIS.PersonalTaskFile (PersonalTaskId, FileId, LevelId, FileType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PersonalTaskFileId
                                VALUES (@PersonalTaskId, @FileId, @LevelId, @FileType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                PersonalTaskId,
                                FileId,
                                LevelId,
                                FileType = "99",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
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

        #region //AddProjectMamo -- 建立MAMO團隊及頻道 -- Chia Yuan 2024.06.05
        public string AddProjectMamo(int ProjectId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    var rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT a.ProjectId
                                , ISNULL(a.TeamId, -1) TeamId, a.CompanyType, a.ProjectType, a.ProjectNo
                                , a.ProjectName, a.ProjectDesc, a.ProjectAttribute, a.WorkTimeStatus, a.WorkTimeInterval, a.CustomerRemark
                                , a.ProjectRemark, a.ProjectStatus
                                , b.*
                                FROM PMIS.Project a 
                                OUTER APPLY (
	                                SELECT bu1.UserId,bu1.UserNo,bu1.UserName,bc.CompanyNo,bc.CompanyName,bd.DepartmentNo,bd.DepartmentName
	                                FROM PMIS.ProjectUser ba
	                                INNER JOIN BAS.[User] bu1 ON bu1.UserId = ba.UserId
	                                INNER JOIN BAS.Department bd ON bd.DepartmentId = bu1.DepartmentId
	                                INNER JOIN BAS.Company bc ON bc.CompanyId = bd.CompanyId
	                                WHERE ba.ProjectId = a.ProjectId
	                                AND ba.AgentSort = 1
                                ) b
                                WHERE a.ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        #endregion

                        int TeamId = result?.TeamId;
                        string companyNo = result?.CompanyNo;
                        string projectName = result?.ProjectName;
                        string projectNo = result?.ProjectNo;
                        string projectStatus = result?.ProjectStatus;

                        if (string.IsNullOrWhiteSpace(companyNo)) throw new SystemException("【公司別】資料錯誤!");

                        #region //取得任務所有成員資料
                        sql = @"SELECT DISTINCT a.TaskId
                                , u.UserId, u.UserNo, u.UserName
                                FROM PMIS.TaskUser a
                                INNER JOIN PMIS.ProjectTask b ON b.ProjectTaskId = a.TaskId
                                INNER JOIN BAS.[User] u ON u.UserId = a.UserId
                                WHERE a.TaskType = 'Project'
                                AND b.ProjectId = @ProjectId
                                AND NOT EXISTS (
	                                SELECT TOP 1 1 
	                                FROM MAMO.ChannelMembers ca
	                                WHERE ca.UserId = a.UserId
	                                AND ca.ChannelId = b.ChannelId
                                )";
                        var resultTaskUsers = sqlConnection.Query(sql, new { ProjectId });
                        #endregion

                        if (TeamId <= 0)
                        {
                            #region //取得專案任務所有成員資料
                            sql = @"SELECT DISTINCT a.UserId, a.UserNo, a.UserName
                                    FROM (
	                                    SELECT au.UserId, au.UserNo, au.UserName
	                                    FROM PMIS.TaskUser aa
	                                    INNER JOIN PMIS.ProjectTask ab ON ab.ProjectTaskId = aa.TaskId
	                                    INNER JOIN BAS.[User] au ON au.UserId = aa.UserId
	                                    WHERE aa.TaskType = 'Project' 
	                                    AND ab.ProjectId = @ProjectId
	                                    UNION
	                                    SELECT au.UserId, au.UserNo, au.UserName
	                                    FROM PMIS.ProjectUser aa
	                                    INNER JOIN PMIS.Project ab ON ab.ProjectId = aa.ProjectId
	                                    INNER JOIN BAS.[User] au ON au.UserId = aa.UserId
	                                    WHERE aa.ProjectId = @ProjectId
                                    ) a";
                            var resultUsers = sqlConnection.Query(sql, new { ProjectId });
                            #endregion

                            #region //MAMO團隊新增
                            var mamoResult = mamoHelper.CreateTeams(companyNo, CurrentUser, string.Format("PM-{0}-{1}", projectName, projectNo), "Project");
                            JObject mamoResultJson = JObject.Parse(mamoResult);

                            if (mamoResultJson["status"].ToString() == "success")
                            {
                                foreach (var data in mamoResultJson["result"])
                                {
                                    TeamId = Convert.ToInt32(data["TeamId"].ToString());
                                }

                                #region //MAMO團隊成員新增
                                if (resultUsers.Any())
                                {
                                    mamoResult = mamoHelper.AddTeamMembers(companyNo, CurrentUser, TeamId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                                    mamoResultJson = JObject.Parse(mamoResult);
                                }
                                #endregion
                            }
                            #region //TeamId更新
                            sql = @"UPDATE a 
                                    SET a.TeamId = @TeamId
				                    FROM PMIS.Project a
				                    WHERE a.ProjectId = @ProjectId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProjectId,
                                    TeamId
                                });
                            #endregion
                            #endregion
                        }
                        else
                        {
                            #region //取得專案任務所有成員資料
                            sql = @"SELECT DISTINCT a.UserId, a.UserNo, a.UserName
                                    FROM (
	                                    SELECT au.UserId, au.UserNo, au.UserName
	                                    FROM PMIS.TaskUser aa
	                                    INNER JOIN PMIS.ProjectTask ab ON ab.ProjectTaskId = aa.TaskId
	                                    INNER JOIN BAS.[User] au ON au.UserId = aa.UserId
	                                    WHERE aa.TaskType = 'Project' 
	                                    AND ab.ProjectId = @ProjectId
	                                    UNION
	                                    SELECT au.UserId, au.UserNo, au.UserName
	                                    FROM PMIS.ProjectUser aa
	                                    INNER JOIN PMIS.Project ab ON ab.ProjectId = aa.ProjectId
	                                    INNER JOIN BAS.[User] au ON au.UserId = aa.UserId
	                                    WHERE aa.ProjectId = @ProjectId
                                    ) a
                                    WHERE NOT EXISTS (
	                                    SELECT TOP 1 1
	                                    FROM MAMO.TeamMembers aa
	                                    WHERE aa.UserId = a.UserId
                                        AND aa.TeamId = @TeamId
                                    )";
                            var resultUsers = sqlConnection.Query(sql, 
                                new
                                {
                                    ProjectId,
                                    TeamId
                                });
                            #endregion

                            #region //MAMO團隊重啟
                            mamoHelper.RestoreTeams(companyNo, CurrentUser, TeamId);
                            #endregion

                            #region //MAMO團隊成員新增
                            if (resultUsers.Any())
                            {
                                var mamoResult = mamoHelper.AddTeamMembers(companyNo, CurrentUser, TeamId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                                JObject mamoResultJson = JObject.Parse(mamoResult);
                            }
                            #endregion
                        }

                        #region //取得專案任務資料
                        sql = @"SELECT DISTINCT a.ProjectTaskId
                                , ISNULL(a.ChannelId, -1) ChannelId, a.TaskName, a.TaskDesc
                                FROM PMIS.ProjectTask a
                                WHERE a.ProjectId = @ProjectId";
                        var resultSubTask = sqlConnection.Query(sql, new { ProjectId }).ToList();
                        #endregion

                        foreach (var item in resultSubTask)
                        {
                            int ChannelId = item.ChannelId;

                            if (ChannelId <= 0)
                            {
                                #region //MAMO頻道新增
                                var mamoResult = mamoHelper.CreateChannels(companyNo, CurrentUser, TeamId, item.TaskName, "描述：" + item.TaskDesc);
                                JObject mamoResultJson = JObject.Parse(mamoResult);

                                if (mamoResultJson["status"].ToString() == "success")
                                {
                                    foreach (var data in mamoResultJson["result"])
                                    {
                                        ChannelId = Convert.ToInt32(data["ChannelId"].ToString());
                                    }

                                    var tempTaskUsers = resultTaskUsers.Where(w => w.TaskId == item.ProjectTaskId).Select(s => { return (string)s.UserNo; }).Distinct().ToList();

                                    #region //MAMO頻道成員新增 (停用)
                                    if (tempTaskUsers.Any())
                                    {
                                        mamoResult = mamoHelper.AddChannelMembers(companyNo, CurrentUser, ChannelId, tempTaskUsers);
                                        mamoResultJson = JObject.Parse(mamoResult);
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //TeamId更新
                                sql = @"UPDATE a SET 
                                        a.ChannelId = @ChannelId
				                        FROM PMIS.ProjectTask a
				                        WHERE a.ProjectTaskId = @ProjectTaskId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        item.ProjectTaskId,
                                        ChannelId
                                    });
                                #endregion
                            }
                            else
                            {
                                #region //頻道重啟
                                mamoHelper.RestoreChannels(companyNo, CurrentUser, ChannelId);
                                #endregion

                                var tempTaskUsers = resultTaskUsers.Where(w => w.TaskId == item.ProjectTaskId).Select(s => { return (string)s.UserNo; }).Distinct().ToList();

                                #region //MAMO頻道成員新增
                                if (tempTaskUsers.Any())
                                {
                                    var mamoResult = mamoHelper.AddChannelMembers(companyNo, CurrentUser, ChannelId, tempTaskUsers);
                                    JObject mamoResultJson = JObject.Parse(mamoResult);
                                }
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
        #region //UpdateProject -- 專案資料更新 -- Ben Ma 2022.08.19
        public string UpdateProject(int ProjectId, int DepartmentId, string ProjectName, string ProjectDesc
            , string ProjectAttribute, string WorkTimeStatus, string WorkTimeInterval, string CustomerRemark
            , string ProjectRemark)
        {
            try
            {
                if (ProjectName.Length <= 0) throw new SystemException("【專案名稱】不能為空!");
                if (ProjectName.Length > 100) throw new SystemException("【專案名稱】長度錯誤!");
                if (ProjectDesc.Length <= 0) throw new SystemException("【專案描述】不能為空!");
                if (ProjectDesc.Length > 300) throw new SystemException("【專案描述】長度錯誤!");
                if (ProjectAttribute.Length <= 0) throw new SystemException("【專案屬性】不能為空!");
                if (WorkTimeStatus.Length <= 0) throw new SystemException("【工作天模式】不能為空!");
                if (WorkTimeStatus == "Y") if (WorkTimeInterval.Length <= 0) throw new SystemException("【工作時段】不能為空!");
                if (WorkTimeInterval.Length > 100) throw new SystemException("【工作時段】長度錯誤!");
                if (CustomerRemark.Length > 500) throw new SystemException("【客戶備註】長度錯誤!");
                if (ProjectRemark.Length > 500) throw new SystemException("【專案備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ProjectStatus
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (result.ProjectStatus != "P") throw new SystemException("【專案】非計畫中，無法修改!");
                        string ProjectStatus = result.ProjectStatus;
                        #endregion

                        #region //判斷部門資料是否正確
                        if (DepartmentId > 0)
                        {
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE DepartmentId = @DepartmentId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("部門資料錯誤!");
                        }
                        #endregion

                        #region //判斷專案成員才能更新
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId
                                AND UserId = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, CurrentUser }) ?? throw new SystemException("專案成員才有更新權限!");
                        #endregion

                        #region //專案資料更新
                        sql = @"UPDATE PMIS.Project SET
                                DepartmentId = @DepartmentId,
                                ProjectName = @ProjectName,
                                ProjectDesc = @ProjectDesc,
                                ProjectAttribute = @ProjectAttribute,
                                WorkTimeStatus = @WorkTimeStatus,
                                WorkTimeInterval = @WorkTimeInterval,
                                CustomerRemark = @CustomerRemark,
                                ProjectRemark = @ProjectRemark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        int rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                DepartmentId = DepartmentId <= 0 ? (int?)null : DepartmentId,
                                ProjectName,
                                ProjectDesc,
                                ProjectAttribute,
                                WorkTimeStatus,
                                WorkTimeInterval,
                                CustomerRemark,
                                ProjectRemark,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectStartUp -- 專案啟動 -- Ben Ma 2022.10.27
        public string UpdateProjectStartUp(int ProjectId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ProjectStatus
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (result.ProjectStatus != "P") throw new SystemException("【專案】已經啟動或已經完成，無法再啟動!");
                        string ProjectStatus = result.ProjectStatus;
                        #endregion

                        #region //判斷是否是專案負責人
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser a
                                WHERE a.ProjectId = @ProjectId
                                AND a.UserId = @CurrentUser
                                AND a.AgentSort = 1";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, CurrentUser }) ?? throw new SystemException("只有專案負責人可以啟動!");
                        #endregion

                        #region //取得專案任務結購
                        sql = @"SELECT a.ProjectTaskId, a.TaskName, a.TaskDesc, a.TaskSort
                                , ISNULL(a.ChannelId, -1) ChannelId
                                , d.SubUser, e.SubTask
                                FROM PMIS.ProjectTask a 
                                OUTER APPLY (
	                                SELECT COUNT(1) SubUser
	                                FROM PMIS.TaskUser da
	                                WHERE da.TaskId = a.ProjectTaskId
                                    AND da.TaskType = 'Project'
                                ) d
                                OUTER APPLY (
                                    SELECT COUNT(1) SubTask
                                    FROM PMIS.ProjectTask ea
                                    WHERE ea.ParentTaskId= a.ProjectTaskId 
                                ) e
                                WHERE a.ProjectId = @ProjectId
                                AND e.SubTask = 0";
                        var resultSubTask = sqlConnection.Query(sql, new { ProjectId });
                        if (resultSubTask.Any(a => a.SubUser == 0)) throw new SystemException(string.Format("以下任務未設定成員</br>{0}</br>無法啟動專案!", string.Join("</br>", resultSubTask.Where(w => w.SubUser == 0).Select(s => "【" + s.TaskName + "】"))));
                        #endregion

                        #region //任務預估時間初始化
                        sql = @"UPDATE PMIS.ProjectTask SET
                                EstimateStart = PlannedStart,
                                EstimateEnd = PlannedEnd,
                                EstimateDuration = PlannedDuration,
                                ActualStart = null,
                                ActualEnd = null,
                                ActualDuration = null,
                                TaskStatus = 'B',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        #region //專案狀態更新
                        sql = @"UPDATE PMIS.Project SET
                                EstimateStart = PlannedStart,
                                EstimateEnd = PlannedEnd,
                                EstimateDuration = PlannedDuration,
                                ActualStart = null,
                                ActualEnd = null,
                                ActualDuration = null,
                                ProjectStatus = 'I',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectReset -- 專案重置 -- Ben Ma 2022.10.27
        public string UpdateProjectReset(int ProjectId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 a.ProjectStatus, ISNULL(a.TeamId, -1) TeamId, c.CompanyNo
                                FROM PMIS.Project a
                                LEFT JOIN BAS.Department d ON d.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                WHERE a.ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (result.ProjectStatus != "I") throw new SystemException("專案尚未啟動或已經完成，無法再重置!");
                        int TeamId = result.TeamId;
                        string ProjectStatus = result.ProjectStatus, CompanyNo = result.CompanyNo;
                        #endregion

                        #region //判斷是否是專案負責人
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser a
                                WHERE a.ProjectId = @ProjectId
                                AND a.UserId = @CurrentUser
                                AND a.AgentSort = 1";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, CurrentUser }) ?? throw new SystemException("只有專案負責人可以重置!");
                        #endregion

                        #region //取得回報附件
                        sql = @"SELECT DISTINCT a.ReplyFile
                                FROM PMIS.TaskReply a
                                WHERE EXISTS (
                                    SELECT TOP 1 1 
                                    FROM PMIS.ProjectTask aa 
                                    WHERE aa.ProjectTaskId = a.TaskId 
                                    AND aa.ProjectId = @ProjectId)
                                AND ISNULL(a.ReplyFile, -1) > 0
                                AND a.ReplyType = 'Project'";
                        var resultReplyFile = sqlConnection.Query(sql, new { ProjectId }).Select(s => s.ReplyFile);
                        #endregion

                        #region //刪除回報內容
                        sql = @"DELETE a
                                FROM PMIS.TaskReply a
                                WHERE EXISTS (
                                    SELECT TOP 1 1 
                                    FROM PMIS.ProjectTask aa 
                                    WHERE aa.ProjectTaskId = a.TaskId 
                                    AND aa.ProjectId = @ProjectId)
                                AND a.ReplyType = 'Project'";
                        rowsAffected += sqlConnection.Execute(sql, new { ProjectId });
                        #endregion

                        #region //刪除檔案
                        sql = @"DELETE a
                                FROM BAS.[File] a
                                WHERE a.FileId IN @FileIds";
                        rowsAffected += sqlConnection.Execute(sql, new { FileIds = resultReplyFile.ToArray() });
                        #endregion

                        #region //任務時間初始化
                        sql = @"UPDATE PMIS.ProjectTask SET
                                EstimateStart = null,
                                EstimateEnd = null,
                                EstimateDuration = null,
                                ActualStart = null,
                                ActualEnd = null,
                                ActualDuration = null,
                                TaskStatus = 'B',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        #region //專案時間初始化
                        sql = @"UPDATE PMIS.Project SET
                                EstimateStart = null,
                                EstimateEnd = null,
                                EstimateDuration = null,
                                ActualStart = null,
                                ActualEnd = null,
                                ActualDuration = null,
                                ProjectStatus = 'P',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        #region //取得專案任務資料
                        sql = @"SELECT a.ProjectTaskId, ISNULL(a.ChannelId, -1) ChannelId
                                FROM PMIS.ProjectTask a
                                WHERE a.ProjectId = @ProjectId";
                        var resultProjectTask = sqlConnection.Query(sql, new { ProjectId });
                        #endregion

                        if (!string.IsNullOrWhiteSpace(CompanyNo))
                        {
                            #region //MAMO頻道刪除 (停用)
                            //foreach (var item in resultProjectTask.Where(w => w.ChannelId > 0))
                            //{
                            //    var mamoResult = mamoHelper.DeleteChannels(CompanyNo, CurrentUser, item.ChannelId);
                            //    JObject mamoResultJson = JObject.Parse(mamoResult);
                            //}
                            #endregion

                            #region //MAMO團隊刪除 (停用)
                            //if (TeamId > 0)
                            //{
                            //    var mamoResult = mamoHelper.DeleteTeams(CompanyNo, CurrentUser, TeamId);
                            //    JObject mamoResultJson = JObject.Parse(mamoResult);
                            //}
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

        #region //UpdateAllTaskRefresh --專案重整 -- Chia Yuan 2023.11.16
        public string UpdateAllTaskRefresh(int ProjectId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ProjectStatus
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        string ProjectStatus = result?.ProjectStatus ?? string.Empty;
                        if (ProjectStatus == "C") throw new SystemException("【專案】已完成，無法再重置!");
                        #endregion

                        #region //判斷是否是專案負責人
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser a
                                WHERE a.ProjectId = @ProjectId
                                AND a.UserId = @CurrentUser
                                AND a.AgentSort = 1";
                        result = sqlConnection.QueryFirstOrDefault(sql, new  { ProjectId, CurrentUser }) ?? throw new SystemException("只有專案負責人可以重整專案!");
                        #endregion

                        #region //專案計畫時間更新為預估時間
                        sql = @"UPDATE PMIS.Project SET 
                                PlannedStart = EstimateStart,
                                PlannedEnd = EstimateEnd,
                                PlannedDuration = EstimateDuration,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        #region //任務計畫時間更新為預估時間
                        sql = @"UPDATE PMIS.ProjectTask SET
                                PlannedStart = EstimateStart,
                                PlannedEnd = EstimateEnd,
                                PlannedDuration = EstimateDuration,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectUserAgentSort -- 專案成員代理順序更新 -- Ben Ma 2022.08.22
        public string UpdateProjectUserAgentSort(int ProjectId, string ProjectUserList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷是否是專案負責人
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser a
                                WHERE a.ProjectId = @ProjectId
                                AND a.UserId = @CurrentUser
                                AND a.AgentSort = 1";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId, CurrentUser }) ?? throw new SystemException("只有專案負責人可以修改!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE PMIS.ProjectUser SET
                                AgentSort = AgentSort * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        totalRowsAffected = sqlConnection.Execute(sql, new 
                        {
                            ProjectId,
                            LastModifiedDate,
                            LastModifiedBy
                        });
                        #endregion

                        string[] usersList = ProjectUserList.Split(',').Distinct().ToArray();

                        #region //更新順序
                        for (int i = 0; i < usersList.Length; i++)
                        {
                            sql = @"UPDATE PMIS.ProjectUser SET
                                    AgentSort = @AgentSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectId = @ProjectId
                                    AND ProjectUserId = @ProjectUserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProjectId,
                                    ProjectUserId = usersList[i],
                                    AgentSort = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

                        //第一順序的成員，擁有所有權限與最高檔案等級
                        #region //抓取第一順序的成員
                        sql = @"SELECT TOP 1 ProjectUserId
                                FROM PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId
                                AND AgentSort = 1";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId });
                        int ProjectUserId = result?.ProjectUserId ?? -1;
                        #endregion

                        #region //取得目前最高檔案等級
                        sql = @"SELECT TOP 1 LevelId
                                FROM PMIS.FileLevel
                                WHERE Status = 'A'
                                ORDER BY SortNumber";
                        result = sqlConnection.QueryFirstOrDefault(sql) ?? throw new SystemException("檔案等級資料不存在，請確認是否有新增!");
                        int maxLevel = result?.LevelId ?? -1;
                        #endregion

                        #region //更新成員檔案等級
                        sql = @"UPDATE PMIS.ProjectUser SET
                                    LevelId = @LevelId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectId = @ProjectId
                                    AND ProjectUserId = @ProjectUserId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectId,
                                ProjectUserId,
                                LevelId = maxLevel,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectUserAuthority -- 專案成員權限更新 -- Ben Ma 2022.08.22
        public string UpdateProjectUserAuthority(int ProjectUserId, int AuthorityId, bool Checked)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案成員資料是否正確
                        sql = @"SELECT TOP 1 AgentSort
                                FROM PMIS.ProjectUser
                                WHERE ProjectUserId = @ProjectUserId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectUserId }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        int AgentSort = result?.AgentSort ?? -1;
                        if (AgentSort == 1 && !Checked) throw new SystemException("專案負責人無法移除權限!");
                        #endregion
                        
                        #region //判斷權限資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityId = @AuthorityId
                                AND AuthorityType = 'P'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { AuthorityId }) ?? throw new SystemException("權限資料錯誤!");
                        #endregion

                        if (Checked)
                        {
                            #region //新增專案成員權限
                            sql = @"INSERT INTO PMIS.ProjectUserAuthority (ProjectUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@ProjectUserId, @AuthorityId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, new
                            {
                                ProjectUserId,
                                AuthorityId,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                            #endregion
                        }
                        else
                        {
                            #region //刪除
                            sql = @"DELETE PMIS.ProjectUserAuthority
                                    WHERE ProjectUserId = @ProjectUserId
                                    AND AuthorityId = @AuthorityId";
                            rowsAffected = sqlConnection.Execute(sql, new { ProjectUserId, AuthorityId });
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

        #region //UpdateProjectAllUserAuthority --批次專案成員權限更新 -- Chia Yuan 2023.12.04
        public string UpdateProjectAllUserAuthority(int ProjectUserId = -1, bool Checked = false)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser
                                WHERE ProjectUserId = @ProjectUserId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectUserId }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        #endregion

                        if (Checked)
                        {
                            #region //新增專案成員權限
                            sql = @"INSERT INTO PMIS.ProjectUserAuthority (ProjectUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    SELECT @ProjectUserId, a.AuthorityId, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy 
                                    FROM PMIS.Authority a 
                                    LEFT JOIN (SELECT b.AuthorityId FROM PMIS.ProjectUserAuthority b where b.ProjectUserId = @ProjectUserId) b ON b.AuthorityId = a.AuthorityId
                                    WHERE a.AuthorityType = @AuthorityType
                                    AND a.[Status] = 'A'
                                    AND b.AuthorityId IS NULL
                                    ORDER BY a.SortNumber";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProjectUserId,
                                    AuthorityType = "P",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion
                        }
                        else
                        {
                            #region //刪除專案成員權限
                            sql = @"DELETE PMIS.ProjectUserAuthority
                                    WHERE ProjectUserId = @ProjectUserId";
                            rowsAffected += sqlConnection.Execute(sql, new { ProjectUserId });
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

        #region //UpdateProjectUserLevel -- 專案成員閱覽檔案等級更新 -- Ben Ma 2022.08.22
        public string UpdateProjectUserLevel(int ProjectUserId, int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案成員資料是否正確
                        sql = @"SELECT TOP 1 AgentSort
                                FROM PMIS.ProjectUser
                                WHERE ProjectUserId = @ProjectUserId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectUserId }) ?? throw new SystemException("【專案成員】資料錯誤!");
                        int AgentSort = result?.AgentSort ?? -1;
                        if (AgentSort == 1) throw new SystemException("專案負責人無法更新閱覽檔案等級!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { LevelId }) ?? throw new SystemException("【檔案等級】資料錯誤!");
                        #endregion

                        #region //更新成員檔案等級
                        sql = @"UPDATE PMIS.ProjectUser SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectUserId = @ProjectUserId";
                        int rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                ProjectUserId,
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectFileLevel -- 專案檔案資閱覽檔案等級更新 -- Ben Ma 2022.08.31
        public string UpdateProjectFileLevel(int ProjectFileId, int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectFile
                                WHERE ProjectFileId = @ProjectFileId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectFileId }) ?? throw new SystemException("專案檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { LevelId }) ?? throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        #region //更新檔案等級
                        sql = @"UPDATE PMIS.ProjectFile SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectFileId = @ProjectFileId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectFileId,
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectFileFileType -- 專案檔案檔案類型更新 -- Ben Ma 2022.08.31
        public string UpdateProjectFileFileType(int ProjectFileId, string FileType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectFile
                                WHERE ProjectFileId = @ProjectFileId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectFileId }) ?? throw new SystemException("專案檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @FileType
                                AND TypeSchema = 'ProjectFile.FileType'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileType }) ?? throw new SystemException("檔案類型資料錯誤!");
                        #endregion

                        #region //更新檔案類型
                        sql = @"UPDATE PMIS.ProjectFile SET
                                FileType = @FileType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectFileId = @ProjectFileId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectFileId,
                                FileType,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectTask -- 專案任務資料更新 -- Ben Ma 2022.09.15
        public string UpdateProjectTask(int ProjectTaskId, string TaskName, string TaskDesc, int ReplyFrequency, bool RequiredFile
            , string PlannedStart, int PlannedDuration, string EstimateStart, int EstimateDuration)
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
                        #region //取得專案任務資料
                        sql = @"SELECT TOP 1 a.ProjectId, a.ProjectStatus, b.RequiredFile
                                FROM PMIS.Project a
                                INNER JOIN PMIS.ProjectTask b ON b.ProjectId = a.ProjectId
                                WHERE b.ProjectTaskId = @ProjectTaskId";
                        var resultProject = sqlConnection.QueryFirstOrDefault<ProjectTask>(sql, new { ProjectTaskId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (resultProject.ProjectStatus == "C") throw new SystemException("專案已經完成，任務資料無法更新!");
                        bool bequiredFile = Convert.ToBoolean(resultProject.RequiredFile);
                        #endregion

                        #region //遞迴取得前置任務資料
                        sql = GetPrevTaskSQL();
                        ProjectTask prevProjectTask = sqlConnection.Query<ProjectTask>(sql, new { ProjectTaskId }).FirstOrDefault();
                        #endregion

                        string returnType = "success";

                        switch (resultProject.ProjectStatus)
                        {
                            case "P":
                                if (string.IsNullOrEmpty(PlannedStart)) throw new SystemException("【計畫開始時間】不能為空!");
                                bool inputPlannedStart = DateTime.TryParse(PlannedStart, out DateTime plannedStart);

                                if (RequiredFile != bequiredFile) returnType = "ChangeStatus";

                                #region 判斷任務時間是否異常
                                if (prevProjectTask != null)
                                {
                                    if (!inputPlannedStart) throw new SystemException("【計畫開始時間】格式有誤!");
                                    if (!prevProjectTask.PlannedEnd.HasValue) throw new SystemException("【前置任務】計畫完成時間錯誤!");
                                    if (plannedStart.CompareTo(prevProjectTask.PlannedEnd.Value) < 0)
                                        throw new SystemException(string.Format("【計畫開始時間】不得小於前置任務{0}的計畫完成時間{1}", prevProjectTask.TaskName, prevProjectTask.PlannedEnd.Value.ToString("yyyy-MM-dd HH:mm")));
                                }
                                #endregion

                                #region //更新專案任務資料
                                sql = @"UPDATE PMIS.ProjectTask SET
                                        TaskName = @TaskName,
                                        TaskDesc = @TaskDesc,
                                        ReplyFrequency = @ReplyFrequency,
                                        PlannedStart = @PlannedStart,
                                        PlannedDuration = @PlannedDuration,
                                        RequiredFile = @RequiredFile,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectTaskId = @ProjectTaskId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProjectTaskId,
                                        TaskName,
                                        TaskDesc,
                                        ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                        PlannedStart = PlannedStart.Length <= 0 ? (DateTime?)null : plannedStart,
                                        PlannedDuration = PlannedDuration <= 0 ? (int?)null : PlannedDuration,
                                        RequiredFile,
                                        LastModifiedDate,
                                        LastModifiedBy
                                    });
                                #endregion
                                break;
                            case "I":
                                if (string.IsNullOrEmpty(EstimateStart)) throw new SystemException("【預計開始時間】不能為空!");
                                bool inputEstimateStart = DateTime.TryParse(EstimateStart, out DateTime estimateStart);

                                if (RequiredFile != bequiredFile) returnType = "ChangeStatus";

                                #region //判斷任務時間是否異常
                                if (prevProjectTask != null)
                                {
                                    if (!inputEstimateStart) throw new SystemException("【預計開始時間】格式有誤!");
                                    if (!prevProjectTask.EstimateEnd.HasValue) throw new SystemException("【前置任務】預計完成時間錯誤!");
                                    if (estimateStart.CompareTo(prevProjectTask.EstimateEnd.Value) < 0)
                                        throw new SystemException(string.Format("【預計開始時間】不得小於前置任務{0}的預計完成時間{1}", prevProjectTask.TaskName, prevProjectTask.EstimateEnd.Value.ToString("yyyy-MM-dd HH:mm")));
                                }
                                #endregion

                                #region //更新專案任務資料
                                sql = @"UPDATE PMIS.ProjectTask SET
                                        TaskName = @TaskName,
                                        TaskDesc = @TaskDesc,
                                        ReplyFrequency = @ReplyFrequency,
                                        EstimateStart = @EstimateStart,
                                        EstimateDuration = @EstimateDuration,
                                        RequiredFile = @RequiredFile,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectTaskId = @ProjectTaskId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProjectTaskId,
                                        TaskName,
                                        TaskDesc,
                                        ReplyFrequency = ReplyFrequency <= 0 ? (int?)null : ReplyFrequency,
                                        EstimateStart = EstimateStart.Length <= 0 ? (DateTime?)null : estimateStart,
                                        EstimateDuration = EstimateDuration <= 0 ? (int?)null : EstimateDuration,
                                        RequiredFile,
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
                            data = returnType,
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

        #region //UpdateTemporaryProjectTask 插入、新增子任務排成更新(僅包含位移任務) --Chia Yuan 2023.10.31
        public string UpdateTemporaryProjectTask(int ProjectId, int ProjectTaskId, string TaskData)
        {
            try
            {
                if (!TaskData.TryParseJson(out JObject tempJObject)) throw new SystemException("【專案任務】格式錯誤!");

                List<ProjectTask> ProjectTasks = JsonConvert.DeserializeObject<List<ProjectTask>>(JObject.Parse(TaskData)["data"].ToString());

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 ProjectStatus
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!!");
                        string projectStatus = result?.ProjectStatus ?? string.Empty;
                        #endregion

                        #region //取得專案任務資料
                        sql = @"SELECT TaskStatus, TaskName
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @ProjectTaskId";
                        var result2 = sqlConnection.Query(sql, new { ProjectTaskId });
                        foreach (var item in result2)
                        {
                            if (item.TaskStatus != "B") throw new SystemException(string.Format("【{0}】非待處理狀態，不得異動!", item.TaskName));
                        }
                        #endregion

                        foreach (var task in ProjectTasks)
                        {
                            if (task.ProjectTaskId.ToString() != "1.1")
                            {
                                #region //更新任務資料
                                sql = @"UPDATE PMIS.ProjectTask SET
                                        EstimateStart = @EstimateStart,
                                        EstimateEnd = @EstimateEnd,
                                        EstimateDuration = @EstimateDuration,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectTaskId = @ProjectTaskId";
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
                            }
                            else
                            {
                                #region //更新任務資料
                                sql = @"UPDATE PMIS.Project SET
                                        EstimateStart = @EstimateStart,
                                        EstimateEnd = @EstimateEnd,
                                        EstimateDuration = @EstimateDuration,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectId = @ProjectId";
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

        #region //UpdateProjectTaskSort -- 專案任務順序調整 -- Ben Ma 2022.09.01
        public string UpdateProjectTaskSort(int ProjectTaskId, string SortType)
        {
            try
            {
                if (!Regex.IsMatch(SortType, "^(top|bottom)$", RegexOptions.IgnoreCase)) throw new SystemException("【順序調整類型】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷專案任務資料是否正確
                        sql = @"SELECT TOP 1 ProjectId, ParentTaskId, TaskSort
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @ProjectTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId }) ?? throw new SystemException("專案任務資料錯誤!");

                        int ProjectId = result.ProjectId;
                        int ParentTaskId = result.ParentTaskId ?? 0;
                        int TaskSort = result.TaskSort;
                        #endregion

                        #region //判斷專案資料是否正確
                        sql = @"SELECT TOP 1 b.ProjectStatus
                                FROM PMIS.ProjectTask a
                                INNER JOIN PMIS.Project b ON a.ProjectId = b.ProjectId
                                WHERE a.ProjectTaskId = @ProjectTaskId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId }) ?? throw new SystemException("【專案】資料錯誤!");
                        string ProjectStatus = result?.ProjectStatus ?? string.Empty;
                        if (ProjectStatus != "P") throw new SystemException("【專案】非計畫中，任務順序無法調整!");
                        #endregion

                        switch (SortType)
                        {
                            case "top":
                                if (TaskSort > 1)
                                {
                                    #region //先將自己排序調整成負數
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = TaskSort * -1
                                            WHERE ProjectTaskId = @ProjectTaskId";
                                    rowsAffected += sqlConnection.Execute(sql, new { ProjectTaskId });
                                    #endregion

                                    #region //被調整的後移
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = TaskSort + 1,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ProjectId = @ProjectId
                                            AND ISNULL(ParentTaskId, 0) = @ParentTaskId
                                            AND TaskSort = @TaskSort";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            ProjectId,
                                            ParentTaskId = ParentTaskId > 0 ? ParentTaskId : 0,
                                            TaskSort = TaskSort - 1,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion

                                    #region //將自己的排序前移
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = @TaskSort,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ProjectTaskId = @ProjectTaskId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            ProjectTaskId,
                                            TaskSort = TaskSort - 1,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("任務為第一順位，無法調整!");
                                }
                                break;
                            case "bottom":
                                #region //搜尋最大序號
                                sql = @"SELECT MAX(TaskSort) MaxSort
                                        FROM PMIS.ProjectTask
                                        WHERE ProjectId = @ProjectId
                                        AND ISNULL(ParentTaskId, 0) = @ParentTaskId";
                                int maxSort = sqlConnection.QueryFirstOrDefault(sql,
                                    new 
                                    {
                                        ProjectId,
                                        ParentTaskId = ParentTaskId > 0 ? ParentTaskId : 0
                                    }).MaxSort;
                                #endregion

                                if (TaskSort < maxSort)
                                {
                                    #region //先將自己排序調整成負數
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = TaskSort * -1
                                            WHERE ProjectTaskId = @ProjectTaskId";
                                    rowsAffected += sqlConnection.Execute(sql, new { ProjectTaskId });
                                    #endregion

                                    #region //被調整的前移
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = TaskSort - 1,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ProjectId = @ProjectId
                                            AND ISNULL(ParentTaskId, 0) = @ParentTaskId
                                            AND TaskSort = @TaskSort";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            ProjectId,
                                            ParentTaskId = ParentTaskId > 0 ? ParentTaskId : 0,
                                            TaskSort = TaskSort + 1,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion

                                    #region //將自己的排序後移
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            TaskSort = @TaskSort,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ProjectTaskId = @ProjectTaskId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            ProjectTaskId,
                                            TaskSort = TaskSort + 1,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                        });
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

        #region //UpdateProjectTaskSchedule -- 專案任務排程更新(單一任務) -- Ben Ma 2022.09.27
        public string UpdateProjectTaskSchedule(int ProjectTaskId, string PlannedStart, string PlannedEnd, int PlannedDuration)
        {
            try
            {
                if (!DateTime.TryParse(PlannedStart, out DateTime tempDate)) throw new SystemException("任務【開始時間】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTask
                                WHERE ProjectTaskId = @ProjectTaskId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId}) ?? throw new SystemException("專案任務資料錯誤!");
                        #endregion

                        #region //更新任務資料
                        sql = @"UPDATE PMIS.ProjectTask SET
                                PlannedStart = @PlannedStart,
                                PlannedEnd = @PlannedEnd,
                                PlannedDuration = @PlannedDuration,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectTaskId = @ProjectTaskId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectTaskId,
                                PlannedStart = Convert.ToDateTime(PlannedStart),
                                PlannedEnd = Convert.ToDateTime(PlannedEnd),
                                PlannedDuration,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectAllTaskSchedule -- 專案任務排程更新(全部任務) -- Ben Ma 2022.09.29
        public string UpdateProjectAllTaskSchedule(int ProjectId, string TaskData)
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
                                FROM PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectId }) ?? throw new SystemException("【專案】資料錯誤!");
                        if (result.ProjectStatus != "P") throw new SystemException("【專案】非計畫中，排程無法更新!");
                        string ProjectStatus = result.ProjectStatus;
                        #endregion

                        foreach (var task in projectTasks)
                        {
                            if (task.ProjectTaskId.ToString() != "1.1")
                            {
                                #region //判斷專案任務資料是否正確
                                sql = @"SELECT TOP 1 1
                                        FROM PMIS.ProjectTask
                                        WHERE ProjectTaskId = @ProjectTaskId";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { task.ProjectTaskId }) ?? throw new SystemException("【專案任務】資料錯誤!");
                                #endregion

                                #region //更新任務資料
                                sql = @"UPDATE PMIS.ProjectTask SET
                                        PlannedStart = @PlannedStart,
                                        PlannedEnd = @PlannedEnd,
                                        PlannedDuration = @PlannedDuration,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectTaskId = @ProjectTaskId";
                                rowsAffected = sqlConnection.Execute(sql,
                                    new
                                    {
                                        task.ProjectTaskId,
                                        PlannedStart = Convert.ToDateTime(task.PlannedStart),
                                        PlannedEnd = Convert.ToDateTime(task.PlannedEnd),
                                        task.PlannedDuration,
                                        LastModifiedDate,
                                        LastModifiedBy
                                    });
                                #endregion
                            }
                            else
                            {
                                sql = @"UPDATE PMIS.Project SET
                                        PlannedStart = @PlannedStart,
                                        PlannedEnd = @PlannedEnd,
                                        PlannedDuration = @PlannedDuration,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectId = @ProjectId";
                                rowsAffected = sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProjectId,
                                        PlannedStart = Convert.ToDateTime(task.PlannedStart),
                                        PlannedEnd = Convert.ToDateTime(task.PlannedEnd),
                                        task.PlannedDuration,
                                        LastModifiedDate,
                                        LastModifiedBy
                                    });
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

        #region //UpdateTaskUserAgentSort -- 任務成員代理順序更新-New -- Chia Yuan 2023.12.06
        public string UpdateTaskUserAgentSort(int TaskId, string TaskType, string TaskUserList)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE PMIS.TaskUser SET
                                AgentSort = AgentSort * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TaskId = @TaskId
                                AND TaskType = @TaskType";
                        totalRowsAffected = sqlConnection.Execute(sql,
                            new 
                            {
                                TaskId,
                                TaskType,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        string[] userList = TaskUserList.Split(',').Distinct().ToArray();

                        #region //更新順序
                        for (int i = 0; i < userList.Length; i++)
                        {
                            sql = @"UPDATE PMIS.TaskUser SET
                                    AgentSort = @AgentSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TaskId = @TaskId
                                    AND TaskUserId = @TaskUserId
                                    AND TaskType = @TaskType";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    TaskId,
                                    TaskUserId = userList[i],
                                    TaskType,
                                    AgentSort = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
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

        #region //UpdateTaskUserAuthority -- 任務成員權限更新-New -- Chia Yuan 2023.12.06
        public string UpdateTaskUserAuthority(int TaskUserId, int AuthorityId, string TaskType, bool Checked)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TaskUser
                                WHERE TaskUserId = @TaskUserId
                                AND TaskType = @TaskType";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { TaskUserId, TaskType }) ?? throw new SystemException("【任務成員】資料錯誤!");
                        #endregion

                        #region //判斷權限資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.Authority
                                WHERE AuthorityId = @AuthorityId
                                AND AuthorityType = 'T'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { AuthorityId }) ?? throw new SystemException("權限資料錯誤!");
                        #endregion

                        if (Checked)
                        {
                            #region //新增任務成員權限
                            sql = @"INSERT INTO PMIS.TaskUserAuthority (TaskUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TaskUserId, @AuthorityId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    TaskUserId,
                                    AuthorityId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion
                        }
                        else
                        {
                            #region //刪除任務成員權限
                            sql = @"DELETE PMIS.TaskUserAuthority
                                    WHERE TaskUserId = @TaskUserId
                                    AND AuthorityId = @AuthorityId";
                            rowsAffected += sqlConnection.Execute(sql, new { TaskUserId, AuthorityId });
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

        #region //UpdateAllTaskUserAuthority --批次任務成員權限更新-New -- Chia Yuan 2023.12.06
        public string UpdateAllTaskUserAuthority(int TaskUserId, string TaskType, bool Checked)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TaskUser
                                WHERE TaskUserId = @TaskUserId
                                AND TaskType = @TaskType";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { TaskUserId, TaskType }) ?? throw new SystemException("【任務成員】資料錯誤!");
                        #endregion

                        #region //判斷權限資料是否正確
                        sql = @"SELECT AuthorityId, AuthorityType, AuthorityCode, AuthorityName, SortNumber, Status
                                FROM PMIS.Authority
                                WHERE AuthorityType = 'T'";
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters) ?? throw new SystemException("權限資料錯誤!");
                        #endregion

                        if (Checked)
                        {
                            #region //新增任務成員權限
                            sql = @"INSERT INTO PMIS.TaskUserAuthority (TaskUserId, AuthorityId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    SELECT @TaskUserId, a.AuthorityId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy 
                                    FROM PMIS.Authority a 
                                    LEFT JOIN (SELECT b.AuthorityId from PMIS.TaskUserAuthority b WHERE b.TaskUserId = @TaskUserId) b on b.AuthorityId = a.AuthorityId
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
                        }
                        else
                        {
                            #region //刪除任務成員權限
                            sql = @"DELETE PMIS.TaskUserAuthority
                                    WHERE TaskUserId = @TaskUserId";
                            rowsAffected += sqlConnection.Execute(sql, new { TaskUserId });
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

        #region //UpdateTaskUserLevel -- 任務成員閱覽檔案等級更新-New -- Chia Yuan 2023.12.06
        public string UpdateTaskUserLevel(int TaskUserId, int LevelId, string TaskType)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.TaskUser
                                WHERE TaskUserId = @TaskUserId
                                AND TaskType = @TaskType";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { TaskUserId, TaskType }) ?? throw new SystemException("【任務成員】資料錯誤!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        dynamicParameters.Add("LevelId", LevelId);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters) ?? throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        #region //更新任務成員檔案等級
                        sql = @"UPDATE PMIS.TaskUser SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TaskUserId = @TaskUserId
                                AND TaskType = @TaskType";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                TaskUserId,
                                TaskType,
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectTaskFileLevel -- 專案任務檔案資閱覽檔案等級更新 -- Ben Ma 2022.09.26
        public string UpdateProjectTaskFileLevel(int ProjectTaskFileId, int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTaskFile
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskFileId }) ?? throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        result = sqlConnection.Query(sql, new { LevelId }) ?? throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        #region //更新檔案等級
                        sql = @"UPDATE PMIS.ProjectTaskFile SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                ProjectTaskFileId,
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdatePersonalTaskFileLevel -- 個人任務檔案資閱覽檔案等級更新-New -- Chia Yuan 2023.12.14
        public string UpdatePersonalTaskFileLevel(int PersonalTaskFileId, int LevelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.PersonalTaskFile
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        dynamicParameters.Add("PersonalTaskFileId", PersonalTaskFileId);

                        var result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskFileId }) ?? throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案等級資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.FileLevel
                                WHERE LevelId = @LevelId";
                        result = sqlConnection.Query(sql, new { LevelId }) ?? throw new SystemException("檔案等級資料錯誤!");
                        #endregion

                        #region //更新檔案等級
                        sql = @"UPDATE PMIS.PersonalTaskFile SET
                                LevelId = @LevelId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                PersonalTaskFileId,
                                LevelId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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

        #region //UpdateProjectTaskFileType -- 專案任務檔案檔案類型更新 -- Ben Ma 2022.09.26
        public string UpdateProjectTaskFileType(int ProjectTaskFileId, string FileType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTaskFile
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskFileId }) ?? throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @FileType
                                AND TypeSchema = 'ProjectFile.FileType'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileType }) ?? throw new SystemException("檔案類型資料錯誤!");
                        #endregion

                        #region //更新檔案類型
                        sql = @"UPDATE PMIS.ProjectTaskFile SET
                                FileType = @FileType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                ProjectTaskFileId,
                                FileType,
                                LastModifiedDate,
                                LastModifiedBy                                
                            });
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

        #region //UpdatePersonalTaskFileType -- 個人任務檔案檔案類型更新 -- Chia Yuan 2023.12.14
        public string UpdatePersonalTaskFileType(int PersonalTaskFileId, string FileType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.PersonalTaskFile
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { PersonalTaskFileId }) ?? throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        #region //判斷檔案類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @FileType
                                AND TypeSchema = 'ProjectFile.FileType'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { FileType }) ?? throw new SystemException("檔案類型資料錯誤!");
                        #endregion

                        #region //更新檔案類型
                        sql = @"UPDATE PMIS.PersonalTaskFile SET
                                FileType = @FileType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                PersonalTaskFileId,
                                FileType,
                                LastModifiedDate,
                                LastModifiedBy
                            });
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
        #region //DeleteProject -- 專案資料刪除 -- Ben Ma 2022.08.19
        public string DeleteProject(int ProjectId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ISNULL(a.TeamId, -1) TeamId , a.ProjectStatus, c.CompanyNo
                                FROM PMIS.Project a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");
                        if (result.Any(a => a.ProjectStatus != "P")) throw new SystemException("【專案】非計畫中，任務無法刪除!");
                        int teamId = -1;
                        string companyNo = string.Empty;
                        foreach (var item in result)
                        {
                            teamId = item.TeamId;
                            companyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //判斷專案成員才能刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId
                                AND UserId = @UserId";
                        dynamicParameters.Add("ProjectId", ProjectId);
                        dynamicParameters.Add("UserId", CurrentUser);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("專案成員才有刪除權限!");
                        #endregion

                        #region //刪除關聯table
                        #region //刪除專案任務成員權限
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TaskUserAuthority a
                                INNER JOIN PMIS.TaskUser b ON a.TaskUserId = b.TaskUserId
                                INNER JOIN PMIS.ProjectTask c ON b.TaskId = c.ProjectTaskId
                                WHERE c.ProjectId = @ProjectId
                                AND b.TaskType = 'Project'";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案任務成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TaskUser a
                                INNER JOIN PMIS.ProjectTask b ON a.TaskId = b.ProjectTaskId
                                WHERE b.ProjectId = @ProjectId
                                AND a.TaskType = 'Project'";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案任務連結資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.ProjectTaskLink a
                                INNER JOIN PMIS.ProjectTask b ON a.SourceTaskId = b.ProjectTaskId
                                WHERE b.ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案任務檔案資料
                        #region //任務檔案改為刪除狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                INNER JOIN PMIS.ProjectTaskFile b ON a.FileId = b.FileId
                                INNER JOIN PMIS.ProjectTask c ON b.ProjectTaskId = c.ProjectTaskId
                                WHERE c.ProjectId = @ProjectId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.ProjectTaskFile a
                                INNER JOIN PMIS.ProjectTask b ON a.ProjectTaskId = b.ProjectTaskId
                                WHERE b.ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案任務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectTask
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案檔案
                        #region //將所有專案檔案改為刪除狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                INNER JOIN PMIS.ProjectFile b ON a.FileId = b.FileId
                                WHERE b.ProjectId = @ProjectId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectFile
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案成員權限
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.ProjectUserAuthority a
                                INNER JOIN PMIS.ProjectUser b ON a.ProjectUserId = b.ProjectUserId
                                WHERE b.ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除專案成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectUser
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.Project
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("ProjectId", ProjectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //MAMO團隊刪除 (停用)
                        //if (!string.IsNullOrWhiteSpace(companyNo) && teamId > 0)
                        //{
                        //    var mamoResult = mamoHelper.DeleteTeams(companyNo, CurrentUser, teamId);
                        //    JObject mamoResultJson = JObject.Parse(mamoResult);
                        //}
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

        #region //DeleteProjectUser -- 專案成員資料刪除 -- Ben Ma 2022.08.22
        public string DeleteProjectUser(int ProjectUserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProjectId
                                , ISNULL(a.TeamId, -1) TeamId
                                , c.CompanyNo, d.AgentSort, u.UserNo
                                FROM PMIS.Project a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                INNER JOIN PMIS.ProjectUser d ON d.ProjectId = a.ProjectId
                                INNER JOIN BAS.[User] u ON u.UserId = d.UserId
                                WHERE d.ProjectUserId = @ProjectUserId";
                        dynamicParameters.Add("ProjectUserId", ProjectUserId);

                        var resultUsers = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUsers.Count() <= 0) throw new SystemException("【專案成員】資料錯誤!");

                        int agentSort = 0;
                        int projectId = -1;
                        int teamId = -1;
                        string companyNo = string.Empty;
                        foreach (var item in resultUsers)
                        {
                            agentSort = Convert.ToInt32(item.AgentSort);
                            projectId = item.ProjectId;
                            teamId = item.TeamId;
                            companyNo = item.CompanyNo;
                        }

                        if (agentSort == 1) throw new SystemException("專案負責人無法刪除!");
                        #endregion

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectUserAuthority
                                WHERE ProjectUserId = @ProjectUserId";
                        dynamicParameters.Add("ProjectUserId", ProjectUserId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectUser
                                WHERE ProjectUserId = @ProjectUserId";
                        dynamicParameters.Add("ProjectUserId", ProjectUserId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新取得專案成員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT a.ProjectUserId, a.AgentSort
                                FROM PMIS.ProjectUser a
                                WHERE a.ProjectId = @ProjectId
                                ORDER BY a.AgentSort";
                        dynamicParameters.Add("ProjectId", projectId);
                        resultUsers = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PMIS.ProjectUser SET
                                AgentSort = AgentSort * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectId = @ProjectId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ProjectId", projectId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int i = 1;
                        foreach (var item in resultUsers)
                        {
                            #region //更新順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PMIS.ProjectUser SET
                                    AgentSort = @AgentSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectId = @ProjectId
                                    AND ProjectUserId = @ProjectUserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AgentSort = i,
                                    ProjectId = projectId,
                                    item.ProjectUserId,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            i++;
                        }

                        #region //MAMO團隊成員刪除 (停用)
                        //if (!string.IsNullOrWhiteSpace(companyNo) && teamId > 0)
                        //{
                        //    var mamoResult = mamoHelper.DeleteTeamMembers(companyNo, CurrentUser, teamId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                        //    JObject mamoResultJson = JObject.Parse(mamoResult);
                        //}
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

        #region //DeleteProjectFile -- 專案檔案刪除 -- Ben Ma 2022.09.26
        public string DeleteProjectFile(int ProjectFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectFile
                                WHERE ProjectFileId = @ProjectFileId";
                        dynamicParameters.Add("ProjectFileId", ProjectFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案檔案資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                INNER JOIN PMIS.ProjectFile b ON a.FileId = b.FileId
                                WHERE b.ProjectFileId = @ProjectFileId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("ProjectFileId", ProjectFileId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectFile
                                WHERE ProjectFileId = @ProjectFileId";
                        dynamicParameters.Add("ProjectFileId", ProjectFileId);

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

        #region //DeleteProjectTask -- 專案任務資料刪除 -- Ben Ma 2022.09.19
        public string DeleteProjectTask(int ProjectTaskId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string projectStatus = string.Empty, companyNo = string.Empty;

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProjectName, a.ProjectStatus
                                , c.CompanyNo, ISNULL(t.ParentTaskId, -1) ParentTaskId, t.TaskStatus, t.TaskName
                                FROM PMIS.Project a
                                INNER JOIN PMIS.ProjectTask t ON t.ProjectId = a.ProjectId
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                                WHERE t.ProjectTaskId = @ProjectTaskId";
                        dynamicParameters.Add("ProjectTaskId", ProjectTaskId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.TaskStatus != "B") throw new SystemException(string.Format("【{0}】非待處理狀態，任務無法刪除!", item.TaskName));
                            projectStatus = item.ProjectStatus;
                            companyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //取得父任務結構
                        dynamicParameters = new DynamicParameters();
                        sql = GetParentTaskStrucSQL();
                        dynamicParameters.Add("ProjectTaskId", ProjectTaskId);
                        var resultParentTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得子任務結構
                        dynamicParameters = new DynamicParameters();
                        sql = GetSubTaskStrucSQL();
                        dynamicParameters.Add("ProjectTaskId", ProjectTaskId);
                        var resultSubTasks = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //刪除關聯table
                        if (resultSubTasks.Count() > 0)
                        {
                            if (resultSubTasks.Any(a => a.TaskStatus == "I")) throw new SystemException(string.Format("子任務【進行中】，任務無法刪除<br/>{0}", string.Join("、", resultSubTasks.Select(s => s.TaskName))));

                            var ProjectTaskIds = resultSubTasks.Select(s => s.ProjectTaskId).ToArray();

                            #region //刪除結構任務連結
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM PMIS.ProjectTaskLink a
                                    WHERE a.SourceTaskId IN @ProjectTaskIds 
                                        OR a.TargetTaskId IN @ProjectTaskIds";
                            dynamicParameters.Add("ProjectTaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除結構任務成員權限
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM PMIS.TaskUserAuthority a
                                    INNER JOIN PMIS.TaskUser b ON a.TaskUserId = b.TaskUserId
                                    WHERE b.TaskId IN @TaskIds
                                    AND b.TaskType = 'Project'";
                            dynamicParameters.Add("TaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除結構任務成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.TaskUser
                                    WHERE TaskId IN @TaskIds
                                    AND TaskType = 'Project'";
                            dynamicParameters.Add("TaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新結構任務刪除狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.DeleteStatus = @DeleteStatus
                                    FROM BAS.[File] a
                                    INNER JOIN PMIS.ProjectTaskFile b ON a.FileId = b.FileId
                                    WHERE b.ProjectTaskId IN @ProjectTaskIds";
                            dynamicParameters.Add("DeleteStatus", "Y");
                            dynamicParameters.Add("ProjectTaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除結構任務檔案資料表
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM PMIS.ProjectTaskFile a
                                    WHERE a.ProjectTaskId IN @ProjectTaskIds";
                            dynamicParameters.Add("ProjectTaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除結構任務
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PMIS.ProjectTask
                                    WHERE ProjectTaskId IN @ProjectTaskIds";
                            dynamicParameters.Add("ProjectTaskIds", ProjectTaskIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //MAMO頻道刪除 (停用)
                            //resultSubTasks.ForEach(f =>
                            //{
                            //    if (!string.IsNullOrWhiteSpace(companyNo) && f.ChannelId > 0)
                            //    {
                            //        var mamoResult = mamoHelper.DeleteChannels(companyNo, CurrentUser, f.ChannelId);
                            //        JObject mamoResultJson = JObject.Parse(mamoResult);
                            //    }
                            //});
                            #endregion
                        }
                        #endregion

                        if (projectStatus == "I")
                        {
                            foreach (var ParentTaskId in resultParentTasks.Where(w => w.ParentTaskId > 0).Select(s => s.ParentTaskId).Distinct())
                            {
                                #region //取得該層所有任務
                                sql = @"SELECT a.ProjectTaskId, a.TaskStatus, a.TaskName
                                        , FORMAT(ISNULL(a.ActualStart, a.EstimateStart), 'yyyy-MM-dd HH:mm') StartDate
                                        , FORMAT(ISNULL(a.ActualEnd, a.EstimateEnd), 'yyyy-MM-dd HH:mm') EndDate
                                        FROM PMIS.ProjectTask a
                                        WHERE a.ParentTaskId = @ParentTaskId";
                                var resultSubTask = sqlConnection.Query(sql, new { ParentTaskId });
                                #endregion

                                if (!resultSubTask.Any())
                                {
                                    sql = @"SELECT TOP 1 ISNULL(FORMAT(aa.EventDate, 'yyyy-MM-dd HH:mm'), '') EventDate
                                            FROM PMIS.TaskReply aa
                                            WHERE aa.TaskId = @ParentTaskId
                                            AND aa.ReplyType = 'Project' 
                                            AND aa.ReplyStatus = 'C'
                                            ORDER BY aa.ReplyId DESC";
                                    var resultParentTask = sqlConnection.QueryFirstOrDefault(sql, new { ParentTaskId });

                                    if (resultParentTask != null)
                                    {
                                        if (DateTime.TryParse(resultParentTask.EventDate, out DateTime maxActualEnd))
                                        {
                                            #region //自動完成上層任務
                                            sql = @"UPDATE PMIS.ProjectTask SET
                                                    ActualEnd = @ActualEnd,
                                                    ActualDuration = DATEDIFF(MINUTE, ActualStart, @ActualEnd),
                                                    TaskStatus = 'C',
                                                    LastModifiedDate = @LastModifiedDate,
                                                    LastModifiedBy = @LastModifiedBy
                                                    WHERE ProjectTaskId = @ParentTaskId
                                                    AND TaskStatus <> 'C'";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    ParentTaskId,
                                                    ActualEnd = maxActualEnd,
                                                    LastModifiedDate,
                                                    LastModifiedBy
                                                });
                                            #endregion
                                        }
                                    }
                                }
                                else if (resultSubTask.Where(w => w.TaskStatus == "C").Count() == resultSubTask.Count())
                                {
                                    bool hasParentActualStart = DateTime.TryParse(resultParentTasks.FirstOrDefault(f => f.ProjectTaskId == ParentTaskId).ActualStart.ToString(), out DateTime minActualStart);
                                    bool hasMaxActualEnd = DateTime.TryParse(resultSubTask.Max(m => m.EndDate), out DateTime maxActualEnd);

                                    #region //自動完成上層任務
                                    sql = @"UPDATE PMIS.ProjectTask SET
                                            ActualEnd = @ActualEnd,
                                            ActualDuration = @ActualDuration,
                                            TaskStatus = 'C',
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ProjectTaskId = @ParentTaskId
                                            AND TaskStatus <> 'C'";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            ParentTaskId,
                                            ActualEnd = hasMaxActualEnd ? maxActualEnd.ToString("yyyy-MM-dd HH:mm") : null,
                                            ActualDuration = hasMaxActualEnd && hasParentActualStart ? (maxActualEnd - minActualStart).TotalMinutes : (double?)null,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion
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

        #region //DeleteProjectTaskLink -- 專案任務連結資料刪除 -- Ben Ma 2022.09.27
        public string DeleteProjectTaskLink(int ProjectTaskLinkId, int SourceTaskId, int TargetTaskId, string LinkType)
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
                                FROM PMIS.ProjectTaskLink
                                WHERE ProjectTaskLinkId = @ProjectTaskLinkId
                                AND SourceTaskId = @SourceTaskId
                                AND TargetTaskId = @TargetTaskId
                                AND LinkType = @LinkType";
                        dynamicParameters.Add("ProjectTaskLinkId", ProjectTaskLinkId);
                        dynamicParameters.Add("SourceTaskId", SourceTaskId);
                        dynamicParameters.Add("TargetTaskId", TargetTaskId);
                        dynamicParameters.Add("LinkType", LinkType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案任務連結資料錯誤!");
                        #endregion

                        #region //判斷專案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.ProjectStatus
                                FROM PMIS.ProjectTask a
                                INNER JOIN PMIS.Project b ON a.ProjectId = b.ProjectId
                                WHERE a.ProjectTaskId = @ProjectTaskId";
                        dynamicParameters.Add("ProjectTaskId", SourceTaskId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【專案】資料錯誤!");

                        string projectStatus = "";
                        foreach (var item in result2)
                        {
                            projectStatus = item.ProjectStatus;
                        }

                        //if (projectStatus != "P") throw new SystemException("【專案】非計畫中，任務連結無法刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectTaskLink
                                WHERE ProjectTaskLinkId = @ProjectTaskLinkId";
                        dynamicParameters.Add("ProjectTaskLinkId", ProjectTaskLinkId);

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

        #region //DeleteTaskUser -- 任務成員刪除-New -- Chia Yuan 2023.12.06
        public string DeleteTaskUser(int TaskId, int TaskUserId, string TaskType)
        {
            try
            {
                if (!Regex.IsMatch(TaskType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("任務類型錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務成員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT a.TaskId, a.TaskUserId, a.TaskType, u.UserNo
                                FROM PMIS.TaskUser a
                                INNER JOIN BAS.[User] u ON u.UserId = a.UserId
                                WHERE a.TaskId = @TaskId
                                AND a.TaskType = @TaskType";
                        dynamicParameters.Add("TaskId", TaskId);
                        dynamicParameters.Add("TaskType", TaskType);
                        var resultUsers = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUsers.Count() <= 0) throw new SystemException("【任務成員】資料錯誤!");
                        if (!resultUsers.Any(a => a.TaskUserId == TaskUserId)) throw new SystemException("【任務成員】資料錯誤!");
                        int taskId = -1;
                        foreach (var item in resultUsers)
                        {
                            taskId = Convert.ToInt32(item.TaskId);
                        }
                        #endregion

                        int channelId = -1;
                        string companyNo = string.Empty;

                        #region //取得專案任務資料
                        switch (TaskType)
                        {
                            case "Project":
                                #region //判斷專案任務資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 c.CompanyNo, ISNULL(a.ChannelId, -1) ChannelId, a.TaskStatus
                                        FROM PMIS.ProjectTask a
                                        INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                        LEFT JOIN BAS.Department d ON d.DepartmentId = b.DepartmentId
                                        LEFT JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                        WHERE a.ProjectTaskId = @TaskId";
                                dynamicParameters.Add("TaskId", taskId);
                                var result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");
                                if (result1.Any(a => a.TaskStatus == "C")) throw new SystemException("【專案任務已完成】不得刪除成員!");
                                foreach (var item in result1)
                                {
                                    channelId = item.ChannelId;
                                    companyNo = item.CompanyNo;
                                }
                                #endregion
                                break;
                            case "Personal":
                                if (resultUsers.Count() == 1) throw new SystemException("【個人任務】成員至少一位!");

                                #region //判斷個人任務資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 c.CompanyNo, a.TaskStatus 
                                        FROM PMIS.PersonalTask a
                                        INNER JOIN BAS.[User] u ON u.UserId = a.UserId
                                        INNER JOIN BAS.Department d ON d.DepartmentId = u.DepartmentId
                                        INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                        WHERE a.PersonalTaskId = @TaskId";
                                dynamicParameters.Add("TaskId", taskId);
                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() <= 0) throw new SystemException("【個人任務】資料錯誤!");
                                if (result2.Any(a => a.TaskStatus == "C")) throw new SystemException("【個人任務已完成】不得刪除成員!");
                                foreach (var item in result2)
                                {                                    
                                    companyNo = item.CompanyNo;
                                }
                                #endregion
                                break;
                            default:
                                break;
                        }
                        #endregion

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TaskUserAuthority
                                WHERE TaskUserId = @TaskUserId";
                        dynamicParameters.Add("TaskUserId", TaskUserId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TaskUser
                                WHERE TaskUserId = @TaskUserId
                                AND TaskType = @TaskType";
                        dynamicParameters.Add("TaskUserId", TaskUserId);
                        dynamicParameters.Add("TaskType", TaskType);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //MAMO頻道成員刪除 (停用)
                        //if (!string.IsNullOrWhiteSpace(companyNo) && TaskType == "Project" && channelId > 0)
                        //{
                        //    var mamoResult = mamoHelper.DeleteChannelMembers(companyNo, CurrentUser, channelId, resultUsers.Select(s => { return (string)s.UserNo; }).Distinct().ToList());
                        //    JObject mamoResultJson = JObject.Parse(mamoResult);
                        //}
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

        #region //DeleteProjectTaskFile -- 專案任務檔案刪除 -- Ben Ma 2022.09.23
        public string DeleteProjectTaskFile(int ProjectTaskFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.ProjectTaskFile
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        dynamicParameters.Add("ProjectTaskFileId", ProjectTaskFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                INNER JOIN PMIS.ProjectTaskFile b ON a.FileId = b.FileId
                                WHERE b.ProjectTaskFileId = @ProjectTaskFileId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("ProjectTaskFileId", ProjectTaskFileId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.ProjectTaskFile
                                WHERE ProjectTaskFileId = @ProjectTaskFileId";
                        dynamicParameters.Add("ProjectTaskFileId", ProjectTaskFileId);

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

        #region //DeletePersonalTask -- 專案任務資料刪除 -- Chia Yuan
        public string DeletePersonalTask(int PersonalTaskId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TaskStatus, a.TaskName
                                FROM PMIS.PersonalTask a
                                WHERE a.PersonalTaskId = @PersonalTaskId";
                        dynamicParameters.Add("PersonalTaskId", PersonalTaskId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【專案任務】資料錯誤!");
                        //if (result.TaskStatus != "B") throw new SystemException(string.Format("【{0}】非待處理狀態，任務無法刪除!", result.TaskName));
                        #endregion

                        #region //刪除任務成員權限
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TaskUserAuthority a
                                INNER JOIN PMIS.TaskUser b ON a.TaskUserId = b.TaskUserId
                                WHERE b.TaskId = @TaskId
                                AND b.TaskType = 'Personal'";
                        dynamicParameters.Add("TaskId", PersonalTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除任務成員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.TaskUser
                                WHERE TaskId = @TaskId
                                AND TaskType = 'Personal'";
                        dynamicParameters.Add("TaskId", PersonalTaskId);                        
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //取得個人任務附件資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PersonalTaskFileId, a.FileId
                                FROM PMIS.PersonalTaskFile a
                                WHERE a.PersonalTaskId = @PersonalTaskId";
                        dynamicParameters.Add("PersonalTaskId", PersonalTaskId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        int fileId = -1;
                        int personalTaskFileId = -1;
                        foreach (var item in result)
                        {
                            fileId = item.FileId;
                            personalTaskFileId = item.PersonalTaskFileId;
                        }
                        #endregion

                        #region //刪除個人任務附件資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.PersonalTaskFile
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        dynamicParameters.Add("PersonalTaskFileId", personalTaskFileId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除個人任務檔案
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("FileId", fileId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //任務回報內容刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM PMIS.TaskReply a
                                WHERE a.TaskId = @TaskId
                                AND a.ReplyType = 'Personal'";
                        dynamicParameters.Add("TaskId", PersonalTaskId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除個人任務
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.PersonalTask
                                WHERE PersonalTaskId = @PersonalTaskId";
                        dynamicParameters.Add("PersonalTaskId", PersonalTaskId);
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

        #region //DeletePersonalTaskFile -- 個人任務檔案刪除-New -- Chia Yuan 2023.12.14
        public string DeletePersonalTaskFile(int PersonalTaskFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷專案任務檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PMIS.PersonalTaskFile
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        dynamicParameters.Add("PersonalTaskFileId", PersonalTaskFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("專案任務檔案資料錯誤!");
                        #endregion

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DeleteStatus = @DeleteStatus
                                FROM BAS.[File] a
                                INNER JOIN PMIS.PersonalTaskFile b ON a.FileId = b.FileId
                                WHERE b.PersonalTaskFileId = @PersonalTaskFileId";
                        dynamicParameters.Add("DeleteStatus", "Y");
                        dynamicParameters.Add("PersonalTaskFileId", PersonalTaskFileId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PMIS.PersonalTaskFile
                                WHERE PersonalTaskFileId = @PersonalTaskFileId";
                        dynamicParameters.Add("PersonalTaskFileId", PersonalTaskFileId);

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

        #region //Api
        #region //ProjectFilter -- API查詢過濾器 -- Chia Yuan 2024.06.19
        private List<dynamic> ProjectFilter(int ProjectId, int CompanyId, int DepartmentId, string TaskUserNos, string ProjectStatus
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

                if (!string.IsNullOrWhiteSpace(ProjectStatus))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectStatus", @" AND a.ProjectStatus IN @ProjectStatus", ProjectStatus.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(TaskUserNos))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "TaskUserNos"
                        , @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM PMIS.ProjectTask ba
                            INNER JOIN PMIS.TaskUser bb ON bb.TaskId = ba.ProjectTaskId AND bb.TaskType = 'Project'
                            INNER JOIN BAS.[User] bu1 ON bu1.UserId = bb.UserId
                            WHERE ba.ProjectId = a.ProjectId
		                    AND bu1.UserNo IN @TaskUserNos
	                    )", TaskUserNos.Split(','));
                }

                if (CompanyId > 0 || DepartmentId > 0)
                {
                    sql += @" AND (EXISTS (
		                    SELECT TOP 1 1
		                    FROM PMIS.ProjectUser da
                            INNER JOIN BAS.[User] db ON db.UserId = da.UserId
                            INNER JOIN BAS.Department dc ON dc.DepartmentId = db.DepartmentId
		                    WHERE da.ProjectId = a.ProjectId AND da.AgentSort = 1";
                    if (CompanyId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", " AND dc.CompanyId = @CompanyId", CompanyId);
                    }
                    if (DepartmentId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", " AND dc.DepartmentId = @DepartmentId", DepartmentId);
                    }
                    sql += @") OR 
                            EXISTS (
                            SELECT TOP 1 1 
                            FROM BAS.Department dd 
                            WHERE dd.DepartmentId = a.DepartmentId";
                    if (CompanyId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", " AND dd.CompanyId = @CompanyId", CompanyId);
                    }
                    if (DepartmentId > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", " AND dd.DepartmentId = @DepartmentId", DepartmentId);
                    }
                    sql += "))";
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

        #region //GetTaskRecursionSqlString -- 遞迴結構語法(串接用) -- Chia Yuan 2024.06.19
        public string GetTaskRecursionSqlString(ref DynamicParameters dynamicParameters, List<dynamic> ProjectIds)
        {
            var recursionString = "";
            recursionString = @"
                DECLARE @rowsAdded int
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
	            AND ProjectId IN @ProjectIds";

            dynamicParameters.Add("ProjectIds", ProjectIds);

            recursionString += @"
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
                END";

            return recursionString;
        }
        #endregion

        #region //GetSingleTaskRecursionSqlString -- 單一任務遞迴結構語法(串接用) -- Chia Yuan 2024.06.20
        public string GetSingleTaskRecursionSqlString(ref DynamicParameters dynamicParameters, int ProjectTaskId)
        {
            var recursionString = "";
            recursionString = @"
                DECLARE @rowsAdded int
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
                WHERE ProjectTaskId = @ProjectTaskId";

            dynamicParameters.Add("ProjectTaskId", ProjectTaskId);

            recursionString += @"
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
                        INNER JOIN PMIS.ProjectTask b ON a.ParentTaskId= b.ProjectTaskId 
                        WHERE a.processed = 1
                    SET @rowsAdded = @@rowcount
                    UPDATE @projectTask SET processed = 2 WHERE processed = 1
                END";

            return recursionString;
        }
        #endregion

        #region //GetReplyTaskApi -- 取得使用者可回報任務列表 -- Chia Yuan 2024.06.19
        public string GetReplyTaskApi(int ProjectId, int CompanyId, int DepartmentId, string TaskUserNos
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //if (ProjectId <= 0 && CompanyId <= 0 && DepartmentId <= 0 && string.IsNullOrWhiteSpace(TaskUserNos))
                //    throw new SystemException("【參數條件】資料錯誤!");
                
                var projects = ProjectFilter(ProjectId, CompanyId, DepartmentId, TaskUserNos, "I"
                    , OrderBy, PageIndex, PageSize);

                var projectIds = projects.Select(s => s.ProjectId).ToList();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = GetTaskRecursionSqlString(ref dynamicParameters, projectIds);
                    sql += @"
                        SELECT
                        a.ProjectTaskId
                        , b.ProjectNo, b.ProjectName, b.ProjectDesc, b.CustomerRemark, b.ProjectRemark
                        , b.ProjectStatus, s1.StatusName ProjectStatusName
                        , b.ProjectType, t1.TypeName ProjectTypeName
                        , b.ProjectAttribute, t2.TypeName ProjectAttributeName
                        , b.WorkTimeStatus, b.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                        , a.TaskName, a.TaskDesc
                        , a.TaskStatus, s2.StatusName TaskStatusName
                        , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') ProjectStart
                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') ProjectEnd
                        , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration) ProjectDuration
                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') TaskStart
                        , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') TaskEnd
                        , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration) TaskDuration
                        , ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                        , ISNULL(e.DepartmentNo, '') DepartmentNo, ISNULL(e.DepartmentName, '') DepartmentName
                        , ISNULL(e.CompanyNo, '') CompanyNo, ISNULL(e.CompanyName, '') CompanyName, ISNULL(e.LogoIcon, -1) LogoIcon
                        , ISNULL(SUBSTRING(d.TaskUsers, 0, LEN(d.TaskUsers)), '') TaskUsers
                        , u.UserNo ProjectMasterNo, u.UserName ProjectMasterName, u.Gender ProjectMasterGender, u.UserName + '(' + u.UserNo + ')' ProjectMaster
                        , 'Project' TaskType
                        , (SELECT cb.UserNo, cb.UserName, cb.Gender
	                        FROM PMIS.ProjectUser ca
	                        INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
	                        WHERE ca.ProjectId = b.ProjectId
	                        AND cb.[Status] = 'A'
	                        ORDER BY ca.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) ProjectUser
                        , (SELECT db.UserNo, db.UserName, db.Gender
	                        FROM PMIS.TaskUser da
	                        INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                        WHERE da.TaskType = 'Project'
	                        AND da.TaskId = a.ProjectTaskId
	                        AND db.[Status] = 'A'
	                        ORDER BY da.AgentSort
	                        FOR JSON PATH, ROOT('data')
                        ) TaskUser
                        , (SELECT fa.FileId, fb.[FileName], fb.FileExtension
	                        FROM PMIS.ProjectTaskFile fa
	                        INNER JOIN BAS.[File] fb ON fb.FileId = fa.FileId
	                        WHERE fa.ProjectTaskId = a.ProjectTaskId
	                        ORDER BY fa.ProjectTaskFileId
	                        FOR JSON PATH, ROOT('data')
                        ) TaskFile
                        , (SELECT fb.UserNo, fb.UserName, fb.Gender
	                        , ISNULL(FORMAT(fa.ReplyDate, 'yyyy-MM-dd HH:mm'), '') ReplyDate, fa.ReplyContent, fa.DeferredDuration, fa.ReplyStatus, fc.StatusName ReplyStatusName
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
                        FROM @projectTask x
                        INNER JOIN PMIS.ProjectTask a ON a.ProjectTaskId = x.ProjectTaskId
                        INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                        INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = b.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                        INNER JOIN BAS.[Type] t2 ON t2.TypeNo = b.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                        OUTER APPLY (
	                        SELECT ea.DepartmentNo, ea.DepartmentName
	                        , eb.CompanyNo, eb.CompanyName , eb.LogoIcon
	                        FROM BAS.[Department] ea
	                        INNER JOIN BAS.[Company] eb ON eb.CompanyId = ea.CompanyId
	                        WHERE ea.DepartmentId = b.DepartmentId
                        ) e
                        OUTER APPLY (
                            SELECT COUNT(1) SubTask
                            FROM PMIS.ProjectTask ca
                            WHERE ca.ParentTaskId = a.ProjectTaskId
                        ) c
                        OUTER APPLY (
                            SELECT (
                                SELECT db.UserName + ','
                                FROM PMIS.TaskUser da
                                INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                WHERE da.TaskId = a.ProjectTaskId
                                AND da.TaskType = 'Project'
                                ORDER BY da.AgentSort
                                FOR XML PATH('')
                            ) TaskUsers
                        ) d
                        OUTER APPLY (
	                        SELECT TOP 1 ub.UserId, ub.UserNo, ub.UserName, ub.Gender
	                        FROM PMIS.ProjectUser ua 
	                        INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                        WHERE ua.ProjectId = a.ProjectId 
	                        AND ua.AgentSort = 1
                        ) u
                        OUTER APPLY (SELECT ba.SourceTaskId, ba.TargetTaskId, bb.TaskStatus SourceStatus FROM PMIS.ProjectTaskLink ba INNER JOIN PMIS.ProjectTask bb ON bb.ProjectTaskId = ba.SourceTaskId WHERE ba.TargetTaskId = a.ProjectTaskId) link
                        OUTER APPLY (SELECT ca.SourceTaskId, ca.TargetTaskId, cb.TaskStatus ParentSourceStatus FROM PMIS.ProjectTaskLink ca INNER JOIN PMIS.ProjectTask cb ON cb.ProjectTaskId = ca.SourceTaskId WHERE ca.TargetTaskId = a.ParentTaskId) pLink
                        WHERE NOT EXISTS (SELECT TOP 1 1 FROM @projectTask aa WHERE aa.ParentTaskId = a.ProjectTaskId) --排除父層任務
                        AND ISNULL(link.SourceStatus, '') IN ('', 'C') --前置任務需完成
                        AND ISNULL(pLink.ParentSourceStatus, '') IN ('', 'C') --父層前置任務需完成
                        AND a.TaskStatus IN ('B', 'I') --待開始、進行中
                        AND (EXISTS (
		                        SELECT TOP 1 ISNULL(da.ProjectTaskLinkId, 1)
		                        FROM PMIS.ProjectTaskLink da 
		                        LEFT JOIN PMIS.ProjectTask db ON db.ProjectTaskId = da.TargetTaskId
		                        LEFT JOIN PMIS.ProjectTask dc ON dc.ProjectTaskId = da.SourceTaskId
		                        WHERE da.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')) AND ISNULL(dc.TaskStatus, '') IN ('', 'C')
	                        ) 
	                        OR NOT EXISTS (SELECT TOP 1 1 FROM PMIS.ProjectTaskLink ea WHERE ea.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')))
                        )";

                    if (!string.IsNullOrWhiteSpace(TaskUserNos))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, 
                            "TaskUserNos",
                            @" AND EXISTS (
                                SELECT TOP 1 1 
                                FROM PMIS.TaskUser fa 
                                INNER JOIN BAS.[User] fb ON fb.UserId = fa.UserId
                                WHERE fa.TaskId = a.ProjectTaskId 
                                AND fa.TaskType = 'Project'
                                AND fb.UserNo IN @TaskUserNos)"
                            , TaskUserNos.Split(','));
                    }

                    sql += " ORDER BY ISNULL(ISNULL(a.ActualStart, a.EstimateStart), a.PlannedStart)";

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

        #region //GetProjectApi -- 取得專案資料 -- Chia Yuan 2024.06.17
        public string GetProjectApi(int ProjectId, int ProjectTaskId, string ProjectUserNos
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (ProjectId <= 0 && ProjectTaskId <= 0 && string.IsNullOrWhiteSpace(ProjectUserNos))
                    throw new SystemException("【參數條件】資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT
                            b.ProjectId
                            , b.ProjectNo, b.ProjectName, b.ProjectDesc, b.CustomerRemark, b.ProjectRemark
                            , b.ProjectStatus, s1.StatusName ProjectStatusName
                            , ISNULL(h.MamoTeamId, '') TeamId
                            , ISNULL(h.TeamNo, '') TeamNo
                            , ISNULL(h.TeamName, '') TeamName
                            , b.ProjectType, t1.TypeName ProjectTypeName
                            , b.ProjectAttribute, t2.TypeName ProjectAttributeName
                            , b.WorkTimeStatus, b.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                            , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualStart, b.EstimateStart), b.PlannedStart), 'yyyy-MM-dd HH:mm'), '') ProjectStart
                            , ISNULL(FORMAT(ISNULL(ISNULL(b.ActualEnd, b.EstimateEnd), b.PlannedEnd), 'yyyy-MM-dd HH:mm'), '') ProjectEnd
                            , ISNULL(ISNULL(b.ActualDuration, b.EstimateDuration), b.PlannedDuration) ProjectDuration
                            , ISNULL(e.DepartmentNo, '') DepartmentNo, ISNULL(e.DepartmentName, '') DepartmentName
                            , ISNULL(e.CompanyNo, '') CompanyNo, ISNULL(e.CompanyName, '') CompanyName, ISNULL(e.LogoIcon, -1) LogoIcon
                            , c.SubTask
                            , ISNULL(SUBSTRING(d.ProjectUsers, 0, LEN(d.ProjectUsers)), '') ProjectUsers
                            , u.UserNo ProjectMasterNo, u.UserName ProjectMasterName, u.Gender ProjectMasterGender, u.UserName + '(' + u.UserNo + ')' ProjectMaster
                            , (SELECT cb.UserNo, cb.UserName, cb.Gender
	                            FROM PMIS.ProjectUser ca
	                            INNER JOIN BAS.[User] cb ON cb.UserId = ca.UserId
	                            WHERE ca.ProjectId = b.ProjectId
	                            AND cb.[Status] = 'A'
	                            ORDER BY ca.AgentSort
	                            FOR JSON PATH, ROOT('data')
                            ) ProjectUser
                            , (SELECT fa.FileId, fb.[FileName], fb.FileExtension
	                            FROM PMIS.ProjectFile fa
	                            INNER JOIN BAS.[File] fb ON fb.FileId = fa.FileId
	                            WHERE fa.ProjectId = b.ProjectId
	                            ORDER BY fa.ProjectFileId
	                            FOR JSON PATH, ROOT('data')
                            ) ProjectFile
                            , ISNULL((
	                            SELECT COUNT(DISTINCT eb.UserId)
	                            FROM PMIS.ProjectTask ea
	                            INNER JOIN PMIS.TaskUser eb ON eb.TaskId = ea.ProjectTaskId AND eb.TaskType = 'Project'
	                            WHERE ea.ProjectId = b.ProjectId), 0
                            ) TaskUserCount
                            FROM PMIS.Project b
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                            INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                            INNER JOIN BAS.[Type] t1 ON t1.TypeNo = b.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                            INNER JOIN BAS.[Type] t2 ON t2.TypeNo = b.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                            OUTER APPLY (
	                            SELECT ea.DepartmentNo, ea.DepartmentName
	                            , eb.CompanyNo, eb.CompanyName , eb.LogoIcon
	                            FROM BAS.[Department] ea
	                            INNER JOIN BAS.[Company] eb ON eb.CompanyId = ea.CompanyId
	                            WHERE ea.DepartmentId = b.DepartmentId
                            ) e
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ProjectId = b.ProjectId
                            ) c
                            OUTER APPLY (
                                SELECT (
                                    SELECT db.UserName + ','
                                    FROM PMIS.ProjectUser da
                                    INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                    WHERE da.ProjectId = b.ProjectId
                                    ORDER BY da.AgentSort
                                    FOR XML PATH('')
                                ) ProjectUsers
                            ) d
                            OUTER APPLY (
	                            SELECT TOP 1 ub.UserId, ub.UserNo, ub.UserName, ub.Gender
	                            FROM PMIS.ProjectUser ua 
	                            INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                            WHERE ua.ProjectId = b.ProjectId 
	                            AND ua.AgentSort = 1
                            ) u
                            OUTER APPLY (
	                            SELECT TOP 1 ha.MamoTeamId, ha.TeamNo, ha.TeamName
	                            FROM MAMO.Teams ha
	                            WHERE ha.TeamId = b.TeamId
                            ) h
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectId", " AND b.ProjectId = @ProjectId", ProjectId);
                    if (!string.IsNullOrWhiteSpace(ProjectUserNos))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                            , "ProjectUserNos"
                            , @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM PMIS.ProjectUser aa
                            INNER JOIN BAS.[User] bu1 ON bu1.UserId = aa.UserId
                            WHERE aa.ProjectId = b.ProjectId
		                    AND bu1.UserNo IN @ProjectUserNos
	                    )", ProjectUserNos.Split(','));
                    }
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "ProjectTaskId"
                        , @" AND EXISTS (
                                SELECT TOP 1 1
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ProjectId = b.ProjectId
		                        AND ca.ProjectTaskId = @ProjectTaskId)"
                        , ProjectTaskId);
                    sql += " ORDER BY b.ProjectId";

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

        #region //GetProjectTaskApi -- 取得專案任務資料 -- Chia Yuan 2024.06.17
        public string GetProjectTaskApi(int ProjectTaskId, string TaskUserNos
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (ProjectTaskId <= 0 && string.IsNullOrWhiteSpace(TaskUserNos))
                    throw new SystemException("【參數條件】資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得專案資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ProjectId
                            FROM PMIS.ProjectTask a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", " AND a.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                    if (!string.IsNullOrWhiteSpace(TaskUserNos))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters,
                            "TaskUserNos",
                            @" AND EXISTS (
                            SELECT TOP 1 1 
                            FROM PMIS.TaskUser aa
                            INNER JOIN BAS.[User] ab ON ab.UserId = aa.UserId 
                            WHERE aa.TaskType = 'Project'
                            AND aa.TaskId = a.ProjectTaskId
                            AND ab.UserNo IN @TaskUserNos)",
                            TaskUserNos.Split(','));
                    }
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");
                    #endregion

                    #region //取得專案任務資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetTaskRecursionSqlString(ref dynamicParameters, result.Select(s => s.ProjectId).Distinct().ToList());
                    sql += @"
                            SELECT a.ProjectTaskId
                            , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                            , b.ProjectId, b.ProjectNo, b.ProjectName, b.ProjectDesc, b.CustomerRemark, b.ProjectRemark
                            , b.ProjectStatus, s1.StatusName ProjectStatusName
                            , b.WorkTimeStatus, b.WorkTimeInterval, s3.StatusName WorkTimeStatusName
                            , a.TaskName, a.TaskDesc
                            , a.TaskStatus, s2.StatusName TaskStatusName
                            , ISNULL(h.MamoChannelId, '') ChannelId
                            , ISNULL(h.ChannelNo, '') ChannelNo
                            , ISNULL(h.ChannelName, '') ChannelName
                            , ISNULL(FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm'), '') PlannedStart
                            , ISNULL(FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm'), '') PlannedEnd
                            , ISNULL(a.PlannedDuration, 0) PlannedDuration
                            , ISNULL(FORMAT(a.EstimateStart, 'yyyy-MM-dd HH:mm'), '') EstimateStart
                            , ISNULL(FORMAT(a.EstimateEnd, 'yyyy-MM-dd HH:mm'), '') EstimateEnd
                            , ISNULL(a.EstimateDuration, 0) EstimateDuration
                            , ISNULL(FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm'), '') ActualStart
                            , ISNULL(FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm'), '') ActualEnd
                            , ISNULL(a.ActualDuration, 0) ActualDuration
                            , ISNULL(a.ReplyFrequency, 0) ReplyFrequency
                            , c.SubTask
                            , 'Project' TaskType
                            , ISNULL(SUBSTRING(d.TaskUsers, 0, LEN(d.TaskUsers)), '') TaskUsers
                            , (SELECT db.UserNo, db.UserName, db.Gender
	                            FROM PMIS.TaskUser da
	                            INNER JOIN BAS.[User] db ON db.UserId = da.UserId
	                            WHERE da.TaskType = 'Project'
	                            AND da.TaskId = a.ProjectTaskId
	                            AND db.[Status] = 'A'
	                            ORDER BY da.AgentSort
	                            FOR JSON PATH, ROOT('data')
                            ) TaskUser
                            , (SELECT fa.FileId, fb.[FileName], fb.FileExtension
	                            FROM PMIS.ProjectTaskFile fa
	                            INNER JOIN BAS.[File] fb ON fb.FileId = fa.FileId
	                            WHERE fa.ProjectTaskId = a.ProjectTaskId
	                            ORDER BY fa.ProjectTaskFileId
	                            FOR JSON PATH, ROOT('data')
                            ) TaskFile
                            , ISNULL((
	                            SELECT COUNT(1)
	                            FROM PMIS.TaskUser ea
                                INNER JOIN BAS.[User] eb ON eb.UserId = ea.UserId
	                            WHERE ea.TaskId = a.ProjectTaskId
	                            AND ea.TaskType = 'Project'
                                AND eb.[Status] = 'A'), 0
                            ) TaskUserCount
                            , (SELECT rb.UserNo, rb.UserName, rb.Gender
	                            , ISNULL(FORMAT(ra.ReplyDate, 'yyyy-MM-dd HH:mm'), '') ReplyDate, ra.ReplyContent, ra.DeferredDuration, ra.ReplyStatus, rc.StatusName ReplyStatusName
	                            FROM PMIS.TaskReply ra
	                            INNER JOIN BAS.[User] rb ON rb.UserId = ra.ReplyUser
	                            INNER JOIN BAS.[Status] rc ON rc.StatusNo = ra.ReplyStatus AND rc.StatusSchema = 'TaskReply.ReplyStatus'
	                            WHERE ra.ReplyType = 'Project'
	                            AND ra.TaskId = a.ProjectTaskId
	                            AND rb.[Status] = 'A'
	                            ORDER BY ra.ReplyDate
	                            FOR JSON PATH, ROOT('data')
                            ) TaskReply
                            , (
	                            SELECT DISTINCT sa.* FROM (
		                            SELECT sa.StatusNo, sa.StatusName, sa.StatusDesc, 'B' TaskStatus
		                            FROM BAS.[Status] sa 
		                            WHERE sa.StatusSchema = 'TaskReply.ReplyStatus'
		                            AND sa.StatusNo IN ('S', 'N', 'D')
		                            UNION
		                            SELECT sa.StatusNo, sa.StatusName, sa.StatusDesc, 'I' TaskStatus
		                            FROM BAS.[Status] sa 
		                            WHERE sa.StatusSchema = 'TaskReply.ReplyStatus'
		                            AND sa.StatusNo IN ('O', 'D', 'C')
	                            ) sa
	                            WHERE sa.TaskStatus = a.TaskStatus
	                            ORDER BY sa.StatusNo DESC
	                            FOR JSON PATH, ROOT('data')
                            ) ReplyOptions
                            FROM @projectTask x
                            INNER JOIN PMIS.ProjectTask a ON a.ProjectTaskId = x.ProjectTaskId
                            INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                            INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                            INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                            OUTER APPLY (
                                SELECT COUNT(1) SubTask
                                FROM PMIS.ProjectTask ca
                                WHERE ca.ParentTaskId = a.ProjectTaskId
                            ) c
                            OUTER APPLY (
                                SELECT (
                                    SELECT db.UserName + ','
                                    FROM PMIS.TaskUser da
                                    INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                    WHERE da.TaskId = a.ProjectTaskId
                                    AND da.TaskType = 'Project'
                                    ORDER BY da.AgentSort
                                    FOR XML PATH('')
                                ) TaskUsers
                            ) d
                            OUTER APPLY (
	                            SELECT TOP 1 ha.MamoChannelId, ha.ChannelNo, ha.ChannelName
	                            FROM MAMO.Channels ha
	                            WHERE ha.ChannelId = a.ChannelId
                            ) h
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProjectTaskId", " AND a.ProjectTaskId = @ProjectTaskId", ProjectTaskId);
                    if (!string.IsNullOrWhiteSpace(TaskUserNos))
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                            , "TaskUserNos"
                            , @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM PMIS.TaskUser aa
                            INNER JOIN BAS.[User] bu1 ON bu1.UserId = aa.UserId
                            WHERE aa.TaskId = a.ProjectTaskId
                            AND TaskType = 'Project'
		                    AND bu1.UserNo IN @TaskUserNos)"
                            , TaskUserNos.Split(','));
                    }
                    sql += " ORDER BY a.ProjectTaskId";
                    result = sqlConnection.Query(sql, dynamicParameters);
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

        #region //UpdateTaskReply -- 任務回報 -- Chia Yuan 2024.06.21
        public string UpdateTaskReply(int ProjectTaskId, string ReplyType, string ReplyStatus, string ActualStart, string ActualEnd
            , int DeferredDuration, string DurationUnit, int ReplyFile, string ReplyContent, string UserNo)
        {
            try
            {
                if (!Regex.IsMatch(ReplyType, "^(Project|Personal)$", RegexOptions.IgnoreCase)) throw new SystemException("【回報類型】錯誤!");
                if (!Regex.IsMatch(ReplyStatus, "^(C|D|N|O|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【回報狀態】錯誤!");
                if (string.IsNullOrWhiteSpace(ReplyType)) throw new SystemException("【回報類型】不能為空!");
                if (string.IsNullOrWhiteSpace(ReplyStatus)) throw new SystemException("【回報狀態】不能為空!");
                if (ReplyStatus == "D")
                {
                    if (DeferredDuration <= 0) throw new SystemException("【延後時間】必填且需大於0");
                    if (string.IsNullOrWhiteSpace(DurationUnit)) throw new SystemException("【時間單位】不能為空!");
                    if (string.IsNullOrWhiteSpace(ReplyContent)) throw new SystemException("【延宕原因】不能為空!");
                }
                if (ReplyStatus == "N")
                {
                    if (string.IsNullOrWhiteSpace(ReplyContent)) throw new SystemException("【尚未開始原因】不能為空!");
                }
                if (!string.IsNullOrWhiteSpace(ReplyContent)) ReplyContent = ReplyContent.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1
                                u.UserId, u.UserName
                                FROM BAS.[User] u
                                WHERE u.UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【任務成員】資料錯誤!");
                        foreach (var item in result)
                        {
                            CurrentUser = item.UserId;
                            CreateBy = item.UserId;
                            LastModifiedBy = item.UserId;
                        }
                        #endregion

                        #region //判斷回報類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusSchema = 'TaskReply.ReplyStatus'
                                AND StatusNo = @ReplyStatus";
                        result = sqlConnection.Query(sql, new { ReplyStatus });
                        if (result.Count() <= 0) throw new SystemException("【回報類型】資料錯誤!");
                        #endregion

                        #region //判斷回報類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusSchema = @StatusSchema
                                AND StatusNo = @ReplyStatus";
                        result = sqlConnection.Query(sql, new { StatusSchema = "TaskReply.ReplyStatus", ReplyStatus });
                        if (result.Count() <= 0) throw new SystemException("【回報類型】資料錯誤!");
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

                        switch (ReplyType)
                        {
                            case "Project":

                                #region //取得專案任務資料
                                sql = @"SELECT TOP 1 a.ProjectId
                                        , ISNULL(a.ParentTaskId, -1) ParentTaskId
                                        , a.TaskName, a.TaskStatus, a.RequiredFile
                                        , FORMAT(a.EstimateStart, 'yyyy-MM-dd HH:mm') EstimateStart
                                        , FORMAT(a.EstimateEnd, 'yyyy-MM-dd HH:mm') EstimateEnd
                                        , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                        , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
                                        , b.ProjectName
                                        , c.UserId TaskUserId
                                        FROM PMIS.ProjectTask a
                                        INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                        INNER JOIN PMIS.TaskUser c ON c.TaskId = a.ProjectTaskId AND c.TaskType = 'Project'
                                        WHERE ProjectTaskId = @ProjectTaskId
                                        AND c.UserId = @CurrentUser";
                                result = sqlConnection.Query(sql, new { ProjectTaskId, CurrentUser });
                                if (result.Count() <= 0) throw new SystemException("【專案】資料錯誤!");

                                foreach (var item in result)
                                {
                                    ProjectId = item.ProjectId;
                                    projectName = item.ProjectName;
                                    taskName = item.TaskName;
                                    taskStatus = item.TaskStatus;
                                    estimateStart = item.EstimateStart;
                                    estimateEnd = item.EstimateEnd;
                                    actualStart = item.ActualStart;
                                    requiredFile = Convert.ToBoolean(item.RequiredFile);
                                }
                                #endregion

                                #region //遞迴取得任務所有前置任務
                                dynamicParameters = new DynamicParameters();
                                sql = GetSingleTaskRecursionSqlString(ref dynamicParameters, ProjectTaskId);
                                sql += @"
                                    SELECT a.ProjectTaskId
                                    , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                                    , a.TaskName, a.TaskStatus
                                    , FORMAT(a.PlannedStart, 'yyyy-MM-dd HH:mm') PlannedStart
                                    , FORMAT(a.PlannedEnd, 'yyyy-MM-dd HH:mm') PlannedEnd
                                    , ISNULL(a.PlannedDuration, 1) PlannedDuration
                                    , FORMAT(a.EstimateStart, 'yyyy-MM-dd HH:mm') EstimateStart
                                    , FORMAT(a.EstimateEnd, 'yyyy-MM-dd HH:mm') EstimateEnd
                                    , ISNULL(a.EstimateDuration, 1) EstimateDuration
                                    , FORMAT(a.ActualStart, 'yyyy-MM-dd HH:mm') ActualStart
                                    , FORMAT(a.ActualEnd, 'yyyy-MM-dd HH:mm') ActualEnd
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
                                                        WHERE a.ParentTaskId = @ProjectTaskId";
                                                var resultSubTask = sqlConnection.Query(sql, new { ProjectTaskId });
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
                                                    new 
                                                    {
                                                        ProjectId,
                                                        ActualStart = resultActualStart.ToString("yyyy-MM-dd HH:mm"), //實際開始時間
                                                    });
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualStart = @ActualStart,
                                                        a.EstimateDuration = PlannedDuration,
                                                        a.TaskStatus = 'I',
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId IN @ProjectTaskIds";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectTaskIds = projectTasks.Where(w => w.TaskStatus == "B").Select(s => s.ProjectTaskId).ToArray(),
                                                        ActualStart = resultActualStart.ToString("yyyy-MM-dd HH:mm"), //實際開始時間
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
                                                        WHERE a.ProjectTaskId = @ProjectTaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectTaskId,
                                                        EstimateStart = Convert.ToDateTime(estimateStart).AddMinutes(DeferredDuration).ToString("yyyy-MM-dd HH:mm"), //預估開始時間
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
                                            default:
                                                throw new SystemException("不允許的【回報類型】!");
                                        }

                                        if (canReply)
                                        {
                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (ProjectTaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@ProjectTaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    ProjectTaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateStart ? resultActualStart : CreateDate,
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
                                                        WHERE a.ProjectTaskId = @ProjectTaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectTaskId,
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
                                                        WHERE a.ParentTaskId = @ProjectTaskId";
                                                var resultSubTask = sqlConnection.Query(sql, new { ProjectTaskId });
                                                if (resultSubTask.Any())
                                                {
                                                    if (resultSubTask.Any(a => a.TaskStatus != "C")) throw new SystemException("【子任務尚未完成】，無法回報完成!");

                                                    if (DateTime.TryParse(resultSubTask.Max(m => m.ActualEnd), out DateTime subMaxActualEnd))
                                                    {
                                                        if (resultActualEnd.CompareTo(subMaxActualEnd) <= 0) throw new SystemException(string.Format("【完成時間】不得小於子任務【最後完成時間】{0}", resultSubTask.Max(m => m.ActualEnd)));
                                                    }
                                                }
                                                #endregion

                                                #region //更新任務資料
                                                sql = @"UPDATE a SET
                                                        a.ActualEnd = @ActualEnd,
                                                        a.TaskStatus = 'C',
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.ProjectTask a
                                                        WHERE a.ProjectTaskId = @ProjectTaskId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectTaskId,
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
                                                                    a.TaskStatus = 'C',
                                                                    a.LastModifiedDate = @LastModifiedDate,
                                                                    a.LastModifiedBy = @LastModifiedBy
                                                                    FROM PMIS.ProjectTask a
                                                                    WHERE a.ProjectTaskId = @TempParentTaskId
                                                                    AND a.TaskStatus <> 'C'";
                                                            dynamicParameters.AddDynamicParams(
                                                                new
                                                                {
                                                                    TempParentTaskId,
                                                                    ActualEnd = hasMaxActualEnd ? maxActualEnd.ToString("yyyy-MM-dd HH:mm") : resultActualEnd.ToString("yyyy-MM-dd HH:mm"),
                                                                    LastModifiedDate,
                                                                    LastModifiedBy
                                                                });
                                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                }

                                                #region //更新專案資訊(必須在任務狀態更新之後執行)
                                                sql = @"UPDATE a SET 
                                                        a.ProjectStatus = 'C',
                                                        a.ActualEnd = @ActualEnd,
                                                        a.LastModifiedDate = @LastModifiedDate,
                                                        a.LastModifiedBy = @LastModifiedBy
                                                        FROM PMIS.Project a
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubTask
	                                                        FROM PMIS.ProjectTask ca
	                                                        WHERE ca.ProjectId = a.ProjectId
                                                        ) c
                                                        OUTER APPLY (
	                                                        SELECT COUNT(1) SubC
	                                                        FROM PMIS.ProjectTask da
	                                                        WHERE da.ProjectId = a.ProjectId
	                                                        AND da.TaskStatus = 'C'
                                                        ) d
                                                        WHERE c.SubTask = d.SubC
                                                        AND a.ProjectStatus <> 'P'
                                                        AND a.ProjectId = @ProjectId";
                                                rowsAffected += sqlConnection.Execute(sql,
                                                    new
                                                    {
                                                        ProjectId,
                                                        ActualEnd = resultActualEnd.ToString("yyyy-MM-dd HH:mm"),
                                                        LastModifiedDate,
                                                        LastModifiedBy,
                                                    });
                                                #endregion

                                                #region //取得通知對象
                                                dynamicParameters = new DynamicParameters();
                                                sql = GetTaskRecursionSqlString(ref dynamicParameters, new List<dynamic> { ProjectId });
                                                sql += @"
                                                    SELECT
                                                    a.ProjectTaskId, a.TaskStatus
                                                    , b.ProjectNo, b.ProjectName, b.ProjectDesc
                                                    , a.TaskName, a.TaskDesc
                                                    , a.TaskStatus, s2.StatusName TaskStatusName
                                                    , ISNULL(x.ParentTaskId, -1) ParentTaskId, x.TaskLevel, x.TaskRoute, x.TaskSort, x.TaskRouteName
                                                    , ISNULL(e.DepartmentNo, '') DepartmentNo, ISNULL(e.DepartmentName, '') DepartmentName
                                                    , ISNULL(e.CompanyNo, '') CompanyNo, ISNULL(e.CompanyName, '') CompanyName, ISNULL(e.LogoIcon, -1) LogoIcon
                                                    , ISNULL(SUBSTRING(d.TaskUsers, 0, LEN(d.TaskUsers)), '') TaskUsers
                                                    , u.UserName ProjectMasterName, u.UserNo ProjectMasterNo, u.Gender ProjectMasterGender, u.UserName + '(' + u.UserNo + ')' ProjectMaster
                                                    , ISNULL(ISNULL(link.TaskName, pLink.TaskName), '-') PrevTaskName
                                                    ,  (
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
                                                            AND ba.TaskId = a.ProjectTaskId
                                                            AND ba.UserId = ub.UserId
                                                        )
                                                        FOR JSON PATH, ROOT('data')
                                                    ) MailUsers
                                                    FROM @projectTask x
                                                    INNER JOIN PMIS.ProjectTask a ON a.ProjectTaskId = x.ProjectTaskId
                                                    INNER JOIN PMIS.Project b ON b.ProjectId = a.ProjectId
                                                    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.ProjectStatus AND s1.StatusSchema = 'Project.ProjectStatus'
                                                    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.TaskStatus AND s2.StatusSchema = 'ProjectTask.TaskStatus'
                                                    INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.WorkTimeStatus AND s3.StatusSchema = 'Boolean'
                                                    INNER JOIN BAS.[Type] t1 ON t1.TypeNo = b.ProjectType AND t1.TypeSchema = 'Project.ProjectType'
                                                    INNER JOIN BAS.[Type] t2 ON t2.TypeNo = b.ProjectAttribute AND t2.TypeSchema = 'Project.ProjectAttribute'
                                                    OUTER APPLY (
	                                                    SELECT ea.DepartmentNo, ea.DepartmentName
	                                                    , eb.CompanyNo, eb.CompanyName , eb.LogoIcon
	                                                    FROM BAS.[Department] ea
	                                                    INNER JOIN BAS.[Company] eb ON eb.CompanyId = ea.CompanyId
	                                                    WHERE ea.DepartmentId = b.DepartmentId
                                                    ) e
                                                    OUTER APPLY (
                                                        SELECT (
                                                            SELECT db.UserName + ','
                                                            FROM PMIS.TaskUser da
                                                            INNER JOIN BAS.[User] db ON da.UserId = db.UserId
                                                            WHERE da.TaskId = a.ProjectTaskId
                                                            AND da.TaskType = 'Project'
                                                            ORDER BY da.AgentSort
                                                            FOR XML PATH('')
                                                        ) TaskUsers
                                                    ) d
                                                    OUTER APPLY (
	                                                    SELECT TOP 1 ub.UserId, ub.UserNo, ub.UserName, ub.Gender
	                                                    FROM PMIS.ProjectUser ua 
	                                                    INNER JOIN BAS.[User] ub ON ub.UserId = ua.UserId
	                                                    WHERE ua.ProjectId = a.ProjectId 
	                                                    AND ua.AgentSort = 1
                                                    ) u
                                                    OUTER APPLY (SELECT ba.SourceTaskId, ba.TargetTaskId, bb.TaskName, bb.TaskStatus SourceStatus FROM PMIS.ProjectTaskLink ba INNER JOIN PMIS.ProjectTask bb ON bb.ProjectTaskId = ba.SourceTaskId WHERE ba.TargetTaskId = a.ProjectTaskId) link
                                                    OUTER APPLY (SELECT ca.SourceTaskId, ca.TargetTaskId, cb.TaskName, cb.TaskStatus ParentSourceStatus FROM PMIS.ProjectTaskLink ca INNER JOIN PMIS.ProjectTask cb ON cb.ProjectTaskId = ca.SourceTaskId WHERE ca.TargetTaskId = a.ParentTaskId) pLink
                                                    WHERE NOT EXISTS (SELECT TOP 1 1 FROM @projectTask aa WHERE aa.ParentTaskId = a.ProjectTaskId) --排除父層任務
                                                    AND ISNULL(link.SourceStatus, '') IN ('', 'C') --前置任務需完成
                                                    AND ISNULL(pLink.ParentSourceStatus, '') IN ('', 'C') --父層前置任務需完成
                                                    AND a.TaskStatus IN ('B') --待開始
                                                    AND (EXISTS (
		                                                    SELECT TOP 1 ISNULL(da.ProjectTaskLinkId, 1)
		                                                    FROM PMIS.ProjectTaskLink da 
		                                                    LEFT JOIN PMIS.ProjectTask db ON db.ProjectTaskId = da.TargetTaskId
		                                                    LEFT JOIN PMIS.ProjectTask dc ON dc.ProjectTaskId = da.SourceTaskId
		                                                    WHERE da.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')) AND ISNULL(dc.TaskStatus, '') IN ('', 'C')
	                                                    ) 
	                                                    OR NOT EXISTS (SELECT TOP 1 1 FROM PMIS.ProjectTaskLink ea WHERE ea.TargetTaskId IN (SELECT REPLACE(value, -1, x.ProjectTaskId) FROM STRING_SPLIT(x.TaskRoute, ',')))
                                                    )
                                                    ORDER BY ISNULL(ISNULL(a.ActualStart, a.EstimateStart), a.PlannedStart)";
                                                targetMailUsers = sqlConnection.Query(sql, dynamicParameters).ToList();
                                                #endregion

                                                #region //完成通知
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
                                                            mailSubject = mailSubject.Replace("[ProjectName]", task.ProjectName)
                                                                                     .Replace("[TaskName]", task.TaskName);
                                                            mailContent = mailContent.Replace("[ProjectName]", task.ProjectName)
                                                                                     .Replace("[PrevTaskName]", task.PrevTaskName)
                                                                                     .Replace("[TaskName]", task.TaskName)
                                                                                     .Replace("[TaskUser]", task.TaskUsers)
                                                                                     .Replace("[MailContent]", string.Format("專案任務【{0}】已可進行，請安排時間開始進行。", task.TaskName));

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
                                                    AND a.TaskId = @ProjectTaskId
                                                    AND ISNULL(a.ReplyFile, -1) <> -1";
                                            var resultReplyFile = sqlConnection.QueryFirstOrDefault(sql, new { ProjectTaskId, CurrentUser });
                                            if (requiredFile && resultReplyFile == null && ReplyFile == -1) throw new SystemException("本任務為必傳附件，請使用BM系統回報並上傳回報附件!");
                                            #endregion

                                            #region //新增回報紀錄
                                            sql = @"INSERT INTO PMIS.TaskReply (ProjectTaskId, ReplyType, ReplyContent, ReplyUser
                                                    , ReplyDate, ReplyStatus, DeferredDuration, EventDate
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.ReplyId
                                                    VALUES (@ProjectTaskId, @ReplyType, @ReplyContent, @ReplyUser
                                                    , @ReplyDate, @ReplyStatus, @DeferredDuration, @EventDate
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            rowsAffected += sqlConnection.Query(sql,
                                                new
                                                {
                                                    ProjectTaskId,
                                                    ReplyType,
                                                    ReplyContent,
                                                    ReplyUser = CurrentUser,
                                                    ReplyDate = CreateDate,
                                                    ReplyStatus,
                                                    DeferredDuration,
                                                    EventDate = inputDateEnd ? resultActualEnd : CreateDate,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                }).Count();
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
                                //待開發
                                break;
                            default:
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
    }
}
