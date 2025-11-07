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

namespace SSODA
{
    public class SsoSystemSettingDA
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

        public SsoSystemSettingDA()
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
        #region //GetSourceFlowSetting -- 需求來源流程設定資料 -- Ben Ma 2023.07.13
        public string GetSourceFlowSetting(int SourceId, int ExcludeSettingId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SettingId, a.SourceId, a.FlowId, a.Xaxis, a.Yaxis
                            , b.FlowName, b.FlowImage
                            , ISNULL(c.SourceSettings, '') SourceSettings
                            , (
                                SELECT ISNULL(aa.RoleId, -1) RoleId, ISNULL(aa.UserId, -1) UserId, aa.MailAdviceStatus, aa.PushAdviceStatus, aa.WorkWeixinStatus
                                , ISNULL(ab.RoleName, ac.UserNo + ' ' + ac.UserName) RoleUserName
                                , CASE WHEN aa.RoleId IS NOT NULL THEN 'Y' ELSE 'N' END RoleStatus
                                , (
                                    SELECT aab.UserNo, aab.UserName
                                    FROM SSO.RoleUser aaa
                                    INNER JOIN BAS.[User] aab ON aaa.UserId = aab.UserId
                                    WHERE aaa.RoleId = aa.RoleId
                                    ORDER BY aab.UserNo
                                    FOR JSON PATH, ROOT('data')
                                ) RoleUser
                                FROM SSO.SourceFlowUser aa
                                LEFT JOIN SSO.DemandRole ab ON aa.RoleId = ab.RoleId
                                LEFT JOIN BAS.[User] ac ON aa.UserId = ac.UserId
                                WHERE aa.SettingId = a.SettingId
                                ORDER BY ab.RoleName, ac.UserNo
                                FOR JSON PATH, ROOT('data')
                            ) SourceFlowUser
                            FROM SSO.SourceFlowSetting a
                            INNER JOIN SSO.DemandFlow b ON a.FlowId = b.FlowId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ',' + CAST(ca.SourceSettingId AS NVARCHAR)
                                    FROM SSO.SourceFlowLink ca
                                    WHERE ca.TargetSettingId = a.SettingId
                                    FOR XML PATH('')
                                ), 1, 1, '') SourceSettings
                            ) c
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SourceId", @" AND a.SourceId = @SourceId", SourceId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ExcludeSettingId", @" AND a.SettingId != @ExcludeSettingId", ExcludeSettingId);

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

        #region //GetSourceFlowUser -- 取得需求來源流程使用者對應資料 -- Ben Ma 2023.07.14
        public string GetSourceFlowUser(int SettingId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SettingId, a.RoleId, a.UserId, a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                            , CASE WHEN a.RoleId IS NOT NULL THEN 'Y' ELSE 'N' END RoleStatus
                            , ISNULL(b.RoleName, c.UserNo + ' ' + c.UserName) RoleUserName
                            FROM SSO.SourceFlowUser a
                            LEFT JOIN SSO.DemandRole b ON a.RoleId = b.RoleId
                            LEFT JOIN BAS.[User] c ON a.UserId = c.UserId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SettingId", @" AND a.SettingId = @SettingId", SettingId);

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

        #region //GetSourceFlowSettingDiagram -- 取得需求來源流程設定資料(流程圖) -- Ben Ma 2023.07.18
        public string GetSourceFlowSettingDiagram(int SourceId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Shape Json組成
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SettingId, a.Xaxis, a.Yaxis
                            , b.FlowName, b.FlowImage
                            , ISNULL(c.SourceFlowUser, '') SourceFlowUser
                            FROM SSO.SourceFlowSetting a
                            INNER JOIN SSO.DemandFlow b ON a.FlowId = b.FlowId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ',' + ISNULL(ab.RoleName, ac.UserName)
                                    FROM SSO.SourceFlowUser aa
                                    LEFT JOIN SSO.DemandRole ab ON aa.RoleId = ab.RoleId
                                    LEFT JOIN BAS.[User] ac ON aa.UserId = ac.UserId
                                    WHERE aa.SettingId = a.SettingId
                                    ORDER BY ab.RoleName, ac.UserNo
                                    FOR XML PATH('')
                                ), 1, 1, '') SourceFlowUser
                            ) c
                            WHERE a.SourceId = @SourceId";
                    dynamicParameters.Add("SourceId", SourceId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    //if (result.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");

                    List<SettingDiagramShape> settingDiagramShapes = new List<SettingDiagramShape>();
                    foreach (var item in result)
                    {
                        settingDiagramShapes.Add(new SettingDiagramShape
                        {
                            id = item.SettingId.ToString(),
                            flowName = item.FlowName,
                            flowImage = item.FlowImage,
                            flowUser = item.SourceFlowUser,
                            x = Convert.ToInt32(item.Xaxis),
                            y = Convert.ToInt32(item.Yaxis)
                        });
                    }
                    #endregion

                    #region //Line Json組成
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SourceSettingId, a.TargetSettingId
                            FROM SSO.SourceFlowLink a
                            INNER JOIN SSO.SourceFlowSetting b ON a.SourceSettingId = b.SettingId
                            INNER JOIN SSO.SourceFlowSetting c ON a.TargetSettingId = c.SettingId
                            WHERE b.SourceId = c.SourceId
                            AND b.SourceId = @SourceId";
                    dynamicParameters.Add("SourceId", SourceId);
                    var result2 = sqlConnection.Query(sql, dynamicParameters);

                    List<DHXDiagramLine> dHXDiagramLines = new List<DHXDiagramLine>();
                    if (result2.Count() > 0)
                    {
                        foreach (var item in result2)
                        {
                            dHXDiagramLines.Add(new DHXDiagramLine
                            {
                                id = "L_" + item.SourceSettingId.ToString() + "_" + item.TargetSettingId.ToString(),
                                from = item.SourceSettingId.ToString(),
                                to = item.TargetSettingId.ToString()
                            });
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        dataShapes = settingDiagramShapes,
                        dataLines = dHXDiagramLines
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
        #region //AddSourceFlowSetting -- 需求來源流程設定資料新增 -- Ben Ma 2023.07.14
        public string AddSourceFlowSetting(int SourceId, int FlowId)
        {
            try
            {
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

                        #region //判斷流程資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlow
                                WHERE FlowId = @FlowId";
                        dynamicParameters.Add("FlowId", FlowId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("流程資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SSO.SourceFlowSetting (SourceId, FlowId, Xaxis, Yaxis
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SettingId
                                VALUES (@SourceId, @FlowId, @Xaxis, @Yaxis
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SourceId,
                                FlowId,
                                Xaxis = 0,
                                Yaxis = 200,
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

        #region //AddSourceFlowLink -- 需求來源流程設定連結新增 -- Ben Ma 2023.07.17
        public string AddSourceFlowLink(string SourceSettings, int TargetSettingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源流程設定資料是否正確
                        List<string> settingsList = new List<string>();

                        if (SourceSettings.Length > 0)
                        {
                            settingsList = SourceSettings.Split(',').ToList();

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(1) TotalSettings
                                    FROM SSO.SourceFlowSetting
                                    WHERE SettingId IN @SettingId";
                            dynamicParameters.Add("SettingId", settingsList);

                            int totalSettings = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalSettings;
                            if (totalSettings != settingsList.Count) throw new SystemException("來源流程設定資料錯誤!");
                        }
                        #endregion

                        #region //判斷需求來源流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", TargetSettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("目標流程設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除所有目標流程設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowLink
                                WHERE TargetSettingId = @TargetSettingId";
                        dynamicParameters.Add("TargetSettingId", TargetSettingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        foreach (var setting in settingsList)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SSO.SourceFlowLink (SourceSettingId, TargetSettingId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@SourceSettingId, @TargetSettingId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SourceSettingId = Convert.ToInt32(setting),
                                    TargetSettingId,
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

        #region //AddSourceFlowUser -- 需求來源流程使用者對應新增 -- Ben Ma 2023.07.14
        public string AddSourceFlowUser(int SettingId, string UserRole, string Users, string Roles)
        {
            try
            {
                if (!Regex.IsMatch(UserRole, "^(user|role)$", RegexOptions.IgnoreCase)) throw new SystemException("【使用者/角色】設定錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        switch (UserRole)
                        {
                            case "user":
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

                                foreach (var user in usersList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.SourceFlowUser (SettingId, UserId
                                            , MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@SettingId, @UserId
                                            , @MailAdviceStatus, @PushAdviceStatus, @WorkWeixinStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SettingId,
                                            UserId = Convert.ToInt32(user),
                                            MailAdviceStatus = "N",
                                            PushAdviceStatus = "N",
                                            WorkWeixinStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                break;
                            case "role":
                                #region //判斷角色資料是否正確
                                string[] rolesList = Roles.Split(',');

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT COUNT(1) TotalRoles
                                        FROM SSO.DemandRole
                                        WHERE RoleId IN @RoleId";
                                dynamicParameters.Add("RoleId", rolesList);

                                int totalRoles = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalRoles;
                                if (totalRoles != rolesList.Length) throw new SystemException("角色資料錯誤!");
                                #endregion

                                foreach (var role in rolesList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.SourceFlowUser (SettingId, RoleId
                                            , MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus
                                            , CreateDate, CreateBy)
                                            VALUES (@SettingId, @RoleId
                                            , @MailAdviceStatus, @PushAdviceStatus, @WorkWeixinStatus
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SettingId,
                                            RoleId = Convert.ToInt32(role),
                                            MailAdviceStatus = "N",
                                            PushAdviceStatus = "N",
                                            WorkWeixinStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
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
        #endregion

        #region //Update
        #region //UpdateSourceFlowUserStatus -- 需求來源流程設定狀態更新 -- Ben Ma 2023.07.18
        public string UpdateSourceFlowUserStatus(int SettingId, int RoleId, int UserId
            , string MailAdviceStatus, string PushAdviceStatus, string WorkWeixinStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");
                        #endregion

                        if (RoleId > 0)
                        {
                            #region //判斷角色資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SSO.DemandRole
                                    WHERE RoleId = @RoleId";
                            dynamicParameters.Add("RoleId", RoleId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("角色資料錯誤!");
                            #endregion
                        }

                        if (UserId > 0)
                        {
                            #region //判斷使用者資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                            #endregion
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.SourceFlowUser SET
                                MailAdviceStatus = @MailAdviceStatus,
                                PushAdviceStatus = @PushAdviceStatus,
                                WorkWeixinStatus = @WorkWeixinStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SettingId = @SettingId
                                AND ISNULL(RoleId, -1) = @RoleId
                                AND ISNULL(UserId, -1) = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MailAdviceStatus,
                                PushAdviceStatus,
                                WorkWeixinStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                SettingId,
                                RoleId,
                                UserId
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

        #region //UpdateSourceFlowSettingCoordinates -- 需求來源流程設定座標更新 -- Ben Ma 2023.07.19
        public string UpdateSourceFlowSettingCoordinates(int SourceId, string DiagramData)
        {
            try
            {
                if (!DiagramData.TryParseJson(out JObject tempJObject)) throw new SystemException("需求來源流程設定資料格式錯誤");

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

                        int rowsAffected = 0;
                        JObject json = JObject.Parse(DiagramData);
                        for (int i = 0; i < json["data"].Count(); i++)
                        {
                            if (json["data"][i]["type"].ToString() == "template")
                            {
                                int settingId = Convert.ToInt32(json["data"][i]["id"]);

                                #region //判斷需求來源流程設定資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SSO.SourceFlowSetting
                                        WHERE SettingId = @SettingId
                                        AND SourceId = @SourceId";
                                dynamicParameters.Add("SettingId", settingId);
                                dynamicParameters.Add("SourceId", SourceId);

                                var resultExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultExist.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");
                                #endregion

                                #region //更新座標
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SSO.SourceFlowSetting SET
                                        Xaxis = @Xaxis,
                                        Yaxis = @Yaxis,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SettingId = @SettingId
                                        AND SourceId = @SourceId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Xaxis = Convert.ToInt32(json["data"][i]["x"]),
                                        Yaxis = Convert.ToInt32(json["data"][i]["y"]),
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        settingId,
                                        SourceId
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
        #region //DeleteSourceFlowUser -- 需求來源流程使用者對應刪除 -- Ben Ma 2023.07.14
        public string DeleteSourceFlowUser(int SettingId, int RoleId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");
                        #endregion

                        if (RoleId > 0)
                        {
                            #region //判斷角色資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SSO.DemandRole
                                    WHERE RoleId = @RoleId";
                            dynamicParameters.Add("RoleId", RoleId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("角色資料錯誤!");
                            #endregion
                        }

                        if (UserId > 0)
                        {
                            #region //判斷使用者資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                            #endregion
                        }

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowUser
                                WHERE SettingId = @SettingId
                                AND ISNULL(RoleId, -1) = @RoleId
                                AND ISNULL(UserId, -1) = @UserId";
                        dynamicParameters.Add("SettingId", SettingId);
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

        #region //DeleteSourceFlowSetting -- 需求來源流程設定資料刪除 -- Ben Ma 2023.07.17
        public string DeleteSourceFlowSetting(int SettingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求來源流程設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //需求來源流程設定連結刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowLink
                                WHERE SourceSettingId = @SourceSettingId";
                        dynamicParameters.Add("SourceSettingId", SettingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowLink
                                WHERE TargetSettingId = @TargetSettingId";
                        dynamicParameters.Add("TargetSettingId", SettingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //需求來源流程使用者對應刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowUser
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.SourceFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

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
