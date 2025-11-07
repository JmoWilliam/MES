using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace BASDA
{
    public class SystemSettingDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
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

        public SystemSettingDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
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
        #region //GetRandomApiKey -- 取得隨機Api金鑰 -- Ben Ma 2022.04.27
        public string GetRandomApiKey()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    bool check = true;
                    string cipherText = "";

                    while (check)
                    {
                        var plainText = Guid.NewGuid();
                        cipherText = BaseHelper.Sha256Encrypt(plainText.ToString());

                        #region //判斷隨機金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", cipherText);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        check = result.Count() > 0;
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = cipherText
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

        #region //GetApiKey -- 取得Api金鑰資料 -- Ben Ma 2022.04.26
        public string GetApiKey(int KeyId, string KeyText, string Purpose, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.KeyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.KeyText, a.Purpose, a.Status";
                    sqlQuery.mainTables =
                        @"FROM BAS.ApiKey a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND CompanyId=@CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "KeyId", @" AND a.KeyId = @KeyId", KeyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "KeyText", @" AND a.KeyText LIKE '%' + @KeyText + '%'", KeyText);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Purpose", @" AND a.Purpose LIKE '%' + @Purpose + '%'", Purpose);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.KeyId";
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

        #region //GetApiKeyFunction -- 取得Api適用功能 -- Ben Ma 2022.04.29
        public string GetApiKeyFunction(int FunctionId, int KeyId, string FunctionCode
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.FunctionId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FunctionCode, a.Purpose
                        , b.KeyId";
                    sqlQuery.mainTables =
                        @"FROM BAS.ApiKeyFunction a
                        INNER JOIN BAS.ApiKey b ON a.KeyId = b.KeyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND CompanyId=@CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionId", @" AND a.FunctionId = @FunctionId", FunctionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "KeyId", @" AND b.KeyId = @KeyId", KeyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionCode", @" AND a.FunctionCode LIKE '%' + @FunctionCode + '%'", FunctionCode);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.FunctionId";
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

        #region //GetApiKeyVerify -- Api金鑰驗證 -- Ben Ma 2022.05.03
        public string GetApiKeyVerify(string CompanyNo, string KeyText, string FunctionCode)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 1
                            FROM BAS.ApiKeyFunction a
                            INNER JOIN BAS.ApiKey b ON a.KeyId = b.KeyId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                            WHERE c.CompanyNo = @CompanyNo
                            AND b.KeyText = @KeyText
                            AND a.FunctionCode = @FunctionCode";
                    dynamicParameters.Add("CompanyNo", CompanyNo);
                    dynamicParameters.Add("KeyText", KeyText);
                    dynamicParameters.Add("FunctionCode", FunctionCode);

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

        #region //GetRole -- 取得角色資料 -- Ben Ma 2022.05.06
        public string GetRole(int RoleId, string RoleName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoleName, a.AdminStatus, a.Status";
                    sqlQuery.mainTables =
                        @"FROM BAS.Role a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND CompanyId=@CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleName", @" AND a.RoleName LIKE '%' + @RoleName + '%'", RoleName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoleId";
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

        #region //GetRoleAuthority -- 取得角色權限資料 -- Ben Ma 2022.05.09
        public string GetRoleAuthority(int RoleId, string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SystemId, a.SystemName, a.SortNumber, c.TotalFunction, d.RoleTotalFunction
                            , (
                                SELECT aa.ModuleId, aa.ModuleName, ab.TotalModuleFunction, ac.RoleTotalModuleFunction
                                FROM BAS.[Module] aa
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM BAS.FunctionDetail aba
                                    INNER JOIN BAS.[Function] abb ON aba.FunctionId = abb.FunctionId
                                    WHERE abb.ModuleId = aa.ModuleId
                                    AND abb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ab
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalModuleFunction
                                    FROM BAS.RoleFunctionDetail aca
                                    INNER JOIN BAS.FunctionDetail acb ON aca.DetailId = acb.DetailId
                                    INNER JOIN BAS.[Function] acc ON acb.FunctionId = acc.FunctionId
                                    WHERE acc.ModuleId = aa.ModuleId
                                    AND acc.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND aca.RoleId = @RoleId
                                ) ac
                                WHERE 1=1
                                AND aa.SystemId = a.SystemId
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) ModuleData
                            FROM BAS.[System] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunction
                                FROM BAS.FunctionDetail ca
                                INNER JOIN BAS.[Function] cb ON ca.FunctionId = cb.FunctionId
                                INNER JOIN BAS.[Module] cc ON cb.ModuleId = cc.ModuleId
                                WHERE cc.SystemId = a.SystemId
                                AND cb.FunctionName LIKE '%' + @SearchKey + '%'
                            ) c
                            OUTER APPLY (
                                SELECT COUNT(1) RoleTotalFunction
                                FROM BAS.RoleFunctionDetail da
                                INNER JOIN BAS.FunctionDetail db ON da.DetailId = db.DetailId
                                INNER JOIN BAS.[Function] dc ON db.FunctionId = dc.FunctionId
                                INNER JOIN BAS.[Module] dd ON dc.ModuleId = dd.ModuleId
                                WHERE dd.SystemId = a.SystemId
                                AND dc.FunctionName LIKE '%' + @SearchKey + '%'
                                AND da.RoleId = @RoleId
                            ) d
                            WHERE 1=1
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("RoleId", RoleId);
                    dynamicParameters.Add("SearchKey", SearchKey);

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

        #region //GetRoleDetailAuthority -- 取得角色詳細權限資料 -- Ben Ma 2022.05.11
        public string GetRoleDetailAuthority(int ModuleId, int RoleId, string SearchKey, bool Grant)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.FunctionId, a.FunctionName, a.FunctionCode, a.SortNumber, b.TotalFunctionDetail, c.RoleTotalFunctionDetail
                            , (
                                SELECT aa.DetailId, aa.DetailName, aa.DetailCode, aa.Status, ab.DetailAuthority
                                FROM BAS.FunctionDetail aa
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM BAS.RoleFunctionDetail aaa
                                        WHERE aaa.DetailId = aa.DetailId
                                        AND aaa.RoleId = @RoleId
                                    ), 0) DetailAuthority
                                ) ab
                                WHERE 1=1
                                AND aa.FunctionId = a.FunctionId
                                AND ab.DetailAuthority >= @DetailAuthority
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) FunctionDetailData
                            FROM BAS.[Function] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunctionDetail
                                FROM BAS.FunctionDetail ba
                                WHERE ba.FunctionId = a.FunctionId
                            ) b
                            OUTER APPLY (
                                SELECT COUNT(1) RoleTotalFunctionDetail
                                FROM BAS.RoleFunctionDetail ca
                                INNER JOIN BAS.FunctionDetail cb ON ca.DetailId = cb.DetailId
                                WHERE cb.FunctionId = a.FunctionId
                                AND ca.RoleId = @RoleId
                            ) c
                            WHERE a.ModuleId = @ModuleId
                            AND a.FunctionName LIKE '%' + @SearchKey + '%'
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("ModuleId", ModuleId);
                    dynamicParameters.Add("RoleId", RoleId);
                    dynamicParameters.Add("SearchKey", SearchKey);
                    dynamicParameters.Add("DetailAuthority", Grant ? 1 : 0);

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

        #region //GetAuthorityUser -- 取得權限使用者資料 -- Ben Ma 2022.12.27
        public string GetAuthorityUser(string Roles, int CompanyId, string Departments, string UserNo, string UserName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = ", a.UserNo, b.Roles";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender
                        , CASE a.Status 
                            WHEN 'S' THEN a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE a.UserNo + ' ' + a.UserName
                        END UserWithNo
                        , b.DepartmentId, b.DepartmentNo, b.DepartmentName
                        , c.CompanyId, c.CompanyNo, c.CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon
                        , (
                            SELECT ab.RoleId, ab.RoleName, ab.AdminStatus
                            FROM BAS.UserRole aa
                            INNER JOIN BAS.Role ab ON aa.RoleId = ab.RoleId
                            WHERE aa.UserId = a.UserId
                            AND ab.CompanyId = @CurrentCompany
                            ORDER BY ab.RoleId
                            FOR JSON PATH, ROOT('data')
                        ) UserRole";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    string queryTable =
                        @"FROM (
                            SELECT a.UserId, a.UserNo, a.UserName, b.DepartmentId, c.CompanyId, a.Status
                            , (
                                SELECT ab.RoleId, ab.RoleName, ab.AdminStatus
                                FROM BAS.UserRole aa
                                INNER JOIN BAS.Role ab ON aa.RoleId = ab.RoleId
                                WHERE aa.UserId = a.UserId
                                AND ab.CompanyId = @CurrentCompany
                                ORDER BY ab.RoleId
                                FOR JSON PATH, ROOT('data')
                            ) UserRole
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                        ) a
                        OUTER APPLY (
                            SELECT CONCAT(STUFF((
                                SELECT ',' + CONVERT(nvarchar(50), x.RoleId)
                                FROM OPENJSON(a.UserRole, '$.data')
                                WITH (
                                    RoleId INT N'$.RoleId'
                                ) x
                                FOR XML PATH('')
                            ), 1, 0, ''), ',') Roles
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    //string queryCondition = " AND a.CompanyId = @CompanyId";
                    string queryCondition = "";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);
                    if (Roles.Length > 0)
                    {
                        string[] roleArray = Roles.Split(',');

                        string sqlRole = "";
                        foreach (var role in roleArray)
                        {
                            sqlRole += string.Format(@" AND b.Roles LIKE '%,{0},%'", role);
                        }

                        queryCondition += sqlRole;
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND a.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserNo, a.UserId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = true;

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

        #region //GetAuthorityUserCompany -- 取得使用者公司權限資料 -- Ben Ma 2023.05.17
        public string GetAuthorityUserCompany(int CompanyId, string Departments, string UserNo, string UserName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender
                        , CASE a.Status 
                            WHEN 'S' THEN a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE a.UserNo + ' ' + a.UserName
                        END UserWithNo
                        , b.DepartmentId, b.DepartmentNo, b.DepartmentName
                        , c.CompanyId, c.CompanyNo, c.CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon
                        , (
                            SELECT aa.CompanyId, ab.CompanyNo, ab.CompanyName, ISNULL(ab.LogoIcon, -1) LogoIcon
                            FROM BAS.AuthorityUserCompany aa
                            INNER JOIN BAS.Company ab ON aa.CompanyId = ab.CompanyId
                            WHERE aa.UserId = a.UserId
                            ORDER BY ab.CompanyId
                            FOR JSON PATH, ROOT('data')
                        ) UserCompany";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    string queryTable =
                        @"FROM (
                            SELECT a.UserId, a.DepartmentId, a.UserNo, a.UserName, b.CompanyId, a.Status
                            , (
                                SELECT aa.CompanyId
                                FROM BAS.AuthorityUserCompany aa
                                WHERE aa.UserId = a.UserId
                                ORDER BY aa.CompanyId
                                FOR JSON PATH, ROOT('data')
                            ) UserCompany
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        ) a
                        OUTER APPLY (
                            SELECT TOP 1 x.CompanyId
                            FROM OPENJSON(a.UserCompany, '$.data')
                            WITH (
                                CompanyId INT N'$.CompanyId'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND b.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND a.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserNo, a.UserId";
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

        #region //GetRoleUser -- 取得角色使用者資料 -- Ben Ma 2022.05.16
        public string GetRoleUser(int RoleId, string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT c.UserId, c.UserNo, c.UserName, c.Status
                            FROM BAS.UserRole a
                            INNER JOIN BAS.Role b ON a.RoleId = b.RoleId
                            INNER JOIN BAS.[User] c ON a.UserId = c.UserId
                            WHERE b.CompanyId = @CompanyId
                            AND a.RoleId = @RoleId
                            AND (c.UserNo LIKE '%' + @SearchKey + '%' OR c.UserName LIKE '%' + @SearchKey + '%')
                            ORDER BY c.Status, c.UserNo";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("RoleId", RoleId);
                    dynamicParameters.Add("SearchKey", SearchKey);

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

        #region //GetUserRole -- 取得使用者角色資料 -- Ben Ma 2022.05.13
        public string GetUserRole(string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT ISNULL(b.RoleNames, '無權限') RoleNames, b.RoleIds
                            , (
                                SELECT aa.UserId, aa.UserNo, aa.UserName
                                FROM BAS.[User] aa
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ',' + CONVERT(nvarchar(50), abb.RoleId)
                                        FROM BAS.UserRole aba
                                        INNER JOIN BAS.Role abb ON aba.RoleId = abb.RoleId
                                        WHERE aba.UserId = aa.UserId
                                        AND abb.CompanyId = @CompanyId
                                        ORDER BY aba.RoleId
                                        FOR XML PATH('')
                                    ), 1, 1, '') RoleIds
                                ) ab
                                INNER JOIN BAS.Department ac ON aa.DepartmentId = ac.DepartmentId
                                WHERE 1=1
                                AND (aa.UserNo LIKE '%' + @SearchKey + '%' OR aa.UserName LIKE '%' + @SearchKey + '%')
                                AND ISNULL(ab.RoleIds, '-1') = ISNULL(b.RoleIds, '-1')
                                ORDER BY aa.UserNo
                                FOR JSON PATH, ROOT('data')
                            ) UserData
                            FROM BAS.[User] a
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ',' + CONVERT(nvarchar(50), bb.RoleId)
                                    FROM BAS.UserRole ba
                                    INNER JOIN BAS.Role bb ON ba.RoleId = bb.RoleId
                                    WHERE ba.UserId = a.UserId
                                    AND bb.CompanyId = @CompanyId
                                    ORDER BY ba.RoleId
                                    FOR XML PATH('')
                                ), 1, 1, '') RoleIds
                                , STUFF((
                                    SELECT ',' + CONVERT(nvarchar(50), bb.RoleName)
                                    FROM BAS.UserRole ba
                                    INNER JOIN BAS.Role bb ON ba.RoleId = bb.RoleId
                                    WHERE ba.UserId = a.UserId
                                    AND bb.CompanyId = @CompanyId
                                    ORDER BY ba.RoleId
                                    FOR XML PATH('')
                                ), 1, 1, '') RoleNames
                            ) b
                            INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                            WHERE 1=1
                            AND (a.UserNo LIKE '%' + @SearchKey + '%' OR a.UserName LIKE '%' + @SearchKey + '%')
                            ORDER BY b.RoleIds";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("SearchKey", SearchKey);

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

        #region //GetUserAuthority -- 取得使用者權限資料 -- Ben Ma 2022.05.17
        public string GetUserAuthority(int UserId, string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SystemId, a.SystemName, a.SortNumber, b.TotalFunction, c.UserTotalFunction
                            , (
                                SELECT aa.ModuleId, aa.ModuleName, ab.TotalModuleFunction, ac.UserTotalModuleFunction
                                FROM BAS.[Module] aa
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM BAS.FunctionDetail aba
                                    INNER JOIN BAS.[Function] abb ON aba.FunctionId = abb.FunctionId
                                    WHERE abb.ModuleId = aa.ModuleId
                                    AND abb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ab
                                OUTER APPLY (
                                    SELECT COUNT(DISTINCT aca.DetailId) UserTotalModuleFunction
                                    FROM BAS.RoleFunctionDetail aca
                                    INNER JOIN BAS.FunctionDetail acb ON aca.DetailId = acb.DetailId
                                    INNER JOIN BAS.[Function] acc ON acb.FunctionId = acc.FunctionId
                                    WHERE acc.ModuleId = aa.ModuleId
                                    AND aca.RoleId IN (
                                        SELECT RoleId
                                        FROM BAS.UserRole
                                        WHERE UserId = @UserId
                                    )
                                    AND acc.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ac
                                WHERE 1=1
                                AND aa.SystemId = a.SystemId
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) ModuleData
                            FROM BAS.[System] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunction
                                FROM BAS.FunctionDetail ba
                                INNER JOIN BAS.[Function] bb ON ba.FunctionId = bb.FunctionId
                                INNER JOIN BAS.[Module] bc ON bb.ModuleId = bc.ModuleId
                                WHERE bc.SystemId = a.SystemId
                                AND bb.FunctionName LIKE '%' + @SearchKey + '%'
                            ) b
                            OUTER APPLY (
                                SELECT COUNT(DISTINCT ca.DetailId) UserTotalFunction
                                FROM BAS.RoleFunctionDetail ca
                                INNER JOIN BAS.FunctionDetail cb ON ca.DetailId = cb.DetailId
                                INNER JOIN BAS.[Function] cc ON cb.FunctionId = cc.FunctionId
                                INNER JOIN BAS.[Module] cd ON cc.ModuleId = cd.ModuleId
                                WHERE cd.SystemId = a.SystemId
                                AND ca.RoleId IN (
                                    SELECT RoleId
                                    FROM BAS.UserRole
                                    WHERE UserId = @UserId
                                )
                                AND cc.FunctionName LIKE '%' + @SearchKey + '%'
                            ) c
                            WHERE 1=1
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("UserId", UserId);
                    dynamicParameters.Add("SearchKey", SearchKey);

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

        #region //GetUserDetailAuthority -- 取得使用者詳細權限資料 -- Ben Ma 2022.05.17
        public string GetUserDetailAuthority(int ModuleId, int UserId, string SearchKey, bool Grant)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.FunctionId, a.FunctionName, a.FunctionCode, a.SortNumber, b.TotalFunctionDetail, c.UserTotalFunctionDetail
                            , (
                                SELECT aa.DetailId, aa.DetailName, aa.DetailCode, ab.DetailAuthority
                                FROM BAS.FunctionDetail aa
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM BAS.RoleFunctionDetail aaa
                                        WHERE aaa.DetailId = aa.DetailId
                                        AND aaa.RoleId IN (
                                            SELECT RoleId
                                            FROM BAS.UserRole
                                            WHERE UserId = @UserId
                                        )
                                    ), 0) DetailAuthority
                                ) ab
                                WHERE 1=1
                                AND aa.FunctionId = a.FunctionId
                                AND ab.DetailAuthority >= @DetailAuthority
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) FunctionDetailData
                            FROM BAS.[Function] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunctionDetail
                                FROM BAS.FunctionDetail ba
                                WHERE ba.FunctionId = a.FunctionId
                            ) b
                            OUTER APPLY (
                                SELECT COUNT(DISTINCT ca.DetailId) UserTotalFunctionDetail
                                FROM BAS.RoleFunctionDetail ca
                                INNER JOIN BAS.FunctionDetail cb ON ca.DetailId = cb.DetailId
                                WHERE cb.FunctionId = a.FunctionId
                                AND ca.RoleId IN (
                                    SELECT RoleId
                                    FROM BAS.UserRole
                                    WHERE UserId = @UserId
                                )
                            ) c
                            WHERE a.ModuleId = @ModuleId
                            AND a.FunctionName LIKE '%' + @SearchKey + '%'
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("ModuleId", ModuleId);
                    dynamicParameters.Add("UserId", UserId);
                    dynamicParameters.Add("SearchKey", SearchKey);
                    dynamicParameters.Add("DetailAuthority", Grant ? 1 : 0);

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

        #region //GetUserCompanyAuthority -- 取得使用者公司權限 -- Ben Ma 2023.05.18
        public string GetUserCompanyAuthority()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT x.CompanyId, y.CompanyName
                            FROM (
                                SELECT a.CompanyId
                                FROM BAS.AuthorityUserCompany a
                                WHERE a.UserId = @CurrentUser
                                UNION ALL
                                SELECT b.CompanyId
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.UserId = @CurrentUser
                            ) x
                            INNER JOIN BAS.Company y ON x.CompanyId = y.CompanyId
                            WHERE y.Status = @Status
                            ORDER BY x.CompanyId";
                    dynamicParameters.Add("CurrentUser", CurrentUser);
                    dynamicParameters.Add("Status", "A");

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

        #region //GetUserVerify -- 使用者驗證 -- Ben Ma 2022.05.20
        public string GetUserVerify(int UserId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 1
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                            WHERE a.UserId = @UserId
                            AND a.Status = @Status
                            AND b.Status = @Status
                            AND c.Status = @Status";
                    dynamicParameters.Add("UserId", UserId);
                    dynamicParameters.Add("Status", "A");

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

        #region //GetUserCookieVerify -- Cookie驗證 -- Zoey 2022.09.30
        public string GetUserCookieVerify(string Key)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷BAS.UserLoginKey是否有資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.UserId
                            FROM BAS.UserLoginKey a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                            INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                            WHERE 1=1
                            AND a.KeyText = @KeyText
                            AND a.LoginIP = @LoginIP
                            AND a.ExpirationDate > @Now
                            AND b.Status = @Status
                            AND c.Status = @Status
                            AND d.Status = @Status";
                    dynamicParameters.Add("KeyText", Key);
                    dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                    dynamicParameters.Add("Now", DateTime.Now);
                    dynamicParameters.Add("Status", "A");
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    int userId = -1;
                    if (result.Count() <= 0)
                    {
                        HttpContext.Current.Response.Cookies["BmLogin"].Expires = DateTime.Now.AddDays(-1);
                        HttpContext.Current.Response.Cookies["BmLoginAcc"].Expires = DateTime.Now.AddDays(-1);
                    }
                    else
                    {
                        foreach (var item in result)
                        {
                            userId = item.UserId;
                        }
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.UserId, a.UserNo, a.UserName, a.Email
                            , b.CompanyId
                            , c.CompanyName
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                            WHERE 1=1
                            AND a.UserId = @UserId";
                    dynamicParameters.Add("UserId", userId);
                    var dataresult = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = dataresult
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

        #region //GetAuthorityVerify -- 權限驗證 -- Ben Ma 2022.05.20
        public string GetAuthorityVerify(string FunctionCode, string DetailCode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SortNumber, a.DetailCode, c.Authority
                            FROM BAS.FunctionDetail a
                            INNER JOIN BAS.[Function] b ON a.FunctionId = b.FunctionId
                            OUTER APPLY (
                                SELECT ISNULL((
                                    SELECT TOP 1 1
                                    FROM BAS.RoleFunctionDetail ca
                                    WHERE ca.DetailId = a.DetailId
                                    AND ca.RoleId IN (
                                        SELECT caa.RoleId
                                        FROM BAS.UserRole caa
                                        INNER JOIN BAS.[Role] cab ON caa.RoleId = cab.RoleId
                                        WHERE caa.UserId = @UserId
                                        AND cab.CompanyId = @CompanyId
                                    )
                                ), 0) Authority
                            ) c
                            WHERE a.[Status] = @Status
                            AND b.[Status] = @Status
                            AND b.FunctionCode = @FunctionCode
                            AND c.Authority > 0";
                    dynamicParameters.Add("UserId", CurrentUser);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("FunctionCode", FunctionCode);

                    if (DetailCode.Length > 0)
                    {
                        var queryCondition = new StringBuilder(sql);
                        var detailCodes = DetailCode.Split(',').ToList();

                        int i = 0;
                        foreach (var code in detailCodes)
                        {
                            queryCondition.Append(i == 0 ? " AND ( " : " OR ").Append("a.DetailCode LIKE '%' + @code").Append(i).Append(" + '%'");
                            dynamicParameters.Add("code" + i, code);
                            i++;

                            queryCondition.Append(i == detailCodes.Count ? " ) " : "");
                        }

                        sql = queryCondition.ToString();
                    }

                    sql += " ORDER BY a.SortNumber";
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

        #region //GetLog -- 取得系統日誌資料 -- Zoey 2022.05.23
        public string GetLog(int LogId, string UserNo, string Message, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.LogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Level, a.Message, a.Host, a.UserNo, a.TriggerUrl, a.ServerIP
                        , a.ClientIP, a.Source, FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') CreateDate";
                    sqlQuery.mainTables =
                        @"FROM BAS.Log a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LogId", @" AND a.LogId = @LogId", LogId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Message", @" AND a.Message LIKE '%' + @Message + '%'", Message);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LogId DESC";
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

        #region //GetMailServer -- 取得郵件伺服器資料 -- Zoey 2022.05.24
        public string GetMailServer(int ServerId, string Host
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ServerId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.Host, a.Port, a.SendMode, a.Account, a.Password";
                    sqlQuery.mainTables =
                        @"FROM BAS.MailServer a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ServerId", @" AND a.ServerId = @ServerId", ServerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Host", @" AND a.Host LIKE '%' + @Host + '%'", Host);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ServerId DESC";
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

        #region //GetMailContact -- 取得郵件聯絡人資料 -- Zoey 2022.05.26
        public string GetMailContact(string Contact
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ContactId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ContactName, a.Email";
                    sqlQuery.mainTables =
                        @"FROM BAS.MailContact a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Contact", @" AND (a.ContactName LIKE '%' + @Contact + '%' OR a.Email LIKE '%' + @Contact + '%')", Contact);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ContactId";
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

        #region //GetContact -- 取得聯絡人資料 -- Zoey 2022.05.27
        public string GetContact(int CompanyId, int DepartmentId, string Contact, string Mode
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string mainKey = "", columns = "", mainTables = "", queryCondition = "", defaultOrder = "";

                    switch (Mode)
                    {
                        case "User":
                            mainKey = "a.UserId";
                            columns = @", a.UserId, a.UserNo, a.UserName, a.Gender, a.Email, CASE a.JobType WHEN '管理制' THEN a.Job ELSE '' END Job
                                        , b.DepartmentId, b.DepartmentNo, b.DepartmentName
                                        , c.CompanyId, c.CompanyNo, c.CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon
                                        , 'User' Source";
                            mainTables =
                                @"FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                            queryCondition =
                                @" AND a.Status = @Status
                                AND a.Email IS NOT NULL
                                AND a.Email <> ''";
                            dynamicParameters.Add("Status", "A");
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND b.DepartmentId = @DepartmentId", DepartmentId);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Contact", @" AND (a.UserNo LIKE '%' + @Contact + '%' OR a.UserName LIKE '%' + @Contact + '%' OR a.Email LIKE '%' + @Contact + '%')", Contact);
                            defaultOrder = "c.CompanyNo, b.DepartmentNo, a.UserNo";
                            break;
                        case "Contact":
                            mainKey = "a.ContactId";
                            columns = @", a.ContactId UserId, '' UserNo, a.ContactName UserName, '' Gender, a.Email, '' Job
                                        , -1 DepartmentId, '' DepartmentNo, '' DepartmentName
                                        , -1 CompanyId, '' CompanyNo, '' CompanyName, -1 LogoIcon
                                        , 'Contact' Source";
                            mainTables =
                                @"FROM BAS.MailContact a";
                            queryCondition = " AND a.CompanyId = @CurrentCompany";
                            dynamicParameters.Add("CurrentCompany", CurrentCompany);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Contact", @" AND (a.ContactName LIKE '%' + @Contact + '%' OR a.Email LIKE '%' + @Contact + '%')", Contact);
                            defaultOrder = "a.ContactName";
                            break;
                    }

                    sqlQuery.mainKey = mainKey;
                    sqlQuery.auxKey = "";
                    sqlQuery.columns = columns;
                    sqlQuery.mainTables = mainTables;
                    sqlQuery.auxTables = "";
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : defaultOrder;
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

        #region //GetMailUser -- 取得聯絡人資料 -- Ben Ma 2023.08.03
        public string GetMailUser()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT *
                            FROM (
                                SELECT 'user:' + CAST(a.UserId AS NVARCHAR) Id, a.Email
                                , CASE a.JobType 
                                    WHEN '管理制' THEN b.DepartmentName + a.Job + '-' + a.UserNo + ' ' + a.UserName
                                    ELSE b.DepartmentName + '-' + a.UserNo + ' ' + a.UserName 
                                END DisplayName
                                , c.CompanyId, a.UserNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                                WHERE a.Status = @Status
                                AND a.Email IS NOT NULL
                                AND a.Email <> ''
                                UNION ALL
                                SELECT 'contact:' + CAST(a.ContactId AS NVARCHAR) Id, a.Email, a.ContactName DisplayName
                                , 999 CompanyId, '' UserNo
                                FROM BAS.MailContact a
                                WHERE 1=1
                                AND a.CompanyId = @CurrentCompany
                            ) x
                            WHERE 1=1";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);
                    sql += @" ORDER BY x.CompanyId, x.UserNo";

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

        #region //GetMailTemplate -- 取得郵件樣板資料 -- Zoey 2022.05.27
        public string GetMailTemplate(int MailId, int ServerId, string MailName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ServerId, a.MailName, a.MailSubject, a.MailContent
                        , b.Host, b.Port, b.SendMode, b.Account, b.Password
                        , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc";
                    sqlQuery.mainTables =
                        @"FROM BAS.Mail a
                        LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                        OUTER APPLY (
                            SELECT STUFF((
                                SELECT ',' + CASE WHEN ca.ContactId IS NULL THEN 'user:' + CAST(cc.UserId AS NVARCHAR) ELSE 'contact:' + CAST(cb.ContactId AS NVARCHAR) END
                                FROM BAS.MailUser ca
                                LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId
                                WHERE ca.MailId = a.MailId
                                AND ca.MailUserType = 'F'
                                FOR XML PATH('')
                            ), 1, 1, '') MailFrom
                        ) c
                        OUTER APPLY (
                            SELECT STUFF((
                                SELECT ',' + CASE WHEN da.ContactId IS NULL THEN 'user:' + CAST(dc.UserId AS NVARCHAR) ELSE 'contact:' + CAST(db.ContactId AS NVARCHAR) END
                                FROM BAS.MailUser da
                                LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId
                                WHERE da.MailId = a.MailId
                                AND da.MailUserType = 'T'
                                FOR XML PATH('')
                            ), 1, 1, '') MailTo
                        ) d
                        OUTER APPLY (
                            SELECT STUFF((
                                SELECT ',' + CASE WHEN ea.ContactId IS NULL THEN 'user:' + CAST(ec.UserId AS NVARCHAR) ELSE 'contact:' + CAST(eb.ContactId AS NVARCHAR) END
                                FROM BAS.MailUser ea
                                LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId
                                WHERE ea.MailId = a.MailId
                                AND ea.MailUserType = 'C'
                                FOR XML PATH('')
                            ), 1, 1, '') MailCc
                        ) e
                        OUTER APPLY (
                            SELECT STUFF((
                                SELECT ',' + CASE WHEN fa.ContactId IS NULL THEN 'user:' + CAST(fc.UserId AS NVARCHAR) ELSE 'contact:' + CAST(fb.ContactId AS NVARCHAR) END
                                FROM BAS.MailUser fa
                                LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId
                                WHERE fa.MailId = a.MailId
                                AND fa.MailUserType = 'B'
                                FOR XML PATH('')
                            ), 1, 1, '') MailBcc
                        ) f";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MailId", @" AND a.MailId = @MailId", MailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ServerId", @" AND a.ServerId = @ServerId", ServerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MailName", @" AND a.MailName LIKE '%' + @MailName + '%'", MailName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MailId";
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

        #region //GetMailList -- 取得郵件寄送資料
        public string GetMailList(int MailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.MailId, a.ServerId, a.MailName, a.MailSubject, a.MailContent
                            , b.Host, b.Port, b.SendMode, b.Account, b.Password
                            , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                            FROM BAS.Mail a
                            LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                        CASE cc.JobType 
                                            WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                            ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                        END
                                    ELSE cb.ContactName + ':' + cb.Email END
                                    FROM BAS.MailUser ca
                                    LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                    LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                    LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                    WHERE ca.MailId = a.MailId
                                    AND ca.MailUserType = 'F'
                                    FOR XML PATH('')
                                ), 1, 1, '') MailFrom
                            ) c
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                        CASE dc.JobType 
                                            WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                            ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                        END
                                    ELSE db.ContactName + ':' + db.Email END
                                    FROM BAS.MailUser da
                                    LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                    LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                    LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                    WHERE da.MailId = a.MailId
                                    AND da.MailUserType = 'T'
                                    FOR XML PATH('')
                                ), 1, 1, '') MailTo
                            ) d
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                        CASE ec.JobType 
                                            WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                            ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                        END
                                    ELSE eb.ContactName + ':' + eb.Email END
                                    FROM BAS.MailUser ea
                                    LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                    LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                    LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                    WHERE ea.MailId = a.MailId
                                    AND ea.MailUserType = 'C'
                                    FOR XML PATH('')
                                ), 1, 1, '') MailCc
                            ) e
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                        CASE fc.JobType 
                                            WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                            ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                        END
                                    ELSE fb.ContactName + ':' + fb.Email END
                                    FROM BAS.MailUser fa
                                    LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                    LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                    LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                    WHERE fa.MailId = a.MailId
                                    AND fa.MailUserType = 'B'
                                    FOR XML PATH('')
                                ), 1, 1, '') MailBcc
                            ) f
                            WHERE 1=1
                            AND a.MailId = @MailId";
                    dynamicParameters.Add("MailId", MailId);

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

        #region //GetMailSendSetting -- 取得郵件寄送設定 -- Ben Ma 2023.04.18
        public string GetMailSendSetting(int SettingId, int MailId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SettingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MailId, a.SettingSchema, a.SettingNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.MailSendSetting a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SettingId", @" AND a.SettingId = @SettingId", SettingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MailId", @" AND a.MailId = @MailId", MailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (
                            a.SettingSchema LIKE '%' + @SearchKey + '%'
                            OR a.SettingNo LIKE '%' + @SearchKey + '%'
                        )", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MailId";
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

        #region //GetCalendar -- 取得行事曆資料 -- Zoey 2022.06.06
        public string GetCalendar(string CalendarDate)
        {
            try
            {
                DateTime StartDate = Convert.ToDateTime(Convert.ToDateTime(CalendarDate).ToString("yyyy-MM-01"));
                DateTime EndDate = Convert.ToDateTime(Convert.ToDateTime(CalendarDate).ToString("yyyy-MM-01")).AddMonths(1).AddDays(-1);

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DateTime currentStart = Convert.ToDateTime(GetCurrentWeekDates(StartDate)[0]);
                    DateTime currentEnd = Convert.ToDateTime(GetCurrentWeekDates(EndDate)[6]);

                    sql = @"SELECT
                            (
                                SELECT a.CalendarId, Format(a.CalendarDate, 'yyyy-MM-dd 00:00:00') start_date
                                , Format(DATEADD(DAY, 1,a.CalendarDate), 'yyyy-MM-dd 00:00:00') end_date
                                , a.CalendarDesc text, a.DateType type
                                FROM BAS.Calendar a
                                WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CurrentStart", @" AND a.CalendarDate >= @CurrentStart", currentStart);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CurrentEnd", @" AND a.CalendarDate <= @CurrentEnd", currentEnd);
                    
                    sql += @"  FOR JSON PATH, ROOT('data')
                             ) CalendarDetail";

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

        #region //GetCalendarInfo -- 取得行事曆資料 -- Ben Ma 2023.11.03
        public string GetCalendarInfo(int CalendarId, string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.CalendarId, FORMAT(a.CalendarDate, 'yyyy-MM-dd') CalendarDate
                            , FORMAT(a.CalendarDate, 'yyyy-MM-dd') + ' (' +
                            CASE DATEPART(weekday, a.CalendarDate)
                                WHEN 1 THEN '日'
                                WHEN 2 THEN '一'
                                WHEN 3 THEN '二'
                                WHEN 4 THEN '三'
                                WHEN 5 THEN '四'
                                WHEN 6 THEN '五'
                                WHEN 7 THEN '六'
                            END + ')' CalendarWeekDate
                            , a.CalendarDesc, a.DateType
                            FROM BAS.Calendar a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CalendarId", @" AND a.CalendarId = @CalendarId", CalendarId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CalendarDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CalendarDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sql += @" ORDER BY a.CalendarDate";

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

        #region //GetFile -- 取得檔案資料 -- Ben Ma 2022.06.13
        public string GetFile(int FileId, int UploadUser, string DeleteStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.FileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.[FileName], a.FileExtension, a.FileSize, a.ClientIP, a.Source
                        , a.DeleteStatus, Format(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') UploadDate, Format(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss') DeleteDate
                        , b.UserNo + ' ' + b.UserName UploadUser, c.UserNo + ' ' + c.UserName DeleteUser";
                    if (FileId > 0) sqlQuery.columns += @", a.FileContent";
                    sqlQuery.mainTables =
                        @"FROM BAS.[File] a
                        LEFT JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        LEFT JOIN BAS.[User] c ON a.LastModifiedBy = c.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FileId", @" AND a.FileId = @FileId", FileId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UploadUser", @" AND a.CreateBy = @UploadUser", UploadUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeleteStatus", @" AND a.DeleteStatus = @DeleteStatus", DeleteStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.FileId DESC";
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

        #region //GetUserLogin -- 取得使用者登入紀錄 -- Zoey 2022.09.30
        public string GetUserLogin(int LogId, int CompanyId, int DepartmentId
            , string UserNo, string UserName, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.LogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.LoginSite, a.LoginIP, FORMAT(a.LoginDate, 'yyyy-MM-dd HH:mm') LoginDate, a.UserId
                          , b.UserNo + '-' + b.UserName UserWithNo
                          , d.LogoIcon";
                    sqlQuery.mainTables =
                          @"FROM BAS.UserLoginLog a
                            INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                            INNER JOIN BAS.Department c ON c.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company d ON d.CompanyId = c.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LogId", @" AND a.LogId = @LogId", LogId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND c.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND b.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND b.UserName LIKE '%' + @UserName + '%'", UserName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.LoginDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.LoginDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LogId DESC";
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

        #region //GetSubSystem -- 取得子系統資料 -- Ben Ma 2022.11.23
        public string GetSubSystem(int SubSystemId, string SubSystemCode, string SubSystemName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SubSystemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SubSystemCode, a.SubSystemName, a.KeyText, a.Status";
                    sqlQuery.mainTables =
                        @"FROM BAS.SubSystem a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubSystemId", @" AND a.SubSystemId = @SubSystemId", SubSystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubSystemCode", @" AND a.SubSystemCode LIKE '%' + @SubSystemCode + '%'", SubSystemCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubSystemName", @" AND a.SubSystemName LIKE '%' + @SubSystemName + '%'", SubSystemName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SubSystemId";
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

        #region //GetRandomSystemKey -- 取得隨機系統金鑰 -- Ben Ma 2022.11.23
        public string GetRandomSystemKey()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    bool check = true;
                    string cipherText = "";

                    while (check)
                    {
                        var plainText = Guid.NewGuid();
                        cipherText = BaseHelper.Sha256Encrypt(plainText.ToString());

                        #region //判斷隨機金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", cipherText);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        check = result.Count() > 0;
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = cipherText
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

        #region //GetSubSystemUser -- 取得子系統使用者資料 -- Ben Ma 2022.11.23
        public string GetSubSystemUser(int SubSystemUserId, int SubSystemId, int CompanyId
            , string Departments, string UserNo, string UserName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SubSystemUserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SubSystemId, a.UserId, a.[Status]
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentId, c.DepartmentNo, c.DepartmentName
                        , d.CompanyId, d.CompanyNo, d.CompanyName, ISNULL(d.LogoIcon, -1) LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM BAS.SubSystemUser a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubSystemUserId", @" AND a.SubSystemUserId = @SubSystemUserId", SubSystemUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SubSystemId", @" AND a.SubSystemId = @SubSystemId", SubSystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND c.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND b.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND b.UserName LIKE '%' + @UserName + '%'", UserName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SubSystemId, d.CompanyId, b.UserNo";
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

        #region //GetPasswordSetting -- 取得密碼參數設定資料 -- Ben Ma 2022.12.26
        public string GetPasswordSetting()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";

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

        #region //GetPushSubscription -- 取得推播訂閱資料 -- Ben Ma 2023.01.04
        public string GetPushSubscription(string UserNo, string Users)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.UserId
                            , (
                                SELECT z.PushSubscriptionId, z.ApiEndpoint, z.Keysp256dh, z.Keysauth
                                FROM BAS.PushSubscription z
                                WHERE z.UserId = a.UserId
                                ORDER BY z.CreateDate
                                FOR JSON PATH, ROOT('data')
                            ) PushInfo
                            FROM BAS.PushSubscription a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserNo", @" AND b.UserNo = @UserNo", UserNo);
                    if (Users.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Users", @" AND b.UserId IN @Users", Users.Split(','));

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

        #region //GetCurrentWeekDates -- 找當週所有日期

        private List<string> GetCurrentWeekDates(DateTime date)
        {
            List<string> dates = new List<string>();

            switch (date.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    for (int i = 0; i < 7; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Monday:
                    for (int i = -1; i < 6; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Tuesday:
                    for (int i = -2; i < 5; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Wednesday:
                    for (int i = -3; i < 4; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Thursday:
                    for (int i = -4; i < 3; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Friday:
                    for (int i = -5; i < 2; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
                case DayOfWeek.Saturday:
                    for (int i = -6; i < 1; i++)
                    {
                        dates.Add(date.AddDays(i).ToString("yyyy-MM-dd"));
                    }
                    break;
            }

            return dates;
        }
        #endregion

        #region //GetNotificationMail -- 取得系統通知信件資料 -- Ben Ma 2023.08.09
        public string GetNotificationMail(string UserNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.MailSubject, a.MailContent
                            , b.Host, b.Port, b.SendMode, b.Account, b.Password
                            , ISNULL(c.MailFrom, '') MailFrom
                            , ISNULL(d.DisplayName, '') DisplayName, ISNULL(d.Email, '') Email
                            FROM BAS.Mail a
                            LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                        CASE cc.JobType 
                                            WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                            ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                        END
                                    ELSE cb.ContactName + ':' + cb.Email END
                                    FROM BAS.MailUser ca
                                    LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                    LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                    LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                    WHERE ca.MailId = a.MailId
                                    AND ca.MailUserType = 'F'
                                    FOR XML PATH('')
                                ), 1, 1, '') MailFrom
                            ) c
                            OUTER APPLY (
                                SELECT a.Email
                                , CASE a.JobType 
                                    WHEN '管理制' THEN b.DepartmentName + a.Job + '-' + a.UserName
                                    ELSE b.DepartmentName + '-' + a.UserName 
                                END DisplayName
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                                WHERE a.Status = @Status
                                AND a.UserNo = @UserNo
                                AND a.Email IS NOT NULL
                                AND a.Email <> ''
                            ) d
                            WHERE a.MailId IN (
                                SELECT z.MailId
                                FROM BAS.MailSendSetting z
                                WHERE z.SettingSchema = @SettingSchema
                                AND z.SettingNo = @SettingNo
                            )";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("UserNo", UserNo);
                    dynamicParameters.Add("SettingSchema", "NotificationMailAdvice");
                    dynamicParameters.Add("SettingNo", "Y");

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

        #region //GetNotificationSetting -- 取得系統通知個人設定資料 -- Ben Ma 2023.08.10
        public string GetNotificationSetting()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                            FROM BAS.NotificationSetting a
                            WHERE a.UserId = @CurrentUser";
                    dynamicParameters.Add("CurrentUser", CurrentUser);

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

        #region //GetNotificationLog -- 取得系統通知紀錄資料 -- Ben Ma 2023.08.10
        public string GetNotificationLog(string TriggerFunction, string ReadStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.LogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.TriggerFunction, a.LogTitle, a.LogContent
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') LogDate, FORMAT(a.CreateDate, 'MM-dd HH:mm') LogShortDate
                        , a.ReadStatus, ISNULL(FORMAT(a.ReadDate, 'yyyy-MM-dd HH:mm:ss'), '') ReadDate, ISNULL(a.ReadIP, '') ReadIP";
                    sqlQuery.mainTables =
                        @"FROM BAS.NotificationLog a
                        LEFT JOIN BAS.[User] b ON a.UserNo = b.UserNo";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.UserId = @CurrentUser";
                    dynamicParameters.Add("CurrentUser", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReadStatus", @" AND a.ReadStatus = @ReadStatus", ReadStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LogId DESC";
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

        #region //GetUser -- 取得新版使用者資料 -- Yi 2023.09.01
        public string GetUser(int UserId, string UserType, int CompanyId, string Departments
            , string InnerCode, string DisplayName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserType, b.UserTypeInfo, a.Account, a.Password, a.InnerCode, a.DisplayName
                        , a.PasswordStatus, c.StatusName PasswordStatusName, FORMAT(a.PasswordExpire, 'yyyy-MM-dd HH:mm:ss') PasswordExpire, a.PasswordMistake
                        , a.Status, d.StatusName
                        , ISNULL(LAG(a.UserId) OVER (ORDER BY a.InnerCode), -1) PreviousId, ISNULL(LEAD(a.UserId) OVER (ORDER BY a.InnerCode), -1) NextId
                        , (
                            SELECT aa.DepartmentId, aa.Gender, aa.Email, aa.Job, aa.JobType, aa.EmployeeStatus
                            , ab.DepartmentNo, ab.DepartmentName
                            , ac.CompanyId, ac.CompanyNo, ac.CompanyName, ac.LogoIcon
                            FROM BAS.UserEmployee aa
                            INNER JOIN BAS.Department ab ON aa.DepartmentId = ab.DepartmentId
                            INNER JOIN BAS.Company ac ON ab.CompanyId = ac.CompanyId
                            WHERE aa.UserId = a.UserId
                            FOR JSON PATH, ROOT('data')
                        ) UserEmployee
                        , (
                            SELECT aa.Contact, aa.ContactPhone, aa.ContactAddress
                            FROM BAS.UserCustomer aa
                            WHERE aa.UserId = a.UserId
                            FOR JSON PATH, ROOT('data')
                        ) UserCustomer
                        , (
                            SELECT aa.Contact, aa.ContactPhone, aa.ContactAddress
                            FROM BAS.UserSupplier aa
                            WHERE aa.UserId = a.UserId
                            FOR JSON PATH, ROOT('data')
                        ) UserSupplier";
                    sqlQuery.mainTables =
                        @"FROM BAS.UserInfo a
                        OUTER APPLY (
                            SELECT STUFF((
                                SELECT ',' + ac.TypeName
                                FROM BAS.UserInfo aa
                                OUTER APPLY STRING_SPLIT(a.UserType, ',') ab
                                LEFT JOIN BAS.[Type] ac ON ac.TypeNo = ab.value AND ac.TypeSchema = 'UserInfo.UserType' 
                                WHERE aa.UserId = a.UserId
                                FOR XML PATH('')
                            ), 1, 1, '') UserTypeInfo
                        ) b
                        LEFT JOIN BAS.[Status] c ON c.StatusNo = a.PasswordStatus AND c.StatusSchema = 'Boolean'
                        LEFT JOIN BAS.[Status] d ON d.StatusNo = a.Status AND d.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserType", @" AND a.UserType = @UserType", UserType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM BAS.UserEmployee aa
                                                                                                            INNER JOIN BAS.Department ab ON aa.DepartmentId = ab.DepartmentId
                                                                                                            INNER JOIN BAS.Company ac ON ab.CompanyId = ac.CompanyId
                                                                                                            WHERE aa.UserId = a.UserId
                                                                                                            AND ab.CompanyId = @CompanyId
                                                                                                        )", CompanyId);
                    if (Departments.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND EXISTS (
                                                                                                                    SELECT TOP 1 1
                                                                                                                    FROM BAS.UserEmployee aa
                                                                                                                    INNER JOIN BAS.Department ab ON aa.DepartmentId = ab.DepartmentId
                                                                                                                    WHERE aa.UserId = a.UserId
                                                                                                                    AND ab.DepartmentId IN @Departments
                                                                                                               )", Departments.Split(','));
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InnerCode", @" AND a.InnerCode LIKE '%' + @InnerCode + '%'", InnerCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DisplayName", @" AND a.DisplayName LIKE '%' + @DisplayName + '%'", DisplayName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));                    
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.InnerCode";
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

        #region //GetUserEmployee -- 取得使用者內部員工設定檔資料 -- Yi 2023.09.05
        public string GetUserEmployee(int UserId, string EmployeeStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DepartmentId, e.DepartmentName, a.Gender, c.TypeName GenderName
                        , a.Email, a.Job, a.JobType, a.EmployeeStatus, d.StatusName EmployeeStatusName
                        , b.UserType, b.Account, b.InnerCode, b.DisplayName, b.[Status]";
                    sqlQuery.mainTables =
                        @"FROM BAS.UserEmployee a
                        INNER JOIN BAS.UserInfo b ON b.UserId = a.UserId
                        INNER JOIN BAS.[Type] c ON c.TypeNo = a.Gender AND c.TypeSchema = 'UserEmployee.Gender'
                        INNER JOIN BAS.[Status] d ON d.StatusNo = a.EmployeeStatus AND d.StatusSchema = 'UserEmployee.EmployeeStatus'
                        INNER JOIN BAS.Department e ON e.DepartmentId = a.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserId DESC";
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

        #region //GetMamoTeams -- 取得Mamo團隊資料 -- Ann 2023-01-17
        public string GetMamoTeams(int TeamId, string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別資料是否有誤
                    sql = @"SELECT a.CompanyId
                            FROM BAS.Company a 
                            WHERE a.CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", Company);

                    var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                    int CompanyId = -1;
                    foreach (var item in CompanyResult)
                    {
                        CompanyId = item.CompanyId;
                    }
                    #endregion

                    sql = @"SELECT a.TeamId, a.MamoTeamId, a.TeamNo, a.TeamName, a.Remark
                            FROM MAMO.Teams a 
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CompanyId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TeamId", @" AND a.TeamId = @TeamId", TeamId);

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

        #region //GetMamoChannels -- 取得Mamo頻道資料 -- Ann 2023-01-17
        public string GetMamoChannels(int ChannelId, string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (ChannelId <= 0) throw new SystemException("頻道不能為空!!");
                    if (Company.Length <= 0) throw new SystemException("公司別不能為空!!");

                    #region //確認公司資訊
                    if (Company.Length > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a 
                                WHERE a.CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", Company);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        foreach (var item in CompanyResult)
                        {
                            CurrentCompany = item.CompanyId;
                        }
                    }
                    #endregion

                    sql = @"SELECT a.ChannelId, a.MamoChannelId, a.TeamId, a.ChannelName, a.ChannelNo
                            FROM MAMO.Channels a 
                            INNER JOIN MAMO.Teams b ON a.TeamId = b.TeamId
                            WHERE a.ChannelId = @ChannelId";
                    dynamicParameters.Add("ChannelId", ChannelId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("查無頻道資料!!");

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

        #region //GetMamoChannels -- 取得Mamo頻道資料 -- Ann 2023-01-17
        public List<string> GetAllUser(string Company)
        {
            List<string> userNoList = new List<string>();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (Company.Length <= 0) throw new SystemException("公司別不能為空!!");

                    #region //確認公司資訊
                    if (Company.Length > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a 
                                WHERE a.CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", Company);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        foreach (var item in CompanyResult)
                        {
                            CurrentCompany = item.CompanyId;
                        }
                    }
                    #endregion

                    sql = @"SELECT a.UserNo
                            FROM BAS.[User] a 
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            WHERE b.CompanyId = @CompanyId
                            AND a.UserStatus = 'F'
                            AND a.UserNo LIKE 'E%'
                            AND a.[Status] = 'A'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in result)
                    {
                        userNoList.Add(item.UserNo);
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

            return userNoList;
        }
        #endregion
        #endregion

        #region //Add
        #region //AddApiKey -- Api金鑰新增 -- Ben Ma 2022.04.27
        public string AddApiKey(string KeyText, string Purpose)
        {
            try
            {
                if (KeyText.Length <= 0) throw new SystemException("【金鑰】不能為空!");
                if (KeyText.Length > 64) throw new SystemException("【金鑰】長度錯誤!");
                if (Purpose.Length <= 0) throw new SystemException("【用途】不能為空!");
                if (Purpose.Length > 50) throw new SystemException("【用途】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷金鑰是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", KeyText);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【金鑰】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.ApiKey (CompanyId, KeyText, Purpose, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.KeyId
                                VALUES (@CompanyId, @KeyText, @Purpose, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                KeyText,
                                Purpose,
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

        #region //AddApiKeyFunction -- Api適用功能新增 -- Ben Ma 2022.04.29
        public string AddApiKeyFunction(int KeyId, string FunctionCode, string Purpose)
        {
            try
            {
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 100) throw new SystemException("【功能代碼】長度錯誤!");
                if (Purpose.Length <= 0) throw new SystemException("【用途】不能為空!");
                if (Purpose.Length > 50) throw new SystemException("【用途】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷Api金鑰資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("Api金鑰資料錯誤!");
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKeyFunction
                                WHERE KeyId = @KeyId
                                AND FunctionCode = @FunctionCode";
                        dynamicParameters.Add("KeyId", KeyId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.ApiKeyFunction (KeyId, FunctionCode, Purpose
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FunctionId
                                VALUES (@KeyId, @FunctionCode, @Purpose
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                KeyId,
                                FunctionCode,
                                Purpose,
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

        #region //AddRole -- 角色資料新增 -- Ben Ma 2022.05.06
        public string AddRole(string RoleName, string AdminStatus)
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
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Role
                                WHERE CompanyId = @CompanyId
                                AND RoleName = @RoleName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoleName", RoleName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Role (CompanyId, RoleName, AdminStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoleId
                                VALUES (@CompanyId, @RoleName, @AdminStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                RoleName,
                                AdminStatus,
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

        #region //AddMailServer -- 郵件伺服器新增 -- Zoey 2022.05.24
        public string AddMailServer(string Host, string SendMode, int Port, string Account, string Password)
        {
            try
            {
                if (Host.Length <= 0) throw new SystemException("【主機】不能為空!");
                if (Host.Length > 50) throw new SystemException("【主機】長度錯誤!");
                if (Port < 0) throw new SystemException("【埠號】不可小於0!");
                if (Port > 65535) throw new SystemException("【埠號】不可大於65535!");
                if (Account.Length <= 0) throw new SystemException("【帳號】不能為空!");
                if (Account.Length > 100) throw new SystemException("【帳號】長度錯誤!");
                if (Password.Length <= 0) throw new SystemException("【密碼】不能為空!");
                if (Password.Length > 100) throw new SystemException("【密碼】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷寄送模式資料是否錯誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "MailServer.SendMode");
                        dynamicParameters.Add("TypeNo", SendMode);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【寄送模式】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.MailServer (CompanyId, Host
                                , Port, SendMode, Account, Password
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ServerId
                                VALUES (@CompanyId, @Host, @Port, @SendMode, @Account, @Password
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                Host,
                                Port,
                                SendMode,
                                Account,
                                Password,
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

        #region //AddMailContact -- 郵件聯絡人資料新增 -- Zoey 2022.05.26
        public string AddMailContact(string ContactName, string Email)
        {
            try
            {
                if (ContactName.Length <= 0) throw new SystemException("【聯絡人姓名】不能為空!");
                if (ContactName.Length > 50) throw new SystemException("【聯絡人姓名】長度錯誤!");
                if (Email.Length <= 0) throw new SystemException("【聯絡人信箱】不能為空!");
                if (Email.Length > 100) throw new SystemException("【聯絡人信箱】長度錯誤!");
                if (!Regex.IsMatch(Email, RegexHelper.Email, RegexOptions.IgnoreCase)) throw new SystemException("【聯絡人信箱】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.MailContact (CompanyId, ContactName, Email
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ContactId
                                VALUES (@CompanyId, @ContactName, @Email
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ContactName,
                                Email,
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

        #region //AddMailTemplate -- 郵件樣板新增 -- Zoey 2022.05.27
        public string AddMailTemplate(int ServerId, string MailName, string MailFrom, string MailTo
            , string MailCc, string MailBcc, string MailSubject, string MailContent)
        {
            try
            {
                if (ServerId < 0) throw new SystemException("【主機】不能為空!");
                if (MailName.Length <= 0) throw new SystemException("【郵件名稱】不能為空!");
                if (MailFrom.Length <= 0) throw new SystemException("【寄件人】不能為空!");
                if (MailTo.Length <= 0) throw new SystemException("【收件人】不能為空!");
                if (MailSubject.Length <= 0) throw new SystemException("【郵件主旨】不能為空!");
                if (MailContent.Length <= 0) throw new SystemException("【郵件內容】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Mail (ServerId, MailName, MailSubject, MailContent
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MailId
                                VALUES (@ServerId, @MailName, @MailSubject, @MailContent
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ServerId,
                                MailName,
                                MailSubject,
                                MailContent,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int MailId = -1;
                        foreach (var item in insertResult)
                        {
                            MailId = Convert.ToInt32(item.MailId);
                        }

                        #region //寄件人
                        string[] mailFrom = MailFrom.Split(',');
                        for (int i = 0; i < mailFrom.Length; i++)
                        {
                            string[] mailUser = mailFrom[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "F",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "F",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //收件人
                        string[] mailTo = MailTo.Split(',');
                        for (int i = 0; i < mailTo.Length; i++)
                        {
                            string[] mailUser = mailTo[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "T",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "T",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //副本
                        string[] mailCc = MailCc.Split(',');
                        for (int i = 0; i < mailCc.Length; i++)
                        {
                            string[] mailUser = mailCc[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "C",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "C",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //秘密副本
                        string[] mailBcc = MailBcc.Split(',');
                        for (int i = 0; i < mailBcc.Length; i++)
                        {
                            string[] mailUser = mailBcc[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "B",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "B",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
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

        #region //AddMailSendSetting -- 郵件寄送設定新增 -- Ben Ma 2023.04.18
        public string AddMailSendSetting(int MailId, string SettingSchema, string SettingNo)
        {
            try
            {
                if (SettingSchema.Length <= 0) throw new SystemException("【寄送綱要】不能為空!");
                if (SettingSchema.Length > 100) throw new SystemException("【寄送綱要】長度錯誤!");
                if (SettingNo.Length <= 0) throw new SystemException("【寄送綱要】不能為空!");
                if (SettingNo.Length > 10) throw new SystemException("【寄送綱要】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷郵件樣板資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Mail
                                WHERE MailId = @MailId";
                        dynamicParameters.Add("MailId", MailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("郵件樣板資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.MailSendSetting (MailId, SettingSchema, SettingNo
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SettingId
                                VALUES (@MailId, @SettingSchema, @SettingNo
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MailId,
                                SettingSchema,
                                SettingNo,
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

        #region //AddCalendar -- 行事曆資料新增 -- Zoey 2022.06.09
        public string AddCalendar(string Year)
        {
            try
            {
                if (Year.Length < 0) throw new SystemException("【年份】不能為空!");

                DateTime StartDate = Convert.ToDateTime(Year + "-01-01 00:00:00.000");
                DateTime EndDate = Convert.ToDateTime(Year + "-12-31 00:00:00.000");
                string CalendarDesc = ""
                    , DateType = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        while (StartDate <= EndDate)
                        {
                            #region //判斷行事曆資料是否存在
                            bool exist = false;

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Calendar
                                    WHERE CalendarDate = @CalendarDate";
                            dynamicParameters.Add("@CalendarDate", StartDate);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            exist = result.Count() > 0;
                            #endregion

                            if (!exist)
                            {
                                switch (StartDate.DayOfWeek)
                                {
                                    case DayOfWeek.Saturday:
                                        CalendarDesc = "休息日";
                                        DateType = "R";
                                        break;
                                    case DayOfWeek.Sunday:
                                        CalendarDesc = "例假日";
                                        DateType = "M";
                                        break;
                                    default:
                                        CalendarDesc = "上班日";
                                        DateType = "W";
                                        break;
                                }

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO BAS.Calendar (CompanyId, CalendarDate, CalendarDesc, DateType
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.CalendarId
                                        VALUES (@CompanyId, @CalendarDate, @CalendarDesc, @DateType
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        CalendarDate = StartDate,
                                        CalendarDesc,
                                        DateType,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }

                            StartDate = StartDate.AddDays(1);
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

        #region //AddFile -- 檔案資料新增 -- Ben Ma 2022.06.10
        public string AddFile(string FileName, byte[] FileContent, string FileExtension, int FileSize, string ClientIP, string Source)
        {
            try
            {
                if (FileName.Length <= 0) throw new SystemException("【檔案名稱】不能為空!");
                if (FileContent.Length <= 0) throw new SystemException("【檔案內容】不能為空!");
                if (FileExtension.Length <= 0) throw new SystemException("【副檔名】不能為空!");
                if (FileSize <= 0) throw new SystemException("【檔案大小】不能為零!");
                if (ClientIP.Length <= 0) throw new SystemException("【使用者IP】不能為空!");
                if (Source.Length <= 0) throw new SystemException("【來源】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                                , ClientIP, Source, DeleteStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FileId
                                VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                                , @ClientIP, @Source, @DeleteStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                FileName,
                                FileContent,
                                FileExtension,
                                FileSize,
                                ClientIP,
                                Source,
                                DeleteStatus = "N",
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

        #region //AddSubSystem -- 子系統資料新增 -- Ben Ma 2022.11.23
        public string AddSubSystem(string SubSystemCode, string SubSystemName, string KeyText)
        {
            try
            {
                if (SubSystemCode.Length <= 0) throw new SystemException("【子系統代碼】不能為空!");
                if (SubSystemCode.Length > 50) throw new SystemException("【子系統代碼】長度錯誤!");
                if (SubSystemName.Length <= 0) throw new SystemException("【子系統名稱】不能為空!");
                if (SubSystemName.Length > 100) throw new SystemException("【子系統名稱】長度錯誤!");
                if (KeyText.Length <= 0) throw new SystemException("【金鑰】不能為空!");
                if (KeyText.Length > 64) throw new SystemException("【金鑰】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE SubSystemCode = @SubSystemCode";
                        dynamicParameters.Add("SubSystemCode", SubSystemCode);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【子系統代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷金鑰是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", KeyText);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【金鑰】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.SubSystem (SubSystemCode, SubSystemName, KeyText, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SubSystemId
                                VALUES (@SubSystemCode, @SubSystemName, @KeyText, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SubSystemCode,
                                SubSystemName,
                                KeyText,
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

        #region //AddSubSystemUser -- 子系統使用者資料新增 -- Ben Ma 2022.11.23
        public string AddSubSystemUser(int SubSystemId, string Users)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【子系統使用者】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統資料錯誤!");
                        #endregion

                        #region //判斷子系統使用者資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("子系統使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.SubSystemUser (SubSystemId, UserId, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.SubSystemUserId
                                    VALUES (@SubSystemId, @UserId, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SubSystemId,
                                    UserId = Convert.ToInt32(user),
                                    Status = "A",
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

        #region //AddPushSubscription -- 推播訂閱資料新增 -- Ben Ma 2022.12.23
        public string AddPushSubscription(string subscription)
        {
            try
            {
                if (!subscription.TryParseJson(out JObject tempJObject)) throw new SystemException("推播訂閱資料格式錯誤");

                JObject subscriptionJson = JObject.Parse(subscription);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        #region //判斷訂閱金鑰是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.PushSubscription
                                WHERE UserId = @UserId
                                AND ApiEndpoint = @ApiEndpoint";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("ApiEndpoint", subscriptionJson["endpoint"].ToString());

                        var resultPush = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 1;
                        if (resultPush.Count() <= 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.PushSubscription (UserId, ApiEndpoint, Keysp256dh, Keysauth
                                    , CreateDate, CreateBy)
                                    OUTPUT INSERTED.PushSubscriptionId
                                    VALUES (@UserId, @ApiEndpoint, @Keysp256dh, @Keysauth
                                    , @CreateDate, @CreateBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId = CurrentUser,
                                    ApiEndpoint = subscriptionJson["endpoint"].ToString(),
                                    Keysp256dh = subscriptionJson["keys"]["p256dh"].ToString(),
                                    Keysauth = subscriptionJson["keys"]["auth"].ToString(),
                                    CreateDate,
                                    CreateBy
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

        #region //AddNotificationLog -- 系統通知紀錄新增 -- Ben Ma 2023.08.09
        public string AddNotificationLog(Notification notification)
        {
            try
            {
                if (notification.UserNo.Length <= 0) throw new SystemException("【使用者】不能為空!");
                if (notification.TriggerFunction.Length <= 0) throw new SystemException("【觸發功能】不能為空!");
                if (notification.LogTitle.Length <= 0) throw new SystemException("【通知標題】不能為空!");
                if (notification.LogContent.Length <= 0) throw new SystemException("【通知內容】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", notification.UserNo);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //系統通知紀錄新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.NotificationLog (UserNo, TriggerFunction, LogTitle, LogContent
                                , ReadStatus, CreateDate)
                                OUTPUT INSERTED.LogId
                                VALUES (@UserNo, @TriggerFunction, @LogTitle, @LogContent
                                , @ReadStatus, @CreateDate)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                notification.UserNo,
                                notification.TriggerFunction,
                                notification.LogTitle,
                                notification.LogContent,
                                ReadStatus = "N",
                                CreateDate
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int LogId = -1;
                        foreach (var item in insertResult)
                        {
                            LogId = Convert.ToInt32(item.LogId);
                        }
                        #endregion

                        #region //系統通知詳細紀錄新增
                        foreach (var mode in notification.NotificationModes)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.NotificationLogDetail (LogId, LogType)
                                    VALUES (@LogId, @LogType)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LogId,
                                    LogType = mode.GetDescription()
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddDownloadLog -- 下載記錄新增 --Ding 2023.02.20
        public string AddDownloadLog(int FileId)
        {
            try
            {
                if (FileId <= 0) throw new SystemException("【FileId】不能為空!");
               
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string ClientComputer = "", ClientIP = "";
                        #region //判斷資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) {
                            throw new SystemException("【File】資料錯誤!");
                        }else {
                            ClientComputer = BaseHelper.ClientComputer();
                            ClientIP = BaseHelper.ClientIP();
                        };                        
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.DownloadLog (FileId,ClientComputer,ClientIP
                                , CreateDate, CreateBy)                                
                                VALUES (@FileId,@ClientComputer,@ClientIP
                                , @CreateDate, @CreateBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileId,
                                ClientComputer,                                
                                ClientIP,
                                CreateDate,
                                CreateBy
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

        #region //AddAuthorityUserCompany -- 使用者公司權限資料新增 -- Ben Ma 2023.05.17
        public string AddAuthorityUserCompany(string Users, string Companys)
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【使用者】不能為空!");
                if (Companys.Length <= 0) throw new SystemException("【公司別】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
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

                        #region //判斷公司資料是否正確
                        string[] companysList = Companys.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalCompanys
                                FROM BAS.Company
                                WHERE CompanyId IN @CompanyId";
                        dynamicParameters.Add("CompanyId", companysList);

                        int totalCompanys = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalCompanys;
                        if (totalCompanys != companysList.Length) throw new SystemException("使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var user in usersList)
                        {
                            foreach (var company in companysList)
                            {
                                #region //判斷該使用者公司權限是否已存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.AuthorityUserCompany a
                                        WHERE UserId = @UserId
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.Add("UserId", Convert.ToInt32(user));
                                dynamicParameters.Add("CompanyId", Convert.ToInt32(company));

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                if (result.Count() <= 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.AuthorityUserCompany (UserId, CompanyId
                                            , CreateDate, CreateBy)
                                            VALUES (@UserId, @CompanyId
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            UserId = Convert.ToInt32(user),
                                            CompanyId = Convert.ToInt32(company),
                                            CreateDate,
                                            CreateBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected = insertResult.Count();
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

        #region //AddUser -- 新版使用者資料新增 -- Yi 2023.09.01
        public string AddUser(string UserJson)
        {
            try
            {
                if (!UserJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                User users = JsonConvert.DeserializeObject<User>(UserJson);

                if (users.UserType.Length <= 0) throw new SystemException("【使用者類型】不能為空!");
                if (users.Account.Length <= 0) throw new SystemException("【帳號】不能為空!");
                if (users.Account.Length > 320) throw new SystemException("【帳號】長度錯誤!");
                if (users.InnerCode.Length <= 0) throw new SystemException("【內碼】不能為空!");
                if (users.InnerCode.Length > 20) throw new SystemException("【內碼】長度錯誤!");
                if (users.DisplayName.Length <= 0) throw new SystemException("【顯示名稱】不能為空!");
                if (users.DisplayName.Length > 50) throw new SystemException("【顯示名稱】長度錯誤!");
                if (users.UserType.Split(',').ToList().IndexOf("E") > -1)
                {
                    if (users.Employee.DepartmentId <= 0) throw new SystemException("【部門】不能為空!");
                    if (users.Employee.Gender.Length <= 0) throw new SystemException("【性別】不能為空!");
                    if (users.Employee.Email.Length > 100) throw new SystemException("【電子郵件】不能為空!");
                    if (users.Employee.Job.Length > 50) throw new SystemException("【職務】不能為空!");
                    if (users.Employee.JobType.Length > 50) throw new SystemException("【職務分類】不能為空!");
                    if (users.Employee.EmployeeStatus.Length <= 0) throw new SystemException("【內部員工狀態】不能為空!");
                }
                if (users.UserType.Split(',').ToList().IndexOf("C") > -1)
                {
                    if (users.Customer.Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                }
                if (users.UserType.Split(',').ToList().IndexOf("S") > -1)
                {
                    if (users.Supplier.Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者帳號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.UserInfo
                                WHERE Account = @Account";
                        dynamicParameters.Add("Account", users.Account);

                        var resultAccount = sqlConnection.Query(sql, dynamicParameters);
                        if (resultAccount.Count() > 0) throw new SystemException("【使用者帳號】重複，請重新輸入!");
                        #endregion

                        #region //判斷使用者類型資料是否正確
                        string[] userTypesList = users.UserType.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUserTypes
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo IN @TypeNo";
                        dynamicParameters.Add("TypeSchema", "UserInfo.UserType");
                        dynamicParameters.Add("TypeNo", userTypesList);

                        int totalUserTypes = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUserTypes;
                        if (totalUserTypes != userTypesList.Length) throw new SystemException("使用者類型資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //使用者資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.UserInfo (UserType, Account, Password, InnerCode, DisplayName
                                , PasswordStatus, PasswordMistake, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UserId
                                VALUES (@UserType, @Account, @Password, @InnerCode, @DisplayName
                                , @PasswordStatus, @PasswordMistake, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                users.UserType,
                                users.Account,
                                Password = BaseHelper.Sha256Encrypt(users.Account.ToLower()),
                                users.InnerCode,
                                users.DisplayName,
                                PasswordStatus = "Y",
                                PasswordMistake = 0,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int UserId = -1;
                        foreach (var item in insertResult)
                        {
                            UserId = Convert.ToInt32(item.UserId);
                        }
                        #endregion

                        #region //內部員工設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("E") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserEmployee (UserId, DepartmentId, Gender, Email
                                    , Job, JobType, EmployeeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @DepartmentId, @Gender, @Email
                                    , @Job, @JobType, @EmployeeStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Employee.DepartmentId,
                                    users.Employee.Gender,
                                    users.Employee.Email,
                                    users.Employee.Job,
                                    users.Employee.JobType,
                                    users.Employee.EmployeeStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //客戶設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("C") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserCustomer (UserId, Contact, ContactPhone, ContactAddress
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @Contact, @ContactPhone, @ContactAddress
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Customer.Contact,
                                    users.Customer.ContactPhone,
                                    users.Customer.ContactAddress,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //供應商設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("S") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserSupplier (UserId, Contact, ContactPhone, ContactAddress
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @Contact, @ContactPhone, @ContactAddress
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Supplier.Contact,
                                    users.Supplier.ContactPhone,
                                    users.Supplier.ContactAddress,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddTeams -- 新增MAMO團隊 -- Ann 2024-01-16
        public string AddTeams(string MamoTeamId, string TeamName, string TeamNo, string Remark, string Company, int UserId)
        {
            try
            {
                if (MamoTeamId.Length <= 0) throw new SystemException("【MAMO團隊ID】不能為空!");
                if (TeamName.Length <= 0) throw new SystemException("【團隊名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認公司別資料是否有誤
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a 
                                WHERE a.CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", Company);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        int CompanyId = -1;
                        foreach (var item in CompanyResult)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //確認此MAMO ID尚未新增在BM系統
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.Teams a 
                                WHERE a.MamoTeamId = @MamoTeamId";
                        dynamicParameters.Add("MamoTeamId", MamoTeamId);

                        var TeamsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TeamsResult.Count() > 0) throw new SystemException("此MAMO ID已存在，不可重複新增!!");
                        #endregion

                        #region //新增團隊
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MAMO.Teams (CompanyId, MamoTeamId, TeamName, TeamNo, [Status], Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TeamId
                                VALUES (@CompanyId, @MamoTeamId, @TeamName, @TeamNo, @Status, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                MamoTeamId,
                                TeamName,
                                TeamNo, 
                                Status = "A", 
                                Remark,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
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

        #region //AddTeamMembers -- 新增MAMO團隊成員 -- Ann 2024-01-16
        public string AddTeamMembers(int TeamId, List<string> UserNoList, int CreateUserId)
        {
            try
            {
                if (UserNoList.Count <= 0) throw new SystemException("至少需指定一位團隊成員!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateUserId);

                        var UserResult2 = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult2.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //確認團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                FROM MAMO.Teams a 
                                WHERE a.TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var TeamsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TeamsResult.Count() <= 0) throw new SystemException("團隊資料錯誤!!");

                        foreach (var item in TeamsResult)
                        {
                            if (item.Status != "A") throw new SystemException("團隊狀態非啟用中，無法更新!!");
                        }
                        #endregion

                        foreach (var userNo in UserNoList)
                        {
                            #region //確認USER相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                            dynamicParameters.Add("UserNo", userNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                            int UserId = -1;
                            foreach (var item in UserResult)
                            {
                                UserId = item.UserId;
                            }
                            #endregion

                            #region //確認此成員尚未存在於此團體
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MAMO.TeamMembers a 
                                    WHERE a.TeamId = @TeamId
                                    AND a.UserId = @UserId";
                            dynamicParameters.Add("TeamId", TeamId);
                            dynamicParameters.Add("UserId", UserId);

                            var TeamMembersResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TeamMembersResult.Count() > 0) throw new SystemException("此成員已在此團隊中，不可重複新增!!");
                            #endregion

                            #region //新增團隊成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MAMO.TeamMembers (TeamId, UserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MemberId
                                    VALUES (@TeamId, @UserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TeamId,
                                    UserId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = CreateUserId,
                                    LastModifiedBy = CreateUserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddMamoLog -- 新增MAMO LOG紀錄 -- Ann 2024-01-16
        public string AddMamoLog(string MamoId, string MamoTeamId, string ControlType, string UserList, string Remark, int UserId)
        {
            try
            {
                if (ControlType.Length <= 0) throw new SystemException("異動類型不能為空!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //新增LOG
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MAMO.MamoLog (MamoId, ControlType, UserList, Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LogId
                                VALUES (@MamoId, @ControlType, @UserList, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MamoId,
                                ControlType,
                                UserList,
                                Remark,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
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

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddChannels -- 新增MAMO頻道 -- Ann 2024-01-16
        public string AddChannels(int TeamId, string MamoChannelId, string ChannelName, string ChannelNo, int UserId)
        {
            try
            {
                if (MamoChannelId.Length <= 0) throw new SystemException("【MAMO團隊ID】不能為空!");
                if (ChannelName.Length <= 0) throw new SystemException("【團隊名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //確認團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                FROM MAMO.Teams a 
                                WHERE a.TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var TeamsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TeamsResult.Count() <= 0) throw new SystemException("團隊資料錯誤!!");

                        foreach (var item in TeamsResult)
                        {
                            if (item.Status != "A") throw new SystemException("團隊狀態非啟用中!!");
                        }
                        #endregion

                        #region //確認此MAMO ID尚未新增在BM系統
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.Channels a 
                                WHERE a.MamoChannelId = @MamoChannelId";
                        dynamicParameters.Add("MamoChannelId", MamoChannelId);

                        var ChannelsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ChannelsResult.Count() > 0) throw new SystemException("此MAMO ID已存在，不可重複新增!!");
                        #endregion

                        #region //新增頻道
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MAMO.Channels (TeamId, MamoChannelId, ChannelName, ChannelNo
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ChannelId, INSERTED.TeamId
                                VALUES (@TeamId, @MamoChannelId, @ChannelName, @ChannelNo
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TeamId,
                                MamoChannelId,
                                ChannelName, 
                                ChannelNo,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
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

        #region //AddChannelMembers -- 新增MAMO頻道成員 -- Ann 2024-01-16
        public string AddChannelMembers(int ChannelId, List<string> UserNoList, int CreateUserId)
        {
            try
            {
                if (UserNoList.Count <= 0) throw new SystemException("至少需指定一位團隊成員!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateUserId);

                        var UserResult2 = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult2.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //確認頻道資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                FROM MAMO.Channels a 
                                WHERE a.ChannelId = @ChannelId";
                        dynamicParameters.Add("ChannelId", ChannelId);

                        var ChannelsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ChannelsResult.Count() <= 0) throw new SystemException("頻道資料錯誤!!");

                        foreach (var item in ChannelsResult)
                        {
                            if (item.Status != "A") throw new SystemException("頻道狀態非啟用中，無法更新!!");
                        }
                        #endregion

                        foreach (var userNo in UserNoList)
                        {
                            #region //確認USER相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                            dynamicParameters.Add("UserNo", userNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                            int UserId = -1;
                            foreach (var item in UserResult)
                            {
                                UserId = item.UserId;
                            }
                            #endregion

                            #region //新增團隊成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MAMO.ChannelMembers (ChannelId, UserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MemberId
                                    VALUES (@ChannelId, @UserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ChannelId,
                                    UserId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = CreateUserId,
                                    LastModifiedBy = CreateUserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddMamoPushNotification -- 新增MAMO推播訊息 -- Ann 2024-01-16
        public string AddMamoPushNotification(string PushType, string SendId, string Content, List<string> Tags, List<int> Files, int UserId)
        {
            try
            {
                if (PushType.Length <= 0) throw new SystemException("推播類型不能為空!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserResult2 = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult2.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        if (PushType == "Personal")
                        {
                            #region //確認USER資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a 
                                    WHERE a.UserNo = @SendId";
                            dynamicParameters.Add("SendId", SendId);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("USER資料錯誤!!");
                            #endregion
                        }
                        else if (PushType == "Channel")
                        {
                            #region //確認頻道資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.[Status]
                                    FROM MAMO.Channels a 
                                    WHERE a.MamoChannelId = @SendId";
                            dynamicParameters.Add("SendId", SendId);

                            var ChannelsResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ChannelsResult.Count() <= 0) throw new SystemException("頻道資料錯誤!!");

                            foreach (var item in ChannelsResult)
                            {
                                if (item.Status != "A") throw new SystemException("頻道狀態非啟用中，無法推播!!");
                            }
                            #endregion
                        }
                        else
                        {
                            throw new SystemException("推播類型錯誤!!");
                        }

                        string TagString = "";
                        if (Tags != null)
                        {
                            TagString = string.Join(", ", Tags);
                        }

                        #region //新增推播紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MAMO.PushNotification (PushType, SendId, Content, Tags
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PnId
                                VALUES (@PushType, @SendId, @Content, @Tags
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PushType,
                                SendId,
                                Content,
                                Tags = TagString,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int PnId = -1;
                        foreach (var item in insertResult)
                        {
                            PnId = item.PnId;
                        }
                        #endregion

                        #region //新增推播檔案
                        if (Files != null)
                        {
                            foreach (var fileId in Files)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MAMO.PnFile (PnId, FileId
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PnFileId
                                        VALUES (@PnId, @FileId
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PnId,
                                        FileId = fileId,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
                                    });

                                var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult2.Count();
                            }
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

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Update
        #region //UpdateApiKey -- Api金鑰更新 -- Ben Ma 2022.04.27
        public string UpdateApiKey(int KeyId, string KeyText, string Purpose)
        {
            try
            {
                if (KeyText.Length <= 0) throw new SystemException("【使用者名稱】不能為空!");
                if (KeyText.Length > 64) throw new SystemException("【使用者名稱】長度錯誤!");
                if (Purpose.Length <= 0) throw new SystemException("【用途】不能為空!");
                if (Purpose.Length > 50) throw new SystemException("【用途】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷金鑰資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("金鑰資料錯誤!");
                        #endregion

                        #region //判斷金鑰是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyText = @KeyText
                                AND KeyId != @KeyId";
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("KeyId", KeyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【金鑰】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.ApiKey SET
                                KeyText = @KeyText,
                                Purpose = @Purpose,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE KeyId = @KeyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                KeyText,
                                Purpose,
                                LastModifiedDate,
                                LastModifiedBy,
                                KeyId
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

        #region //UpdateApiKeyStatus -- Api金鑰狀態更新 -- Ben Ma 2022.04.27
        public string UpdateApiKeyStatus(int KeyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷金鑰資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.ApiKey
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("金鑰資料錯誤!");

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
                        sql = @"UPDATE BAS.ApiKey SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE KeyId = @KeyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                KeyId
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

        #region //UpdateApiKeyFunction -- Api適用功能更新 -- Ben Ma 2022.04.29
        public string UpdateApiKeyFunction(int FunctionId, string FunctionCode, string Purpose)
        {
            try
            {
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 100) throw new SystemException("【功能代碼】長度錯誤!");
                if (Purpose.Length <= 0) throw new SystemException("【用途】不能為空!");
                if (Purpose.Length > 50) throw new SystemException("【用途】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷適用功能資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 KeyId
                                FROM BAS.ApiKeyFunction
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("適用功能資料錯誤!");

                        int keyId = 0;
                        foreach (var item in result)
                        {
                            keyId = Convert.ToInt32(item.KeyId);
                        }
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKeyFunction
                                WHERE KeyId = @KeyId
                                AND FunctionCode = @FunctionCode
                                AND FunctionId != @FunctionId";
                        dynamicParameters.Add("KeyId", keyId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.ApiKeyFunction SET
                                FunctionCode = @FunctionCode,
                                Purpose = @Purpose,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FunctionCode,
                                Purpose,
                                LastModifiedDate,
                                LastModifiedBy,
                                FunctionId
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

        #region //UpdateRole -- 角色資料更新 -- Ben Ma 2022.05.06
        public string UpdateRole(int RoleId, string RoleName, string AdminStatus)
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
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        #region //判斷角色名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Role
                                WHERE CompanyId = @CompanyId
                                AND RoleName = @RoleName
                                AND RoleId != @RoleId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoleName", RoleName);
                        dynamicParameters.Add("RoleId", RoleId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Role SET
                                RoleName = @RoleName,
                                AdminStatus = @AdminStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoleId = @RoleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleName,
                                AdminStatus,
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

        #region //UpdateRoleStatus -- 角色狀態更新 -- Ben Ma 2022.05.06
        public string UpdateRoleStatus(int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.Role
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
                        sql = @"UPDATE BAS.Role SET
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

        #region //UpdateRoleFunctionDetail -- 角色權限更新 -- Ben Ma 2022.05.12
        public string UpdateRoleFunctionDetail(int RoleId, int DetailId, bool Checked, string SearchKey)
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
                                FROM BAS.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        if (Checked)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.RoleFunctionDetail (RoleId, DetailId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@RoleId, @DetailId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoleId,
                                    DetailId,
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
                            sql = @"DELETE BAS.RoleFunctionDetail
                                    WHERE RoleId = @RoleId
                                    AND DetailId = @DetailId";
                            dynamicParameters.Add("RoleId", RoleId);
                            dynamicParameters.Add("DetailId", DetailId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //目前權限數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModuleId, a.SystemId, b.TotalModuleFunction, c.RoleTotalModuleFunction, d.TotalFunction, e.RoleTotalFunction
                                FROM BAS.[Module] a
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM BAS.FunctionDetail ba
                                    INNER JOIN BAS.[Function] bb ON ba.FunctionId = bb.FunctionId
                                    WHERE bb.ModuleId = a.ModuleId
                                    AND bb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) b
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalModuleFunction
                                    FROM BAS.RoleFunctionDetail ca
                                    INNER JOIN BAS.FunctionDetail cb ON ca.DetailId = cb.DetailId
                                    INNER JOIN BAS.[Function] cc ON cb.FunctionId = cc.FunctionId
                                    WHERE cc.ModuleId = a.ModuleId
                                    AND cc.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ca.RoleId = @RoleId
                                ) c
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalFunction
                                    FROM BAS.FunctionDetail da
                                    INNER JOIN BAS.[Function] db ON da.FunctionId = db.FunctionId
                                    INNER JOIN BAS.[Module] dc ON db.ModuleId = dc.ModuleId
                                    WHERE dc.SystemId = a.SystemId
                                    AND db.FunctionName LIKE '%' + @SearchKey + '%'
                                ) d
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalFunction
                                    FROM BAS.RoleFunctionDetail ea
                                    INNER JOIN BAS.FunctionDetail eb ON ea.DetailId = eb.DetailId
                                    INNER JOIN BAS.[Function] ec ON eb.FunctionId = ec.FunctionId
                                    INNER JOIN BAS.[Module] ed ON ec.ModuleId = ed.ModuleId
                                    WHERE ed.SystemId = a.SystemId
                                    AND ec.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ea.RoleId = @RoleId
                                ) e
                                WHERE 1=1
                                AND a.ModuleId IN (
                                    SELECT TOP 1 xb.ModuleId
                                    FROM BAS.FunctionDetail xa
                                    INNER JOIN BAS.[Function] xb ON xa.FunctionId = xb.FunctionId 
                                    WHERE xa.DetailId = @DetailId
                                )";
                        dynamicParameters.Add("RoleId", RoleId);
                        dynamicParameters.Add("DetailId", DetailId);
                        dynamicParameters.Add("SearchKey", SearchKey);

                        var returnData = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = returnData
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

        #region //UpdateRoleUser -- 角色人員更新 -- Ben Ma 2022.05.16
        public string UpdateRoleUser(int RoleId, string UserList)
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
                                FROM BAS.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        string[] userArray = UserList.Split(',');

                        int rowsAffected = 0;
                        for (int i = 0; i < userArray.Length; i++)
                        {
                            int userId = Convert.ToInt32(userArray[i]);

                            #region //判斷使用者資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", userId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                            #endregion

                            #region //判斷角色人員是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.UserRole
                                    WHERE UserId = @UserId
                                    AND RoleId = @RoleId";
                            dynamicParameters.Add("UserId", userId);
                            dynamicParameters.Add("RoleId", RoleId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) throw new SystemException("使用者已有該角色!");
                            #endregion

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserRole (UserId, RoleId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @RoleId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId = userId,
                                    RoleId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateUserFunctionDetail -- 使用者權限更新 -- Ben Ma 2022.05.17
        public string UpdateUserFunctionDetail(int UserId, int DetailId, bool Checked, string SearchKey)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        if (Checked)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserFunctionDetail (UserId, DetailId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @DetailId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    DetailId,
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
                            #region //判斷是否為該角色基本權限
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.UserRole a
                                    INNER JOIN BAS.RoleFunctionDetail b ON a.RoleId = b.RoleId
                                    WHERE a.UserId = @UserId
                                    AND b.DetailId = @DetailId";
                            dynamicParameters.Add("UserId", UserId);
                            dynamicParameters.Add("DetailId", DetailId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0) throw new SystemException("此功能為該角色基本權限，無法刪除!");
                            #endregion

                            #region //刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE BAS.UserFunctionDetail
                                    WHERE UserId = @UserId
                                    AND DetailId = @DetailId";
                            dynamicParameters.Add("UserId", UserId);
                            dynamicParameters.Add("DetailId", DetailId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion                            
                        }

                        #region //目前權限數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModuleId, a.SystemId, b.TotalModuleFunction, c.UserTotalModuleFunction, d.TotalFunction, e.UserTotalFunction
                                FROM BAS.[Module] a
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM BAS.FunctionDetail ba
                                    INNER JOIN BAS.[Function] bb ON ba.FunctionId = bb.FunctionId
                                    WHERE bb.ModuleId = a.ModuleId
                                    AND bb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) b
                                OUTER APPLY (
                                    SELECT COUNT(1) UserTotalModuleFunction
                                    FROM BAS.UserFunctionDetail ca
                                    INNER JOIN BAS.FunctionDetail cb ON ca.DetailId = cb.DetailId
                                    INNER JOIN BAS.[Function] cc ON cb.FunctionId = cc.FunctionId
                                    WHERE cc.ModuleId = a.ModuleId
                                    AND cc.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ca.UserId = @UserId
                                ) c
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalFunction
                                    FROM BAS.FunctionDetail da
                                    INNER JOIN BAS.[Function] db ON da.FunctionId = db.FunctionId
                                    INNER JOIN BAS.[Module] dc ON db.ModuleId = dc.ModuleId
                                    WHERE dc.SystemId = a.SystemId
                                    AND db.FunctionName LIKE '%' + @SearchKey + '%'
                                ) d
                                OUTER APPLY (
                                    SELECT COUNT(1) UserTotalFunction
                                    FROM BAS.UserFunctionDetail ea
                                    INNER JOIN BAS.FunctionDetail eb ON ea.DetailId = eb.DetailId
                                    INNER JOIN BAS.[Function] ec ON eb.FunctionId = ec.FunctionId
                                    INNER JOIN BAS.[Module] ed ON ec.ModuleId = ed.ModuleId
                                    WHERE ed.SystemId = a.SystemId
                                    AND ec.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ea.UserId = @UserId
                                ) e
                                WHERE 1=1
                                AND a.ModuleId IN (
                                    SELECT TOP 1 xb.ModuleId
                                    FROM BAS.FunctionDetail xa
                                    INNER JOIN BAS.[Function] xb ON xa.FunctionId = xb.FunctionId 
                                    WHERE xa.DetailId = @DetailId
                                )";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("DetailId", DetailId);
                        dynamicParameters.Add("SearchKey", SearchKey);

                        var returnData = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = returnData
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

        #region //UpdateMailServer -- 郵件伺服器更新 -- Zoey 2022.05.25
        public string UpdateMailServer(int ServerId, string Host, int Port, string SendMode, string Account, string Password)
        {
            try
            {
                if (Host.Length <= 0) throw new SystemException("【主機】不能為空!");
                if (Host.Length > 50) throw new SystemException("【主機】長度錯誤!");
                if (Port < 0) throw new SystemException("【埠號】不可小於0!");
                if (Port > 65535) throw new SystemException("【埠號】不可大於65535!");
                if (SendMode.Length < 0) throw new SystemException("【寄送模式】不能為空!");
                if (Account.Length <= 0) throw new SystemException("【帳號】不能為空!");
                if (Account.Length > 100) throw new SystemException("【帳號】長度錯誤!");
                if (Password.Length <= 0) throw new SystemException("【密碼】不能為空!");
                if (Password.Length > 100) throw new SystemException("【密碼】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷伺服器資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.MailServer
                                WHERE ServerId = @ServerId";
                        dynamicParameters.Add("ServerId", ServerId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("伺服器資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.MailServer SET
                                Port = @Port,
                                SendMode = @SendMode,
                                Account = @Account,
                                Password = @Password,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ServerId = @ServerId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Port,
                                SendMode,
                                Account,
                                Password,
                                LastModifiedDate,
                                LastModifiedBy,
                                ServerId
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

        #region //UpdateMailContact -- 郵件聯絡人資料更新 -- Zoey 2022.05.26
        public string UpdateMailContact(int ContactId, string ContactName, string Email)
        {
            try
            {
                if (ContactName.Length <= 0) throw new SystemException("【聯絡人姓名】不能為空!");
                if (ContactName.Length > 50) throw new SystemException("【聯絡人姓名】長度錯誤!");
                if (Email.Length <= 0) throw new SystemException("【聯絡人信箱】不能為空!");
                if (Email.Length > 100) throw new SystemException("【聯絡人信箱】長度錯誤!");
                if (!Regex.IsMatch(Email, RegexHelper.Email, RegexOptions.IgnoreCase)) throw new SystemException("【聯絡人信箱】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聯絡人資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.MailContact
                                WHERE ContactId = @ContactId";
                        dynamicParameters.Add("ContactId", ContactId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.MailContact SET
                                ContactName = @ContactName,
                                Email = @Email,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ContactId = @ContactId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ContactName,
                                Email,
                                LastModifiedDate,
                                LastModifiedBy,
                                ContactId
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

        #region //UpdateMailTemplate -- 郵件樣板更新 -- Zoey 2022.05.27
        public string UpdateMailTemplate(int MailId, int ServerId, string MailName, string MailFrom, string MailTo
            , string MailCc, string MailBcc, string MailSubject, string MailContent)
        {
            try
            {
                if (ServerId < 0) throw new SystemException("【主機】不能為空!");
                if (MailName.Length <= 0) throw new SystemException("【郵件名稱】不能為空!");
                if (MailFrom.Length <= 0) throw new SystemException("【寄件人】不能為空!");
                if (MailTo.Length <= 0) throw new SystemException("【收件人】不能為空!");
                if (MailSubject.Length <= 0) throw new SystemException("【郵件主旨】不能為空!");
                if (MailContent.Length <= 0) throw new SystemException("【郵件內容】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷郵件樣板是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Mail
                                WHERE MailId = @MailId";
                        dynamicParameters.Add("MailId", MailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("郵件樣板錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Mail SET
                                ServerId = @ServerId,
                                MailName = @MailName,
                                MailSubject = @MailSubject,
                                MailContent = @MailContent,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MailId = @MailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ServerId,
                                MailName,
                                MailSubject,
                                MailContent,
                                LastModifiedDate,
                                LastModifiedBy,
                                MailId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        #region //刪除郵件人員
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.MailUser
                                WHERE MailId = @MailId";
                        dynamicParameters.Add("MailId", MailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //寄件人
                        string[] mailFrom = MailFrom.Split(',');
                        for (int i = 0; i < mailFrom.Length; i++)
                        {
                            string[] mailUser = mailFrom[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "F",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "F",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //收件人
                        string[] mailTo = MailTo.Split(',');
                        for (int i = 0; i < mailTo.Length; i++)
                        {
                            string[] mailUser = mailTo[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "T",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "T",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //副本
                        string[] mailCc = MailCc.Split(',');
                        for (int i = 0; i < mailCc.Length; i++)
                        {
                            string[] mailUser = mailCc[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "C",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "C",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
                        }
                        #endregion

                        #region //秘密副本
                        string[] mailBcc = MailBcc.Split(',');
                        for (int i = 0; i < mailBcc.Length; i++)
                        {
                            string[] mailUser = mailBcc[i].Split(':');

                            switch (mailUser[0])
                            {
                                case "user":
                                    #region //判斷使用者資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", Convert.ToInt32(mailUser[1]));

                                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, UserId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @UserId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            UserId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "B",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                case "contact":
                                    #region //判斷聯絡人資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.MailContact
                                            WHERE ContactId = @ContactId";
                                    dynamicParameters.Add("ContactId", Convert.ToInt32(mailUser[1]));

                                    var resultContact = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultContact.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.MailUser (MailId, ContactId, MailUserType
                                            , CreateDate, CreateBy)
                                            VALUES (@MailId, @ContactId, @MailUserType
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MailId,
                                            ContactId = Convert.ToInt32(mailUser[1]),
                                            MailUserType = "B",
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                            }
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

        #region //UpdateMailSendSetting -- 郵件寄送設定更新 -- Ben Ma 2023.04.18
        public string UpdateMailSendSetting(int SettingId, string SettingSchema, string SettingNo)
        {
            try
            {
                if (SettingSchema.Length <= 0) throw new SystemException("【寄送綱要】不能為空!");
                if (SettingSchema.Length > 100) throw new SystemException("【寄送綱要】長度錯誤!");
                if (SettingNo.Length <= 0) throw new SystemException("【寄送綱要】不能為空!");
                if (SettingNo.Length > 10) throw new SystemException("【寄送綱要】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷郵件寄送設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.MailSendSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("郵件寄送設定資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.MailSendSetting SET
                                SettingSchema = @SettingSchema,
                                SettingNo = @SettingNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SettingId = @SettingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SettingSchema,
                                SettingNo,
                                LastModifiedDate,
                                LastModifiedBy,
                                SettingId
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

        #region //UpdateCalendar -- 行事曆資料更新 -- Zoey 2022.06.06
        public string UpdateCalendar(string CalendarDate, string CalendarDesc, string DateType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷行事曆資料是否存在
                        bool exist = false;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Calendar
                                WHERE CalendarDate = @CalendarDate
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CalendarDate", CalendarDate);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        exist = result.Count() > 0;
                        #endregion

                        int rowsAffected = 0;
                        if (!exist)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.Calendar (CompanyId, CalendarDate, CalendarDesc, DateType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CalendarId
                                VALUES (@CompanyId, @CalendarDate, @CalendarDesc, @DateType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    CalendarDate,
                                    CalendarDesc,
                                    DateType,
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
                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.Calendar SET
                                    CalendarDesc = @CalendarDesc,
                                    DateType = @DateType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE CalendarDate = @CalendarDate
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CalendarDesc,
                                    DateType,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CalendarDate,
                                    CompanyId = CurrentCompany
                                });

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

        #region //UpdateSubSystem -- 子系統資料更新 -- Ben Ma 2022.11.23
        public string UpdateSubSystem(int SubSystemId, string SubSystemCode, string SubSystemName, string KeyText)
        {
            try
            {
                if (SubSystemCode.Length <= 0) throw new SystemException("【子系統代碼】不能為空!");
                if (SubSystemCode.Length > 50) throw new SystemException("【子系統代碼】長度錯誤!");
                if (SubSystemName.Length <= 0) throw new SystemException("【子系統名稱】不能為空!");
                if (SubSystemName.Length > 100) throw new SystemException("【子系統名稱】長度錯誤!");
                if (KeyText.Length <= 0) throw new SystemException("【金鑰】不能為空!");
                if (KeyText.Length > 64) throw new SystemException("【金鑰】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統資料錯誤!");
                        #endregion

                        #region //判斷子系統代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE SubSystemCode = @SubSystemCode
                                AND SubSystemId != @SubSystemId";
                        dynamicParameters.Add("SubSystemCode", SubSystemCode);
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【子系統代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷金鑰是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE KeyText = @KeyText
                                AND SubSystemId != @SubSystemId";
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【金鑰】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.SubSystem SET
                                SubSystemCode = @SubSystemCode,
                                SubSystemName = @SubSystemName,
                                KeyText = @KeyText,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SubSystemCode,
                                SubSystemName,
                                KeyText,
                                LastModifiedDate,
                                LastModifiedBy,
                                SubSystemId
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

        #region //UpdateSubSystemStatus -- 子系統狀態更新 -- Ben Ma 2022.11.23
        public string UpdateSubSystemStatus(int SubSystemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.SubSystem
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統資料錯誤!");

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
                        sql = @"UPDATE BAS.SubSystem SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SubSystemId
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

        #region //UpdateSubSystemUserStatus -- 子系統使用者狀態更新 -- Ben Ma 2022.11.23
        public string UpdateSubSystemUserStatus(int SubSystemUserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.SubSystemUser
                                WHERE SubSystemUserId = @SubSystemUserId";
                        dynamicParameters.Add("SubSystemUserId", SubSystemUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統使用者資料錯誤!");

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
                        sql = @"UPDATE BAS.SubSystemUser SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SubSystemUserId = @SubSystemUserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SubSystemUserId
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

        #region //UpdatePasswordSetting -- 密碼參數設定資料更新 -- Ben Ma 2022.12.26
        public string UpdatePasswordSetting(int PasswordExpiration, string PasswordFormat, int PasswordWrongCount)
        {
            try
            {
                if (PasswordExpiration <= 0) throw new SystemException("【密碼過期時間】不能為空!");
                if (PasswordFormat.Length <= 0) throw new SystemException("【密碼格式】不能為空!");
                if (PasswordWrongCount <= 0) throw new SystemException("【密碼允許錯誤次數】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.PasswordSetting SET
                                PasswordExpiration = @PasswordExpiration,
                                PasswordFormat = @PasswordFormat,
                                PasswordWrongCount = @PasswordWrongCount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PasswordExpiration,
                                PasswordFormat,
                                PasswordWrongCount,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateNotificationSetting -- 系統通知個人設定更新 -- Ben Ma 2023.08.10
        public string UpdateNotificationSetting(string MailAdviceStatus, string PushAdviceStatus, string WorkWeixinStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");
                        #endregion

                        #region //設定資料更新
                        #region //判斷設定資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.NotificationSetting
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var resultSetting = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        if (resultSetting.Count() <= 0)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.NotificationSetting (UserId
                                    , MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId
                                    , @MailAdviceStatus, @PushAdviceStatus, @WorkWeixinStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserId = CurrentUser,
                                MailAdviceStatus,
                                PushAdviceStatus,
                                WorkWeixinStatus,
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
                            #region //更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.NotificationSetting SET
                                    MailAdviceStatus = @MailAdviceStatus,
                                    PushAdviceStatus = @PushAdviceStatus,
                                    WorkWeixinStatus = @WorkWeixinStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UserId = @UserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MailAdviceStatus,
                                    PushAdviceStatus,
                                    WorkWeixinStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    UserId = CurrentUser
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateNotificationLogReadStatus -- 系統通知讀取狀態更新 -- Ben Ma 2023.08.10
        public string UpdateNotificationLogReadStatus(int LogId, string TriggerFunction)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統通知紀錄是否正確
                        bool checkUpdate = true;

                        if (LogId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.NotificationLog a
                                    LEFT JOIN BAS.[User] b ON a.UserNo = b.UserNo
                                    WHERE a.LogId = @LogId
                                    AND b.UserId = @UserId
                                    AND a.ReadStatus = @ReadStatus";
                            dynamicParameters.Add("LogId", LogId);
                            dynamicParameters.Add("UserId", CurrentUser);
                            dynamicParameters.Add("ReadStatus", "N");

                            var resultLog = sqlConnection.Query(sql, dynamicParameters);
                            checkUpdate = resultLog.Count() > 0;
                        }
                        #endregion

                        int rowsAffected = 0;
                        if (checkUpdate)
                        {
                            if (LogId > 0)
                            {
                                #region //單一已讀
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE BAS.NotificationLog SET
                                        ReadStatus = @ReadStatus,
                                        ReadDate = @ReadDate,
                                        ReadIP = @ReadIP
                                        WHERE LogId = @LogId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ReadStatus = "Y",
                                        ReadDate = CreateDate,
                                        ReadIP = BaseHelper.ClientIP(),
                                        LogId
                                    });
                                #endregion
                            }
                            else
                            {
                                #region //全部已讀
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE a SET
                                        a.ReadStatus = @ReadStatus,
                                        a.ReadDate = @ReadDate,
                                        a.ReadIP = @ReadIP
                                        FROM BAS.NotificationLog a
                                        INNER JOIN BAS.[User] b ON a.UserNo = b.UserNo
                                        WHERE a.ReadStatus = @CurrentStatus
                                        AND b.UserId = @UserId
                                        AND a.TriggerFunction = @TriggerFunction";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ReadStatus = "Y",
                                        ReadDate = CreateDate,
                                        ReadIP = BaseHelper.ClientIP(),
                                        CurrentStatus = "N",
                                        UserId = CurrentUser,
                                        TriggerFunction
                                    });
                                #endregion
                            }

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateUser -- 新版使用者資料更新 -- Yi 2023.09.01
        public string UpdateUser(int UserId, string UserJson)
        {
            try
            {
                if (!UserJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                User users = JsonConvert.DeserializeObject<User>(UserJson);

                if (users.UserType.Length <= 0) throw new SystemException("【使用者類型】不能為空!");
                if (users.InnerCode.Length <= 0) throw new SystemException("【內碼】不能為空!");
                if (users.InnerCode.Length > 20) throw new SystemException("【內碼】長度錯誤!");
                if (users.DisplayName.Length <= 0) throw new SystemException("【顯示名稱】不能為空!");
                if (users.DisplayName.Length > 50) throw new SystemException("【顯示名稱】長度錯誤!");
                if (users.UserType.Split(',').ToList().IndexOf("E") > -1)
                {
                    if (users.Employee.DepartmentId <= 0) throw new SystemException("【部門】不能為空!");
                    if (users.Employee.Gender.Length <= 0) throw new SystemException("【性別】不能為空!");
                    if (users.Employee.Email.Length > 100) throw new SystemException("【電子郵件】不能為空!");
                    if (users.Employee.Job.Length > 50) throw new SystemException("【職務】不能為空!");
                    if (users.Employee.JobType.Length > 50) throw new SystemException("【職務分類】不能為空!");
                    if (users.Employee.EmployeeStatus.Length <= 0) throw new SystemException("【內部員工狀態】不能為空!");
                }
                if (users.UserType.Split(',').ToList().IndexOf("C") > -1)
                {
                    if (users.Customer.Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                }
                if (users.UserType.Split(',').ToList().IndexOf("S") > -1)
                {
                    if (users.Supplier.Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.UserInfo
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //使用者資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.UserInfo SET
                                UserType = @UserType,
                                InnerCode = @InnerCode,
                                DisplayName = @DisplayName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                users.UserType,
                                users.InnerCode,
                                users.DisplayName,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //設定檔資料刪除
                        #region //內部員工設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserEmployee
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //客戶設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserCustomer
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //供應商設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserSupplier
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //內部員工設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("E") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserEmployee (UserId, DepartmentId, Gender, Email
                                    , Job, JobType, EmployeeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @DepartmentId, @Gender, @Email
                                    , @Job, @JobType, @EmployeeStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Employee.DepartmentId,
                                    users.Employee.Gender,
                                    users.Employee.Email,
                                    users.Employee.Job,
                                    users.Employee.JobType,
                                    users.Employee.EmployeeStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //客戶設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("C") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserCustomer (UserId, Contact, ContactPhone, ContactAddress
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @Contact, @ContactPhone, @ContactAddress
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Customer.Contact,
                                    users.Customer.ContactPhone,
                                    users.Customer.ContactAddress,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //供應商設定檔新增
                        if (users.UserType.Split(',').ToList().IndexOf("S") > -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserSupplier (UserId, Contact, ContactPhone, ContactAddress
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @Contact, @ContactPhone, @ContactAddress
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    users.Supplier.Contact,
                                    users.Supplier.ContactPhone,
                                    users.Supplier.ContactAddress,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
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

        #region //UpdateUserStatus -- 新版使用者狀態更新 -- Yi 2023.09.01
        public string UpdateUserStatus(int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.UserInfo
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

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
                        sql = @"UPDATE BAS.UserInfo SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            UserId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdateUserPasswordReset -- 新版使用者密碼重置 -- Ben Ma 2023.09.18
        public string UpdateUserPasswordReset(int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Account
                                FROM BAS.UserInfo
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

                        string account = "";
                        foreach (var item in resultUserInfo)
                        {
                            account = item.Account;
                        }
                        #endregion

                        bool adVerify = false, classicVerify = false;
                        #region //判斷是否要用AD登入驗證
                        string companyNo = "";

                        switch (true)
                        {
                            case bool _ when Regex.IsMatch(account, @"^(DC|DD|ZY)[0-9]{5}$"): //中揚
                                adVerify = true;
                                companyNo = "JMO";
                                break;
                            case bool _ when Regex.IsMatch(account, @"^E[0-9]{4}$"): //紘立
                                adVerify = true;
                                companyNo = "Eterge";
                                break;
                            default:
                                classicVerify = true;
                                break;
                        }
                        #endregion

                        if (adVerify)
                        {
                            #region //AD主機資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT c.ConnectionServer, c.ConnectionAccount, c.ConnectionPassword
                                    FROM BAS.Interfacing a
                                    INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                                    OUTER APPLY (
                                        SELECT TOP 1 ca.ConnectionServer, ca.ConnectionAccount, ca.ConnectionPassword
                                        FROM BAS.InterfacingConnection ca
                                        WHERE ca.InterfacingId = a.InterfacingId
                                        AND ca.[Status] = @Status
                                    ) c
                                    WHERE a.InterfacingType = @InterfacingType
                                    AND b.CompanyNo = @CompanyNo";
                            dynamicParameters.Add("Status", "A");
                            dynamicParameters.Add("InterfacingType", "AD_Main");
                            dynamicParameters.Add("CompanyNo", companyNo);

                            var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInterfacing.Count() <= 0) throw new SystemException("AD主機資料錯誤!");

                            string connectionServer = "", connectionAccount = "", connectionPassword = "";
                            foreach (var item in resultInterfacing)
                            {
                                connectionServer = item.ConnectionServer;
                                connectionAccount = item.connectionAccount;
                                connectionPassword = item.connectionPassword;
                            }
                            #endregion

                            #region //檢驗AD帳號狀態
                            using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, connectionServer, connectionAccount, connectionPassword))
                            {
                                using (UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, account))
                                {
                                    if (userPrincipal == null)
                                    {
                                        classicVerify = true;
                                    }
                                    else
                                    {
                                        if (!(bool)userPrincipal.Enabled) throw new Exception("AD帳號已關閉!");
                                        if (userPrincipal.IsAccountLockedOut()) throw new Exception("AD帳號已鎖定!");
                                    }
                                }
                            }
                            #endregion
                        }

                        int rowsAffected = 0;
                        if (classicVerify)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.UserInfo SET
                                    Password = @Password,
                                    PasswordStatus = @PasswordStatus,
                                    PasswordMistake = @PasswordMistake,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UserId = @UserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Password = BaseHelper.Sha256Encrypt(account.ToString().ToLower()),
                                    PasswordStatus = "Y",
                                    PasswordMistake = 0,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    UserId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "密碼重置成功!"
                            });
                            #endregion
                        }
                        else
                        {
                            throw new SystemException("該帳號採用AD登入，無法重置!");
                        }
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

        #region //UpdateUserInfoCopy -- 新舊使用者介接程式 -- Yi 2023.09.06
        public string UpdateUserInfoCopy()
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //清空使用者內部員工設定檔表格資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserEmployee";
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //清空使用者表格資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserInfo";
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        //#region //SQL識別規格改成"否"(可跳號編碼)
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SET IDENTITY_INSERT BAS.UserInfo ON;";
                        //rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        //#endregion

                        #region //自動編號(identity)歸零設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"DBCC CHECKIDENT('BAS.UserInfo', RESEED, 0)";
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //複製舊表格對應欄位資料(BAS.User)至新表格(BAS.UserInfo)
                        dynamicParameters = new DynamicParameters();
                        sql = @"SET IDENTITY_INSERT BAS.UserInfo ON;
                                INSERT INTO BAS.UserInfo
                                (UserId, UserType, Account, [Password], InnerCode, DisplayName, PasswordStatus, PasswordExpire, PasswordMistake, [Status], CreateDate)
                                SELECT a.UserId, 'E', a.UserNo, a.[Password], a.UserNo, a.UserName, a.PasswordStatus, a.PasswordExpire, a.PasswordMistake, a.[Status], a.CreateDate
                                FROM BAS.[User] a
                                SET IDENTITY_INSERT BAS.UserInfo OFF;";
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //複製舊表格對應欄位資料(BAS.User)至新表格(BAS.UserEmployee)
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.UserEmployee
                                (UserId, DepartmentId, Gender, Email, Job, JobType, EmployeeStatus, CreateDate, CreateBy)
                                SELECT a.UserId, a.DepartmentId, a.Gender, a.Email, ISNULL(a.Job, ''), ISNULL(a.JobType, ''), a.UserStatus, a.CreateDate, a.UserId
                                FROM BAS.[User] a";
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        //#region //SQL識別規格改成"是"(不可跳號編碼)
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SET IDENTITY_INSERT BAS.UserInfo OFF;";
                        //rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        //#endregion

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

        #region //UpdateDepartmentCopy -- 新舊部門介接程式 -- Ben Ma 2023.10.06
        public string UpdateDepartmentCopy()
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //清空部門資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE HRM.Department";
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //自動編號(identity)歸零設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"DBCC CHECKIDENT('HRM.Department', RESEED, 0)";
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //複製舊表格對應欄位資料(BAS.Department)至新表格(HRM.Department)
                        dynamicParameters = new DynamicParameters();
                        sql = @"SET IDENTITY_INSERT HRM.Department ON;
                                INSERT INTO HRM.Department
                                (DepartmentId, CorporationId, DepartmentNo, ShortName, FullName, [Status], CreateDate, CreateBy)
                                SELECT a.DepartmentId, c.CorporationId, a.DepartmentNo, a.DepartmentName, a.DepartmentName, a.[Status], a.CreateDate, 749
                                FROM BAS.Department a
                                INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                                INNER JOIN HRM.Corporation c ON b.CompanyId = c.CompanyId
                                SET IDENTITY_INSERT HRM.Department OFF;";
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTeamStatus -- 更新團隊狀態 -- Ann 2024-01-16
        public string UpdateTeamStatus(int TeamId, string Status, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");
                        #endregion

                        #region //判斷團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.Teams a 
                                WHERE a.TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var TeamsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TeamsResult.Count() <= 0) throw new SystemException("團隊資料有誤!!");
                        #endregion

                        #region //更新團隊狀態
                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MAMO.Teams SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TeamId = @TeamId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                TeamId
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

        #region //UpdateChannelStatus -- 更新頻道狀態 -- Ann 2024-01-16
        public string UpdateChannelStatus(int ChannelId, string Status, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷頻道資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.Channels a 
                                WHERE a.ChannelId = @ChannelId";
                        dynamicParameters.Add("ChannelId", ChannelId);

                        var ChannelsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ChannelsResult.Count() <= 0) throw new SystemException("頻道資料有誤!!");
                        #endregion

                        #region //更新團隊狀態
                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MAMO.Channels SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ChannelId = @ChannelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                ChannelId
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
        #region //DeleteApiKey -- Api金鑰刪除 -- Ben Ma 2022.04.27
        public string DeleteApiKey(int KeyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷金鑰資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKey
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("金鑰資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.ApiKeyFunction
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.ApiKey
                                WHERE KeyId = @KeyId";
                        dynamicParameters.Add("KeyId", KeyId);

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

        #region //DeleteApiKeyFunction -- Api適用功能刪除 -- Ben Ma 2022.04.29
        public string DeleteApiKeyFunction(int FunctionId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷適用功能資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.ApiKeyFunction
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("適用功能資料錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.ApiKeyFunction
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRoleUser -- 角色使用者刪除 -- Ben Ma 2022.05.16
        public string DeleteRoleUser(int UserId, int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.UserRole
                                WHERE UserId = @UserId
                                AND RoleId = @RoleId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserRole
                                WHERE UserId = @UserId
                                AND RoleId = @RoleId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("RoleId", RoleId);

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

        #region //DeleteMailServer -- 郵件伺服器刪除 -- Zoey 2022.05.25
        public string DeleteMailServer(int ServerId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷伺服器資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.MailServer
                                WHERE ServerId = @ServerId";
                        dynamicParameters.Add("ServerId", ServerId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("伺服器資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Mail
                                WHERE ServerId = @ServerId";
                        dynamicParameters.Add("ServerId", ServerId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.MailServer
                                WHERE ServerId = @ServerId";
                        dynamicParameters.Add("ServerId", ServerId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteMailContact -- 郵件聯絡人資料刪除 -- Zoey 2022.05.26
        public string DeleteMailContact(int ContactId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷聯絡人資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.MailContact
                                WHERE ContactId = @ContactId";
                        dynamicParameters.Add("ContactId", ContactId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("聯絡人資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.MailContact
                                WHERE ContactId = @ContactId";
                        dynamicParameters.Add("ContactId", ContactId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteMailTemplate -- 郵件樣板刪除 -- Zoey 2022.05.27
        public string DeleteMailTemplate(int MailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷郵件樣板是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Mail
                                WHERE MailId = @MailId";
                        dynamicParameters.Add("MailId", MailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("郵件樣板錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Mail
                                WHERE MailId = @MailId";
                        dynamicParameters.Add("MailId", MailId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteCalendar -- 行事曆資料刪除 -- Zoey 2022.06.07
        public string DeleteCalendar(string CalendarDate)
        {
            try
            {
                if (!DateTime.TryParse(CalendarDate, out DateTime tempDate)) throw new SystemException("【日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Calendar
                                WHERE CalendarDate = @CalendarDate";
                        dynamicParameters.Add("CalendarDate", CalendarDate);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteFile -- 檔案資料刪除 -- Ben Ma 2022.06.13
        public string DeleteFile(int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[File] SET
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                        var parametersObject = new
                        {
                            DeleteStatus = "Y",
                            LastModifiedDate,
                            LastModifiedBy,
                            FileId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //DeleteFilePermanent -- 檔案資料永久刪除 -- Ben Ma 2022.06.15
        public string DeleteFilePermanent(int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //紀錄參數
                        string logMsg = "",
                            fileName = "",
                            fileExtension = "",
                            deleteUser = "";
                        #endregion

                        #region //判斷檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 [FileName], FileExtension
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案資料錯誤!");

                        foreach (var item in result)
                        {
                            fileName = item.FileName;
                            fileExtension = item.FileExtension;
                        }
                        #endregion

                        #region //刪除者資料
                        sql = @"SELECT TOP 1 UserNo, UserName
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("刪除者資料錯誤!");

                        foreach (var item in result2)
                        {
                            deleteUser = item.UserNo + " " + item.UserName;
                        }
                        #endregion

                        logMsg = string.Format("【{0}{1}】已由【{2}】永久刪除", fileName, fileExtension, deleteUser);

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        logger.Trace(logMsg);
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

        #region //DeleteLotFilePermanent -- 檔案資料永久刪除(批次) -- Ben Ma 2023.08.16
        public string DeleteLotFilePermanent()
        {
            try
            {
                IEnumerable<dynamic> resultForeignKey;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //尋找相關foreign key
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT obj.name AS FK_NAME,
                                    sch.name AS [schema_name],
                                    tab1.name AS [table],
                                    col1.name AS [column],
                                    tab2.name AS [referenced_table],
                                    col2.name AS [referenced_column]
                                FROM sys.foreign_key_columns fkc
                                INNER JOIN sys.objects obj
                                    ON obj.object_id = fkc.constraint_object_id
                                INNER JOIN sys.tables tab1
                                    ON tab1.object_id = fkc.parent_object_id
                                INNER JOIN sys.schemas sch
                                    ON tab1.schema_id = sch.schema_id
                                INNER JOIN sys.columns col1
                                    ON col1.column_id = parent_column_id AND col1.object_id = tab1.object_id
                                INNER JOIN sys.tables tab2
                                    ON tab2.object_id = fkc.referenced_object_id
                                INNER JOIN sys.columns col2
                                    ON col2.column_id = referenced_column_id AND col2.object_id = tab2.object_id
                                WHERE tab2.name = 'File'
                                AND col2.name = 'FileId'";
                    resultForeignKey = sqlConnection.Query(sql, dynamicParameters);
                    if (resultForeignKey.Count() <= 0) throw new SystemException("Foreign Key資料錯誤!");
                    #endregion
                }

                int rowsAffected = 0, currentRowsAffected = 1;

                while (currentRowsAffected != 0)
                {
                    using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 500, 0)))
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //刪除檔案資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM BAS.[File] a
                                    WHERE a.FileId IN (
                                        SELECT TOP 10 a.FileId
                                        FROM BAS.[File] a
                                        WHERE 1=1";

                            int index = 0;
                            sql += @" AND (";
                            foreach (var item in resultForeignKey)
                            {
                                if (index == 0)
                                {
                                    sql += @" NOT EXISTS (SELECT TOP 1 1 FROM " + item.schema_name + @"." + item.table + @" WHERE " + item.column + @"=a.FileId)";
                                }
                                else
                                {
                                    sql += @" AND NOT EXISTS (SELECT TOP 1 1 FROM " + item.schema_name + @"." + item.table + @" WHERE " + item.column + @"=a.FileId)";
                                }

                                index++;
                            }
                            sql += @" )
                                    )";
                            currentRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            rowsAffected += currentRowsAffected;
                            #endregion
                        }

                        transactionScope.Complete();
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

        #region //DeleteSubSystem -- 子系統資料刪除 -- Ben Ma 2022.11.23
        public string DeleteSubSystem(int SubSystemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystem
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.SubSystemUser
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.SubSystem
                                WHERE SubSystemId = @SubSystemId";
                        dynamicParameters.Add("SubSystemId", SubSystemId);

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

        #region //DeleteSubSystemUser -- 子系統使用者資料刪除 -- Ben Ma 2022.11.23
        public string DeleteSubSystemUser(int SubSystemUserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.SubSystemUser
                                WHERE SubSystemUserId = @SubSystemUserId";
                        dynamicParameters.Add("SubSystemUserId", SubSystemUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統使用者資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.SubSystemUser
                                WHERE SubSystemUserId = @SubSystemUserId";
                        dynamicParameters.Add("SubSystemUserId", SubSystemUserId);

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

        #region //DeletePushSubscription -- 推播訂閱資料刪除 -- Ben Ma 2023.01.31
        public string DeletePushSubscription(int PushSubscriptionId, string ApiEndpoint)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷推播訂閱資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.PushSubscription
                                WHERE PushSubscriptionId = @PushSubscriptionId
                                AND ApiEndpoint = @ApiEndpoint";
                        dynamicParameters.Add("PushSubscriptionId", PushSubscriptionId);
                        dynamicParameters.Add("ApiEndpoint", ApiEndpoint);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("推播訂閱資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.PushSubscription
                                WHERE PushSubscriptionId = @PushSubscriptionId";
                        dynamicParameters.Add("PushSubscriptionId", PushSubscriptionId);

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

        #region //DeleteAuthorityUserCompany -- 使用者公司權限資料刪除 -- Ben Ma 2023.05.17
        public string DeleteAuthorityUserCompany(int UserId, int CompanyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.AuthorityUserCompany
                                WHERE UserId = @UserId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("CompanyId", CompanyId);

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

        #region //DeleteUser -- 刪除新版使用者資料 -- Yi 2023-09-01
        public string DeleteUser(int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.UserInfo
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯Table
                        #region //內部員工設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserEmployee
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //客戶設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserCustomer
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //供應商設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserSupplier
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //尋找相關foreign key
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT obj.name AS FK_NAME,
                                    sch.name AS [schema_name],
                                    tab1.name AS [table],
                                    col1.name AS [column],
                                    tab2.name AS [referenced_table],
                                    col2.name AS [referenced_column]
                                FROM sys.foreign_key_columns fkc
                                INNER JOIN sys.objects obj
                                    ON obj.object_id = fkc.constraint_object_id
                                INNER JOIN sys.tables tab1
                                    ON tab1.object_id = fkc.parent_object_id
                                INNER JOIN sys.schemas sch
                                    ON tab1.schema_id = sch.schema_id
                                INNER JOIN sys.columns col1
                                    ON col1.column_id = parent_column_id AND col1.object_id = tab1.object_id
                                INNER JOIN sys.tables tab2
                                    ON tab2.object_id = fkc.referenced_object_id
                                INNER JOIN sys.columns col2
                                    ON col2.column_id = referenced_column_id AND col2.object_id = tab2.object_id
                                WHERE tab2.name = 'UserInfo'
                                AND col2.name = 'UserId'";
                        var resultForeignKey = sqlConnection.Query(sql, dynamicParameters);
                        if (resultForeignKey.Count() <= 0) throw new SystemException("Foreign Key資料錯誤!");
                        #endregion

                        #region //查找相關資料
                        if (resultForeignKey.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.UserInfo a
                                    WHERE a.UserId = @UserId";

                            int index = 0;
                            sql += @" AND (";
                            foreach (var item in resultForeignKey)
                            {
                                if (index == 0)
                                {
                                    sql += @" NOT EXISTS (SELECT TOP 1 1 FROM " + item.schema_name + @"." + item.table + @" WHERE " + item.column + @"=a.UserId)";
                                }
                                else
                                {
                                    sql += @" AND NOT EXISTS (SELECT TOP 1 1 FROM " + item.schema_name + @"." + item.table + @" WHERE " + item.column + @"=a.UserId)";
                                }

                                index++;
                            }
                            sql += @" )";
                            dynamicParameters.Add("UserId", UserId);
                            var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                            if (resultUserInfo.Count() <= 0) throw new SystemException("已有相關使用者紀錄，無法刪除!");
                        }
                        #endregion
                        
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UserInfo
                                WHERE UserId = @UserId";
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

        #region //DeleteTeamMembers -- 刪除團隊成員 -- Ann 2024-01-16
        public string DeleteTeamMembers(int TeamId, List<string> UserNoList)
        {
            try
            {
                if (UserNoList.Count <= 0) throw new SystemException("至少需指定一位團隊成員!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                FROM MAMO.Teams a 
                                WHERE a.TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var TeamsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TeamsResult.Count() <= 0) throw new SystemException("團隊資料錯誤!!");

                        foreach (var item in TeamsResult)
                        {
                            if (item.Status != "A") throw new SystemException("團隊狀態非啟用中，無法更新!!");
                        }
                        #endregion

                        foreach (var userNo in UserNoList)
                        {
                            #region //確認USER相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                            dynamicParameters.Add("UserNo", userNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                            int UserId = -1;
                            foreach (var item in UserResult)
                            {
                                UserId = item.UserId;
                            }
                            #endregion

                            #region //確認團隊成員資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MAMO.TeamMembers a 
                                    WHERE a.TeamId = @TeamId
                                    AND a.UserId = @UserId";
                            dynamicParameters.Add("TeamId", TeamId);
                            dynamicParameters.Add("UserId", UserId);

                            var TeamMembersResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TeamMembersResult.Count() <= 0) throw new SystemException("使用者【" + userNo + "】非此團隊成員!!");
                            #endregion

                            #region //刪除團隊成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MAMO.TeamMembers
                                    WHERE TeamId = @TeamId
                                    AND UserId = @UserId";
                            dynamicParameters.Add("TeamId", TeamId);
                            dynamicParameters.Add("UserId", UserId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteChannelMembers -- 刪除頻道成員 -- Ann 2024-01-16
        public string DeleteChannelMembers(int ChannelId, List<string> UserNoList)
        {
            try
            {
                if (UserNoList.Count <= 0) throw new SystemException("至少需指定一位團隊成員!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //確認頻道資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                FROM MAMO.Channels a 
                                WHERE a.ChannelId = @ChannelId";
                        dynamicParameters.Add("ChannelId", ChannelId);

                        var ChannelsResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ChannelsResult.Count() <= 0) throw new SystemException("頻道資料錯誤!!");

                        foreach (var item in ChannelsResult)
                        {
                            if (item.Status != "A") throw new SystemException("頻道狀態非啟用中，無法更新!!");
                        }
                        #endregion

                        foreach (var userNo in UserNoList)
                        {
                            #region //確認USER相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                            dynamicParameters.Add("UserNo", userNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                            int UserId = -1;
                            foreach (var item in UserResult)
                            {
                                UserId = item.UserId;
                            }
                            #endregion

                            #region //確認頻道成員資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MAMO.ChannelMembers a 
                                    WHERE a.ChannelId = @ChannelId
                                    AND a.UserId = UserId";
                            dynamicParameters.Add("ChannelId", ChannelId);
                            dynamicParameters.Add("UserId", UserId);

                            var ChannelMembersResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ChannelMembersResult.Count() <= 0) throw new SystemException("使用者【" + userNo + "】非此團隊成員!!");
                            #endregion

                            #region //刪除頻道成員
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MAMO.ChannelMembers
                                    WHERE ChannelId = @ChannelId
                                    AND UserId = @UserId";
                            dynamicParameters.Add("ChannelId", ChannelId);
                            dynamicParameters.Add("UserId", UserId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteNLog -- 刪除NLog紀錄 -- Ann 2024-06-28
        public string DeleteNLog(string DeleteRange)
        {
            try
            {
                if (DeleteRange.Length <= 0) throw new SystemException("刪除時間範圍錯誤!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM BAS.[Log] WHERE CreateDate <= DATEADD(DAY, " + DeleteRange + ", GETDATE())";

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
    }
}
