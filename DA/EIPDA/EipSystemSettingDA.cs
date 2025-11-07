using System;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using Newtonsoft.Json.Linq;
using Dapper;
using Helpers;
using NLog;
using System.Transactions;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace EIPDA
{
    public class EipSystemSettingDA
    {
        public string MainConnectionStrings = string.Empty;
        public string OfficialConnectionStrings = string.Empty;

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

        public EipSystemSettingDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            OfficialConnectionStrings = ConfigurationManager.AppSettings["OfficialDb"];

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
        #region //GetStatus -- 取得狀態資料 -- Chia Yuan
        public string GetStatus(string StatusSchema, string StatusNo, string StatusName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.StatusId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.StatusSchema, a.StatusNo, a.StatusName, a.StatusDesc
                        , ISNULL(b.StatusName, a.StatusName) ApplyStatusName, ISNULL(b.StatusDesc, a.StatusDesc) ApplyStatusDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Status] a
                        OUTER APPLY (
                            SELECT ba.StatusName, ba.StatusDesc
                            FROM BAS.StatusCompany ba
                            WHERE ba.StatusId = a.StatusId
                            AND ba.CompanyId = @CompanyId
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusNo", @" AND a.StatusNo = @StatusNo", StatusNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusName", @" AND (a.StatusName LIKE '%' + @StatusName + '%' OR a.StatusDesc LIKE '%' + @StatusName + '%')", StatusName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.StatusSchema, a.StatusNo";
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

        #region //GetStatusSchema -- 取得狀態綱要 -- Chia Yuan
        public string GetStatusSchema(string StatusSchema)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.StatusSchema
                            FROM BAS.[Status] a";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);

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

        #region //GetStatusCompany -- 取得狀態公司對應資料 -- Chia Yuan
        public string GetStatusCompany(int StatusId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, ISNULL(a.LogoIcon, -1) AS LogoIcon
                        , ISNULL(b.StatusName, '') StatusName, ISNULL(b.StatusDesc, '') StatusDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a
                        OUTER APPLY (
                            SELECT ba.StatusName, ba.StatusDesc
                            FROM BAS.StatusCompany ba
                            WHERE ba.CompanyId = a.CompanyId
                            AND ba.StatusId = @StatusId
                        ) b";
                    sqlQuery.auxTables = "";
                    dynamicParameters.Add("StatusId", StatusId);
                    string queryCondition = "";
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "a.CompanyId";

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

        #region //GetType -- 取得類別資料 -- Chia Yuan
        public string GetType(string TypeSchema, string TypeNo, string TypeName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.TypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TypeSchema, a.TypeNo, a.TypeName, a.TypeDesc
                        , ISNULL(b.TypeName, a.TypeName) ApplyTypeName, ISNULL(b.TypeDesc, a.TypeDesc) ApplyTypeDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Type] a
                        OUTER APPLY (
                            SELECT ba.TypeName, ba.TypeDesc
                            FROM BAS.TypeCompany ba
                            WHERE ba.TypeId = a.TypeId
                            AND ba.CompanyId = @CompanyId
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeSchema", @" AND a.TypeSchema = @TypeSchema", TypeSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeNo", @" AND a.TypeNo = @TypeNo", TypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeName", @" AND (a.TypeName LIKE '%' + @TypeName + '%' OR a.TypeDesc LIKE '%' + @TypeName + '%')", TypeName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TypeSchema, a.TypeNo";
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

        #region //GetTypeSchema -- 取得類別綱要 -- Chia Yuan
        public string GetTypeSchema(string TypeSchema
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.TypeSchema
                            FROM BAS.[Type] a";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeSchema", @" AND a.TypeSchema = @TypeSchema", TypeSchema);

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

        #region //GetTypeCompany -- 取得類別公司對應資料 -- Chia Yuan
        public string GetTypeCompany(int TypeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, ISNULL(a.LogoIcon, -1) AS LogoIcon
                        , ISNULL(b.TypeName, '') TypeName, ISNULL(b.TypeDesc, '') TypeDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a
                        OUTER APPLY (
                            SELECT ba.TypeName, ba.TypeDesc
                            FROM BAS.TypeCompany ba
                            WHERE ba.CompanyId = a.CompanyId
                            AND ba.TypeId = @TypeId
                        ) b";
                    sqlQuery.auxTables = "";
                    dynamicParameters.Add("TypeId", TypeId);
                    string queryCondition = "";
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "a.CompanyId";

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

        #region //GetCompany -- 取得公司資料 -- Chia Yuan
        public string GetCompany(int CompanyId, string CompanyNo, string CompanyName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, a.Telephone, a.Fax, a.Address
                        , a.AddressEn, ISNULL(a.LogoIcon, -1) LogoIcon, a.Status
                        , a.CompanyNo + ' ' + a.CompanyName CompanyWithNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyNo", @" AND a.CompanyNo LIKE '%' + @CompanyNo + '%'", CompanyNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyName", @" AND a.CompanyName LIKE '%' + @CompanyName + '%'", CompanyName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CompanyId";
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

        #region //GetSystem -- 取得系統別資料 -- Chia Yuan
        public string GetSystem(int SystemId, string SystemCode, string SystemName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.SystemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber, a.Status
                        , a.SystemCode + ' ' + a.SystemName SystemWithCode";
                    sqlQuery.mainTables =
                        @"FROM BAS.[System] a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemId", @" AND a.SystemId = @SystemId", SystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemCode", @" AND a.SystemCode LIKE '%' + @SystemCode + '%'", SystemCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemName", @" AND a.SystemName LIKE '%' + @SystemName + '%'", SystemName);
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

        #region //GetSystemMenu -- 取得系統選單資料 -- Chia Yuan
        public string GetSystemMenu(int MemberId)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    #region //取得會員權限
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT d.SystemId
                            FROM EIP.RoleFunctionDetail a
                            INNER JOIN EIP.FunctionDetail b ON a.FnDetailId = b.FnDetailId
                            INNER JOIN EIP.[Function] c ON b.FunctionId = c.FunctionId
                            INNER JOIN EIP.[Module] d ON c.ModuleId = d.ModuleId
                            WHERE a.RoleId IN (
                                SELECT aa.RoleId
                                FROM EIP.UserRole aa
                                INNER JOIN EIP.[Role] ab ON aa.RoleId = ab.RoleId
                                WHERE aa.MemberId = @MemberId
                            )
                            AND b.FnDetailCode IN ('read')
                            AND b.[Status] = @Status
                            AND c.[Status] = @Status
                            AND d.[Status] = @Status";
                    dynamicParameters.Add("MemberId", MemberId);
                    dynamicParameters.Add("Status", "A");
                    var resultDetail = officialConnection.Query(sql, dynamicParameters);
                    #endregion

                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status
                                AND a.SystemId IN @SystemIds";
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("SystemIds", resultDetail.Select(s => s.SystemId));
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                  join s2 in resultDetail
                                    on s1.SystemId equals s2.SystemId
                                  select new
                                  {
                                      s1.SystemId,
                                      s1.SystemCode,
                                      s1.SystemName,
                                      s1.ThemeIcon,
                                      s1.SortNumber
                                  })
                                  .OrderBy(o => o.SortNumber )
                                  .ToList();
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

        #region //GetModule -- 取得模組別資料 -- Chia Yuan
        public string GetModule(int ModuleId, int SystemId, string ModuleCode, string ModuleName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status";
                        dynamicParameters.Add("Status", "A");
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SystemId", @" AND a.SystemId = @SystemId", SystemId);
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    #region //取得模組資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ModuleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SystemId, a.ModuleCode, a.ModuleName, a.ThemeIcon, a.[Status], a.SortNumber
                        , a.ModuleCode + ' ' + a.ModuleName AS ModuleWithCode";
                    sqlQuery.mainTables =
                        @"FROM EIP.Module a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.SystemId IN @SystemIds";
                    dynamicParameters.Add("SystemIds", resultSystem.Select(s => s.SystemId).Distinct().ToArray());
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleId", @" AND a.ModuleId = @ModuleId", ModuleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleCode", @" AND a.ModuleCode = @ModuleCode", ModuleCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleName", @" AND a.ModuleName LIKE '%' + @ModuleName + '%'", ModuleName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var resultModule = BaseHelper.SqlQuery(officialConnection, dynamicParameters, sqlQuery).ToList();
                    #endregion

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                 join s2 in resultModule
                                   on s1.SystemId equals s2.SystemId
                                 select new
                                 {
                                     s1.SystemId,
                                     s1.SystemCode,
                                     s1.SystemName,
                                     s2.ModuleId,
                                     s2.ModuleCode,
                                     s2.ModuleName,
                                     s2.ThemeIcon,
                                     s2.Status,
                                     s2.ModuleWithCode,
                                     SortNumber1 = s1.SortNumber,
                                     SortNumber2 = s2.SortNumber,
                                     s2.TotalCount
                                 })
                                 .OrderBy(o => o.SortNumber1)
                                 .ThenBy(t => t.SortNumber2)
                                 .ToList();
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

        #region //GetModuleMenu -- 取得模組選單資料 -- Chia Yuan
        public string GetModuleMenu(string SystemCode, int MemberId)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status";
                        dynamicParameters.Add("Status", "A");
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SystemCode", @" AND a.SystemCode = @SystemCode", SystemCode);
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT d.ModuleId, d.SystemId, d.ModuleCode, d.ModuleName, d.ThemeIcon, d.[Status], d.SortNumber
                        , (
                            SELECT aa.FunctionId, aa.FunctionCode, aa.FunctionName, ISNULL(aa.UrlTarget, '') AS UrlTarget
                            FROM EIP.[Function] aa
                            INNER JOIN EIP.[Module] ab ON aa.ModuleId = ab.ModuleId
                            WHERE ab.ModuleId = d.ModuleId
                            AND aa.FunctionId IN (
                                SELECT y.FunctionId
                                FROM EIP.RoleFunctionDetail x
                                INNER JOIN EIP.FunctionDetail y ON x.FnDetailId = y.FnDetailId
                                WHERE x.RoleId IN (
                                    SELECT xa.RoleId
                                    FROM EIP.UserRole xa
                                    INNER JOIN EIP.[Role] xb ON xa.RoleId = xb.RoleId
                                    WHERE xa.MemberId = @MemberId
                                )
                                AND y.FnDetailCode IN ('read')
                                AND y.[Status] = @Status
                            )
                            AND aa.[Status] = @Status
                            AND ab.[Status] = @Status
                            ORDER BY aa.SortNumber
                            FOR JSON PATH, ROOT('data')
                        ) FunctionData
                        FROM EIP.RoleFunctionDetail a
                        INNER JOIN EIP.FunctionDetail b ON a.FnDetailId = b.FnDetailId
                        INNER JOIN EIP.[Function] c ON b.FunctionId = c.FunctionId
                        INNER JOIN EIP.[Module] d ON c.ModuleId = d.ModuleId
                        WHERE a.RoleId IN (
                            SELECT aa.RoleId
                            FROM EIP.UserRole aa
                            INNER JOIN EIP.[Role] ab ON aa.RoleId = ab.RoleId
                            WHERE aa.MemberId = @MemberId
                        )
                        AND b.FnDetailCode IN ('read')
                        AND b.[Status] = @Status
                        AND c.[Status] = @Status
                        AND d.[Status] = @Status
                        ORDER BY d.SortNumber";
                    dynamicParameters.Add("MemberId", MemberId);
                    dynamicParameters.Add("Status", "A");

                    var resultModule = officialConnection.Query(sql, dynamicParameters);

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                 join s2 in resultModule
                                   on s1.SystemId equals s2.SystemId
                                 select new
                                 {
                                     s1.SystemId,
                                     s1.SystemCode,
                                     s1.SystemName,
                                     s2.ModuleId,
                                     s2.ModuleCode,
                                     s2.ModuleName,
                                     s2.ThemeIcon,
                                     s2.FunctionData,
                                     SortNumber1 = s1.SortNumber,
                                     SortNumber2 = s2.SortNumber
                                 })
                                 .OrderBy(o => o.SortNumber1)
                                 .ThenBy(t => t.SortNumber2)
                                 .ToList();
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

        #region //GetFunction -- 取得功能別資料 -- Chia Yuan
        public string GetFunction(int FunctionId, int ModuleId, int SystemId, string FunctionCode, string FunctionName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status";
                        dynamicParameters.Add("Status", "A");
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SystemId", @" AND a.SystemId = @SystemId", SystemId);
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.FunctionId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FunctionCode, a.FunctionName, a.UrlTarget, a.[Status], a.SortNumber AS SortNumber3
                        , b.ModuleId, b.SystemId, b.ModuleCode, b.ModuleName, b.SortNumber AS SortNumber2";
                    sqlQuery.mainTables =
                        @"FROM EIP.[Function] a
                        INNER JOIN EIP.Module b ON a.ModuleId = b.ModuleId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND b.SystemId IN @SystemIds";
                    dynamicParameters.Add("SystemIds", resultSystem.Select(s => s.SystemId).Distinct().ToArray());
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionId", @" AND a.FunctionId = @FunctionId", FunctionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleId", @" AND b.ModuleId = @ModuleId", ModuleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionCode", @" AND a.FunctionCode LIKE '%' + @FunctionCode + '%'", FunctionCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionName", @" AND a.FunctionName LIKE '%' + @FunctionName + '%'", FunctionName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.SortNumber, a.SortNumber";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var resultFunction = BaseHelper.SqlQuery(officialConnection, dynamicParameters, sqlQuery);

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                 join s2 in resultFunction
                                   on s1.SystemId equals s2.SystemId
                                 select new
                                 {
                                     s1.SystemId,
                                     s1.SystemCode,
                                     s1.SystemName,
                                     s2.FunctionId,
                                     s2.FunctionCode,
                                     s2.FunctionName,
                                     s2.UrlTarget,
                                     s2.Status,
                                     s2.ModuleId,
                                     s2.ModuleCode,
                                     s2.ModuleName,
                                     SortNumber1 = s1.SortNumber,
                                     s2.SortNumber2,
                                     s2.SortNumber3,
                                     s2.TotalCount
                                 })
                                 .OrderBy(o => o.SortNumber1)
                                 .ThenBy(t => t.SortNumber2)
                                 .ThenBy(t => t.SortNumber3)
                                 .ToList();
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

        #region //GetFunctionDetail -- 取得詳細功能別資料 -- Chia Yuan
        public string GetFunctionDetail(int FnDetailId, int FunctionId, string FnDetailCode, string FnDetailName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.FnDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FnDetailCode, a.FnDetailName, a.SortNumber, a.Status
                        , b.FunctionId, b.FunctionCode, b.FunctionName";
                    sqlQuery.mainTables =
                        @"FROM EIP.FunctionDetail a
                        INNER JOIN EIP.[Function] b ON a.FunctionId = b.FunctionId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FnDetailId", @" AND a.FnDetailId = @FnDetailId", FnDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionId", @" AND b.FunctionId = @FunctionId", FunctionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FnDetailCode", @" AND a.FnDetailCode LIKE '%' + @FnDetailCode + '%'", FnDetailCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FnDetailName", @" AND a.FnDetailName LIKE '%' + @FnDetailName + '%'", FnDetailName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.FunctionId, a.SortNumber";
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

        #region //GetFunctionDetailCode -- 取得功能詳細代碼 -- Chia Yuan
        public string GetFunctionDetailCode()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.FnDetailCode, a.FnDetailName
                            FROM EIP.FunctionDetail a
                            WHERE 1=1
                            ORDER BY a.FnDetailCode";

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

        #region //GetRole -- 取得角色資料 -- Chia Yuan
        public string GetRole(int RoleId, string RoleName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoleName, a.AdminStatus, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EIP.Role a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
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

        #region //GetAuthorityUser -- 取得權限使用者資料 -- Chia Yuan
        public string GetAuthorityMember(string Roles, string MemberEmail, string MemberName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = ", a.MemberName, b.Roles";
                    sqlQuery.columns =
                        @", a.MemberEmail, a.MemberName, a.Gender, ISNULL(a.MemberIcon, -1) AS MemberIcon
                        , ISNULL(a.OrgShortName, '') AS OrgShortName
                        , CASE a.[Status] 
                            WHEN 'S' THEN a.MemberName + ' ' + a.MemberEmail + '(已停用)'
                            ELSE a.MemberName + ' ' + a.MemberEmail
                        END EmailWithMember
                        , ISNULL(b.CustomerName, '') AS CustomerName
                        , ISNULL(b.CustomerEnglishName, '') AS CustomerEnglishName
                        , ISNULL(b.CustomerName + ISNULL(' ' + b.CustomerEnglishName, ''), '') AS CNameAndEName
                        , ISNULL(a.MemberIcon, -1) AS MemberIcon
                        , (
                            SELECT ab.RoleId, ab.RoleName, ab.AdminStatus
                            FROM EIP.UserRole aa
                            INNER JOIN EIP.[Role] ab ON aa.RoleId = ab.RoleId
                            WHERE aa.MemberId = a.MemberId
                            ORDER BY ab.RoleId
                            FOR JSON PATH, ROOT('data')
                        ) UserRole";
                    sqlQuery.mainTables =
                        @"FROM EIP.[Member] a
                        LEFT JOIN EIP.CsCustomer b ON b.CsCustId = a.CsCustId";
                    string queryTable =
                        @"FROM (
                            SELECT a.MemberId, a.MemberEmail, a.MemberName, a.Gender
	                        , a.[Status]
                            , (
                                SELECT ab.RoleId, ab.RoleName, ab.AdminStatus
                                FROM EIP.UserRole aa
                                INNER JOIN EIP.[Role] ab ON aa.RoleId = ab.RoleId
                                WHERE aa.MemberId = a.MemberId
                                ORDER BY ab.RoleId
                                FOR JSON PATH, ROOT('data')
                            ) UserRole
                            FROM EIP.[Member] a
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
                    string queryCondition = "";
                    //dynamicParameters.Add("CurrentCompany", CurrentCompany);
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
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    //if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND a.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberEmail", @" AND a.MemberEmail LIKE '%' + @MemberEmail + '%'", MemberEmail);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND a.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MemberName";
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

        #region //GetMember -- 取得使用者資料 -- Chia Yuan
        public string GetMember(int MemberId, int CsCustId, string MemberEmail, string MemberName, string Status
            , string Gender, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CsCustId, a.MemberName, a.MemberEmail, a.Gender, ISNULL(a.MemberIcon, -1) AS MemberIcon
                        , ISNULL(a.OrgShortName, '') AS OrgShortName, a.[Address], a.ContactName, a.ContactPhone, a.[Description], a.[Status]
                        , CASE a.[Status] 
                            WHEN 'S' THEN a.MemberName + ' ' + a.MemberEmail + '(已停用)'
                            ELSE a.MemberName + ' ' + a.MemberEmail
                        END EmailWithMember
                        , ISNULL(b.CustomerName, '') AS CustomerName
                        , ISNULL(b.CustomerEnglishName, '') AS CustomerEnglishName
                        , ISNULL(b.CustomerName + ISNULL(' ' + b.CustomerEnglishName, ''), '未綁定') AS CNameAndEName";

                    if (MemberId > 0)
                    {
                        sqlQuery.columns += @", (
	                                            SELECT aa.CustomerName, aa.CustomerEnglishName
	                                            FROM EIP.CsCustomer aa
	                                            WHERE aa.CsCustId = a.CsCustId
	                                            ORDER BY aa.CsCustId
	                                            FOR JSON PATH, ROOT('data')
                                            ) AS Customers";
                    }

                    sqlQuery.mainTables =
                        @"FROM EIP.[Member] a
                        LEFT JOIN EIP.CsCustomer b ON b.CsCustId = a.CsCustId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CsCustId", @" AND b.CsCustId = @CsCustId", CsCustId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberEmail", @" AND a.MemberEmail LIKE '%' + @MemberEmail + '%'", MemberEmail);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND a.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    if (Gender.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Gender", @" AND a.Gender IN @Gender", Gender.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.MemberEmail LIKE '%' + @SearchKey + '%' OR a.MemberName LIKE '%' + @SearchKey + '%')", SearchKey);
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

        #region //GetCsCustomer -- 會員客戶對照表 -- Chia Yuan
        public string GetCsCustomer(int CsCustId, string CustomerName, string CustomerEnglishName, string SearchKey, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CsCustId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CustomerName
                        , ISNULL(a.CustomerEnglishName, '') AS CustomerEnglishName
                        , a.CustomerName + ISNULL(' ' + a.CustomerEnglishName, '') AS CNameAndEName
                        , a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EIP.CsCustomer a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CsCustId", @" AND a.CsCustId = @CsCustId", CsCustId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerEnglishName", @" AND a.CustomerEnglishName LIKE '%' + @CustomerEnglishName + '%'", CustomerEnglishName);
                    if (SearchKey.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND a.CustomerName LIKE  '%' + @SearchKey + '%' OR a.CustomerEnglishName LIKE  '%' + @SearchKey + '%'", SearchKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CsCustId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(officialConnection, dynamicParameters, sqlQuery);

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

        #region //GetCustomer
        public string GetCustomer(int CustomerId, int CompanyId, string CustomerName, string CustomerEnglishName, string SearchKey, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CustomerId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CustomerName
                        , ISNULL(a.CustomerEnglishName, '') AS CustomerEnglishName
                        , a.CustomerName + ISNULL(' ' + a.CustomerEnglishName, '') AS CNameAndEName
                        , b.CompanyName + ' ' + a.CustomerName AS CustWithCorp
                        , a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM SCM.Customer a
                        INNER JOIN BAS.Company b ON b.CompanyId = a.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerEnglishName", @" AND a.CustomerEnglishName LIKE '%' + @CustomerEnglishName + '%'", CustomerEnglishName);
                    if (SearchKey.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND a.CustomerName LIKE  '%' + @SearchKey + '%' OR a.CustomerEnglishName LIKE  '%' + @SearchKey + '%'", SearchKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CustomerId";
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

        #region //GetCsCustomerDetail -- 客戶明細表 -- Chia Yuan
        public string GetCsCustomerDetail(int CsCustId, int CustomerId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CsCustId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", CustomerId, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EIP.CsCustomerDetail a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CsCustId", @" AND a.CsCustId = @CsCustId", CsCustId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CsCustId, a.CustomerId";
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

        #region //GetRoleAuthority -- 取得角色權限資料 -- Chia Yuan
        public string GetRoleAuthority(int RoleId, string SearchKey)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    #region //取得功能角色資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SystemId, c.TotalFunction, d.RoleTotalFunction
                            , (
                                SELECT aa.ModuleId, aa.ModuleName, ab.TotalModuleFunction, ac.RoleTotalModuleFunction
                                FROM EIP.[Module] aa
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM EIP.FunctionDetail aba
                                    INNER JOIN EIP.[Function] abb ON aba.FunctionId = abb.FunctionId
                                    WHERE abb.ModuleId = aa.ModuleId
                                    AND abb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ab
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalModuleFunction
                                    FROM EIP.RoleFunctionDetail aca
                                    INNER JOIN EIP.FunctionDetail acb ON aca.FnDetailId = acb.FnDetailId
                                    INNER JOIN EIP.[Function] acc ON acb.FunctionId = acc.FunctionId
                                    WHERE acc.ModuleId = aa.ModuleId
                                    AND acc.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND aca.RoleId = @RoleId
                                ) ac
                                WHERE 1=1
                                AND aa.SystemId = a.SystemId
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) ModuleData
                            FROM (SELECT DISTINCT SystemId FROM EIP.Module) a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunction
                                FROM EIP.FunctionDetail ca
                                INNER JOIN EIP.[Function] cb ON ca.FunctionId = cb.FunctionId
                                INNER JOIN EIP.[Module] cc ON cb.ModuleId = cc.ModuleId
                                WHERE cc.SystemId = a.SystemId
                                AND cb.FunctionName LIKE '%' + @SearchKey + '%'
                            ) c
                            OUTER APPLY (
                                SELECT COUNT(1) RoleTotalFunction
                                FROM EIP.RoleFunctionDetail da
                                INNER JOIN EIP.FunctionDetail db ON da.FnDetailId = db.FnDetailId
                                INNER JOIN EIP.[Function] dc ON db.FunctionId = dc.FunctionId
                                INNER JOIN EIP.[Module] dd ON dc.ModuleId = dd.ModuleId
                                WHERE dd.SystemId = a.SystemId
                                AND dc.FunctionName LIKE '%' + @SearchKey + '%'
                                AND da.RoleId = @RoleId
                            ) d
                            WHERE 1=1";
                    dynamicParameters.Add("RoleId", RoleId);
                    dynamicParameters.Add("SearchKey", SearchKey);
                    var resultRole = officialConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status
                                AND a.SystemId IN @SystemIds";
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("SystemIds", resultRole.Select(s => s.SystemId));
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                  join s2 in resultRole
                                    on s1.SystemId equals s2.SystemId
                                  select new
                                  {
                                      s1.SystemId,
                                      s1.SystemCode,
                                      s1.SystemName,
                                      s1.SortNumber,
                                      s2.TotalFunction,
                                      s2.RoleTotalFunction,
                                      s2.ModuleData
                                  })
                                  .OrderBy(o => o.SortNumber )
                                 .ToList();
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

        #region //GetRoleDetailAuthority -- 取得角色詳細權限資料 -- Chia Yuan
        public string GetRoleDetailAuthority(int ModuleId, int RoleId, string SearchKey, bool Grant)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FunctionId, a.FunctionName, a.FunctionCode, a.SortNumber, b.TotalFunctionDetail, c.RoleTotalFunctionDetail
                            , (
                                SELECT aa.FnDetailId, aa.FnDetailName, aa.FnDetailCode, aa.Status, ab.DetailAuthority
                                FROM EIP.FunctionDetail aa
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM EIP.RoleFunctionDetail aaa
                                        WHERE aaa.FnDetailId = aa.FnDetailId
                                        AND aaa.RoleId = @RoleId
                                    ), 0) DetailAuthority
                                ) ab
                                WHERE 1=1
                                AND aa.FunctionId = a.FunctionId
                                AND ab.DetailAuthority >= @DetailAuthority
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) FunctionDetailData
                            FROM EIP.[Function] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunctionDetail
                                FROM EIP.FunctionDetail ba
                                WHERE ba.FunctionId = a.FunctionId
                            ) b
                            OUTER APPLY (
                                SELECT COUNT(1) RoleTotalFunctionDetail
                                FROM EIP.RoleFunctionDetail ca
                                INNER JOIN EIP.FunctionDetail cb ON ca.FnDetailId = cb.FnDetailId
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

        #region //GetRoleUser -- 取得角色使用者資料 -- Chia Yuan
        public string GetRoleUser(int RoleId, string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT c.MemberId, c.MemberEmail, c.MemberName, c.Status
                            FROM EIP.UserRole a
                            INNER JOIN EIP.Role b ON a.RoleId = b.RoleId
                            INNER JOIN EIP.[Member] c ON a.MemberId = c.MemberId
                            WHERE a.RoleId = @RoleId
                            AND (c.MemberEmail LIKE '%' + @SearchKey + '%' OR c.MemberName LIKE '%' + @SearchKey + '%')
                            ORDER BY c.Status, c.CsCustId";
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

        #region //GetUserAuthority -- 取得使用者權限資料 -- Chia Yuan
        public string GetUserAuthority(int MemberId, string SearchKey)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SystemId, b.TotalFunction, c.UserTotalFunction
                            , (
                                SELECT aa.ModuleId, aa.ModuleName, ab.TotalModuleFunction, ac.UserTotalModuleFunction
                                FROM EIP.[Module] aa
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM EIP.FunctionDetail aba
                                    INNER JOIN EIP.[Function] abb ON aba.FunctionId = abb.FunctionId
                                    WHERE abb.ModuleId = aa.ModuleId
                                    AND abb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ab
                                OUTER APPLY (
                                    SELECT COUNT(DISTINCT aca.FnDetailId) UserTotalModuleFunction
                                    FROM EIP.RoleFunctionDetail aca
                                    INNER JOIN EIP.FunctionDetail acb ON aca.FnDetailId = acb.FnDetailId
                                    INNER JOIN EIP.[Function] acc ON acb.FunctionId = acc.FunctionId
                                    WHERE acc.ModuleId = aa.ModuleId
                                    AND aca.RoleId IN (
                                        SELECT RoleId
                                        FROM EIP.UserRole
                                        WHERE MemberId = @MemberId
                                    )
                                    AND acc.FunctionName LIKE '%' + @SearchKey + '%'
                                ) ac
                                WHERE 1=1
                                AND aa.SystemId = a.SystemId
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) ModuleData
                            FROM (SELECT DISTINCT SystemId FROM EIP.[Module]) a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunction
                                FROM EIP.FunctionDetail ba
                                INNER JOIN EIP.[Function] bb ON ba.FunctionId = bb.FunctionId
                                INNER JOIN EIP.[Module] bc ON bb.ModuleId = bc.ModuleId
                                WHERE bc.SystemId = a.SystemId
                                AND bb.FunctionName LIKE '%' + @SearchKey + '%'
                            ) b
                            OUTER APPLY (
                                SELECT COUNT(DISTINCT ca.FnDetailId) UserTotalFunction
                                FROM EIP.RoleFunctionDetail ca
                                INNER JOIN EIP.FunctionDetail cb ON ca.FnDetailId = cb.FnDetailId
                                INNER JOIN EIP.[Function] cc ON cb.FunctionId = cc.FunctionId
                                INNER JOIN EIP.[Module] cd ON cc.ModuleId = cd.ModuleId
                                WHERE cd.SystemId = a.SystemId
                                AND ca.RoleId IN (
                                    SELECT RoleId
                                    FROM EIP.UserRole
                                    WHERE MemberId = @MemberId
                                )
                                AND cc.FunctionName LIKE '%' + @SearchKey + '%'
                            ) c";
                    dynamicParameters.Add("MemberId", MemberId);
                    dynamicParameters.Add("SearchKey", SearchKey);
                    var resultRole = officialConnection.Query(sql, dynamicParameters);

                    var resultSystem = new List<dynamic>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得系統資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SystemId, a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber
                                FROM BAS.[System] a
                                WHERE a.Status = @Status
                                AND a.SystemId IN @SystemIds";
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("SystemIds", resultRole.Select(s => s.SystemId));
                        resultSystem = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!resultSystem.Any()) throw new SystemException("無法取得系統列表!");
                        #endregion
                    }

                    #region //資料集合處理
                    var result = (from s1 in resultSystem
                                  join s2 in resultRole
                                    on s1.SystemId equals s2.SystemId
                                  select new
                                  {
                                      s1.SystemId,
                                      s1.SystemCode,
                                      s1.SystemName,
                                      s1.SortNumber,
                                      s2.TotalFunction,
                                      s2.UserTotalFunction,
                                      s2.ModuleData
                                  })
                                  .OrderBy(o =>  o.SortNumber )
                                  .ToList();
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

        #region //GetUserDetailAuthority -- 取得使用者詳細權限資料 -- Chia Yuan
        public string GetUserDetailAuthority(int ModuleId, int MemberId, int RoleId, string SearchKey, bool Grant)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FunctionId, a.FunctionName, a.FunctionCode, a.SortNumber, b.TotalFunctionDetail, c.UserTotalFunctionDetail
                            , (
	                            SELECT aa.RoleId, aa.RoleName
	                            , (
		                            SELECT ba.FnDetailId, ba.FnDetailCode, ba.FnDetailName, bb.DetailAuthority
		                            FROM EIP.FunctionDetail ba
		                            OUTER APPLY (
			                            SELECT ISNULL((
				                            SELECT TOP 1 1
				                            FROM EIP.UserRole bba
				                            INNER JOIN EIP.RoleFunctionDetail bbb ON bbb.RoleId = bba.RoleId
				                            WHERE bba.MemberId = @MemberId
				                            AND bbb.RoleId = aa.RoleId
				                            AND bbb.FnDetailId = ba.FnDetailId
			                            ), 0) DetailAuthority
		                            ) bb
		                            WHERE ba.FunctionId = a.FunctionId
                                    AND bb.DetailAuthority >= @DetailAuthority
		                            ORDER BY ba.SortNumber
		                            FOR JSON PATH, ROOT('data')
	                            ) AS FnRoles
	                            FROM EIP.[Role] aa
	                            ORDER BY aa.RoleId
	                            FOR JSON PATH, ROOT('data')
                            ) AS Roles
                            FROM EIP.[Function] a
                            OUTER APPLY (
                                SELECT COUNT(1) TotalFunctionDetail
                                FROM EIP.FunctionDetail ba
                                WHERE ba.FunctionId = a.FunctionId
                            ) b
                            OUTER APPLY (
	                            SELECT COUNT(cb.FnDetailId) UserTotalFunctionDetail
	                            FROM EIP.FunctionDetail ca
	                            INNER JOIN EIP.RoleFunctionDetail cb ON cb.FnDetailId = ca.FnDetailId
	                            WHERE ca.FunctionId = a.FunctionId
	                            AND cb.RoleId = @RoleId
	                            AND EXISTS (
		                            SELECT TOP 1 1 
		                            FROM EIP.UserRole cc 
		                            WHERE cc.RoleId = cb.RoleId 
		                            AND cc.MemberId = @MemberId
	                            )
                            ) c
                            WHERE a.ModuleId = @ModuleId
                            AND a.FunctionName LIKE '%' + @SearchKey + '%'
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("ModuleId", ModuleId);
                    dynamicParameters.Add("MemberId", MemberId);
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

        #region //GetUserRole -- 取得使用者角色資料 -- Chai Yuan
        public string GetUserRole(string SearchKey)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT ISNULL(b.RoleNames, '無權限') RoleNames, b.RoleIds
                            , (
                                SELECT aa.MemberId, aa.MemberEmail, aa.MemberName
                                FROM EIP.[Member] aa
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ',' + CONVERT(nvarchar(50), abb.RoleId)
                                        FROM EIP.UserRole aba
                                        INNER JOIN EIP.Role abb ON aba.RoleId = abb.RoleId
                                        WHERE aba.MemberId = aa.MemberId
                                        ORDER BY aba.RoleId
                                        FOR XML PATH('')
                                    ), 1, 1, '') RoleIds
                                ) ab
                                LEFT JOIN EIP.CsCustomer ac ON aa.CsCustId = ac.CsCustId
                                WHERE 1=1
                                AND (aa.MemberEmail LIKE '%' + @SearchKey + '%' OR aa.MemberName LIKE '%' + @SearchKey + '%')
                                AND ISNULL(ab.RoleIds, '-1') = ISNULL(b.RoleIds, '-1')
                                ORDER BY aa.CsCustId
                                FOR JSON PATH, ROOT('data')
                            ) UserData
                            FROM EIP.[Member] a
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ',' + CONVERT(nvarchar(50), bb.RoleId)
                                    FROM EIP.UserRole ba
                                    INNER JOIN EIP.Role bb ON ba.RoleId = bb.RoleId
                                    WHERE ba.MemberId = a.MemberId
                                    ORDER BY ba.RoleId
                                    FOR XML PATH('')
                                ), 1, 1, '') RoleIds
                                , STUFF((
                                    SELECT ',' + CONVERT(nvarchar(50), bb.RoleName)
                                    FROM EIP.UserRole ba
                                    INNER JOIN EIP.Role bb ON ba.RoleId = bb.RoleId
                                    WHERE ba.MemberId = a.MemberId
                                    ORDER BY ba.RoleId
                                    FOR XML PATH('')
                                ), 1, 1, '') RoleNames
                            ) b
                            LEFT JOIN EIP.CsCustomer c ON a.CsCustId = c.CsCustId
                            WHERE 1=1
                            AND (a.MemberEmail LIKE '%' + @SearchKey + '%' OR a.MemberName LIKE '%' + @SearchKey + '%')
                            ORDER BY b.RoleIds";
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

        #region //GetMemberCustCorp -- 取得會員客戶公司資料 -- Chia Yuan
        public string GetMemberCustCorp(string CompanyIds, string CsCustIds, string MemberEmail, string MemberName, string Status, string CsCustStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (!Regex.IsMatch(CsCustStatus, "^(A|Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【綁定狀態】選擇錯誤!");


                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CsCustId, a.MemberName, a.MemberEmail, a.Gender, ISNULL(a.MemberIcon, -1) AS MemberIcon
                        , ISNULL(a.OrgShortName, '') AS OrgShortName
                        , CASE a.[Status] 
                            WHEN 'S' THEN a.MemberName + ' ' + a.MemberEmail + '(已停用)'
                            ELSE a.MemberName + ' ' + a.MemberEmail
                        END EmailWithMember
                        , ISNULL(b.CustomerName, '') AS CustomerName
                        , ISNULL(b.CustomerEnglishName, '') AS CustomerEnglishName
                        , ISNULL(b.CustomerName + ISNULL(' ' + b.CustomerEnglishName, ''), '未綁定') AS CNameAndEName
                        , a.Customers
                        , c.CustomerIds";
                    sqlQuery.mainTables =
                        @"FROM (
	                        SELECT a.MemberId
	                        , a.CsCustId, a.MemberEmail, a.MemberName, a.Gender, a.MemberIcon, a.OrgShortName
	                        , a.[Status]
	                        , (
		                        SELECT DISTINCT ba.CustomerId 
		                        FROM EIP.CsCustomerDetail ba 
		                        WHERE ba.CsCustId = a.CsCustId 
		                        ORDER BY ba.CustomerId
		                        FOR JSON PATH, ROOT('data')
	                        ) AS Customers
	                        FROM EIP.[Member] a
                        ) a
                        LEFT JOIN EIP.CsCustomer b ON b.CsCustId = a.CsCustId
                        OUTER APPLY (
	                        SELECT ISNULL(SUBSTRING(ba.CustomerIds, 0, LEN(ba.CustomerIds)), '') AS CustomerIds 
	                        FROM (
		                        SELECT (
			                        SELECT CONVERT(VARCHAR(255), x.CustomerId) + ','
			                        FROM OPENJSON(a.Customers, '$.data')
			                        WITH (
				                        CustomerId INT N'$.CustomerId'
			                        ) x
			                        FOR XML PATH('')
		                        ) AS CustomerIds
	                        ) ba
                        ) c";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    switch (CsCustStatus)
                    {
                        case "Y":
                            queryCondition += " AND b.CsCustId IS NOT NULL";
                            break;
                        case "N":
                            queryCondition += " AND b.CsCustId IS NULL";
                            break;
                        default:
                            break;
                    }
                    if (CsCustIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CsCustIds", @" AND a.CsCustId IN @CsCustIds", CsCustIds.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberEmail", @" AND a.MemberEmail LIKE '%' + @MemberEmail + '%'", MemberEmail);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND a.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MemberId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var resultMemberCustCrop = BaseHelper.SqlQuery<MemberCustCorpVM>(officialConnection, dynamicParameters, sqlQuery).ToList();

                    var customerArray = resultMemberCustCrop.Where(w => !string.IsNullOrWhiteSpace(w.CustomerIds)).Select(s1 => s1.CustomerIds).ToList();

                    List<CustWithCorpVM> resultCorp = new List<CustWithCorpVM>();
                    List<int> companyIds = new List<int>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得公司資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT a.CompanyId, a.CompanyNo, a.CompanyName
                                , ISNULL(a.LogoIcon, -1) AS LogoIcon
                                , a.SystemStatus, a.Status
                                , b.CustomerId
                                FROM BAS.Company a
                                INNER JOIN SCM.Customer b ON b.CompanyId = a.CompanyId
                                WHERE a.Status = @Status";
                        dynamicParameters.Add("Status", "A");

                        if (customerArray.Count > 0)
                        {
                            var customerIds = customerArray.SelectMany(s1 => s1.Split(',').Select(s2 => { return Convert.ToInt32(s2); })).Distinct().ToList();
                            if (customerIds.Count > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", @" AND b.CustomerId IN @CustomerIds", customerIds);
                        }
                        if (CompanyIds.Length > 0)
                        {
                            companyIds = CompanyIds.Split(',').Select(s => { return Convert.ToInt32(s); }).Distinct().ToList();
                            if (companyIds.Count > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyIds", @" AND a.CompanyId IN @CompanyIds", companyIds);
                        }
                        resultCorp = sqlConnection.Query<CustWithCorpVM>(sql, dynamicParameters).ToList();
                        //if (!resultCompany.Any()) throw new SystemException("無法取得公司列表!");
                        #endregion
                    }

                    #region //資料集合處理
                    var result = (from s2 in resultMemberCustCrop
                                    .Select(ss1 => new {
                                        ss1,
                                        CustomerIds = ss1.CustomerIds.Split(',').Where(ww1 => !string.IsNullOrWhiteSpace(ww1)).Select(ss2 => { return Convert.ToInt32(ss2); }).ToList() //取得客戶到照表的CustomerId
                                    })
                                  select new
                                  {
                                      Companys = resultCorp.Where(w => s2.CustomerIds.Contains(w.CustomerId)).ToList(),
                                      s2.CustomerIds,
                                      s2.ss1.MemberId,
                                      s2.ss1.CsCustId,
                                      s2.ss1.MemberEmail,
                                      s2.ss1.MemberName,
                                      s2.ss1.Gender,
                                      s2.ss1.OrgShortName,
                                      s2.ss1.MemberIcon,
                                      s2.ss1.EmailWithMember,
                                      s2.ss1.CustomerEnglishName,
                                      s2.ss1.CustomerName,
                                      s2.ss1.CNameAndEName,
                                      s2.ss1.Customers,
                                      s2.ss1.TotalCount
                                  });
                    if (companyIds.Any()) result = result.Where(w => w.Companys.Select(s => s.CompanyId).Intersect(companyIds).Any());
                    result = result
                        .OrderBy(o => o.CsCustId)
                        .ThenBy(t => t.MemberName)
                        .ToList();
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

        #region //GetCorpCust -- 取得公司客戶資料 -- Chia Yuan
        public string GetCorpCust(int CompanyId, int CsCustId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                var resultCorp = new List<dynamic>();
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得公司資料 (停用)
                    //dynamicParameters = new DynamicParameters();
                    //sqlQuery.mainKey = "a.CompanyId, b.CustomerId";
                    //sqlQuery.auxKey = "";
                    //sqlQuery.columns =
                    //    @", a.CompanyNo, a.CompanyName
                    //        , a.CompanyNo + ' ' + a.CompanyName CompanyWithNo
                    //        , ISNULL(a.LogoIcon, -1) AS LogoIcon
                    //        , a.SystemStatus, a.[Status]
                    //        , b.CustomerName
                    //        , ISNULL(b.CustomerEnglishName, '') AS CustomerEnglishName
                    //        , ISNULL(b.CustomerName + ISNULL(' ' + b.CustomerEnglishName, ''), '未綁定') AS CNameAndEName
                    //        --, (
                    //     --    SELECT DISTINCT aa.CompanyId, aa.CustomerId
                    //        --    , aa.CustomerName
                    //        --    , ISNULL(aa.CustomerEnglishName, '') AS CustomerEnglishName
                    //        --    , ISNULL(aa.CustomerName + ISNULL(' ' + aa.CustomerEnglishName, ''), '未綁定') AS CNameAndEName
                    //        --    , aa.CustomerNo
                    //     --    FROM SCM.Customer aa
                    //     --    WHERE aa.CustomerId IN @CustomerIds
                    //     --    AND aa.CompanyId = a.CompanyId
                    //     --    ORDER BY aa.CompanyId
                    //     --    FOR JSON PATH, ROOT('data')
                    //        --) AS Customers";
                    //sqlQuery.mainTables =
                    //    @"FROM BAS.Company a
                    //    INNER JOIN SCM.Customer b ON b.CompanyId = a.CompanyId";
                    //sqlQuery.auxTables = "";
                    //queryCondition = " AND a.[Status] = @Status";
                    //dynamicParameters.Add("Status", "A");
                    //dynamicParameters.Add("CustomerIds", customerIds);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    //sqlQuery.conditions = queryCondition;
                    //sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CompanyId";
                    //sqlQuery.pageIndex = PageIndex;
                    //sqlQuery.pageSize = PageSize;
                    //resultCorp = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
                    #endregion

                    #region //取得公司資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.CompanyId, b.CustomerId
                            , a.CompanyNo, a.CompanyName
                            , a.CompanyNo + ' ' + a.CompanyName CompanyWithNo
                            , ISNULL(a.LogoIcon, -1) AS LogoIcon
                            , a.SystemStatus, a.[Status]
                            , b.CustomerName
                            , ISNULL(b.CustomerEnglishName, '') AS CustomerEnglishName
                            , ISNULL(b.CustomerName + ISNULL(' ' + b.CustomerEnglishName, ''), '未綁定') AS CNameAndEName
                            FROM BAS.Company a
                            INNER JOIN SCM.Customer b ON b.CompanyId = a.CompanyId
                            WHERE b.Status = @Status";
                    dynamicParameters.Add("Status", "A");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    resultCorp = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion
                }

                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    var bmCustomerIds = resultCorp.Select(s => { return Convert.ToInt32(s.CustomerId); }).Distinct();

                    #region //取得客戶資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT CustomerId
                            FROM EIP.CsCustomerDetail 
                            WHERE CustomerId IN @CustomerIds";
                    dynamicParameters.Add("CustomerIds", bmCustomerIds);
                    var customerIds = officialConnection.Query(sql, dynamicParameters).Select(s => s.CustomerId);
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CsCustId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", CASE a.[Status]
                            WHEN 'S' THEN a.CustomerName + '(已停用)'
                            ELSE a.CustomerName
                        END StatusWithName
                        , a.CustomerName
                        , ISNULL(a.CustomerEnglishName, '') AS CustomerEnglishName
                        , ISNULL(a.CustomerName + ISNULL(' ' + a.CustomerEnglishName, ''), '未綁定') AS CNameAndEName
                        , a.Customers
                        , c.CustomerIds
                        , a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM (
	                        SELECT a.CsCustId
	                        , a.CustomerName
                            , a.CustomerEnglishName
	                        , a.[Status]
	                        , (
		                        SELECT DISTINCT ba.CustomerId, ba.CsCustId
		                        FROM EIP.CsCustomerDetail ba 
		                        WHERE ba.CsCustId = a.CsCustId 
		                        ORDER BY ba.CustomerId
		                        FOR JSON PATH, ROOT('data')
	                        ) Customers
	                        FROM EIP.CsCustomer a
                        ) a
                        OUTER APPLY (
	                        SELECT ISNULL(SUBSTRING(ba.CustomerIds, 0, LEN(ba.CustomerIds)), '') AS CustomerIds 
	                        FROM (
		                        SELECT (
			                        SELECT CONVERT(VARCHAR(255), x.CustomerId) + ','
			                        FROM OPENJSON(a.Customers, '$.data')
			                        WITH (
				                        CustomerId INT N'$.CustomerId'
			                        ) x
			                        FOR XML PATH('')
		                        ) AS CustomerIds
	                        ) ba
                        ) c";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    if (CompanyId > 0 && customerIds.Any())
                    {
                        queryCondition += @" AND EXISTS (
                                                SELECT TOP 1 1
                                                FROM EIP.CsCustomerDetail da
                                                WHERE da.CsCustId = a.CsCustId 
                                                AND da.CustomerId IN @CustomerIds
                                            )";
                        dynamicParameters.Add("CustomerIds", customerIds);
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CsCustId", @" AND a.CsCustId = @CsCustId", CsCustId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CsCustId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var resultCust = BaseHelper.SqlQuery<CorpCustVM>(officialConnection, dynamicParameters, sqlQuery).ToList();

                    #region //資料集合處理
                    var result = from s2 in resultCust.Select(ss1 => new
                                 {
                                     ss1,
                                     CustomerIds = ss1.CustomerIds.Split(',').Where(ww1 => !string.IsNullOrWhiteSpace(ww1)).Select(s2 => { return Convert.ToInt32(s2); }).ToList(), //取得客戶到照表的CustomerId
                                 })
                                 select new
                                 {
                                     Companys = resultCorp.Where(w => s2.CustomerIds.Contains(w.CustomerId)).ToList(),
                                     s2.CustomerIds,
                                     s2.ss1.CsCustId,
                                     s2.ss1.StatusWithName,
                                     s2.ss1.CustomerEnglishName,
                                     s2.ss1.CustomerName,
                                     s2.ss1.CNameAndEName,
                                     s2.ss1.Customers,
                                     s2.ss1.Status,
                                     s2.ss1.TotalCount
                                 };
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

        #region //GetAuthorityVerify -- 權限驗證 -- Chia Yuan
        public string GetAuthorityVerify(int MemberId, int[] CustomerIds, string FunctionCode, string FnDetailCode)
        {
            try
            {
                using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SortNumber, a.FnDetailCode, c.Authority
                            FROM EIP.FunctionDetail a
                            INNER JOIN EIP.[Function] b ON a.FunctionId = b.FunctionId
                            OUTER APPLY (
                                SELECT ISNULL((
                                    SELECT TOP 1 1
                                    FROM EIP.RoleFunctionDetail ca
                                    WHERE ca.FnDetailId = a.FnDetailId
                                    AND ca.RoleId IN (
                                        SELECT caa.RoleId
                                        FROM EIP.UserRole caa
                                        INNER JOIN EIP.[Role] cab ON caa.RoleId = cab.RoleId
                                        WHERE caa.MemberId = @MemberId
                                    )
                                ), 0) Authority
                            ) c
                            WHERE a.[Status] = @Status
                            AND b.[Status] = @Status
                            AND b.FunctionCode = @FunctionCode
                            AND c.Authority > 0";
                    dynamicParameters.Add("MemberId", MemberId);
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("FunctionCode", FunctionCode);

                    if (FnDetailCode.Length > 0)
                    {
                        var queryCondition = new StringBuilder(sql);
                        var detailCodes = FnDetailCode.Split(',').ToList();

                        int i = 0;
                        foreach (var code in detailCodes)
                        {
                            queryCondition.Append(i == 0 ? " AND ( " : " OR ").Append("a.FnDetailCode LIKE '%' + @code").Append(i).Append(" + '%'");
                            dynamicParameters.Add("code" + i, code);
                            i++;

                            queryCondition.Append(i == detailCodes.Count ? " ) " : "");
                        }

                        sql = queryCondition.ToString();
                    }

                    sql += " ORDER BY a.SortNumber";
                    var result = officialConnection.Query(sql, dynamicParameters);

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
        #region //AddModule -- 模組別資料新增 -- Chia Yuan
        public string AddModule(int SystemId, string ModuleCode, string ModuleName, string ThemeIcon)
        {
            try
            {
                if (SystemId < 0) throw new SystemException("【所屬系統】不能為空!");
                if (ModuleCode.Length <= 0) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleCode.Length > 50) throw new SystemException("【模組代碼】長度錯誤!");
                if (string.IsNullOrWhiteSpace(ModuleCode)) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 100) throw new SystemException("【模組名稱】長度錯誤!");
                if (string.IsNullOrWhiteSpace(ModuleName)) throw new SystemException("【模組名稱】不能為空!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                ModuleCode = ModuleCode.Trim();
                ModuleName = ModuleName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE SystemId = @SystemId
                                AND ModuleCode = @ModuleCode";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleCode", ModuleCode);

                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【模組代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE SystemId = @SystemId
                                AND ModuleName = @ModuleName";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleName", ModuleName);

                        result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【模組名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM EIP.Module
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        int maxSort = officialConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.Module (SystemId, ModuleCode, ModuleName, ThemeIcon
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ModuleId
                                VALUES (@SystemId, @ModuleCode, @ModuleName, @ThemeIcon
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemId,
                                ModuleCode,
                                ModuleName,
                                ThemeIcon = ThemeIcon.Length > 0 ? ThemeIcon : "",
                                SortNumber = maxSort + 1,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = officialConnection.Query(sql, dynamicParameters);

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

        #region //AddFunction -- 功能別資料新增 -- Chia Yuan
        public string AddFunction(int ModuleId, string FunctionCode, string FunctionName, string UrlTarget)
        {
            try
            {
                if (ModuleId < 0) throw new SystemException("【所屬模組】不能為空!");
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 50) throw new SystemException("【功能代碼】長度錯誤!");
                if (string.IsNullOrWhiteSpace(FunctionCode)) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionName.Length <= 0) throw new SystemException("【功能名稱】不能為空!");
                if (FunctionName.Length > 100) throw new SystemException("【功能名稱】長度錯誤!");
                if (string.IsNullOrWhiteSpace(FunctionName)) throw new SystemException("【功能名稱】不能為空!");

                FunctionCode = FunctionCode.Trim();
                FunctionName = FunctionName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionCode = @FunctionCode";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionName = @FunctionName";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionName", FunctionName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【功能名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM EIP.[Function]
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.[Function] (ModuleId, FunctionCode, FunctionName
                                , UrlTarget, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FunctionId
                                VALUES (@ModuleId, @FunctionCode, @FunctionName
                                , @UrlTarget, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModuleId,
                                FunctionCode,
                                FunctionName,
                                UrlTarget,
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

        #region //AddFunctionDetail -- 詳細功能別資料新增 -- Chia Yuan
        public string AddFunctionDetail(int FunctionId, string FnDetailCode, string FnDetailName)
        {
            try
            {
                if (FunctionId < 0) throw new SystemException("【所屬功能】不能為空!");
                if (FnDetailCode.Length <= 0) throw new SystemException("【功能詳細代碼】不能為空!");
                if (FnDetailCode.Length > 50) throw new SystemException("【功能詳細代碼】長度錯誤!");
                if (FnDetailName.Length <= 0) throw new SystemException("【功能詳細名稱】不能為空!");
                if (FnDetailName.Length > 100) throw new SystemException("【功能詳細名稱】長度錯誤!");

                FnDetailCode = FnDetailCode.Trim();
                FnDetailName = FnDetailName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        #region //判斷功能詳細代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND FnDetailCode = @FnDetailCode";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("FnDetailCode", FnDetailCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能詳細代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能詳細名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND FnDetailName = @FnDetailName";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("FnDetailName", FnDetailName);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【功能詳細名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM EIP.FunctionDetail
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.FunctionDetail (FunctionId, FnDetailCode, FnDetailName
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FnDetailId
                                VALUES (@FunctionId, @FnDetailCode, @FnDetailName
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FunctionId,
                                FnDetailCode,
                                FnDetailName,
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

        #region //AddRole -- 角色資料新增 -- Chia Yuan
        public string AddRole(string RoleName, string AdminStatus)
        {
            try
            {
                if (RoleName.Length <= 0) throw new SystemException("【角色名稱】不能為空!");
                if (RoleName.Length > 50) throw new SystemException("【角色名稱】長度錯誤!");

                RoleName = RoleName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
                                WHERE RoleName = @RoleName";
                        dynamicParameters.Add("RoleName", RoleName);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.Role (RoleName, AdminStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoleId
                                VALUES (@RoleName, @AdminStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
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

        #region //AddCsCustomerDetail 會員客戶明細資料新增 -- Chia Yuan
        public string AddCsCustomerDetail(int CompanyId, string CustomerName, string CustomerEnglishName, string Customers)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CustomerName)) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");

                if (!string.IsNullOrWhiteSpace(CustomerEnglishName))
                {
                    if (CustomerName.Length > 80) throw new SystemException("【客戶名稱】長度錯誤!");
                    CustomerEnglishName = CustomerEnglishName.Trim();
                }
                else
                {
                    CustomerEnglishName = null;
                }

                CustomerName = CustomerName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //會員客戶資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.CsCustomer (CustomerName, CustomerEnglishName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CsCustId
                                VALUES (@CustomerName, @CustomerEnglishName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerName,
                                CustomerEnglishName,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = officialConnection.Query(sql, dynamicParameters);
                        int csCustId = -1;
                        foreach (var item in insertResult)
                        {
                            csCustId = item.CsCustId;
                        }
                        rowsAffected += insertResult.Count();
                        #endregion

                        if (csCustId < 0) throw new SystemException("【會員客戶】資料錯誤!");

                        List<int> customerIds = Customers.Split(',').Select(s => { return Convert.ToInt32(s); }).ToList();

                        List<dynamic> resultCustomer = new List<dynamic>();
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //取得客戶資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DISTINCT CompanyId, CustomerId
                                    FROM SCM.Customer
                                    WHERE CustomerId IN @CustomerIds
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("CustomerIds", customerIds);
                            dynamicParameters.Add("CompanyId", CompanyId);
                            resultCustomer = sqlConnection.Query(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        resultCustomer.ForEach(customer =>
                        {
                            #region //客戶會員明細資料新增
                            dynamicParameters = new DynamicParameters();

                            sql = @"INSERT INTO EIP.CsCustomerDetail (CsCustId, CustomerId, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CsCustId, @CustomerId, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CsCustId = csCustId,
                                    customer.CustomerId,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += officialConnection.Execute(sql, dynamicParameters);
                            #endregion
                        });

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
        #region //UpdateModule -- 模組別資料更新 -- Chai Yuan
        public string UpdateModule(int ModuleId, int SystemId, string ModuleCode, string ModuleName, string ThemeIcon)
        {
            try
            {
                if (SystemId < 0) throw new SystemException("【所屬系統】不能為空!");
                if (ModuleCode.Length <= 0) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleCode.Length > 50) throw new SystemException("【模組代碼】長度錯誤!");
                if (string.IsNullOrWhiteSpace(ModuleCode)) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 100) throw new SystemException("【模組名稱】長度錯誤!");
                if (string.IsNullOrWhiteSpace(ModuleName)) throw new SystemException("【模組名稱】不能為空!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                ModuleCode = ModuleCode.Trim();
                ModuleName = ModuleName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷模組代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE SystemId = @SystemId
                                AND ModuleCode = @ModuleCode
                                AND ModuleId != @ModuleId";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleCode", ModuleCode);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【模組代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE SystemId = @SystemId
                                AND ModuleName = @ModuleName
                                AND ModuleId != @ModuleId";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleName", ModuleName);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【模組名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM EIP.Module
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        int maxSort = officialConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.Module SET
                                SystemId = @SystemId,
                                ModuleCode = @ModuleCode,
                                ModuleName = @ModuleName,
                                ThemeIcon = @ThemeIcon,
                                SortNumber = @SortNumber,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemId,
                                ModuleCode,
                                ModuleName,
                                ThemeIcon,
                                SortNumber = maxSort + 1,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModuleId
                            });

                        int rowsAffected = officialConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateModuleStatus -- 模組別狀態更新 -- Chia Yuan
        public string UpdateModuleStatus(int ModuleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");

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
                        sql = @"UPDATE EIP.Module SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModuleId
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

        #region //UpdateModuleSort -- 模組別順序調整 -- Chia Yuan
        public string UpdateModuleSort(int SystemId, string ModuleList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE EIP.Module SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("SystemId", SystemId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] moduleSort = ModuleList.Split(',');

                        for (int i = 0; i < moduleSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EIP.Module SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ModuleId = @ModuleId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ModuleId = Convert.ToInt32(moduleSort[i])
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

        #region //UpdateFunction -- 功能別資料更新 -- Chia Yuan
        public string UpdateFunction(int FunctionId, int ModuleId, string FunctionCode, string FunctionName, string UrlTarget)
        {
            try
            {
                if (ModuleId < 0) throw new SystemException("【所屬模組】不能為空!");
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 50) throw new SystemException("【功能代碼】長度錯誤!");
                if (string.IsNullOrWhiteSpace(FunctionCode)) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionName.Length <= 0) throw new SystemException("【功能名稱】不能為空!");
                if (FunctionName.Length > 100) throw new SystemException("【功能名稱】長度錯誤!");
                if (string.IsNullOrWhiteSpace(FunctionName)) throw new SystemException("【功能名稱】不能為空!");

                FunctionCode = FunctionCode.Trim();
                FunctionName = FunctionName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionCode = @FunctionCode
                                AND FunctionId != @FunctionId";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionName = @FunctionName
                                AND FunctionId != @FunctionId";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionName", FunctionName);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【功能名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Function] SET
                                ModuleId = @ModuleId,
                                FunctionCode = @FunctionCode,
                                FunctionName = @FunctionName,
                                UrlTarget = @UrlTarget,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModuleId,
                                FunctionCode,
                                FunctionName,
                                UrlTarget,
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

        #region //UpdateFunctionStatus -- 功能別狀態更新 -- Chia Yuan
        public string UpdateFunctionStatus(int FunctionId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");

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
                        sql = @"UPDATE EIP.[Function] SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
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

        #region //UpdateFunctionSort -- 功能別順序調整 -- Chia Yuan
        public string UpdateFunctionSort(int ModuleId, string FunctionList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Function] SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] functionSort = FunctionList.Split(',');

                        for (int i = 0; i < functionSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EIP.[Function] SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE FunctionId = @FunctionId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    FunctionId = Convert.ToInt32(functionSort[i])
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

        #region //UpdateFunctionDetail -- 詳細功能別資料更新 -- Chia Yuan
        public string UpdateFunctionDetail(int FnDetailId, int FunctionId, string FnDetailCode, string FnDetailName)
        {
            try
            {
                if (FunctionId < 0) throw new SystemException("【所屬功能】不能為空!");
                if (FnDetailCode.Length <= 0) throw new SystemException("【詳細功能代碼】不能為空!");
                if (FnDetailCode.Length > 50) throw new SystemException("【詳細功能代碼】長度錯誤!");
                if (FnDetailName.Length <= 0) throw new SystemException("【詳細功能名稱】不能為空!");
                if (FnDetailName.Length > 100) throw new SystemException("【詳細功能名稱】長度錯誤!");

                FnDetailCode = FnDetailCode.Trim();
                FnDetailName = FnDetailName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        #region //判斷詳細功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND FnDetailCode = @FnDetailCode
                                AND FnDetailId != @FnDetailId";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("FnDetailCode", FnDetailCode);
                        dynamicParameters.Add("FnDetailId", FnDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【詳細功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷詳細功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND FnDetailName = @FnDetailName
                                AND FnDetailId != @FnDetailId";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("FnDetailName", FnDetailName);
                        dynamicParameters.Add("FnDetailId", FnDetailId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【詳細功能名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.FunctionDetail SET
                                FunctionId = @FunctionId,
                                FnDetailCode = @FnDetailCode,
                                FnDetailName = @FnDetailName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FnDetailId = @FnDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FunctionId,
                                FnDetailCode,
                                FnDetailName,
                                LastModifiedDate,
                                LastModifiedBy,
                                FnDetailId
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

        #region //UpdateFunctionDetailStatus -- 詳細功能別狀態更新 -- Chia Yuan
        public string UpdateFunctionDetailStatus(int FnDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷詳細功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.FunctionDetail
                                WHERE FnDetailId = @FnDetailId";
                        dynamicParameters.Add("FnDetailId", FnDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【詳細功能別】資料錯誤!");

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
                        sql = @"UPDATE EIP.FunctionDetail SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FnDetailId = @FnDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                FnDetailId
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

        #region //UpdateFunctionDetailSort -- 詳細功能別順序調整 -- Chia Yuan
        public string UpdateFunctionDetailSort(int FunctionId, string FunctionDetailList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE EIP.FunctionDetail SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] functionDetailSort = FunctionDetailList.Split(',');

                        for (int i = 0; i < functionDetailSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EIP.FunctionDetail SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE FnDetailId = @FnDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    FnDetailId = Convert.ToInt32(functionDetailSort[i])
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

        #region //UpdateRole -- 角色資料更新 -- Chia Yuan
        public string UpdateRole(int RoleId, string RoleName, string AdminStatus)
        {
            try
            {
                if (RoleName.Length <= 0) throw new SystemException("【角色名稱】不能為空!");
                if (RoleName.Length > 50) throw new SystemException("【角色名稱】長度錯誤!");

                RoleName = RoleName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        #region //判斷角色名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
                                WHERE RoleName = @RoleName
                                AND RoleId != @RoleId";
                        dynamicParameters.Add("RoleName", RoleName);
                        dynamicParameters.Add("RoleId", RoleId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【角色名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.Role SET
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

        #region //UpdateRoleStatus -- 角色狀態更新 -- Chia Yuan
        public string UpdateRoleStatus(int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.Role
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
                        sql = @"UPDATE EIP.Role SET
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

        #region //UpdateRoleUser -- 角色人員更新 -- Chia Yuan
        public string UpdateRoleUser(int RoleId, string UserList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        string[] userArray = UserList.Split(',');

                        int rowsAffected = 0;
                        for (int i = 0; i < userArray.Length; i++)
                        {
                            int memberId = Convert.ToInt32(userArray[i]);

                            #region //判斷使用者資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MemberName
                                    FROM EIP.[Member]
                                    WHERE MemberId = @MemberId";
                            dynamicParameters.Add("MemberId", memberId);
                            string memberName = string.Empty;
                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("【會員】資料錯誤!");
                            foreach (var item in result2)
                            {
                                memberName = item.MemberName;
                            }
                            #endregion

                            #region //判斷角色人員是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EIP.UserRole
                                    WHERE MemberId = @MemberId
                                    AND RoleId = @RoleId";
                            dynamicParameters.Add("MemberId", memberId);
                            dynamicParameters.Add("RoleId", RoleId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) throw new SystemException(string.Format("【會員:{0}】已有該角色", memberName));
                            #endregion

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EIP.UserRole (RoleId, MemberId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@RoleId, @MemberId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoleId,
                                    MemberId = memberId,
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

        #region //UpdateRoleFunctionDetail -- 角色權限更新 -- Chia Yuan
        public string UpdateRoleFunctionDetail(int RoleId, int FnDetailId, bool Checked, string SearchKey)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
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
                            sql = @"INSERT INTO EIP.RoleFunctionDetail (RoleId, FnDetailId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@RoleId, @FnDetailId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoleId,
                                    FnDetailId,
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
                            sql = @"DELETE EIP.RoleFunctionDetail
                                    WHERE RoleId = @RoleId
                                    AND FnDetailId = @FnDetailId";
                            dynamicParameters.Add("RoleId", RoleId);
                            dynamicParameters.Add("FnDetailId", FnDetailId);
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //目前權限數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModuleId, a.SystemId, b.TotalModuleFunction, c.RoleTotalModuleFunction, d.TotalFunction, e.RoleTotalFunction
                                FROM EIP.[Module] a
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalModuleFunction
                                    FROM EIP.FunctionDetail ba
                                    INNER JOIN EIP.[Function] bb ON ba.FunctionId = bb.FunctionId
                                    WHERE bb.ModuleId = a.ModuleId
                                    AND bb.FunctionName LIKE '%' + @SearchKey + '%'
                                ) b
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalModuleFunction
                                    FROM EIP.RoleFunctionDetail ca
                                    INNER JOIN EIP.FunctionDetail cb ON ca.FnDetailId = cb.FnDetailId
                                    INNER JOIN EIP.[Function] cc ON cb.FunctionId = cc.FunctionId
                                    WHERE cc.ModuleId = a.ModuleId
                                    AND cc.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ca.RoleId = @RoleId
                                ) c
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalFunction
                                    FROM EIP.FunctionDetail da
                                    INNER JOIN EIP.[Function] db ON da.FunctionId = db.FunctionId
                                    INNER JOIN EIP.[Module] dc ON db.ModuleId = dc.ModuleId
                                    WHERE dc.SystemId = a.SystemId
                                    AND db.FunctionName LIKE '%' + @SearchKey + '%'
                                ) d
                                OUTER APPLY (
                                    SELECT COUNT(1) RoleTotalFunction
                                    FROM EIP.RoleFunctionDetail ea
                                    INNER JOIN EIP.FunctionDetail eb ON ea.FnDetailId = eb.FnDetailId
                                    INNER JOIN EIP.[Function] ec ON eb.FunctionId = ec.FunctionId
                                    INNER JOIN EIP.[Module] ed ON ec.ModuleId = ed.ModuleId
                                    WHERE ed.SystemId = a.SystemId
                                    AND ec.FunctionName LIKE '%' + @SearchKey + '%'
                                    AND ea.RoleId = @RoleId
                                ) e
                                WHERE 1=1
                                AND a.ModuleId IN (
                                    SELECT TOP 1 xb.ModuleId
                                    FROM EIP.FunctionDetail xa
                                    INNER JOIN EIP.[Function] xb ON xa.FunctionId = xb.FunctionId 
                                    WHERE xa.FnDetailId = @FnDetailId
                                )";
                        dynamicParameters.Add("RoleId", RoleId);
                        dynamicParameters.Add("FnDetailId", FnDetailId);
                        dynamicParameters.Add("SearchKey", SearchKey);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = result
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

        #region //UpdateUserFunctionDetail -- 使用者權限更新 -- Chia Yuan
        public string UpdateUserFunctionDetail(int RoleId, string UserList, bool Checked)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Role
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色資料錯誤!");
                        #endregion

                        string[] userArray = UserList.Split(',');

                        int rowsAffected = 0;
                        for (int i = 0; i < userArray.Length; i++)
                        {
                            int memberId = Convert.ToInt32(userArray[i]);

                            #region //判斷使用者資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MemberName
                                    FROM EIP.[Member]
                                    WHERE MemberId = @MemberId";
                            dynamicParameters.Add("MemberId", memberId);
                            string memberName = string.Empty;
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【會員】資料錯誤!");
                            foreach (var item in result)
                            {
                                memberName = item.MemberName;
                            }
                            #endregion

                            #region //判斷角色人員是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EIP.UserRole
                                    WHERE MemberId = @MemberId
                                    AND RoleId = @RoleId";
                            dynamicParameters.Add("MemberId", memberId);
                            dynamicParameters.Add("RoleId", RoleId);
                            var resultUserRole = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            #endregion

                            if (Checked)
                            {
                                if (resultUserRole == null)
                                {
                                    #region //新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO EIP.UserRole (RoleId, MemberId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@RoleId, @MemberId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RoleId,
                                            MemberId = memberId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                if (resultUserRole != null)
                                {
                                    #region //刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE EIP.UserRole
                                            WHERE RoleId = @RoleId
                                            AND MemberId = @MemberId";
                                    dynamicParameters.Add("RoleId", RoleId);
                                    dynamicParameters.Add("MemberId", memberId);
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = result
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

        #region //UpdateMemberCustomer -- 會員客戶資料更新 -- Chia Yuan
        public string UpdateMemberCustomer(string UserList, int CsCustId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CsCustId
                                FROM EIP.CsCustomer
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員客戶】資料錯誤!");
                        #endregion

                        int[] userArray = UserList.Split(',').Select(s => { return Convert.ToInt32(s); }).ToArray();

                        int rowsAffected = 0;

                        #region //更新會員客戶資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET a.CsCustId = @CsCustId
                                FROM EIP.Member a
                                WHERE a.MemberId IN @MemberIds";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CsCustId,
                                MemberIds = userArray
                            });
                        rowsAffected = officialConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateCsCustomerDetail -- 會員客戶資料更新
        public string UpdateCsCustomerDetail(int CompanyId, int CsCustId, string CustomerName, string CustomerEnglishName, string Customers)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CustomerName)) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");

                if (!string.IsNullOrWhiteSpace(CustomerEnglishName))
                {
                    if (CustomerName.Length > 80) throw new SystemException("【客戶名稱】長度錯誤!");
                    CustomerEnglishName = CustomerEnglishName.Trim();
                }
                else
                {
                    CustomerEnglishName = null;
                }

                CustomerName = CustomerName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CsCustId
                                FROM EIP.CsCustomer
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員客戶】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //會員客戶資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.CustomerName = @CustomerName, 
                                a.CustomerEnglishName = @CustomerEnglishName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM EIP.CsCustomer a
                                WHERE a.CsCustId = @CsCustId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CsCustId,
                                CustomerName,
                                CustomerEnglishName,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //取得客戶會員明細資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT CustomerId
                                FROM EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        result = officialConnection.Query(sql, dynamicParameters);
                        #endregion

                        List<dynamic> resultCustomer = new List<dynamic>();
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //取得客戶資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DISTINCT CompanyId, CustomerId
                                    FROM SCM.Customer
                                    WHERE CustomerId IN @CustomerIds";
                            dynamicParameters.Add("CustomerIds", result.Select(s => s.CustomerId));
                            resultCustomer = sqlConnection.Query(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        List<int> customerIds = Customers.Split(',').Select(s => { return Convert.ToInt32(s); }).ToList();

                        var deleteCustomerIds = resultCustomer.Where(w => w.CompanyId == CompanyId && !customerIds.Contains(w.CustomerId)).Select(s => s.CustomerId).ToList();

                        #region //客戶會員明細資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId
                                AND CustomerId IN @CustomerIds";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        dynamicParameters.Add("CustomerIds", deleteCustomerIds);
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
                        #endregion

                        var addCustomerIds = customerIds.Where(w => !resultCustomer.Where(w1 => w1.CompanyId == CompanyId).Select(s => s.CustomerId).ToList().Contains(w)).ToList();

                        addCustomerIds.ForEach(customerId =>
                        {
                            #region //客戶會員明細資料新增
                            dynamicParameters = new DynamicParameters();

                            sql = @"INSERT INTO EIP.CsCustomerDetail (CsCustId, CustomerId, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CsCustId, @CustomerId, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CsCustId,
                                    CustomerId = customerId,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += officialConnection.Execute(sql, dynamicParameters);
                            #endregion
                        });

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

        #region //UpdateCsCustomerStatus -- 會員客戶狀態更新
        public string UpdateCsCustomerStatus(int CsCustId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.CsCustomer
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員客戶】資料錯誤!");

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
                        sql = @"UPDATE EIP.CsCustomer SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                CsCustId
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

        #region //UpdateMemberStatus -- 會員狀態更新
        public string UpdateMemberStatus(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.Member
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員】資料錯誤!");

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
                        sql = @"UPDATE EIP.Member SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
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
        #region //DeleteRoleUser -- 角色使用者刪除 -- Chia Yuan
        public string DeleteRoleUser(int MemberId, int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷角色人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.UserRole
                                WHERE MemberId = @MemberId
                                AND RoleId = @RoleId";
                        dynamicParameters.Add("MemberId", MemberId);
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("角色人員資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.UserRole
                                WHERE MemberId = @MemberId
                                AND RoleId = @RoleId";
                        dynamicParameters.Add("MemberId", MemberId);
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

        #region //DeleteMemberCustomer -- 會員客戶資料刪除 -- Chia Yuan
        public string DeleteMemberCustomer(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Member
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);
                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET a.CsCustId = null 
                                FROM EIP.Member a
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteCsCustomer -- 會員客戶資料刪除 -- Chia Yuan
        public string DeleteCsCustomer(int CsCustId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.CsCustomer
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);

                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員客戶】資料錯誤!");
                        #endregion

                        #region //取得會員客戶明細資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT CsCustId, CustomerId
                                FROM EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);

                        result = officialConnection.Query(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;

                        #region //會員客戶明細資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.CsCustomer
                                WHERE CsCustId = @CsCustId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteCsCustomerDetail -- 會員客戶明細資料刪除 -- Chia Yuan
        public string DeleteCsCustomerDetail(int CsCustId, int CustomerId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        #region //判斷會員客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        dynamicParameters.Add("CustomerId", CustomerId);

                        var result = officialConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【會員客戶明細】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.CsCustomerDetail
                                WHERE CsCustId = @CsCustId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CsCustId", CsCustId);
                        dynamicParameters.Add("CustomerId", CustomerId);
                        rowsAffected += officialConnection.Execute(sql, dynamicParameters);
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
