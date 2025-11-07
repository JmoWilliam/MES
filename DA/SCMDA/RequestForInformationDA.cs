using System;
using System.Web;
using System.Linq;
using System.Transactions;
using System.Configuration;
using System.Data.SqlClient;
using Dapper;
using Helpers;
using NLog;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Collections.Generic;

namespace EIPDA
{
    public class RequestForInformationDA
    {
        public string MainConnectionStrings = string.Empty;
        public string ErpConnectionStrings = string.Empty;

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
        public MamoHelper mamoHelper = new MamoHelper();

        public RequestForInformationDA()
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

                //CurrentUser = 6637;

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
        #region //GetProdTerminal -- 取得終端資料 -- Chia Yuan 2023.12.25
        public string GetProdTerminal(int ProdTerminalId, string TerminalName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProdTerminalId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TerminalName, a.Status, s.StatusName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.Status WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM RFI.ProdTerminal a
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.Status AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProdTerminalId", @" AND a.ProdTerminalId = @ProdTerminalId", ProdTerminalId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TerminalName", @" AND a.TerminalName LIKE '%' + @TerminalName + '%'", TerminalName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TerminalName";
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

        #region //GetProdSystem -- 取得系統資料 -- Chia Yuan 2023.12.25
        public string GetProdSystem(int ProdSysId, string SystemName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProdSysId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SystemName, a.[Status], s.StatusName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.Status WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM RFI.ProdSystem a
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProdSysId", @" AND a.ProdSysId = @ProdSysId", ProdSysId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemName", @" AND a.SystemName LIKE '%' + @SystemName + '%'", SystemName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SystemName";
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

        #region //GetProdModule -- 取得模組資料 -- Chia Yuan 2023.12.25
        public string GetProdModule(int ProdModId, string ModuleName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProdModId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ModuleName, a.[Status], s.StatusName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.Status WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM RFI.ProdModule a
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProdModId", @" AND a.ProdModId = @ProdModId", ProdModId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleName", @" AND a.ModuleName LIKE '%' + @ModuleName + '%'", ModuleName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModuleName";
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

        #region //GetProductUse -- 取得產品應用資料 -- Chia Yuan 2023.12.25
        public string GetProductUse(int ProductUseId, string ProductUseName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProductUseId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProductUseName, a.[Status], s.StatusName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.Status WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , a.TypeOne, t1.TypeName AS ProductUseTypeName
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM SCM.ProductUse a
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = a.TypeOne AND t1.TypeSchema = 'ProductUse.Type'
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseName", @" AND a.ProductUseName LIKE '%' + @ProductUseName + '%'", ProductUseName.Trim());
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProductUseName";
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

        #region //GetTemplateProdSpec -- 取得評估樣板主資料 -- Chia Yuan 2023.12.27
        public string GetTemplateProdSpec(int TempProdSpecId, int ProTypeGroupId, string SpecName, string SpecEName, string ControlType, string Status, string FeatureName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.TempProdSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", LTRIM(RTRIM(ISNULL(a.SpecName, ''))) AS SpecName, LTRIM(RTRIM(ISNULL(a.SpecEName, ''))) AS SpecEName
                        , a.SortNumber, a.ControlType, a.[Required], a.[Status], s.StatusName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.[Status] WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , f.FeatureCount
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate
                        , (
                            SELECT aa.TempProdSpecId, aa.TempPsDetailId
                            , LTRIM(RTRIM(aa.FeatureName)) AS FeatureName
                            , ISNULL(aa.DataType, '') AS DataType
                            , ISNULL(aa.[Description], '') AS [Description]
                            , ISNULL(ac.TypeName,'') AS TypeName
                            , aa.SortNumber
                            , aa.[Status]
                            , ab.StatusName
                            FROM RFI.TemplateProdSpecDetail aa
                            INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND ab.StatusSchema = 'Status'
                            LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.DataType AND ac.TypeSchema = 'RfiFormItem.OtherType'
                            WHERE aa.TempProdSpecId = a.TempProdSpecId
                            ORDER BY aa.SortNumber
                            FOR JSON PATH
                        ) AS TempProdSpecDetail";
                    sqlQuery.mainTables =
                        @"FROM RFI.TemplateProdSpec a
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        OUTER APPLY (
	                        SELECT COUNT(1) FeatureCount
	                        FROM RFI.TemplateProdSpecDetail fa
	                        WHERE fa.TempProdSpecId = a.TempProdSpecId
                        ) f
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TempProdSpecId", @" AND a.TempProdSpecId = @TempProdSpecId", TempProdSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SpecName", @" AND a.SpecName LIKE '%' + @SpecName + '%'", SpecName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SpecEName", @" AND a.SpecEName LIKE '%' + @SpecEName + '%'", SpecEName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FeatureName", @" AND EXISTS (
                                                                                                        SELECT TOP 1 1
                                                                                                        FROM RFI.TemplateProdSpecDetail ba
                                                                                                        WHERE ba.TempProdSpecId = a.TempProdSpecId
                                                                                                        AND ba.FeatureName LIKE '%' + @FeatureName + '%')", FeatureName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (ControlType.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlType", @" AND a.ControlType IN @ControlType", ControlType.Split(','));
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

        #region //GetTemplateSpecParameter -- 取得新設計項目主資料 -- Chia Yuan 2024.01.25
        public string GetTemplateSpecParameter(int TempSpId, int ProTypeGroupId, string ParameterName, string ControlType, string Status, string PmtDetailName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.TempSpId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", LTRIM(RTRIM(ISNULL(a.ParameterName, ''))) AS ParameterName, LTRIM(RTRIM(ISNULL(a.ParameterEName, ''))) AS ParameterEName
                        , a.SortNumber, a.ControlType, a.[Required], a.[Status], s.StatusName
                        , ISNULL(d.TypeName,'') AS TypeName
                        , ISNULL(b.UserName, '') AS UserName
                        , CASE a.[Status] WHEN 'A' THEN 'true' ELSE 'false' END AS Checked
                        , f.FeatureCount
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') AS LastModifiedDate
                        , (
                            SELECT aa.TempSpDetailId, aa.TempSpId
	                        , LTRIM(RTRIM(aa.PmtDetailName)) AS PmtDetailName
                            , ISNULL(aa.DataType, '') AS DataType
                            , ISNULL(aa.[Description], '') AS [Description]
	                        , ISNULL(ac.TypeName,'') AS TypeName
	                        , aa.SortNumber
                            , aa.[Required]
                            , aa.[Status]
                            , ab.StatusName
                            FROM RFI.TemplateSpecParameterDetail aa
                            INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND ab.StatusSchema = 'Status'
	                        LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.DataType AND ac.TypeSchema = 'RfiFormItem.OtherType'
                            WHERE aa.TempSpId = a.TempSpId
                            ORDER BY aa.SortNumber
                            FOR JSON PATH
                        ) AS TempSpecParameterDetail";
                    sqlQuery.mainTables =
                        @"FROM RFI.TemplateSpecParameter a
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        LEFT JOIN BAS.[Type] d ON d.TypeNo = a.ControlType AND d.TypeSchema = 'RfiFormItem.ItemType'
                        OUTER APPLY (
	                        SELECT COUNT(1) FeatureCount
	                        FROM RFI.TemplateSpecParameterDetail fa
	                        WHERE fa.TempSpId = a.TempSpId
                        ) f
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TempSpId", @" AND a.TempSpId = @TempSpId", TempSpId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterName", @" AND a.ParameterName LIKE '%' + @ParameterName + '%'", ParameterName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PmtDetailName", @" AND EXISTS (
                                                                                                        SELECT TOP 1 1
                                                                                                        FROM RFI.TemplateSpecParameterDetail ba
                                                                                                        WHERE ba.TempSpId = a.TempSpId
                                                                                                        AND ba.PmtDetailName LIKE '%' + @PmtDetailName + '%')", PmtDetailName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (ControlType.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlType", @" AND a.ControlType IN @ControlType", ControlType.Split(','));
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

        #region //GetType -- 取得類別資料 -- Chia Yuan 2023.12.27
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
                        @", a.TypeSchema, a.TypeNo
                        , CASE a.TypeNo WHEN 'T' THEN N'📝' WHEN 'R' THEN N'🔘' WHEN 'C' THEN N'✅' WHEN 'D' THEN N'🔻' WHEN 'A' THEN N'📝' END + ' ' + a.TypeName AS TypeName
                        , a.TypeDesc
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

        #region //GetAuthority -- 取得權限資料 -- Chia Yuan 2024.5.3
        public string GetAuthority(string FunctionCode, string DetailCode, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.RoleId, a.DetailId, c.RoleName
	                        , u.UserNo, u.UserName
	                        , s.StatusNo AS RoleStatus, s.StatusName AS RoleStatusName
	                        FROM BAS.RoleFunctionDetail a 
	                        INNER JOIN BAS.[Role] c ON c.RoleId = a.RoleId
	                        INNER JOIN BAS.UserRole d ON d.RoleId = a.RoleId
	                        INNER JOIN BAS.[Status] s ON s.StatusNo = c.[Status] AND s.StatusSchema = 'Status'
	                        INNER JOIN BAS.[User] u ON u.UserId = d.UserId
	                        WHERE EXISTS (
		                        SELECT TOP 1 1
		                        FROM BAS.[Function] ba
		                        INNER JOIN BAS.FunctionDetail bb ON bb.FunctionId = ba.FunctionId
		                        WHERE bb.DetailId = a.DetailId
		                        AND ba.FunctionCode = @FunctionCode
		                        AND bb.DetailCode IN @DetailCodes
	                        )
	                        AND d.UserId = @UserId
	                        AND c.CompanyId = @CompanyId";

                    dynamicParameters.Add("FunctionCode", FunctionCode);
                    dynamicParameters.Add("DetailCodes", DetailCode.Split(','));
                    dynamicParameters.Add("UserId", CurrentUser);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND c.Status IN @Status", Status.Split(','));

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

        #region //GetSignAuthority -- 取得審批權限資料 -- Chia Yuan 2024.6.12
        public string GetDesignSignAuthority(int DesignId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1 
                            FROM RFI.DesignSignFlow a
                            WHERE a.DesignId = @DesignId
                            AND ISNULL(a.SignUser, a.FlowUser) = @Flower";
                    dynamicParameters.Add("DesignId", DesignId);
                    dynamicParameters.Add("Flower", CurrentUser);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0)
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = "無此單據審批權限，不得查看!"
                        });
                    }

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

        #region //GetRequestForInformation -- 取得RFI單頭 -- Chia Yuan 2024.01.04
        public string GetRequestForInformation(int RfiId, string RfiNo, int ProTypeGroupId, int ProductUseId, int SalesId, string StartDate, string EndDate
            , string Status, string RfiDetailStatus, string SignStatus, string FlowStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RfiId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfiNo, a.ProTypeGroupId, a.ProductUseId, b.ProductUseName
                        , a.SalesId, d.UserNo SalesNo, d.UserName SalesName, d.Gender SalesGender
                        , a.OrganizaitonType, i.TypeName OrganizaitonTypeName, a.CustomerId, a.CustomerName, ISNULL(a.CustomerEName, '') CustomerEName
                        , ISNULL(FORMAT(a.AnalysisDate, 'yyyy-MM-dd HH:mm'), '') AnalysisDate
                        , a.[Status] RfiStatus, a.ExistFlag
                        , ISNULL(FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm'), '') CreateDate
                        , ISNULL(c.AssemblyName, '') AssemblyName
                        , e.StatusName RfiStatusName
                        , h.StatusName ExistFlagName
                        , g.CompanyName
                        , j.ProTypeGroupName, j.CoatingFlag, k.StatusName CoatingFlagName
                        , ISNULL(ab.PreSignJobName, '') 'PreSignJobName', ISNULL(ab.NextSignJobName, '') NextSignJobName
                        , flow.*
                        , (
	                        SELECT bb.*
	                        FROM RFI.RfiDetail ba
	                        OUTER APPLY (
		                        SELECT ca.RfiSfId
                                , ISNULL(ca.Approved, '') Approved
                                --, CASE ISNULL(ca.Approved, '') WHEN 'Y' THEN '核准' WHEN 'N' THEN '否決' WHEN 'B' THEN '退回' ELSE '待審批' END ApprovedName
                                , ISNULL(s2.StatusName, '待審批') ApprovedName
                                , ISNULL(ca.SignUser, ca.FlowUser) Flower
                                , ISNULL(cu2.UserNo, cu1.UserNo) FlowerNo
                                , ISNULL(cu2.UserName, cu1.UserName) FlowerName
                                , ISNULL(cu2.Gender, cu1.Gender) FlowerGender
                                , ISNULL(cu2.Email, cu1.Email) FlowerEmail
	                            , ISNULL(ca.SignJobName, '') SignJobName
	                            , ISNULL(ca.SignDesc, '') SignDesc
	                            , ISNULL(FORMAT(ca.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime
	                            , ISNULL(FORMAT(ca.SignDateTime, 'yyyy-MM-dd HH:mm'), '') SignDateTime
	                            , DATEDIFF(MINUTE, ca.ArriveDateTime, ISNULL(ca.SignDateTime, @dtNow)) / 1440 StayDays
	                            , (DATEDIFF(MINUTE, ca.ArriveDateTime, ISNULL(ca.SignDateTime, @dtNow)) % 1440) / 60 StayHours
	                            , (DATEDIFF(MINUTE, ca.ArriveDateTime, ISNULL(ca.SignDateTime, @dtNow)) % 60) StayMinutes
		                        , DENSE_RANK () OVER (PARTITION BY ca.RfiDetailId ORDER BY ca.Edition DESC) SortEdition
	                            , Edition, ca.SortNumber
		                        FROM RFI.RfiSignFlow ca
                                LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = ca.Approved AND s2.StatusSchema = 'RfiSign.Status'
	                            INNER JOIN BAS.[User] cu1 ON cu1.UserId = ca.FlowUser
	                            LEFT JOIN BAS.[User] cu2 ON cu2.UserId = ca.SignUser
		                        WHERE ca.RfiDetailId = ba.RfiDetailId
	                        ) bb
	                        WHERE ba.RfiId = a.RfiId
                            ORDER BY bb.Edition, bb.SortNumber
	                        --AND bb.SortEdition = 1
                            FOR JSON PATH, ROOT('data')
                        ) SignFlow
                        , (
                            SELECT aa.RfiId, aa.RfiDetailId, aa.CompanyId, aa.RfiSequence, aa.Edition
                            , aa.SalesId, ad.UserNo SalesNo, ad.UserName SalesName, ad.Gender SalesGender
                            , FORMAT(aa.ProdDate, 'yyyy-MM-dd') ProdDate
                            , FORMAT(aa.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart
                            , FORMAT(aa.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
	                        , aa.TargetUnitPrice, aa.LifeCycleQty, aa.Revenue, aa.RevenueFC
                            , aa.ExchangeRate, aa.Currency, ac.TypeNo 'CurrencyTypeName', ac.TypeName CurrencyName
	                        , aa.[Status] RfiDetailStatus, ae.StatusName RfiDetailStatusName
                            , ISNULL(FORMAT(aa.ConfirmFinalTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmFinalTime
                            , ISNULL(aa.AdditionalFile, '') AdditionalFile
                            , (
						        SELECT TOP 1 1
						        FROM RFI.ProductSpec ab
						        WHERE ab.RfiDetailId = aa.RfiDetailId
					        ) ProductSpecExists
                            , (
                                SELECT TOP 1 1 
                                FROM RFI.RfiSignFlow af
                                WHERE af.RfiDetailId = aa.RfiDetailId
                            ) SignFlowExists
                            FROM RFI.RfiDetail aa
                            LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.Currency AND ac.TypeSchema = 'ExchangeRate.Currency'
                            INNER JOIN BAS.[User] ad ON ad.UserId = aa.SalesId
                            INNER JOIN BAS.[Status] ae ON ae.StatusNo = aa.[Status] AND ae.StatusSchema = 'RfiDetail.Status'
                            WHERE aa.RfiId = a.RfiId";
                    if (SalesId > 0)
                    {
                        //用於判斷單身多筆不同指派人員，需濾掉其他業務，只能看自己的單據
                        sqlQuery.columns += @" AND EXISTS (
                                                    SELECT TOP 1 1
                                                    FROM RFI.RfiDetail ba
                                                    WHERE ba.RfiId = aa.RfiId
                                                    AND ba.SalesId = @SalesId
                                               )";
                    }
                    sqlQuery.columns += 
                        @" ORDER BY aa.RfiDetailId
                            FOR JSON PATH, ROOT('data')
                        ) RfiDetail";
                    sqlQuery.mainTables =
                        @" FROM RFI.RequestForInformation a
                        INNER JOIN SCM.ProductUse b ON b.ProductUseId = a.ProductUseId
                        INNER JOIN SCM.ProductTypeGroup j ON j.ProTypeGroupId = a.ProTypeGroupId
                        OUTER APPLY (
	                        SELECT TOP 1 ca.AssemblyName
	                        FROM RFI.DemandDesign ca
	                        WHERE ca.RfiId = a.RfiId
	                        ORDER BY ca.Edition DESC
                        ) c
	                    INNER JOIN BAS.[User] d ON d.UserId = a.SalesId
	                    INNER JOIN BAS.[Status] e ON e.StatusNo = a.[Status] AND e.StatusSchema = 'RequestForInformation.Status'
                        INNER JOIN BAS.[Status] h ON h.StatusNo = a.ExistFlag AND h.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Status] k ON k.StatusNo = j.CoatingFlag AND k.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] i ON i.TypeNo = a.OrganizaitonType AND i.TypeSchema = 'MemberOrganization.OrganizaitonType'
	                    INNER JOIN BAS.Department f ON f.DepartmentId = d.DepartmentId
	                    INNER JOIN BAS.Company g ON g.CompanyId = f.CompanyId
                        OUTER APPLY (
	                        SELECT *
	                        FROM (
		                        SELECT ISNULL(aa.SignUser, aa.FlowUser) Flower, aa.RfiSfId, aa.SignJobName, aa.SortNumber, ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime, aa.Approved
		                        --, ROW_NUMBER() OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.SortNumber) SortRow
		                        --虛擬狀態 W=等待審批, I=進行中, D=否決, C=核准(完成)
		                        , CASE ISNULL(aa.ArriveDateTime, '') WHEN '' THEN 'W' ELSE (CASE ISNULL(aa.Approved, '') WHEN '' THEN 'I' WHEN 'N' THEN 'D' WHEN 'B' THEN 'B' ELSE 'C' END) END SignStatus
                                , aa.[Status] FlowStatus, ab.StatusName FlowStatusName
                                , DENSE_RANK () OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.Edition DESC) SortEdition
		                        FROM RFI.RfiSignFlow aa
                                INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND StatusSchema = 'RfiDetail.FlowStatus'
		                        INNER JOIN RFI.RfiDetail ac ON ac.RfiDetailId = aa.RfiDetailId
		                        WHERE ac.RfiId = a.RfiId
	                        ) x WHERE x.SignStatus IN ('I', 'D')
                            AND x.SortEdition = 1
                        ) flow
                        LEFT JOIN (
	                        SELECT ba.RfiSfId
	                        , LAG(ba.SignJobName) OVER (PARTITION BY ba.Edition, ba.RfiDetailId ORDER BY ba.SortNumber) 'PreSignJobName'
	                        , LEAD(ba.SignJobName) OVER (PARTITION BY ba.Edition, ba.RfiDetailId ORDER BY ba.SortNumber) 'NextSignJobName'
	                        FROM RFI.RfiSignFlow ba
                        ) ab ON ab.RfiSfId = flow.RfiSfId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND (EXISTS (
                                            SELECT TOP 1 1
                                            FROM RFI.RfiDetail da
                                            INNER JOIN RFI.RfiSignFlow db ON da.RfiDetailId = db.RfiDetailId
                                            OUTER APPLY (
                                                SELECT DISTINCT ISNULL(ab.SignUser, ab.FlowUser) 'DesignFlower'
                                                FROM RFI.DemandDesign aa
                                                INNER JOIN RFI.DesignSignFlow ab ON ab.DesignId = aa.DesignId
                                                WHERE aa.RfiId = da.RfiId
                                            ) dc
                                            WHERE da.RfiId = a.RfiId
                                            AND (ISNULL(db.SignUser, db.FlowUser) = @Flower OR dc.DesignFlower = @Flower)) OR a.SalesId = @Flower)";
                    dynamicParameters.Add("dtNow", CreateDate);
                    dynamicParameters.Add("Flower", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiId", @" AND a.RfiId = @RfiId", RfiId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiNo", @" AND a.RfiNo = @RfiNo", RfiNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    if (DateTime.TryParse(StartDate, out DateTime startDate))
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", " AND a.AnalysisDate >= @StartDate", startDate.ToString("yyyy-MM-dd 00:00:00"));
                    if (DateTime.TryParse(EndDate, out DateTime endDate))
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", " AND a.AnalysisDate <= @EndDate", endDate.ToString("yyyy-MM-dd 23:59:59"));

                    //只能看自己的單據
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesId", " AND a.SalesId = @SalesId", SalesId);

                    //待審批
                    if (SignStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SignStatus", @" AND flow.SignStatus IN @SignStatus", SignStatus.Split(','));
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Flower", @" AND flow.Flower = @Flower", CurrentUser);
                    }

                    //單據狀態
                    if (Status.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.[Status] IN @Status", Status.Split(','));

                    //關卡狀態
                    if (FlowStatus.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowStatus", @" AND flow.FlowStatus IN @FlowStatus", FlowStatus.Split(','));

                    //單身狀態
                    if (RfiDetailStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DetailStatus", 
                            @" AND EXISTS (
                                SELECT TOP 1 1
                                FROM RFI.RfiDetail ba
                                WHERE ba.RfiId = a.RfiId
                                AND ba.Status IN @DetailStatus)", RfiDetailStatus.Split(','));
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfiId DESC";
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

        #region //GetRfiDetail -- 取得RFI單身 -- Chia Yuan 2024.01.08
        public string GetRfiDetail(int RfiId, int RfiDetailId, int SalesId, string RfiDetailStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RfiDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfiId, a.CompanyId, a.RfiSequence, a.Edition, a.TargetUnitPrice, a.LifeCycleQty, a.Revenue, a.RevenueFC
					    , a.ExchangeRate, a.Currency, f.TypeName CurrencyName, ISNULL(k.AssemblyName, '') AssemblyName
					    , ISNULL(FORMAT(a.ProdDate, 'yyyy-MM-dd'), '') ProdDate
					    , ISNULL(FORMAT(a.ProdLifeCycleStart, 'yyyy-MM-dd'), '') ProdLifeCycleStart
					    , ISNULL(FORMAT(a.ProdLifeCycleEnd, 'yyyy-MM-dd'), '') ProdLifeCycleEnd
					    , a.SalesId, d.UserNo SalesNo, d.UserName SalesName, d.Gender SalesGender
					    , a.[Status] RfiDetailStatus
					    , b.[Status] RfiStatus
					    , e.StatusName RfiDetailStatusName
					    , ISNULL(FORMAT(a.ConfirmFinalTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmFinalTime
					    , flow.*
					    , ISNULL(ab.PreSignJobName, '') 'PreSignJobName', ISNULL(ab.NextSignJobName, '') NextSignJobName
					    , (
						    SELECT aa.RfiSfId
						    , ISNULL(aa.Approved, '') Approved
						    --, CASE ISNULL(aa.Approved, '') WHEN 'Y' THEN '核准' WHEN 'N' THEN '否決' WHEN 'B' THEN '退回' ELSE '待審批' END ApprovedName
                            , ISNULL(s2.StatusName, '待審批') ApprovedName
						    , ISNULL(aa.SignUser, aa.FlowUser) Flower
						    , ISNULL(au2.UserNo, au1.UserNo) FlowerNo
						    , ISNULL(au2.UserName, au1.UserName) FlowerName
						    , ISNULL(au2.Gender, au1.Gender) FlowerGender
						    , ISNULL(au2.Email, au1.Email) FlowerEmail
						    , ISNULL(aa.SignJobName, '') SignJobName
						    , ISNULL(aa.SignDesc, '') SignDesc
                            , s1.StatusDesc FlowStatusDesc
						    , ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime
						    , ISNULL(FORMAT(aa.SignDateTime, 'yyyy-MM-dd HH:mm'), '') SignDateTime
						    , DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) / 1440 StayDays
						    , (DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
						    , (DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) % 60) StayMinutes
                            , DENSE_RANK () OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.Edition DESC) SortEdition
						    , aa.Edition, aa.SortNumber
                            , ISNULL(af.FileId, '') 'FileId', ISNULL(af.[FileName], '') [FileName], ISNULL(af.FileExtension, '') FileExtension
						    FROM RFI.RfiSignFlow aa
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = aa.[Status] AND s1.StatusSchema = 'RfiDetail.FlowStatus'
                            LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = aa.Approved AND s2.StatusSchema = 'RfiSign.Status'
						    INNER JOIN BAS.[User] au1 ON au1.UserId = aa.FlowUser
						    LEFT JOIN BAS.[User] au2 ON au2.UserId = aa.SignUser
                            OUTER APPLY (SELECT FileId, [FileName], FileExtension FROM BAS.[File] WHERE FileId = aa.AdditionalFile AND DeleteStatus = 'N') af
						    WHERE aa.RfiDetailId = a.RfiDetailId
                            ORDER BY aa.Edition, aa.SortNumber
						    FOR JSON PATH, ROOT('data')
					    ) SignFlow
					    , (
						    SELECT TOP 1 1
						    FROM RFI.ProductSpec aa
						    WHERE aa.RfiDetailId = a.RfiDetailId
					    ) ProductSpecExists
                        , (
                            SELECT TOP 1 1 
                            FROM RFI.RfiSignFlow aa
                            WHERE aa.RfiDetailId = a.RfiDetailId
                        ) SignFlowExists";
                    if (RfiDetailId > 0)
                    {
                        sqlQuery.columns += @", ISNULL(j.FileId, '') AdditionalFile, j.[FileName] AdditionalFileName, j.FileExtension AdditionalFileExtension";

                        //單筆查詢才撈出規格相關資料
                        sqlQuery.columns += @", (SELECT aa.ProdSpecId, LTRIM(RTRIM(aa.SpecName)) SpecName, LTRIM(RTRIM(aa.SpecEName)) SpecEName, aa.SortNumber, aa.ControlType, aa.[Required], aa.[Status]
											    , (SELECT aa.ProdSpecId, ba.PsDetailId, LTRIM(RTRIM(ba.FeatureName)) FeatureName, ba.SortNumber
												    , ISNULL(ba.DataType, '') DataType
												    , ISNULL(ba.[Description], '') [Description]
												    , ba.[Status]
												    , bb.StatusName
												    FROM RFI.ProductSpecDetail ba
												    INNER JOIN BAS.[Status] bb ON bb.StatusNo = ba.[Status] AND bb.StatusSchema = 'Status'
												    WHERE ba.ProdSpecId = aa.ProdSpecId
												    ORDER BY ba.SortNumber
												    FOR JSON PATH, ROOT('data')
											    ) ProductSpecDetail
											    FROM RFI.ProductSpec aa
											    WHERE aa.RfiDetailId = a.RfiDetailId
											    ORDER BY aa.SortNumber
											    FOR JSON PATH, ROOT('data')
										    ) ProductSpec";
                    }
                    sqlQuery.mainTables =
                        @"FROM RFI.RfiDetail a
					    INNER JOIN RFI.RequestForInformation b ON b.RfiId = a.RfiId
					    INNER JOIN BAS.[User] d ON d.UserId = a.SalesId
					    INNER JOIN BAS.[Status] e ON e.StatusNo = a.[Status] AND e.StatusSchema = 'RfiDetail.Status'
					    LEFT JOIN BAS.[Type] f ON f.TypeNo = a.Currency AND f.TypeSchema = 'ExchangeRate.Currency'
					    OUTER APPLY (
						    SELECT TOP 1 ka.AssemblyName
						    FROM RFI.DemandDesign ka
						    WHERE ka.RfiId = b.RfiId
						    ORDER BY ka.Edition DESC
					    ) k
					    OUTER APPLY (
						    SELECT *
						    FROM (
							    SELECT ISNULL(aa.SignUser, aa.FlowUser) Flower, aa.RfiSfId, aa.SignJobName, aa.SortNumber, ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime, aa.Approved
							    --, ROW_NUMBER() OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.SortNumber) SortRow
							    --虛擬狀態 W=等待審批, I=進行中, D=否決, C=核准(完成)
							    , CASE ISNULL(aa.ArriveDateTime, '') WHEN '' THEN 'W' ELSE (CASE ISNULL(aa.Approved, '') WHEN '' THEN 'I' WHEN 'N' THEN 'D' WHEN 'B' THEN 'B' ELSE 'C' END) END SignStatus
							    , aa.[Status] FlowStatus, ab.StatusName FlowStatusName
                                , DENSE_RANK () OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.Edition DESC) SortEdition
							    FROM RFI.RfiSignFlow aa
							    INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND StatusSchema = 'RfiDetail.FlowStatus'
							    INNER JOIN RFI.RfiDetail ac ON ac.RfiDetailId = aa.RfiDetailId
							    WHERE ac.RfiId = a.RfiId
						    ) x WHERE x.SignStatus IN ('I', 'D')
                            AND SortEdition = 1
					    ) flow
					    LEFT JOIN (
						    SELECT ba.RfiSfId, ba.RfiDetailId
						    , LAG(ba.SignJobName) OVER (PARTITION BY ba.Edition, ba.RfiDetailId ORDER BY ba.SortNumber) PreSignJobName
						    , LEAD(ba.SignJobName) OVER (PARTITION BY ba.Edition, ba.RfiDetailId ORDER BY ba.SortNumber) NextSignJobName
						    FROM RFI.RfiSignFlow ba
					    ) ab ON ab.RfiSfId = flow.RfiSfId";
                    if (RfiDetailId > 0) sqlQuery.mainTables += " OUTER APPLY (SELECT FileId, [FileName], FileExtension FROM BAS.[File] WHERE FileId = a.AdditionalFile AND DeleteStatus = 'N') j";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CreateDate", CreateDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiId", @" AND a.RfiId = @RfiId", RfiId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiDetailId", @" AND a.RfiDetailId = @RfiDetailId", RfiDetailId);
                    if (RfiDetailStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", RfiDetailStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfiDetailId DESC";
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

        #region //GetRfiSignFlow -- 取得RFI審批流程 -- Chia Yuan 2024.06.25
        public string GetRfiSignFlow(int RfiSfId, int RfiDetailId, string Edition, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RfiSfId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfiDetailId, a.SignJobName, a.Approved, a.Edition, a.SortNumber
                        , e.ProTypeGroupId, e.ProductUseId
                        , u1.UserId, u1.UserName, u1.UserNo, u1.Email
                        , ISNULL(f.FileId, '') AdditionalFile                        
                        , c.CompanyId,c.CompanyName,c.CompanyNo
                        , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
                        , DENSE_RANK () OVER (PARTITION BY a.RfiDetailId ORDER BY a.Edition DESC) SortEdition";
                    //, LEAD(a.SortNumber) OVER (PARTITION BY a.RfiDetailId ORDER BY a.SortNumber) NextSortNumber
                    sqlQuery.mainTables =
                        @"FROM RFI.RfiSignFlow a
                        INNER JOIN RFI.RfiDetail b ON b.RfiDetailId = a.RfiDetailId
                        INNER JOIN RFI.RequestForInformation e ON e.RfiId = b.RfiId
                        INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                        INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                        INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                        OUTER APPLY (SELECT TOP 1 fa.FileId, fa.[FileName], fa.FileExtension FROM BAS.[File] fa WHERE fa.FileId = a.AdditionalFile AND fa.DeleteStatus = 'N') f";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiSfId", @" AND a.RfiSfId = @RfiSfId", RfiSfId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition = @Edition", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiDetailId", @" AND a.RfiDetailId = @RfiDetailId", RfiDetailId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();

                    sql = @"SELECT RfiSfId 
                            , LEAD(SortNumber) OVER (PARTITION BY RfiDetailId ORDER BY SortNumber) NextSortNumber
                            FROM RFI.RfiSignFlow
                            WHERE RfiDetailId IN @RfiDetailIds";
                    var resultRfiSignFlow = sqlConnection.Query(sql, new { RfiDetailIds = result.Select(x => x.RfiDetailId).Distinct().ToArray() }).ToList();

                    result.ForEach(x =>
                    {
                        var RfiSignFlow = resultRfiSignFlow.FirstOrDefault(f => f.RfiSfId == x.RfiSfId);
                        x.NextSortNumber = RfiSignFlow?.NextSortNumber ?? null;
                    });

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

        #region //GetRfiDetailByTemplate -- 取得RFI單身(樣板單筆查詢) (停用) -- Chia Yuan 2024.01.15
        //public string GetRfiDetailByTemplate(int RfiId, int RfiDetailId)
        //{
        //    try
        //    {
        //        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
        //        {
        //            dynamicParameters = new DynamicParameters();
        //            sql = @"SELECT a.TempProdSpecId,ISNULL(b.ProdSpecId, -1) AS ProdSpecId
        //                    , ISNULL(b.SpecName, a.SpecName) AS SpecName
        //                    , ISNULL(b.SpecEName,a.SpecEName) AS SpecEName
        //                    , ISNULL(b.ControlType,a.ControlType) AS ControlType
        //                    , ISNULL(b.[Required],a.[Required]) AS [Required]
        //                    , ISNULL(b.SortNumber,a.SortNumber) AS SortNumber
        //                    , (
        //                     SELECT aa.TempPsDetailId, aa.FeatureName
        //                        , ISNULL(aa.DataType,'') AS DataType
        //                        , ISNULL(aa.[Description],'') AS [Description]
        //                        , aa.SortNumber, ac.*
        //                     FROM RFI.TemplateProdSpecDetail aa
        //                     OUTER APPLY (
        //                      SELECT ca.RfiDetailId, ca.RfiId, ca.RfiSequence, ca.Edition, ca.TargetUnitPrice, ca.LifeCycleQty, ca.Currency, ca.Revenue, ca.RevenueFC, ca.ExchangeRate
        //                            , ISNULL(FORMAT(ca.ProdDate, 'yyyy-MM-dd'), '') ProdDate
        //                            , ISNULL(FORMAT(ca.ProdLifeCycleStart, 'yyyy-MM-dd'), '') ProdLifeCycleStart
        //                            , ISNULL(FORMAT(ca.ProdLifeCycleEnd, 'yyyy-MM-dd'), '') ProdLifeCycleEnd
        //                            , ca.SalesId, cc.UserNo AS SalesNo, cc.UserName AS SalesName, cc.Gender AS SalesGender
        //                            , ca.[Status] AS RfiDetailStatus
        //                            , ce.StatusName AS RfiDetailStatusName
        //                      , ISNULL((
        //                       SELECT aca.PsDetailId
        //                        FROM RFI.ProductSpecDetail aca 
        //                        INNER JOIN RFI.ProductSpec acb ON acb.ProdSpecId = aca.ProdSpecId 
        //                        WHERE acb.RfiDetailId = ca.RfiDetailId  
        //                        AND aca.FeatureName = aa.FeatureName)
        //                       , -1) AS PsDetailId
        //                      , ISNULL((
        //                       SELECT TOP 1 1 FROM RFI.ProductSpecDetail aca 
        //                        INNER JOIN RFI.ProductSpec acb ON acb.ProdSpecId = aca.ProdSpecId 
        //                        WHERE acb.RfiDetailId = ca.RfiDetailId 
        //                        AND aca.FeatureName = aa.FeatureName)
        //                       , 0) AS IsExists
        //                      FROM RFI.RfiDetail ca
        //                      INNER JOIN RFI.RequestForInformation cb ON cb.RfiId = ca.RfiId
        //                      INNER JOIN BAS.[User] cc ON cc.UserId = ca.SalesId
        //                      LEFT JOIN BAS.[Type] cd ON cd.TypeNo = ca.Currency AND cd.TypeSchema = 'ExchangeRate.Currency'
        //                      INNER JOIN BAS.[Status] ce ON ce.StatusNo = ca.[Status] AND ce.StatusSchema = 'RfiDetail.Status'
        //                      WHERE ca.RfiDetailId = @RfiDetailId
        //                      AND ca.CompanyId = @CompanyId
        //                     ) ac
        //                     WHERE aa.TempProdSpecId = a.TempProdSpecId
        //                     FOR JSON PATH, ROOT('data')
        //                    ) AS ProductSpecDetail
        //                    FROM RFI.TemplateProdSpec a
        //                    OUTER APPLY (
        //                     SELECT ba.ProdSpecId, ba.SpecName, ba.SpecEName, ba.ControlType, ba.[Required], ba.SortNumber
        //                     FROM RFI.ProductSpec ba 
        //                        INNER JOIN RFI.RfiDetail bb ON bb.RfiDetailId = ba.RfiDetailId
        //                        INNER JOIN RFI.RequestForInformation bc ON bc.RfiId = bb.RfiId AND bc.ProTypeGroupId = ba.ProTypeGroupId
        //                     WHERE ba.SpecName = a.SpecName
        //                    ) b
        //                    ORDER BY a.SortNumber";
        //            dynamicParameters.Add("CompanyId", CurrentCompany);
        //            dynamicParameters.Add("RfiDetailId", RfiDetailId);
        //            var result = sqlConnection.Query(sql, dynamicParameters);

        //            #region //Response
        //            jsonResponse = JObject.FromObject(new
        //            {
        //                status = "success",
        //                data = result
        //            });
        //            #endregion
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "errorForDA",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    return jsonResponse.ToString();
        //}
        #endregion

        #region //GetExchangeRateERP -- 取得四捨五入位數資料 -- Chia Yuan 2024.01.11
        public string GetExchangeRateERP(string Currency)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //取得共用參數
                        sql = @"SELECT TOP 1 MA003 LocalCurrency FROM CMSMA";
                        var resultCMSMA = erpConnection.QueryFirstOrDefault(sql) ?? throw new SystemException(string.Format("【ERP共用參數檔】資料錯誤!"));
                        #endregion

                        //if (string.IsNullOrWhiteSpace(Currency)) Currency = resultCMSMA.LocalCurrency;

                        string[] Currencys = new string[] { resultCMSMA.LocalCurrency, Currency };

                        #region //目前匯率+幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 LTRIM(RTRIM(MF001)) Currency, LTRIM(RTRIM(MF002)) CurrencyName, MF003 UnitRound, MF004 TotalRound, MF005, MF006
                                , b.MG002, b.MG004 ExchangeRate
                                FROM CMSMF a
                                INNER JOIN CMSMG b ON LTRIM(RTRIM(b.MG001)) = LTRIM(RTRIM(MF001))
                                WHERE b.MG002 <= @MG002";
                        dynamicParameters.Add("MG002", CreateDate.ToString("yyyyMMdd"));
                        if (Currencys.Length > 0)
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MG001", @" AND LTRIM(RTRIM(b.MG001)) IN @MG001", Currencys);
                        sql += " ORDER BY b.MG002 DESC";
                        var result = erpConnection.Query(sql, dynamicParameters).ToList();
                        //throw new SystemException(string.Format("【ERP交易幣別匯率】{0}不存在!", Currency));

                        //double exchangeRate = Convert.ToDouble(resultCMSMG.MG004);

                        #region //停用
                        //switch (exchangeRateSource)
                        //{
                        //    case "I": //銀行買進匯率
                        //        exchangeRate = Convert.ToDouble(item.MG003);
                        //        break;
                        //    case "O": //銀行賣出匯率
                        //        exchangeRate = Convert.ToDouble(item.MG004);
                        //        break;
                        //    case "E": //報關買進匯率
                        //        exchangeRate = Convert.ToDouble(item.MG005);
                        //        break;
                        //    case "W": //報關賣出匯率
                        //        exchangeRate = Convert.ToDouble(item.MG006);
                        //        break;
                        //}
                        #endregion
                        #endregion

                        result.ForEach(x =>
                        {
                            x.IsLocal = x.Currency == resultCMSMA.LocalCurrency;
                        });

                        #region //交易幣別設定
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT LTRIM(RTRIM(MF001)) Currency, LTRIM(RTRIM(MF002)) CurrencyName, MF003 UnitRound, MF004 TotalRound, MF005, MF006
                        //        FROM CMSMF
                        //        WHERE 1=1";
                        //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MF001", @" AND LTRIM(RTRIM(MF001)) = @MF001", Currency);
                        //var result = erpConnection.Query(sql, new { Currency });

                        //throw new SystemException(string.Format("【ERP交易幣別設定】{0} 不存在!", Currency));
                        //int unitRound = Convert.ToInt32(result.MF003); //單價取位
                        //int totalRound = Convert.ToInt32(result.MF004); //金額取位
                        //string currency = result.MF001;
                        //string currencyName = result.MF002;
                        #endregion

                        //JObject data = JObject.FromObject(new
                        //{
                        //    UnitRound = unitRound,
                        //    TotalRound = totalRound,
                        //    ExchangeRate = exchangeRate,
                        //    CurrencyName = currencyName,
                        //    Currency = currency
                        //});

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
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

        #region //GetDemandDesign -- 取得新設計申請單 -- Chia Yuan 2024.01.22
        public string GetDemandDesign(int DesignId, int RfiId, string DesignNo, int SalesId, int ProTypeGroupId, int ProductUseId
            , string CustomerName, string AssemblyName, string StartDate, string EndDate
            , string SignStatus, string DesignStatus, string FlowStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DesignId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfiId, a.DesignNo, a.CompanyId, a.Edition, a.AssemblyName
                        , b.RfiNo, b.CustomerName, ISNULL(b.CustomerEName, '') CustomerEName, b.ProTypeGroupId, f.ProTypeGroupName, b.ProductUseId, g.ProductUseName
                        , a.[Status] DesignStatus, s.StatusName DesignStatusName
                        , a.SalesId, ISNULL(u.UserNo,'') SalesNo, ISNULL(u.UserName,'') SalesName, u.Gender SalesGender
                        , a.SuperSalesId, ISNULL(t.UserNo,'') SuperSalesNo, ISNULL(t.UserName,'') SuperSalesName, t.Gender SuperSalesGender
                        , h.DepartmentName, h2.DepartmentName DesignDeptName, i.CompanyName
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                        , ISNULL(ab.PreSignJobName, '') PreSignJobName, ISNULL(ab.NextSignJobName, '') NextSignJobName
                        , flow.*
                        , (
	                        SELECT aa.DesignSfId, aa.DesignId
                            , ISNULL(aa.Approved, '') Approved
                            --, CASE ISNULL(aa.Approved, '') WHEN 'Y' THEN '核准' WHEN 'N' THEN '否決' ELSE '待審批' END 'ApprovedName'
                            , ISNULL(s2.StatusName, '待審批') ApprovedName
                            , ISNULL(aa.SignUser, aa.FlowUser) Flower
                            , ISNULL(au2.UserNo, au1.UserNo) FlowerNo
                            , ISNULL(au2.UserName, au1.UserName) FlowerName
                            , ISNULL(au2.Gender, au1.Gender) FlowerGender
                            , ISNULL(au2.Email, au1.Email) FlowerEmail
	                        , ISNULL(aa.SignJobName, '') SignJobName
	                        , ISNULL(aa.SignDesc, '') SignDesc
                            , s1.StatusDesc FlowStatusDesc
	                        , ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime
	                        , ISNULL(FORMAT(aa.SignDateTime, 'yyyy-MM-dd HH:mm'), '') SignDateTime
	                        , DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) / 1440 StayDays
	                        , (DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
	                        , (DATEDIFF(MINUTE, aa.ArriveDateTime, ISNULL(aa.SignDateTime, @CreateDate)) % 60) StayMinutes
	                        , aa.SortNumber
                            , ISNULL(af.FileId, '') FileId, ISNULL(af.[FileName], '') [FileName], ISNULL(af.FileExtension, '') FileExtension
	                        FROM RFI.DesignSignFlow aa
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = aa.[Status] AND s1.StatusSchema = 'RfiDesign.FlowStatus'
                            LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = aa.Approved AND s2.StatusSchema = 'RfiSign.Status'
	                        INNER JOIN BAS.[User] au1 ON au1.UserId = aa.FlowUser
	                        LEFT JOIN BAS.[User] au2 ON au2.UserId = aa.SignUser
                            OUTER APPLY (SELECT FileId, [FileName], FileExtension FROM BAS.[File] WHERE FileId = aa.AdditionalFile AND DeleteStatus = 'N') af
	                        WHERE aa.DesignId = a.DesignId
                            ORDER BY aa.SortNumber
	                        FOR JSON PATH, ROOT('data')
                        ) SignFlow";

                    if (DesignId > 0) {
                        sqlQuery.columns += @", ISNULL(j.FileId, '') 'AdditionalFile', j.[FileName] 'AdditionalFileName', j.FileExtension 'AdditionalFileExtension'
                                              , ISNULL(k.FileId, '') 'DesignFile', k.[FileName] 'DesignFileName', k.FileExtension 'DesignFileExtension'";
                    }

                        sqlQuery.columns += @", (SELECT aa.DdSpecId, aa.ParameterName, aa.SortNumber, aa.ControlType, aa.[Required], aa.[Status], ab.StatusName
                                            , ISNULL(SUBSTRING(ac.FeatureNames, 0, LEN(ac.FeatureNames)), '') 'FeatureNames'";

                    if (DesignId > 0)
                    {
                        //單筆查詢才撈出規格相關資料
                        sqlQuery.columns += @", (SELECT ba.DdsDetailId, ba.PmtDetailName, ba.SortNumber
                                                , ISNULL(ba.DataType,'') 'DataType'
                                                , ISNULL(ba.RequireFlag,'') 'RequireFlag'
                                                , ISNULL(ba.DesignInput,'') 'DesignInput'
                                                , ISNULL(ba.[Description],'') [Description]
                                                , ISNULL(bc.TypeName,'') 'DataTypeName'
                                                , ba.[Status]
                                                , bb.StatusName
		                                        FROM RFI.DemandDesignSpecDetail ba
		                                        INNER JOIN BAS.[Status] bb ON bb.StatusNo = ba.[Status] AND bb.StatusSchema = 'Status'
                                                LEFT JOIN BAS.[Type] bc ON bc.TypeNo = ba.DataType AND bc.TypeSchema = 'RfiFormItem.OtherType'
		                                        WHERE ba.DdSpecId = aa.DdSpecId
                                                AND ba.[Status] = 'A'
		                                        ORDER BY ba.SortNumber
		                                        FOR JSON PATH, ROOT('data')
	                                        ) DemandDesignSpecDetail";
                    }
                    sqlQuery.columns += @" FROM RFI.DemandDesignSpec aa
	                                       INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND ab.StatusSchema = 'Status'
                                           OUTER APPLY (
                                               SELECT (
                                                   SELECT ca.PmtDetailName + CASE ISNULL(ca.DesignInput, '') WHEN '' THEN '' ELSE ':' + ca.DesignInput END + ','
                                                   FROM RFI.DemandDesignSpecDetail ca
                                                   WHERE ca.DdSpecId = aa.DdSpecId
                                                   AND ca.[Status] = 'A'
                                                   ORDER BY ca.SortNumber
                                                   FOR XML PATH('')
                                               ) FeatureNames
                                           ) ac
	                                       WHERE aa.DesignId = a.DesignId
                                           AND aa.[Status] = 'A'
	                                       ORDER BY aa.SortNumber
	                                       FOR JSON PATH, ROOT('data')
                                        ) DemandDesignSpec";
                    //if (SalesId > 0)
                    //{
                    //    //用於判斷單身多筆不同指派人員，需濾掉其他業務，只能看自己的單據
                    //    sqlQuery.columns += @" AND EXISTS (
                    //                                SELECT TOP 1 1
                    //                                FROM RFI.RfiDetail ba
                    //                                WHERE ba.RfiId = aa.RfiId
                    //                                AND ba.SalesId = @SalesId
                    //                           )";
                    //}
                    //sqlQuery.columns +=
                    //    @" ORDER BY aa.RfiDetailId
                    //        FOR JSON PATH, ROOT('data')
                    //    ) RfiDetail";
                    sqlQuery.mainTables =
                        @"FROM RFI.DemandDesign a
                        INNER JOIN RFI.RequestForInformation b ON b.RfiId = a.RfiId
                        INNER JOIN BAS.[User] u ON u.UserId = a.SalesId
                        INNER JOIN BAS.[User] t ON t.UserId = a.SuperSalesId
                        INNER JOIN SCM.ProductTypeGroup f ON f.ProTypeGroupId = b.ProTypeGroupId
                        INNER JOIN SCM.ProductUse g ON g.ProductUseId = b.ProductUseId
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'RfiDesign.Status'
                        INNER JOIN BAS.Department h ON h.DepartmentId = u.DepartmentId
                        INNER JOIN BAS.Department h2 ON h2.DepartmentId = a.DepartmentId
                        INNER JOIN BAS.Company i ON i.CompanyId = h.CompanyId
                        OUTER APPLY (
	                        SELECT *
	                        FROM (
		                        SELECT ISNULL(aa.SignUser, aa.FlowUser) Flower, aa.DesignSfId, aa.SignJobName, aa.SortNumber, ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime, aa.Approved
		                        --, ROW_NUMBER() OVER (PARTITION BY aa.DesignId ORDER BY aa.SortNumber) SortRow
		                        --虛擬狀態 W=等待審批, I=進行中, D=否決, C=核准(完成)
		                        , CASE ISNULL(aa.ArriveDateTime, '') WHEN '' THEN 'W' ELSE (CASE ISNULL(aa.Approved, '') WHEN '' THEN 'I' WHEN 'N' THEN 'D' ELSE 'C' END) END SignStatus
                                , aa.[Status] FlowStatus, ab.StatusName FlowStatusName
		                        FROM RFI.DesignSignFlow aa
                                INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND StatusSchema = 'RfiDesign.FlowStatus'
		                        WHERE aa.DesignId = a.DesignId
	                        ) x WHERE x.SignStatus IN ('I', 'D')
                        ) flow
                        LEFT JOIN (
	                        SELECT ba.DesignSfId
	                        , LAG(ba.SignJobName) OVER (PARTITION BY ba.DesignId ORDER BY ba.SortNumber) PreSignJobName
	                        , LEAD(ba.SignJobName) OVER (PARTITION BY ba.DesignId ORDER BY ba.SortNumber) NextSignJobName
	                        FROM RFI.DesignSignFlow ba
                        ) ab ON ab.DesignSfId = flow.DesignSfId";
                    if (DesignId > 0) {
                        sqlQuery.mainTables += @" OUTER APPLY (SELECT FileId, [FileName], FileExtension FROM BAS.[File] WHERE FileId = a.AdditionalFile AND DeleteStatus = 'N') j
                                                OUTER APPLY (SELECT FileId, [FileName], FileExtension FROM BAS.[File] WHERE FileId = a.DesignFile AND DeleteStatus = 'N') k";
                    }
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND (EXISTS (
                                            SELECT TOP 1 1
                                            FROM RFI.DesignSignFlow da
                                            OUTER APPLY (
                                                SELECT DISTINCT ISNULL(ab.SignUser, ab.FlowUser) AS FlowUser
                                                FROM RFI.RfiDetail aa
                                                INNER JOIN RFI.RfiSignFlow ab ON ab.RfiDetailId = aa.RfiDetailId
	                                            WHERE aa.RfiId = a.RfiId
                                            ) AS db
                                            WHERE da.DesignId = a.DesignId
                                            AND (ISNULL(da.SignUser, da.FlowUser) = @Flower OR db.FlowUser = @Flower)) OR a.SalesId = @Flower)";
                    dynamicParameters.Add("CreateDate", CreateDate);
                    dynamicParameters.Add("Flower", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignNo", @" AND a.DesignNo = @DesignNo", DesignNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiId", @" AND b.RfiId = @RfiId", RfiId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND b.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND b.ProductUseId = @ProductUseId", ProductUseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND b.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssemblyName", @" AND a.AssemblyName LIKE '%' + @AssemblyName + '%'", AssemblyName);
                    if (DateTime.TryParse(StartDate, out DateTime startDate))
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", " AND a.CreateDate >= @StartDate", startDate.ToString("yyyy-MM-dd 00:00:00"));
                    if (DateTime.TryParse(EndDate, out DateTime endDate))
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", " AND a.CreateDate <= @EndDate", endDate.ToString("yyyy-MM-dd 23:59:59"));

                    //只能看自己的單據
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesId", " AND a.SalesId = @SalesId", SalesId);

                    //待審批
                    if (SignStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SignStatus", @" AND flow.SignStatus = @SignStatus", SignStatus.Split(','));
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Flower", @" AND flow.Flower = @Flower", CurrentUser);
                    }

                    //單據狀態
                    if (DesignStatus.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.[Status] IN @Status", DesignStatus.Split(','));

                    //關卡狀態
                    if (FlowStatus.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowStatus", @" AND flow.FlowStatus IN @FlowStatus", FlowStatus.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DesignNo DESC, a.Edition DESC";
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

        #region //GetDesignSignFlow -- 取得設計申請審批流程 -- Chia Yuan 2024.06.25
        public string GetDesignSignFlow(int DesignSfId, int DesignId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DesignSfId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignId, a.SignJobName, a.Approved, b.Edition, a.SortNumber, b.DepartmentId
                        , e.ProTypeGroupId, e.ProductUseId, ISNULL(f.FileId, '') AdditionalFile
                        , u1.UserId, u1.UserName, u1.UserNo, u1.Email
                        , c.CompanyId, c.CompanyName, c.CompanyNo
                        , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo";
                    sqlQuery.mainTables =
                        @"FROM RFI.DesignSignFlow a
                        INNER JOIN RFI.DemandDesign b ON b.DesignId = a.DesignId
                        INNER JOIN RFI.RequestForInformation e ON e.RfiId = b.RfiId
                        INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                        INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                        INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                        OUTER APPLY (SELECT TOP 1 fa.FileId, fa.[FileName], fa.FileExtension FROM BAS.[File] fa WHERE fa.FileId = a.AdditionalFile AND fa.DeleteStatus = 'N') f";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignSfId", @" AND a.DesignSfId = @DesignSfId", DesignSfId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();

                    sql = @"SELECT DesignSfId 
                            , LEAD(SortNumber) OVER (PARTITION BY DesignId ORDER BY SortNumber) NextSortNumber
                            FROM RFI.DesignSignFlow
                            WHERE DesignId IN @DesignIds";
                    var resultDesignSignFlow = sqlConnection.Query(sql, new { DesignIds = result.Select(x => x.DesignId).Distinct().ToArray() }).ToList();

                    result.ForEach(x =>
                    {
                        var DesignSignFlow = resultDesignSignFlow.FirstOrDefault(f => f.DesignSfId == x.DesignSfId);
                        x.NextSortNumber = DesignSignFlow?.NextSortNumber ?? null;
                    });

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

        #region //GetDemandDesignSpec -- 取得新設計規格資料 -- Chia Yuan 2024.01.24
        public string GetDemandDesignSpec(int DdSpecId, int DesignId, string ParameterName, string PmtDetailName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DdSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignId,a.ParameterName,a.SortNumber,a.ControlType,a.[Required],a.[Status],s.StatusName AS DesignStatusName
                        , ISNULL(SUBSTRING(c.FeatureNames, 0, LEN(c.FeatureNames)), '') AS FeatureNames
                        , (
	                        SELECT aa.DdSpecId,aa.DdsDetailId,aa.PmtDetailName,aa.SortNumber
	                        , ISNULL(aa.DataType,'') AS DataType
	                        , ISNULL(aa.RequireFlag,'') AS RequireFlag
	                        , ISNULL(aa.DesignInput,'') AS DesignInput
	                        , ISNULL(aa.[Description],'') AS [Description]
	                        , ISNULL(ac.TypeName,'') AS DataTypeName
                            , aa.[Required]
	                        , aa.[Status]
                            , ab.StatusName
	                        FROM RFI.DemandDesignSpecDetail aa
	                        INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND ab.StatusSchema = 'Status'
	                        LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.DataType AND ac.TypeSchema = 'RfiFormItem.OtherType'
	                        WHERE aa.DdSpecId = a.DdSpecId
	                        ORDER BY aa.SortNumber
	                        FOR JSON PATH, ROOT('data')
                        ) AS DemandDesignSpecDetail";
                    sqlQuery.mainTables =
                        @"FROM RFI.DemandDesignSpec a
                        INNER JOIN RFI.DemandDesign b ON b.DesignId = a.DesignId
                        OUTER APPLY (
                            SELECT (
                                SELECT ca.PmtDetailName + ','
                                FROM RFI.DemandDesignSpecDetail ca
                                WHERE ca.DdSpecId = a.DdSpecId
                                AND ca.[Status] = 'A'
                                ORDER BY ca.SortNumber
                                FOR XML PATH('')
                            ) FeatureNames
                        ) c
                        INNER JOIN BAS.[Type] t ON t.TypeNo = a.ControlType AND t.TypeSchema = 'RfiFormItem.ItemType'
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DdSpecId", @" AND a.DdSpecId = @DdSpecId", DdSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterName", @" AND a.ParameterName LIKE '%' + @ParameterName + '%'", ParameterName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PmtDetailName", @" AND EXISTS (
                                                                                                        SELECT TOP 1 1
                                                                                                        FROM RFI.DemandDesignSpecDetail da
                                                                                                        WHERE ba.DdSpecId = a.DdSpecId
                                                                                                        AND da.PmtDetailName LIKE '%' + @PmtDetailName + '%')", PmtDetailName);

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

        #region //GetDemandDesignSpecDetail -- 取得設計規格子資料 -- Chia Yuan -- 2024.02.01
        public string GetDemandDesignSpecDetail(int DdsDetailId, int DdSpecId, string PmtDetailName, string DesignInput, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DdSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.DdsDetailId,a.PmtDetailName,a.SortNumber
	                    , ISNULL(a.DataType,'') AS DataType
	                    , ISNULL(a.RequireFlag,'') AS RequireFlag
	                    , ISNULL(a.DesignInput,'') AS DesignInput
	                    , ISNULL(a.[Description],'') AS [Description]
	                    , ISNULL(c.TypeName,'') AS DataTypeName
                        , a.[Required]
	                    , a.[Status]
                        , b.StatusName";
                    sqlQuery.mainTables =
                        @"FROM RFI.DemandDesignSpecDetail a
	                    INNER JOIN BAS.[Status] b ON b.StatusNo = a.[Status] AND b.StatusSchema = 'Status'
	                    LEFT JOIN BAS.[Type] c ON c.TypeNo = a.DataType AND c.TypeSchema = 'RfiFormItem.OtherType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DdsDetailId", @" AND a.DdsDetailId = @DdsDetailId", DdsDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DdSpecId", @" AND a.DdSpecId = @DdSpecId", DdSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PmtDetailName", @" AND a.PmtDetailName LIKE '%' + @PmtDetailName + '%'", PmtDetailName);
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

        #region //GetDesignInput -- 取得設計規格輸入資料 -- Chia Yuan -- 2024.02.02
        public string GetDesignInput(int DdsDetailId, int DdSpecId, string PmtDetailName, string DesignInput, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.DesignInput
	                        FROM RFI.DemandDesignSpecDetail a
	                        INNER JOIN BAS.[Status] b ON b.StatusNo = a.[Status] AND b.StatusSchema = 'Status'
	                        LEFT JOIN BAS.[Type] c ON c.TypeNo = a.DataType AND c.TypeSchema = 'RfiFormItem.OtherType'
                            WHERE ISNULL(a.DesignInput, '') <> ''";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DdsDetailId", @" AND a.DdsDetailId = @DdsDetailId", DdsDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DdSpecId", @" AND a.DdSpecId = @DdSpecId", DdSpecId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PmtDetailName", @" AND a.PmtDetailName LIKE '%' + @PmtDetailName + '%'", PmtDetailName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sql += @" ORDER BY a.DesignInput";

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

        #region //GetDesignHistory -- 取得設計規格歷史紀錄 -- Chia Yuan -- 2024.02.16
        public string GetDesignHistory(int DesignId, string SchemaSpecEName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @tempFeatureName TABLE (FeatureName NVARCHAR(50), ProTypeGroupId INT, ProductUseId INT)
                            INSERT @tempFeatureName
                            SELECT DISTINCT ag.FeatureName, ad.ProTypeGroupId, ad.ProductUseId
                            FROM RFI.DemandDesign aa
                            INNER JOIN RFI.RequestForInformation ad ON ad.RfiId = aa.RfiId
                            INNER JOIN RFI.RfiDetail ae ON ae.RfiId = ad.RfiId
                            INNER JOIN RFI.ProductSpec af ON af.RfiDetailId = ae.RfiDetailId
                            INNER JOIN RFI.ProductSpecDetail ag ON ag.ProdSpecId = af.ProdSpecId AND ag.[Status] = af.[Status]
                            WHERE aa.DesignId = @DesignId
                            AND ae.[Status] = '5' --新設計申請單
                            AND ag.[Status] = 'A'";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SchemaSpecEName", @" AND af.SpecEName LIKE '%' + @SchemaSpecEName + '%'", SchemaSpecEName);

                    sql += @"
                            DECLARE @tempTable TABLE (mName NVARCHAR(255), DesignInput NVARCHAR(512), mSort INT, dSort INT)
                            INSERT @tempTable
                            SELECT DISTINCT a.ParameterName,b.DesignInput,a.SortNumber,b.SortNumber
                            FROM RFI.DemandDesignSpec a
                            INNER JOIN RFI.DemandDesignSpecDetail b ON b.DdSpecId = a.DdSpecId AND b.[Status] = a.[Status]
                            INNER JOIN RFI.DemandDesign d ON d.DesignId = a.DesignId
                            AND EXISTS (
	                            SELECT TOP 1 1
                                FROM RFI.RequestForInformation aa
	                            INNER JOIN RFI.RfiDetail ab ON ab.RfiId = aa.RfiId
	                            INNER JOIN RFI.ProductSpec ac ON ac.RfiDetailId = ab.RfiDetailId
	                            INNER JOIN RFI.ProductSpecDetail ad ON ad.ProdSpecId = ac.ProdSpecId
	                            WHERE ab.[Status] = '5' --新設計申請單
	                            AND aa.RfiId = d.RfiId
	                            AND EXISTS (
		                            SELECT TOP 1 1
		                            FROM @tempFeatureName ae
		                            WHERE ae.ProTypeGroupId = aa.ProTypeGroupId
		                            AND ae.ProductUseId = aa.ProductUseId
		                            AND ae.FeatureName = ad.FeatureName
	                            )
                            )
                            ORDER BY a.SortNumber, b.SortNumber

                            SELECT DISTINCT mName FROM @tempTable";
                    dynamicParameters.Add("DesignId", DesignId);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    string columns = string.Join(",", result.Select(s => string.Format("[{0}]", s.mName)));

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT *
                            FROM (
	                            SELECT c.RfiId,y.col,y.val
	                            FROM RFI.DemandDesignSpec b
	                            INNER JOIN RFI.DemandDesign c ON c.DesignId = b.DesignId
	                            CROSS APPLY (
		                            VALUES (b.ParameterName
		                            , (SELECT 
			                            aa.PmtDetailName,aa.[Required],ISNULL(aa.RequireFlag, '') AS RequireFlag,ISNULL(aa.DesignInput, '') AS DesignInput,ISNULL(aa.[Description], '') AS [Description]
			                            FROM RFI.DemandDesignSpecDetail aa 
			                            WHERE aa.DdSpecId = b.DdSpecId 
                                        AND aa.[Status] = 'A'
			                            ORDER BY aa.CreateDate DESC, aa.SortNumber
			                            FOR JSON PATH, ROOT('data'))
		                            )
	                            ) y (col, val)
                            ) t 
                            PIVOT (MAX(t.val) FOR t.col IN (" + columns + ")) p";

                    result = sqlConnection.Query(sql, dynamicParameters);

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

        #region //GetRfiToExcel --市場評估單報表資料 -- Chia Yuan -- 2024.05.22
        public string GetRfiToExcel(int RfiId, string RfiNo, int ProTypeGroupId, int ProductUseId, int SalesId, string StartDate, string EndDate, string SignStatus, string RfiDetailStatus, string FlowStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得市場評估單資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RfiId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns = "";
                    sqlQuery.mainTables =
                        @"FROM RFI.RequestForInformation a
                        OUTER APPLY (
	                        SELECT *
	                        FROM (
		                        SELECT ISNULL(aa.SignUser, aa.FlowUser) Flower, aa.RfiSfId, aa.SignJobName, aa.SortNumber, ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') ArriveDateTime, aa.Approved
		                        --, ROW_NUMBER() OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.SortNumber) SortRow
		                        --虛擬狀態 W=等待審批, I=進行中, D=否決, C=核准(完成)
		                        , CASE ISNULL(aa.ArriveDateTime, '') WHEN '' THEN 'W' ELSE (CASE ISNULL(aa.Approved, '') WHEN '' THEN 'I' WHEN 'N' THEN 'D' WHEN 'B' THEN 'B' ELSE 'C' END) END SignStatus
                                , aa.[Status] FlowStatus, ab.StatusName FlowStatusName
                                , DENSE_RANK () OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.Edition DESC) SortEdition
		                        FROM RFI.RfiSignFlow aa
                                INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND StatusSchema = 'RfiDetail.FlowStatus'
		                        INNER JOIN RFI.RfiDetail ac ON ac.RfiDetailId = aa.RfiDetailId
		                        WHERE ac.RfiId = a.RfiId
	                        ) x WHERE x.SignStatus IN ('I', 'D')
                            AND x.SortEdition = 1
                        ) flow";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND (EXISTS (
                                            SELECT TOP 1 1
                                            FROM RFI.RfiDetail da
                                            INNER JOIN RFI.RfiSignFlow db ON da.RfiDetailId = db.RfiDetailId
                                            OUTER APPLY (
                                                SELECT DISTINCT ISNULL(ab.SignUser, ab.FlowUser) AS DesignFlower
                                                FROM RFI.DemandDesign aa
                                                INNER JOIN RFI.DesignSignFlow ab ON ab.DesignId = aa.DesignId
                                                WHERE aa.RfiId = da.RfiId
                                            ) AS dc
                                            WHERE da.RfiId = a.RfiId
                                            AND (ISNULL(db.SignUser, db.FlowUser) = @Flower OR dc.DesignFlower = @Flower)) OR a.SalesId = @Flower)";

                    dynamicParameters.Add("Flower", CurrentUser);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiId", " AND a.RfiId = @RfiId", RfiId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfiNo", " AND a.RfiNo LIKE '%' + @RfiNo + '%'", RfiNo.Trim());
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters,
                        "SalesId",
                        @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM RFI.RfiDetail ba
                            WHERE ba.RfiId = a.RfiId
                            AND ba.SalesId = @SalesId
                        )", SalesId);
                    if (DateTime.TryParse(StartDate, out DateTime startDate))
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", " AND a.AnalysisDate >= @StartDate", startDate.ToString("yyyy-MM-dd 00:00:00"));
                    if (DateTime.TryParse(EndDate, out DateTime endDate))
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", " AND a.AnalysisDate <= @EndDate", endDate.ToString("yyyy-MM-dd 23:59:59"));

                    //待審批
                    if (SignStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SignStatus", @" AND flow.SignStatus = @SignStatus", SignStatus.Split(','));
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Flower", @" AND flow.Flower = @Flower", CurrentUser);
                    }

                    //關卡狀態
                    if (FlowStatus.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowStatus", @" AND flow.FlowStatus IN @FlowStatus", FlowStatus.Split(','));

                    //單身狀態
                    if (RfiDetailStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters,
                            "RfiDetailStatus",
                            @" AND EXISTS (
                                SELECT TOP 1 1
                                FROM RFI.RfiDetail ba
                                WHERE ba.RfiId = a.RfiId
                                AND ba.[Status] IN @RfiDetailStatus
                            )", RfiDetailStatus.Split(','));
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfiId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    if (result.Count() <= 0) throw new SystemException("無法取得市場評估單資料!");
                    #endregion

                    string columns = string.Empty;

                    #region //取得欄位資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @tempRfi TABLE (col NVARCHAR(255), mSort INT, dSort INT, sortNum INT)
                            INSERT @tempRfi
                            SELECT DISTINCT a.SpecName, a.SortNumber, b.SortNumber
                            , ROW_NUMBER() OVER (PARTITION BY a.SpecName ORDER BY a.SortNumber, c.RfiId DESC)
                            FROM RFI.ProductSpec a
                            INNER JOIN RFI.ProductSpecDetail b ON b.ProdSpecId = a.ProdSpecId AND b.[Status] = a.[Status]
                            INNER JOIN RFI.RfiDetail c ON c.RfiDetailId = a.RfiDetailId
                            WHERE b.[Status] = 'A'
                            AND c.RfiId IN @RfiIds

                            DECLARE @columns VARCHAR(MAX)
                            SELECT
                                @columns = CASE
                                    WHEN @columns IS NULL
                                    THEN '[' + T.col + ']'
                                    ELSE @columns + ',[' + T.col + ']'
                                END
                            FROM @tempRfi T
                            WHERE sortNum = 1
                            ORDER BY mSort, col
                            SELECT @columns AS [Columns]";
                    dynamicParameters.Add("RfiIds", result.Select(s => s.RfiId).Distinct().ToArray());
                    columns = sqlConnection.QueryFirst(sql, dynamicParameters).Columns;

                    if (columns == null) throw new SystemException("沒有可用的規格項目!");
                    #endregion

                    #region //取得主要資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT *
                            FROM (
	                            SELECT d.RfiNo AS 需求單號
                                , u.UserName AS 業務擔當, FORMAT(d.AnalysisDate, 'yyyy-MM-dd HH:mm') AS 分析日期, s.StatusName AS 單據狀態
                                , g.ProTypeGroupName AS 產品類別, e.ProductUseName AS 產品應用, s2.StatusName AS 是否鍍膜, b.AssemblyName AS 設計代號
                                , y.col,SUBSTRING(y.val, 0, LEN(y.val)) AS val
	                            FROM RFI.ProductSpec a
	                            INNER JOIN RFI.RfiDetail c ON c.RfiDetailId = a.RfiDetailId
	                            INNER JOIN RFI.RequestForInformation d ON d.RfiId = c.RfiId
                                INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = d.ProTypeGroupId
                                INNER JOIN SCM.ProductUse e ON e.ProductUseId = d.ProductUseId
                                INNER JOIN BAS.[Status] s2 ON s2.StatusNo = g.CoatingFlag AND s2.StatusSchema = 'Boolean'
	                            INNER JOIN BAS.[Status] s ON s.StatusNo = c.[Status] AND s.StatusSchema = 'RfiDetail.Status'
	                            INNER JOIN BAS.[User] u ON u.UserId = d.SalesId
                                OUTER APPLY (
	                                SELECT TOP 1 ba.AssemblyName
	                                FROM RFI.DemandDesign ba
	                                WHERE ba.RfiId = d.RfiId
	                                ORDER BY ba.Edition DESC
                                ) AS b
	                            CROSS APPLY (
		                            VALUES (a.SpecName
		                            , (SELECT aa.FeatureName + 
			                            CASE WHEN ISNULL(aa.DataType, '') = 'O' THEN ':' 
			                            WHEN a.ControlType = 'A' OR a.ControlType = 'T' THEN ':'
			                            ELSE '' END + ISNULL(aa.[Description], '') + ','
			                            FROM RFI.ProductSpecDetail aa 
			                            WHERE aa.ProdSpecId = a.ProdSpecId AND aa.[Status] = 'A'
			                            ORDER BY aa.SortNumber FOR XML PATH(''))
		                            --, CASE WHEN b.[Status] = 'A' THEN '●' ELSE '○' END + ISNULL(b.[Description], '')
		                            )
	                            ) y (col, val)
	                            WHERE d.RfiId IN @RfiIds 
                                AND d.RfiNo <> 'RFI20240315001'
                            ) t 
                            PIVOT (MAX(t.val) FOR t.col IN (" + columns + ")) p";
                    dynamicParameters.Add("RfiIds", result.Select(s => s.RfiId).Distinct().ToArray());
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

        #region //GetDesignToExcel --新設計申請單報表資料 -- Chia Yuan -- 2024.05.22
        public string GetDesignToExcel(int DesignId, string DesignNo, int ProTypeGroupId, int ProductUseId, int SalesId, string StartDate, string EndDate, string SignStatus, string DesignStatus, string FlowStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得市場評估單資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DesignId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignNo";
                    sqlQuery.mainTables =
                        @"FROM RFI.DemandDesign a
                        INNER JOIN RFI.RequestForInformation b ON b.RfiId = a.RfiId
                        OUTER APPLY (
	                        SELECT *
	                        FROM (
		                        SELECT ISNULL(aa.SignUser, aa.FlowUser) AS Flower, aa.DesignSfId, aa.SignJobName, aa.SortNumber, ISNULL(FORMAT(aa.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '') AS ArriveDateTime, aa.Approved
		                        --, ROW_NUMBER() OVER (PARTITION BY aa.DesignId ORDER BY aa.SortNumber) AS SortRow
		                        --虛擬狀態 W=等待審批, I=進行中, D=否決, C=核准(完成)
		                        , CASE ISNULL(aa.ArriveDateTime, '') WHEN '' THEN 'W' ELSE (CASE ISNULL(aa.Approved, '') WHEN '' THEN 'I' WHEN 'N' THEN 'D' ELSE 'C' END) END AS SignStatus
                                , aa.[Status] AS FlowStatus, ab.StatusName AS FlowStatusName
		                        FROM RFI.DesignSignFlow aa
                                INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status] AND StatusSchema = 'RfiDesign.FlowStatus'
		                        WHERE aa.DesignId = a.DesignId
	                        ) x WHERE x.SignStatus IN ('I', 'D')
                        ) flow";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND (EXISTS (
                                            SELECT TOP 1 1
                                            FROM RFI.DesignSignFlow da
                                            OUTER APPLY (
                                                SELECT DISTINCT ISNULL(ab.SignUser, ab.FlowUser) AS FlowUser
                                                FROM RFI.RfiDetail aa
                                                INNER JOIN RFI.RfiSignFlow ab ON ab.RfiDetailId = aa.RfiDetailId
	                                            WHERE aa.RfiId = a.RfiId
                                            ) AS db
                                            WHERE da.DesignId = a.DesignId
                                            AND (ISNULL(da.SignUser, da.FlowUser) = @Flower OR db.FlowUser = @Flower)) OR a.SalesId = @Flower)";

                    dynamicParameters.Add("Flower", CurrentUser);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", " AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignNo", " AND a.DesignNo LIKE '%' + @DesignNo + '%'", DesignNo.Trim());
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesId", @" AND a.SalesId = @SalesId", SalesId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND b.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND b.ProductUseId = @ProductUseId", ProductUseId);
                    if (DateTime.TryParse(StartDate, out DateTime startDate))
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", " AND a.CreateDate >= @StartDate", startDate.ToString("yyyy-MM-dd 00:00:00"));
                    if (DateTime.TryParse(EndDate, out DateTime endDate))
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", " AND a.CreateDate <= @EndDate", endDate.ToString("yyyy-MM-dd 23:59:59"));

                    //待審批
                    if (SignStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SignStatus", @" AND flow.SignStatus IN @SignStatus", SignStatus.Split(','));
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Flower", @" AND flow.Flower = @Flower", CurrentUser);
                    }

                    //關卡狀態
                    if (FlowStatus.Length > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowStatus", @" AND flow.FlowStatus IN @FlowStatus", FlowStatus.Split(','));

                    //單據狀態
                    if (DesignStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignStatus", @" AND a.Status IN @DesignStatus", DesignStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DesignId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    if (result.Count() <= 0) throw new SystemException("無法取得新設計申請單資料!");
                    #endregion

                    #region //取得欄位資料
                    sql = @"DECLARE @tempSensor TABLE (FeatureName NVARCHAR(50))
                            INSERT @tempSensor
                            SELECT DISTINCT ag.FeatureName 
                            FROM RFI.DemandDesign aa
                            INNER JOIN RFI.RequestForInformation ad ON ad.RfiId = aa.RfiId
                            INNER JOIN RFI.RfiDetail ae ON ae.RfiId = ad.RfiId
                            INNER JOIN RFI.ProductSpec af ON af.RfiDetailId = ae.RfiDetailId
                            INNER JOIN RFI.ProductSpecDetail ag ON ag.ProdSpecId = af.ProdSpecId AND ag.[Status] = af.[Status]
                            WHERE ag.[Status] = 'A'
                            --AND af.SpecEName LIKE N'%' + @SpecEName + '%'
                            DECLARE @tempDesign TABLE (col NVARCHAR(255), mSort INT, dSort INT, sortNum INT)
                            INSERT @tempDesign
                            SELECT DISTINCT a.ParameterName,a.SortNumber,b.SortNumber
                            , ROW_NUMBER() OVER (PARTITION BY a.ParameterName ORDER BY a.SortNumber,b.SortNumber, d.RfiId DESC)
                            FROM RFI.DemandDesignSpec a
                            INNER JOIN RFI.DemandDesignSpecDetail b ON b.DdSpecId = a.DdSpecId and a.[Status] = b.[Status]
                            INNER JOIN RFI.DemandDesign d ON d.DesignId = a.DesignId
                            WHERE b.[Status] = 'A'
                            AND a.DesignId IN @DesignIds
                            AND EXISTS (
	                            SELECT TOP 1 1
	                            FROM RFI.RequestForInformation aa
	                            INNER JOIN RFI.RfiDetail ab ON ab.RfiId = aa.RfiId
	                            INNER JOIN RFI.ProductSpec ac ON ac.RfiDetailId = ab.RfiDetailId
	                            INNER JOIN RFI.ProductSpecDetail ad ON ad.ProdSpecId = ac.ProdSpecId
	                            WHERE ab.[Status] = '5'
	                            AND aa.RfiId = d.RfiId
	                            AND ad.FeatureName IN (SELECT * FROM @tempSensor)
                            )

                            DECLARE @columns VARCHAR(MAX)
                            SELECT
                                @columns = CASE
                                    WHEN @columns IS NULL
                                    THEN '[' + T.col + ']'
                                    ELSE @columns + ',[' + T.col + ']'
                                END
                            FROM @tempDesign T
                            WHERE sortNum = 1
                            ORDER BY mSort, col
                            SELECT @columns AS [Columns]";
                    var resultDesignSpec = sqlConnection.QueryFirstOrDefault(sql, new { DesignIds = result.Select(s => s.DesignId).Distinct().ToArray() }) ?? throw new SystemException("沒有可用的規格項目!");
                    var columns = resultDesignSpec.Columns;
                    #endregion

                    if (!string.IsNullOrWhiteSpace(columns))
                    {
                        #region //取得主要資料
                        sql = @"SELECT *
                                FROM (
	                                SELECT c.DesignNo AS 需求單號, c.Edition AS 版次
                                    , u.UserName AS 業務擔當, FORMAT(c.CreateDate, 'yyyy-MM-dd HH:mm') AS 建單時間, s.StatusName AS 單據狀態
                                    , g.ProTypeGroupName AS 產品類別, e.ProductUseName AS 產品應用, s2.StatusName AS 是否鍍膜, c.AssemblyName AS 設計代號
                                    , y.col,SUBSTRING(y.val, 0, LEN(y.val)) AS val
	                                FROM RFI.DemandDesignSpecDetail a
	                                INNER JOIN RFI.DemandDesignSpec b ON b.DdSpecId = a.DdSpecId
	                                INNER JOIN RFI.DemandDesign c ON c.DesignId = b.DesignId
	                                INNER JOIN RFI.RequestForInformation d ON d.RfiId = c.RfiId
                                    INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = d.ProTypeGroupId
	                                INNER JOIN SCM.ProductUse e ON e.ProductUseId = d.ProductUseId
                                    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = g.CoatingFlag AND s2.StatusSchema = 'Boolean'
	                                INNER JOIN BAS.[Status] s ON s.StatusNo = c.[Status] AND s.StatusSchema = 'RfiDesign.Status'
	                                INNER JOIN BAS.[User] u ON u.UserId = c.SalesId
	                                CROSS APPLY (
		                                VALUES (b.ParameterName
		                                , (SELECT aa.PmtDetailName + 
			                                CASE WHEN ISNULL(aa.DataType, '') = 'O' THEN ':' 
			                                WHEN b.ControlType = 'A' OR b.ControlType = 'T' THEN ':' + aa.DesignInput
			                                ELSE '' END + CASE WHEN a.Description IS NULL THEN '' ELSE ' (' + aa.[Description] + ')'  END + ','
			                                FROM RFI.DemandDesignSpecDetail aa 
			                                WHERE aa.DdSpecId = a.DdSpecId AND aa.[Status] = 'A'
			                                ORDER BY aa.SortNumber FOR XML PATH(''))
		                                --, CASE WHEN b.[Status] = 'A' THEN '●' ELSE '○' END + ISNULL(b.[Description], '')
		                                )
	                                ) y (col, val)
	                                --CROSS APPLY (
		                            --    VALUES (a.PmtDetailName, ISNULL(a.DesignInput, '') + CASE WHEN a.Description IS NULL THEN '' ELSE '(' + ISNULL(a.Description, '') + ')' END)
	                                --) y (col, val)
                                    WHERE c.DesignId IN @DesignIds
                                ) t 
                                PIVOT (MAX(t.val) FOR t.col IN (" + columns + ")) p";
                        result = sqlConnection.Query(sql, new { DesignIds = result.Select(s => s.DesignId).Distinct().ToArray() });
                        #endregion
                    }

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
        #region //AddProdTerminal -- 終端資料新增 -- Chia Yuan -- 2023.12.26
        public string AddProdTerminal(string TerminalName, string Status)
        {
            try
            {
                if (TerminalName.Length <= 0) throw new SystemException("【終端名稱】不能為空!");
                if (TerminalName.Length > 50) throw new SystemException("【終端名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string terminalName = TerminalName.Trim();

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //判斷終端名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdTerminal a
                                WHERE a.TerminalName = @TerminalName";
                        dynamicParameters.Add("TerminalName", terminalName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【終端名稱】資料重複!");
                        #endregion

                        #region //終端資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO RFI.ProdTerminal(TerminalName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProdTerminalId
                                VALUES (@TerminalName, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TerminalName = terminalName,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += 1;
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

        #region //AddProdSystem -- 系統資料新增 -- Chia Yuan -- 2023.12.26
        public string AddProdSystem(string SystemName, string Status)
        {
            try
            {
                if (SystemName.Length <= 0) throw new SystemException("【系統名稱】不能為空!");
                if (SystemName.Length > 50) throw new SystemException("【系統名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string systemName = SystemName.Trim();

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //判斷系統名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdSystem a
                                WHERE a.SystemName = @SystemName";
                        dynamicParameters.Add("SystemName", systemName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【系統名稱】資料重複!");
                        #endregion

                        #region //系統資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO RFI.ProdSystem(SystemName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProdSysId
                                VALUES (@SystemName, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemName = systemName,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += 1;
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

        #region //AddProdModule -- 模組資料新增 -- Chia Yuan -- 2023.12.26
        public string AddProdModule(string ModuleName, string Status)
        {
            try
            {
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 50) throw new SystemException("【模組名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string moduleName = ModuleName.Trim();

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdModule a
                                WHERE a.ModuleName = @ModuleName";
                        dynamicParameters.Add("ModuleName", moduleName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【模組名稱】資料重複!");
                        #endregion

                        #region //模組資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO RFI.ProdModule(ModuleName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProdModId
                                VALUES (@ModuleName, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModuleName = moduleName,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += 1;
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

        #region //AddTemplateProdSpec -- 評估樣板主資料新增 -- Chia Yuan -- 2023.12.26
        public string AddTemplateProdSpec(int ProTypeGroupId,string SpecName, string SpecEName, string ControlType, string Required, string Status)
        {
            try
            {
                if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("狀態資料錯誤!");

                if (SpecName.Length <= 0) throw new SystemException("【評估項目名稱】不能為空!");
                if (SpecName.Length > 100) throw new SystemException("【評估項目名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(SpecEName))
                {
                    if (SpecName.Length > 100) throw new SystemException("【規格(英文)種類名稱】長度錯誤!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        SpecName = SpecName.Trim();

                        if (!string.IsNullOrWhiteSpace(SpecEName)) SpecEName = SpecEName.Trim();

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //判斷控制類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a 
                                WHERE a.TypeSchema = 'RfiFormItem.ItemType'
                                AND a.TypeNo = @ControlType";
                        dynamicParameters.Add("ControlType", ControlType);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【控制類別】資料錯誤!");
                        #endregion

                        #region //判斷類別群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("Status", "A");
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別群組】資料錯誤!");
                        #endregion

                        #region //判斷評估項目名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SpecName
                                FROM RFI.TemplateProdSpec a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any(a => a.SpecName == SpecName)) throw new SystemException("【設計項目名稱】資料重複!");
                        #endregion

                        #region //取得目前最大排序號碼
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
                                FROM RFI.TemplateProdSpec
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        int maxSortNumber = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }).MaxSortNumber;
                        #endregion

                        #region //評估項目資料新增
                        sql = @"INSERT INTO RFI.TemplateProdSpec(ProTypeGroupId, SpecName, SpecEName, SortNumber, ControlType
                                , Required, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TempProdSpecId
                                VALUES (@ProTypeGroupId, @SpecName, @SpecEName, @SortNumber, @ControlType
                                , @Required, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                ProTypeGroupId,
                                SpecName,
                                SpecEName = string.IsNullOrWhiteSpace(SpecEName) ? null : SpecEName,
                                SortNumber = maxSortNumber + 1,
                                ControlType,
                                Required,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        rowsAffected += 1;
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

        #region //AddTemplateSpecParameter -- 設計樣板主資料新增 -- Chia Yuan -- 2024.05.24
        public string AddTemplateSpecParameter(int ProTypeGroupId, string ParameterName, string ParameterEName, string ControlType, string Required, string Status)
        {
            try
            {
                if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("狀態資料錯誤!");

                if (ParameterName.Length <= 0) throw new SystemException("【設計項目名稱】不能為空!");
                if (ParameterName.Length > 100) throw new SystemException("【設計項目名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(ParameterEName))
                {
                    if (ParameterName.Length > 100) throw new SystemException("【設計項目名稱(英文)】長度錯誤!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        ParameterName = ParameterName.Trim();

                        if (!string.IsNullOrWhiteSpace(ParameterEName)) ParameterEName = ParameterEName.Trim();

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //判斷控制類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a 
                                WHERE a.TypeSchema = 'RfiFormItem.ItemType'
                                AND a.TypeNo = @ControlType";
                        dynamicParameters.Add("ControlType", ControlType);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【控制類別】資料錯誤!");
                        #endregion

                        #region //判斷類別群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("Status", "A");
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品群組】資料錯誤!");
                        #endregion

                        #region //判斷設計項目是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.ParameterName = @ParameterName";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("ParameterName", ParameterName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【設計項目名稱】資料重複!");
                        #endregion

                        #region //取得目前最大排序號碼
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
                                FROM RFI.TemplateSpecParameter
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        int maxSortNumber = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }).MaxSortNumber;
                        #endregion

                        #region //評估項目資料新增
                        sql = @"INSERT INTO RFI.TemplateSpecParameter(ProTypeGroupId, ParameterName, ParameterEName, SortNumber, ControlType
                                , Required, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TempSpId
                                VALUES (@ProTypeGroupId, @ParameterName, @ParameterEName, @SortNumber, @ControlType
                                , @Required, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                ProTypeGroupId,
                                ParameterName,
                                ParameterEName = string.IsNullOrWhiteSpace(ParameterEName) ? null : ParameterEName,
                                SortNumber = maxSortNumber + 1,
                                ControlType,
                                Required,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        rowsAffected += 1;
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

        #region //AddTemplateProdSpecDetail -- 評估樣板子資料新增 -- Chia Yuan -- 2024.01.02
        public string AddTemplateProdSpecDetail(int TempProdSpecId, string FeatureName, string DataType, string Description, string Status)
        {
            try
            {
                if (FeatureName.Length <= 0) throw new SystemException("【特徵名稱】不能為空!");
                if (FeatureName.Length > 50) throw new SystemException("【特徵名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    if (FeatureName.Length > 50) throw new SystemException("【描述】長度錯誤!");
                }
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    Description = Description.Trim();
                }
                else {
                    Description = null;
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        FeatureName = FeatureName.Trim();

                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        #endregion

                        #region //判斷特徵名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.FeatureName = @FeatureName
                                AND a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("FeatureName", FeatureName);
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【特徵名稱】資料重複!");
                        #endregion

                        if (!string.IsNullOrWhiteSpace(DataType))
                        {
                            #region //判斷規格特徵資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type] a 
                                    WHERE a.TypeSchema = 'RfiFormItem.OtherType'
                                    AND a.TypeNo = @DataType";
                            dynamicParameters.Add("DataType", DataType);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【規格特徵】資料錯誤!");
                            #endregion
                        }

                        #region //取得目前最大排序號碼
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
                                FROM RFI.TemplateProdSpecDetail
                                WHERE TempProdSpecId = @TempProdSpecId";
                        int maxSortNumber = sqlConnection.QueryFirstOrDefault(sql, new { TempProdSpecId }).MaxSortNumber;
                        #endregion

                        #region //特徵資料新增
                        sql = @"INSERT INTO RFI.TemplateProdSpecDetail(TempProdSpecId, FeatureName, SortNumber, DataType, Description, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TempPsDetailId
                                VALUES (@TempProdSpecId, @FeatureName, @SortNumber, @DataType, @Description, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                TempProdSpecId,
                                FeatureName,
                                SortNumber = maxSortNumber + 1,
                                DataType = string.IsNullOrWhiteSpace(DataType) ? null : DataType,
                                Description = string.IsNullOrWhiteSpace(Description) ? null : Description,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        rowsAffected += 1;
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

        #region //AddTemplateSpecParameterDetail -- 設計樣板子資料新增 -- Chia Yuan -- 2024.01.02
        public string AddTemplateSpecParameterDetail(int TempSpId, string PmtDetailName, string DataType, string Description, string Status)
        {
            try
            {
                if (PmtDetailName.Length <= 0) throw new SystemException("【設計選項名稱】不能為空!");
                if (PmtDetailName.Length > 50) throw new SystemException("【設計選項名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    if (PmtDetailName.Length > 50) throw new SystemException("【描述】長度錯誤!");
                }
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    Description = Description.Trim();
                }
                else
                {
                    Description = null;
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        PmtDetailName = PmtDetailName.Trim();

                        #region //判斷設計項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【設計項目】資料錯誤!");
                        #endregion

                        #region //判斷設計選項是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.PmtDetailName = @PmtDetailName
                                AND a.TempSpId = @TempSpId";
                        dynamicParameters.Add("PmtDetailName", PmtDetailName);
                        dynamicParameters.Add("TempSpId", TempSpId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【設計選項】資料重複!");
                        #endregion

                        if (!string.IsNullOrWhiteSpace(DataType))
                        {
                            #region //判斷選項資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type] a 
                                    WHERE a.TypeSchema = 'RfiFormItem.OtherType'
                                    AND a.TypeNo = @DataType";
                            dynamicParameters.Add("DataType", DataType);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【設計選項】資料錯誤!");
                            #endregion
                        }

                        #region //取得目前最大排序號碼
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
                                FROM RFI.TemplateSpecParameterDetail
                                WHERE TempSpId = @TempSpId";
                        int maxSortNumber = sqlConnection.QueryFirstOrDefault(sql, new { TempSpId }).MaxSortNumber;
                        #endregion

                        #region //主資料新增
                        sql = @"INSERT INTO RFI.TemplateSpecParameterDetail(TempSpId, PmtDetailName, SortNumber, DataType, Description, Required, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TempSpDetailId
                                VALUES (@TempSpId, @PmtDetailName, @SortNumber, @DataType, @Description, @Required, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                TempSpId,
                                PmtDetailName,
                                SortNumber = maxSortNumber + 1,
                                DataType = string.IsNullOrWhiteSpace(DataType) ? null : DataType,
                                Description = string.IsNullOrWhiteSpace(Description) ? null : Description,
                                Required = 0,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        rowsAffected += 1;
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

        #region //AddRequestForInformation -- RFI單頭新增 -- Chia Yuan -- 2024.01.05
        public string AddRequestForInformation(string RfiNo, int ProTypeGroupId, int ProductUseId, int SalesId, string ExistFlag
            , string CustomerName, string CustomerEName, int CustomerId, string OrganizaitonType, string Status)
        {
            try
            {
                #region //判斷資料長度
                if (!Regex.IsMatch(ExistFlag, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【是否存在】注記錯誤!");
                if (string.IsNullOrWhiteSpace(CustomerName)) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerName.Length > 100) throw new SystemException("【客戶名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(CustomerEName))
                {
                    if (CustomerEName.Length > 100) throw new SystemException("【客戶名稱(英文)】長度錯誤!");
                    CustomerEName = CustomerEName.Trim();
                }
                else
                {
                    CustomerEName = null;
                }
                #endregion

                //bool autoAssemblyNum = true; //預設自動編碼
                //ex: ETG2401X
                //if (!autoAssemblyNum) if (!Regex.IsMatch(, @"^[A-Z]{3}\d{2}([A-Z]|\d){1}\d{1}(Y|X){1}")) throw new SystemException("【設計代號】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        CustomerName = CustomerName.Trim();

                        #region //單號取號
                        sql = @"SELECT TOP 1 CONVERT(int, RIGHT(ISNULL(MAX(RfiNo), '000'), 3)) + 1 CurrentNum
                                FROM RFI.RequestForInformation
                                WHERE RfiNo LIKE @RfiNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { RfiNo = string.Format("{0}{1}___", "RFI", DateTime.Now.ToString("yyyyMMdd")) });
                        int currentNum = result.CurrentNum;
                        string rfiNo = string.Format("{0}{1}{2}", "RFI", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        #region //機種名稱取號 (停用)
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT MAX(SUBSTRING(AssemblyName, 1, LEN(AssemblyName)-1)) AS CurrentAssemblyNum
                        //        FROM RFI.RequestForInformation";
                        //resultFirst = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        //string currentAssemblyNum = resultFirst.CurrentAssemblyNum;
                        #endregion

                        #region //判斷產品應用資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse a
                                WHERE a.ProductUseId = @ProductUseId
                                AND a.Status = 'A'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProductUseId }) ?? throw new SystemException("【產品應用】資料錯誤!");
                        #endregion

                        #region //判斷產品類別群組資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = 'A'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }) ?? throw new SystemException("【產品類別群組】資料錯誤!");
                        #endregion

                        #region //判斷客戶類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a
                                WHERE a.TypeNo = @OrganizaitonType
                                AND a.TypeSchema = 'MemberOrganization.OrganizaitonType'";            
                        result = sqlConnection.QueryFirstOrDefault(sql, new { OrganizaitonType }) ?? throw new SystemException("【客戶類型】資料錯誤!");
                        #endregion

                        #region //判斷存在資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a
                                WHERE a.StatusNo = @ExistFlag
                                AND a.StatusSchema = 'Boolean'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ExistFlag }) ?? throw new SystemException("【是否存在】資料錯誤!");
                        #endregion

                        #region //RFI單頭資料新增
                        sql = @"INSERT INTO RFI.RequestForInformation (RfiNo, ProTypeGroupId, ProductUseId, SalesId, ExistFlag
                                , CustomerName, CustomerEName, CustomerId, OrganizaitonType, AnalysisDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfiId
                                VALUES (@RfiNo, @ProTypeGroupId, @ProductUseId, @SalesId, @ExistFlag
                                , @CustomerName, @CustomerEName, @CustomerId, @OrganizaitonType, @AnalysisDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.Query(sql,
                            new
                            {
                                RfiNo = rfiNo,
                                ProTypeGroupId,
                                ProductUseId,
                                SalesId = CreateBy,
                                ExistFlag,
                                CustomerName,
                                CustomerEName,
                                CustomerId = CustomerId <= 0 ? (int?)null : CustomerId,
                                OrganizaitonType,
                                AnalysisDate = CreateDate,
                                Status = "0", //新單據
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected += 1;
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

        #region //AddRfiDetail -- RFI單身新增 -- Chia Yuan -- 2024.01.10
        public string AddRfiDetail(int RfiId, string RfiSequence, int SalesId, string ProdDate, string ProdLifeCycleStart, string ProdLifeCycleEnd
            , decimal TargetUnitPrice, int LifeCycleQty, decimal Revenue, decimal RevenueFC
            , string Currency, decimal ExchangeRate, int AdditionalFile
            , string ProdSpecs, string ProdSpecDetails)
        {
            try
            {
                #region //判斷資料長度
                if (string.IsNullOrWhiteSpace(ProdDate)) throw new SystemException("【量產時間】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleStart)) throw new SystemException("【產品生命起日】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleEnd)) throw new SystemException("【產品生命迄日】不能為空!");
                if (TargetUnitPrice <= 0) throw new SystemException("【目標單價】資料錯誤!");
                if (LifeCycleQty <= 0) throw new SystemException("【預估數量】資料錯誤!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】資料錯誤!");
                if (string.IsNullOrWhiteSpace(Currency)) throw new SystemException("【貨幣】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdSpecs)) throw new SystemException("【評估項目】資料錯誤!");
                if (string.IsNullOrWhiteSpace(ProdSpecDetails)) throw new SystemException("【規格特徵】資料錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        int unitRound = 0, totalRound = 0, unitRoundFC = 0, totalRoundFC = 0;

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得共用參數
                            sql = @"SELECT TOP 1 MA003 LocalCurrency FROM CMSMA";
                            var erpResult = erpConnection.QueryFirstOrDefault(sql) ?? throw new SystemException(string.Format("【ERP共用參數檔】資料錯誤!"));
                            #endregion

                            string LocalCurrency = erpResult.LocalCurrency;

                            #region //目前匯率
                            sql = @"SELECT TOP 1 MG003, MG004, MG005, MG006
                                    FROM CMSMG
                                    WHERE MG001 = @Currency
                                    AND MG002 <= @DateNow
                                    ORDER BY MG002 DESC";
                            erpResult = erpConnection.QueryFirstOrDefault(sql, new { Currency, DateNow = DateTime.Now.ToString("yyyyMMdd") }) ?? throw new SystemException(string.Format("【ERP交易幣別匯率】{0}不存在!", Currency));

                            double exchangeRate = Convert.ToDouble(erpResult.MG004);

                            #region //停用
                            //switch (exchangeRateSource)
                            //{
                            //    case "I": //銀行買進匯率
                            //        exchangeRate = Convert.ToDouble(item.MG003);
                            //        break;
                            //    case "O": //銀行賣出匯率
                            //        exchangeRate = Convert.ToDouble(item.MG004);
                            //        break;
                            //    case "E": //報關買進匯率
                            //        exchangeRate = Convert.ToDouble(item.MG005);
                            //        break;
                            //    case "W": //報關賣出匯率
                            //        exchangeRate = Convert.ToDouble(item.MG006);
                            //        break;
                            //}
                            #endregion
                            #endregion

                            #region //交易幣別設定
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MF001)) AS MF001, LTRIM(RTRIM(MF002)) AS MF002, MF003, MF004, MF005, MF006
                                    FROM CMSMF
                                    WHERE MF001 = @Currency";
                            erpResult = erpConnection.QueryFirstOrDefault(sql, new { Currency = LocalCurrency }) ?? throw new SystemException(string.Format("【ERP交易幣別設定】{0} 不存在!", LocalCurrency));
                            unitRound = Convert.ToInt32(erpResult.MF003); //單價取位
                            totalRound = Convert.ToInt32(erpResult.MF004); //金額取位
                            #endregion

                            #region //交易幣別設定(外幣)
                            erpResult = erpConnection.QueryFirstOrDefault(sql, new { Currency }) ?? throw new SystemException(string.Format("【ERP交易幣別設定】{0} 不存在!", Currency));
                            unitRoundFC = Convert.ToInt32(erpResult.MF003); //單價取位
                            totalRoundFC = Convert.ToInt32(erpResult.MF004); //金額取位
                            #endregion
                        }

                        #region //判斷RFI單頭資料是否正確
                        sql = @"SELECT TOP 1 a.RfiNo, a.[Status] RfiStatus, s2.StatusName RfiStatusName
                                , s.StatusName ExistFlagName
                                , a.CustomerName, a.CustomerEName
                                , a.ProTypeGroupId, g.ProTypeGroupName
                                , a.ProductUseId, b.ProductUseName, b.ProductUseNo
                                , c.UserName SalesName, c.UserNo SalesNo
                                , ISNULL(FORMAT(a.CreateDate, 'yyyy-MM-dd'), '') AnalysisDate
                                FROM RFI.RequestForInformation a
                                INNER JOIN SCM.ProductUse b ON b.ProductUseId = a.ProductUseId
                                INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = a.ProTypeGroupId
                                INNER JOIN BAS.[User] c ON c.UserId = a.SalesId
                                INNER JOIN BAS.[Status] s ON s.StatusNo = a.ExistFlag AND s.StatusSchema = 'Boolean'
                                INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.[Status] AND s2.StatusSchema = 'RequestForInformation.Status'
                                WHERE a.RfiId = @RfiId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId }) ?? throw new SystemException("【市場評估單】資料錯誤!");
                        int ProTypeGroupId = result.ProTypeGroupId;
                        string rfiStatus = result.RfiStatus;
                        string rfiStatusName = result.RfiStatusName;
                        string salesName = result.SalesName;
                        string customerName = result.CustomerName;
                        string customerEName = result.CustomerEName;
                        string proTypeGroupName = result.ProTypeGroupName;
                        string productUseName = result.ProductUseName;
                        string productUseNo = result.ProductUseNo;
                        string rfiNo = result.RfiNo;
                        string existFlagName = result.ExistFlagName;
                        string analysisDate = result.AnalysisDate;
                        #endregion

                        #region //判斷送單人資料是否正確
                        sql = @"SELECT TOP 1 DepartmentId FROM BAS.[User] a WHERE a.UserId = @CreateBy";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CreateBy }) ?? throw new SystemException("【送單人】資料錯誤!");
                        int DepartmentId = result.DepartmentId;
                        #endregion

                        #region //判斷貨幣資料是否正確
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.[Type] a 
                                WHERE a.TypeNo = @Currency
                                AND a.TypeSchema = 'ExchangeRate.Currency'";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { Currency }) ?? throw new SystemException("【貨幣】資料錯誤!");
                        #endregion

                        #region //取得審批審批流程
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateRfiSignFlow a
                                WHERE EXISTS (SELECT TOP 1 1 FROM SCM.ProductTypeGroup aa WHERE aa.ProTypeGroupId = a.ProTypeGroupId AND aa.[Status] = 'A')
                                AND a.ProTypeGroupId = @ProTypeGroupId
                                AND a.DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, DepartmentId }) ?? throw new SystemException("【市場評估單審批流程】資料錯誤!");
                        #endregion

                        #region //單身目前序號
                        sql = @"SELECT TOP 1 ISNULL(CAST(MAX(RfiSequence) AS INT), 0) + 1 MaxSequence
                                FROM RFI.RfiDetail
                                WHERE RfiId = @RfiId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                        int maxSequence = Convert.ToInt32(result.MaxSequence);
                        #endregion

                        #region //單身目前版次號
                        sql = @"SELECT TOP 1 ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
                                FROM RFI.RfiDetail
                                WHERE RfiId = @RfiId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                        int maxEdition = Convert.ToInt32(result.MaxEdition);
                        #endregion

                        if (!DateTime.TryParse(ProdDate, out DateTime prodDate)) throw new SystemException("【量產時間】資料錯誤!");
                        if (!DateTime.TryParse(ProdLifeCycleStart, out DateTime prodLifeCycleStart)) throw new SystemException("【產品生命起日】資料錯誤!");
                        if (!DateTime.TryParse(ProdLifeCycleEnd, out DateTime prodLifeCycleEnd)) throw new SystemException("【產品生命迄日】資料錯誤!");

                        #region //RFI單身資料新增
                        sql = @"INSERT INTO RFI.RfiDetail (RfiId, RfiSequence, CompanyId, SalesId, ProdDate, ProdLifeCycleStart, ProdLifeCycleEnd
                                , TargetUnitPrice, LifeCycleQty, Revenue, RevenueFC, Currency, ExchangeRate, AdditionalFile
                                , Edition, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfiDetailId
                                VALUES (@RfiId, @RfiSequence, @CompanyId, @SalesId, @ProdDate, @ProdLifeCycleStart, @ProdLifeCycleEnd
                                , @TargetUnitPrice, @LifeCycleQty, @Revenue, @RevenueFC, @Currency, @ExchangeRate, @AdditionalFile
                                , @Edition, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                            new
                            {
                                RfiId,
                                RfiSequence = string.Format("{0:0000}", maxSequence), //取4位數
                                CompanyId = CurrentCompany,
                                SalesId = CreateBy,
                                ProdDate = prodDate.ToString("yyyy-MM-dd"),
                                ProdLifeCycleStart = prodLifeCycleStart.ToString("yyyy-MM-dd"),
                                ProdLifeCycleEnd = prodLifeCycleEnd.ToString("yyyy-MM-dd"),
                                TargetUnitPrice,
                                LifeCycleQty,
                                Revenue = Math.Round(TargetUnitPrice * LifeCycleQty * ExchangeRate, totalRound, MidpointRounding.AwayFromZero),
                                RevenueFC = Math.Round(TargetUnitPrice * LifeCycleQty, totalRoundFC, MidpointRounding.AwayFromZero),
                                Currency,
                                ExchangeRate,
                                AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                                Edition = string.Format("{0:00}", maxEdition), //取2位數
                                Status = "1", //24.03.22 業務反映於建單後要可返回修改，確認無誤後再送出，因此改為預設狀態1
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected += 1;
                        int RfiDetailId = insertResult?.RfiDetailId ?? -1;
                        #endregion

                        #region //RFI單頭狀態更新
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.RequestForInformation a
                                WHERE a.RfiId = @RfiId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                RfiId,
                                Status = "1", //審批中
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        if (!ProdSpecs.TryParseJson(out JObject prodSpecs)) throw new SystemException("【評估項目】資料錯誤!");
                        if (!ProdSpecDetails.TryParseJson(out JObject prodSpecDetails)) throw new SystemException("【規格特徵】資料錯誤!");
                        if (prodSpecs["data"].Count() <= 0) throw new SystemException("【評估項目】資料必須維護!");
                        if (prodSpecDetails["data"].Count() <= 0) throw new SystemException("【規格特徵】資料必須維護!");

                        if (prodSpecs["data"].Count() > 0)
                        {
                            var TempProdSpecIds = prodSpecs["data"].Select(s => { return Convert.ToInt32(s["TempProdSpecId"]); }).ToArray();

                            #region //取得評估項目資料
                            sql = @"SELECT a.TempProdSpecId, a.SpecName, a.SpecEName, a.SortNumber, a.ControlType, a.[Required]
                                    FROM RFI.TemplateProdSpec a 
                                    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.TempProdSpecId IN @TempProdSpecIds
                                    ORDER BY a.SortNumber";
                            var resultProdSpec = sqlConnection.Query(sql, new { ProTypeGroupId, TempProdSpecIds });
                            #endregion

                            #region //取得規格特徵資料
                            sql = @"SELECT a.TempPsDetailId, a.FeatureName, a.SortNumber, a.DataType, a.[Description]
                                    FROM RFI.TemplateProdSpecDetail a 
                                    WHERE a.TempProdSpecId IN @TempProdSpecIds
                                    ORDER BY a.SortNumber";
                            var resultProdSpecDetail = sqlConnection.Query(sql, new { TempProdSpecIds });
                            #endregion

                            #region //判斷狀態資料是否正確
                            sql = @"SELECT a.StatusNo
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo IN @Status";
                            var resultStatus = sqlConnection.Query(sql, new { Status = prodSpecDetails["data"].Select(s => { return s["Status"].ToString(); }).Distinct().ToArray() });
                            if (resultStatus.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                            #endregion

                            foreach (var data in prodSpecs["data"])
                            {
                                if (!int.TryParse(data["TempProdSpecId"].ToString(), out int tempProdSpecId)) throw new SystemException("【評估項目】資料錯誤!");
                                var templateProdSpec = resultProdSpec.FirstOrDefault(f => f.TempProdSpecId == tempProdSpecId) ?? throw new SystemException("【評估項目】資料錯誤!");

                                #region //評估項目新增
                                sql = @"INSERT INTO RFI.ProductSpec (RfiDetailId, SpecName, SpecEName, SortNumber, ControlType, Required, Status	
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.ProdSpecId
                                        VALUES (@RfiDetailId, @SpecName, @SpecEName, @SortNumber, @ControlType, @Required, @Status	
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                    new
                                    {
                                        RfiDetailId,
                                        templateProdSpec.SpecName,
                                        templateProdSpec.SpecEName,
                                        templateProdSpec.SortNumber,
                                        templateProdSpec.ControlType,
                                        templateProdSpec.Required,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += 1;
                                int ProdSpecId = insertResult?.ProdSpecId ?? -1;
                                #endregion

                                foreach (var detail in prodSpecDetails["data"].Where(w => w["TempProdSpecId"].ToString() == tempProdSpecId.ToString()))
                                {
                                    if (!int.TryParse(detail["TempPsDetailId"].ToString(), out int tempPsDetailId)) throw new SystemException("【規格特徵】資料錯誤!");
                                    if (resultStatus.FirstOrDefault(f => f.StatusNo == detail["Status"].ToString()) == null) throw new SystemException("【狀態】資料錯誤!");
                                    var tempProdSpecDetail = resultProdSpecDetail.FirstOrDefault(f => f.TempPsDetailId == tempPsDetailId) ?? throw new SystemException("【規格特徵】資料錯誤!");

                                    #region //規格特徵新增
                                    sql = @"INSERT INTO RFI.ProductSpecDetail (ProdSpecId, FeatureName, SortNumber, DataType, Description, Status
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PsDetailId
                                            VALUES (@ProdSpecId, @FeatureName, @SortNumber, @DataType, @Description, @Status
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                        new
                                        {
                                            ProdSpecId,
                                            FeatureName = detail["DataType"].ToString() == "O" ? detail["FeatureName"].ToString() : tempProdSpecDetail.FeatureName,
                                            tempProdSpecDetail.SortNumber,
                                            DataType = string.IsNullOrWhiteSpace(detail["DataType"].ToString()) ? null : detail["DataType"].ToString().Trim(),
                                            Description = string.IsNullOrWhiteSpace(detail["Description"].ToString()) ? null : detail["Description"].ToString().Trim(),
                                            Status = detail["Status"].ToString(),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    rowsAffected += 1;
                                    int PsDetailId = insertResult?.PsDetailId ?? -1;
                                    #endregion
                                }
                            }
                        }

                        #region //單身目前版次號
                        sql = @"SELECT TOP 1 ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
                                FROM RFI.RfiSignFlow
                                WHERE RfiDetailId = @RfiDetailId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId });
                        int maxEdition1 = Convert.ToInt32(result.MaxEdition);
                        #endregion

                        #region //建立審批流程檔 --new
                        sql = @"INSERT RFI.RfiSignFlow
                                SELECT @RfiDetailId, ISNULL(a.FlowUser, @FlowUser), NULL AS AdditionalFile
                                , CASE a.SortNumber WHEN 1 THEN GETDATE() ELSE NULL END
                                , NULL, a.FlowJobName, NULL, NULL, a.SortNumber, NULL, 0, @Edition, a.FlowStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                FROM RFI.TemplateRfiSignFlow a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.DepartmentId = @DepartmentId
                                AND a.[Status] = 'A'
                                ORDER BY a.SortNumber";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                RfiDetailId,
                                ProTypeGroupId,
                                DepartmentId,
                                FlowUser = CreateBy,
                                Edition = string.Format("{0:00}", maxEdition1), //取2位數
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        #endregion

                        string RfiEdition = string.Empty;
                        string mailFlowContent = "【市場評估單】已建立-單號{0}，請查看內容確認!";
                        string detailHtml = string.Empty;
                        string flowsHtml = string.Empty;
                        string hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/AddRfiDetail", "RequestForInformation/RfiManagement");
                        string pageUrl = hyperLink;
                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));

                        #region //取得單身資料
                        sql = @"SELECT aa.RfiId, aa.RfiDetailId, aa.RfiSequence, aa.Edition
                                , aa.SalesId, ad.UserNo SalesNo, ad.UserName SalesName, ad.Gender SalesGender
                                , FORMAT(aa.ProdDate, 'yyyy-MM-dd') ProdDate
                                , FORMAT(aa.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart
                                , FORMAT(aa.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
                                , aa.TargetUnitPrice, aa.LifeCycleQty, aa.Revenue, aa.RevenueFC
                                , aa.ExchangeRate, aa.Currency, ac.TypeNo CurrencyTypeName, ac.TypeName CurrencyName
                                , aa.[Status] RfiDetailStatus, ae.StatusName RfiDetailStatusName
                                , ISNULL(FORMAT(aa.ConfirmFinalTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmFinalTime
                                , ab.*
                                FROM RFI.RfiDetail aa
                                LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.Currency AND ac.TypeSchema = 'ExchangeRate.Currency'
                                INNER JOIN BAS.[User] ad ON ad.UserId = aa.SalesId
                                INNER JOIN BAS.[Status] ae ON ae.StatusNo = aa.[Status] AND ae.StatusSchema = 'RfiDetail.Status'
                                OUTER APPLY (
	                                SELECT a.RfiSfId
                                    , a.SignJobName, a.Approved, a.Edition
	                                , u1.UserId ,u1.UserName, u1.UserNo, u1.Email
	                                , c.CompanyId, c.CompanyName, c.CompanyNo
	                                , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                                , DENSE_RANK () OVER (PARTITION BY a.RfiDetailId ORDER BY a.Edition DESC) SortEdition
	                                FROM RFI.RfiSignFlow a
	                                INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
	                                INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
	                                INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
	                                WHERE a.RfiDetailId = aa.RfiDetailId
	                                AND a.SortNumber =1
                                ) ab
                                WHERE ab.SortEdition = 1
                                AND aa.RfiDetailId = @RfiDetailId";
                        var resultRfiDetail = sqlConnection.Query(sql, new { RfiDetailId }).ToList();
                        #endregion

                        #region //Mail內容明細
                        resultRfiDetail.ForEach(x =>
                        {
                            detailHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>",
                                x.RfiSequence, x.Edition, x.ProdDate, x.ProdLifeCycleStart, x.ProdLifeCycleEnd
                                , string.Format("{0:#,##0.######}", x.TargetUnitPrice)
                                , x.LifeCycleQty.ToString("N0")
                                , string.Format("{0:#,##0.######}", x.ExchangeRate)
                                , string.Format("{0:#,##0.######}", x.Revenue)
                                , x.RfiDetailStatusName);
                        });

                        detailHtml = @"<table border>
                                        <thead>
                                            <tr>
                                                <th>流水號</th><th>評估版次</th><th>量產時間</th><th>生產生命(起)</th><th>生產生命(迄)</th><th>目標單價</th><th>預估數量</th><th>匯率</th><th>營收</th><th>單據狀態</th>
                                            </tr>
                                        </thead>
                                      <tbody>" + detailHtml + "</tbody></table>";
                        #endregion

                        #region //取得流程檔
                        sql = @"SELECT a.RfiSfId
                                , a.SignJobName, a.Approved, a.Edition, a.SortNumber
                                , u1.UserId, u1.UserName, u1.UserNo, u1.Email
                                , c.CompanyId, c.CompanyName, c.CompanyNo
                                , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                            , ISNULL(a.SignDesc, '-') SignDesc
	                            , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
	                            , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
	                            , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
	                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
	                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                                , ISNULL(s2.StatusName, '待審批') ApprovedName
                                FROM RFI.RfiSignFlow a
                                INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.Approved AND s2.StatusSchema = 'RfiSign.Status'
                                WHERE a.RfiDetailId = @RfiDetailId
                                ORDER BY a.SortNumber";
                        var resultRfiSignFlow = sqlConnection.Query(sql, new { RfiDetailId, CreateDate }).ToList();
                        var firstFlow = resultRfiSignFlow.FirstOrDefault(f => f.SortNumber == 1);
                        string signJobName = firstFlow.SignJobName;
                        string flowerName = firstFlow.UserName;
                        string rfiFlowEdition = firstFlow.Edition;
                        string CompanyNo = firstFlow.CompanyNo;
                        string UserNo = firstFlow.UserNo;
                        #endregion

                        #region //Mail內容明細
                        resultRfiSignFlow.ForEach(x =>
                        {
                            RfiEdition = x.Edition;
                            flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                x.SignJobName,
                                x.FlowerName,
                                x.ApprovedName,
                                string.Format("{0}天{1}時{2}分", x.StayDays, x.StayHours, x.StayMinutes),
                                x.ArriveDateTime,
                                x.SignDateTime,
                                x.SignDesc);
                        });
                        flowsHtml = @"<table border>
                                        <thead>
                                            <tr>
                                                <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                            </tr>
                                        </thead>
                                      <tbody>" + flowsHtml + "</tbody></table>";
                        #endregion

                        #region //Mail資料
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(a.MailFrom, '') 'MailFrom'
                                , ISNULL(a.MailTo, '') 'MailTo'
                                , ISNULL(a.MailCc, '') 'MailCc'
                                , ISNULL(a.MailBcc, '') 'MailBcc'
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                WHERE EXISTS (
                                    SELECT ca.MailId
                                    FROM BAS.MailSendSetting ca
                                    WHERE ca.MailId = a.MailId
                                    AND ca.SettingSchema = @SettingSchema
                                    AND ca.SettingNo = @SettingNo
                                )";
                        var mailTemplate = sqlConnection.Query(sql, new { SettingSchema = "RfiSignFlow", SettingNo = "Y" });
                        if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in mailTemplate)
                        {
                            string mailSubject = item.MailSubject;
                            string mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfiNo]", rfiNo)
                                                     .Replace("[Edition]", rfiFlowEdition)
                                                     .Replace("[SignJobName]", signJobName)
                                                     .Replace("[FlowerName]", flowerName)
                                                     .Replace("[CustomerName]", customerName)
                                                     .Replace("[CustomerEName]", customerEName)
                                                     .Replace("[AnalysisDate]", analysisDate)
                                                     .Replace("[ProTypeGroupName]", proTypeGroupName)
                                                     .Replace("[ProductUseName]", productUseName)
                                                     .Replace("[ExistFlagName]", existFlagName)
                                                     .Replace("[RfiStatusName]", rfiStatusName)
                                                     .Replace("[SalesName]", salesName)
                                                     .Replace("[MailContent]", string.Format(mailFlowContent, rfiNo));
                                                    //.Replace("[FlowContent]", flowsHtml)
                                                    //.Replace("[RfiDetailContent]", detailHtml)
                                                    //.Replace("[hyperlink]", hyperLink)

                            #region //MAMO個人訊息推播
                            string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                            ConvertToMarkdown(detailHtml, out string detailMamo);
                            ConvertToMarkdown(flowsHtml, out string flowsMamo);
                            mamoContent = mamoContent.Replace("[RfiDetailContent]", "\r\n" + detailMamo)
                                                     .Replace("[FlowContent]", "\r\n" + flowsMamo)
                                                     .Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                            mamoHelper.SendMessage(CompanyNo, CurrentUser, "Personal", UserNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());

                            //var companyNos = resultRfiDetail.Where(w => w.CompanyNo != null).Select(s => { return (string)s.CompanyNo; }).Distinct().ToList();
                            //foreach (var companyNo in companyNos)
                            //{
                            //    var userNos = resultRfiDetail.Where(w => w.CompanyNo == companyNo && w.UserNo != null).Select(s => { return (string)s.UserNo; }).Distinct().ToList();
                            //    foreach (var userNo in userNos)
                            //    {
                            //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                            //    }
                            //}
                            #endregion

                            #region //發送
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = string.Join(";", resultRfiDetail.Select(s => s.MailTo).Distinct()),
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent.Replace("[RfiDetailContent]", detailHtml).Replace("[FlowContent]", flowsHtml).Replace("[hyperlink]", hyperLink),
                                TextBody = "-"
                            };
                            BaseHelper.MailSend(mailConfig);
                            #endregion
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

        #region //AddDemandDesign -- 新設計申請單新增 -- Chia Yuan -- 2024.01.25
        public string AddDemandDesign(int RfiId, int AdditionalFile
            , string ProdSpecs, string ProdSpecDetails)
        {
            //尚未開放
            return "";
        }
        #endregion

        #region //AddDemandDesignSpec -- 新設計申請單規格新增 -- Chia Yuan -- 2024.01.25
        public string AddDemandDesignSpec(int DesignId, string ProdSpecs, string ProdSpecDetails)
        {
            try
            {
                #region /判斷資料長度
                if (string.IsNullOrWhiteSpace(ProdSpecs)) throw new SystemException("【設計規格】資料錯誤!");
                if (string.IsNullOrWhiteSpace(ProdSpecDetails)) throw new SystemException("【設計規格特徵】資料錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷新設計申請單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.DemandDesign a
                                WHERE a.DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【新設計申請單】資料錯誤!");
                        #endregion

                        if (!ProdSpecs.TryParseJson(out JObject prodSpecs)) throw new SystemException("【設計規格】資料錯誤!");
                        if (!ProdSpecDetails.TryParseJson(out JObject prodSpecDetails)) throw new SystemException("【設計規格特徵】資料錯誤!");
                        if (prodSpecs["data"].Count() <= 0) throw new SystemException("【設計規格】資料必須維護!");
                        if (prodSpecDetails["data"].Count() <= 0) throw new SystemException("【設計規格特徵】資料必須維護!");

                        if (prodSpecs["data"].Count() > 0)
                        {
                            var tempSpIds = prodSpecs["data"].Select(s => { return Convert.ToInt32(s["TempSpId"]); }).ToArray();

                            #region //取得設計規格資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TempSpId, a.ParameterName, a.SortNumber, a.ControlType, a.[Required], a.[Status]
                                    FROM RFI.TemplateSpecParameter a 
                                    ORDER BY a.SortNumber";
                            var resultProdSpec = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //取得設計規格子資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TempSpDetailId, a.PmtDetailName, a.SortNumber, a.DataType, a.[Description], a.[Required]
                                    FROM RFI.TemplateSpecParameterDetail a 
                                    WHERE a.TempSpId IN @TempSpIds
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("TempSpIds", tempSpIds);
                            var resultProdSpecDetail = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //判斷狀態資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StatusNo
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo IN @Status";
                            dynamicParameters.Add("Status", prodSpecDetails["data"].Select(s => { return s["Status"].ToString(); }).Distinct().ToArray());
                            var resultStatus = sqlConnection.Query(sql, dynamicParameters);
                            if (resultStatus.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                            #endregion

                            foreach (var data in prodSpecs["data"])
                            {
                                if (!int.TryParse(data["TempSpId"].ToString(), out int tempSpId)) throw new SystemException("【設計規格】資料錯誤!");
                                var tempSpecParameter = resultProdSpec.FirstOrDefault(f => f.TempSpId == tempSpId) ?? throw new SystemException("【設計規格】資料錯誤!");

                                #region //設計規格新增
                                sql = @"INSERT INTO RFI.DemandDesignSpec (DesignId, ParameterName, SortNumber, ControlType, Required, Status	
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DdSpecId
                                        VALUES (@DesignId, @ParameterName, @SortNumber, @ControlType, @Required, @Status	
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                    new
                                    {
                                        DesignId,
                                        tempSpecParameter.ParameterName,
                                        tempSpecParameter.SortNumber,
                                        tempSpecParameter.ControlType,
                                        tempSpecParameter.Required,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += 1;
                                int NewDdSpecId = insertResult?.DdSpecId ?? -1;
                                #endregion

                                foreach (var detail in prodSpecDetails["data"].Where(w => w["TempSpId"].ToString() == tempSpId.ToString()))
                                {
                                    if (!int.TryParse(detail["TempSpDetailId"].ToString(), out int tempSpDetailId)) throw new SystemException("【設計規格特徵】資料錯誤!");
                                    if (resultStatus.FirstOrDefault(f => f.StatusNo == detail["Status"].ToString()) == null) throw new SystemException("【狀態】資料錯誤!");
                                    var tempSpecParameterDetail = resultProdSpecDetail.FirstOrDefault(f => f.TempSpDetailId == tempSpDetailId) ?? throw new SystemException("【設計規格特徵】資料錯誤!");

                                    #region //設計規格子資料新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO RFI.DemandDesignSpecDetail (DdSpecId, PmtDetailName, SortNumber, DataType, Required, RequireFlag, DesignInput, Description, Status
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.DdsDetailId
                                            VALUES (@NewDdSpecId, @PmtDetailName, @SortNumber, @DataType, @Required, @RequireFlag, @DesignInput, @Description, @Status
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                        new
                                        {
                                            NewDdSpecId,
                                            PmtDetailName = detail["DataType"].ToString() == "O" ? detail["PmtDetailName"].ToString() : tempSpecParameterDetail.PmtDetailName,
                                            tempSpecParameterDetail.SortNumber,
                                            DataType = string.IsNullOrWhiteSpace(detail["DataType"].ToString()) ? null : detail["DataType"].ToString().Trim(),
                                            tempSpecParameterDetail.Required,
                                            RequireFlag = string.IsNullOrWhiteSpace(detail["RequireFlag"].ToString()) ? null : detail["RequireFlag"].ToString().Trim(),
                                            DesignInput = string.IsNullOrWhiteSpace(detail["DesignInput"].ToString()) ? null : detail["DesignInput"].ToString().Trim(),
                                            Description = string.IsNullOrWhiteSpace(detail["Description"].ToString()) ? null : detail["Description"].ToString().Trim(),
                                            Status = detail["Status"].ToString(),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    rowsAffected += 1;
                                    int NewDdsDetailId = insertResult?.DdsDetailId ?? -1;
                                    #endregion
                                }
                            }
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                            //data = insertResult
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
        #region //UpdateProdTerminal -- 終端資料更新 -- Chia Yuan 2023.12.25
        public string UpdateProdTerminal(int ProdTerminalId, string TerminalName, string Status)
        {
            try
            {
                if (TerminalName.Length <= 0) throw new SystemException("【終端名稱】不能為空!");
                if (TerminalName.Length > 50) throw new SystemException("【終端名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string terminalName = TerminalName.Trim();

                        #region //判斷終端資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TerminalName
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.Add("ProdTerminalId", ProdTerminalId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【終端】資料錯誤!");
                        #endregion

                        #region 判斷終端名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId <> @ProdTerminalId
                                AND a.TerminalName = @TerminalName";
                        dynamicParameters.Add("ProdTerminalId", ProdTerminalId);
                        dynamicParameters.Add("TerminalName", terminalName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【終端名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //終端資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.TerminalName = @TerminalName,
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdTerminalId,
                                Status,
                                TerminalName = terminalName,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProdTerminalStatus -- 終端資料狀態更新 -- Chia Yuan 2023.12.25
        public string UpdateProdTerminalStatus(int ProdTerminalId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷終端資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.Add("ProdTerminalId", ProdTerminalId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【終端】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdTerminalId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProdSystem -- 系統資料更新 -- Chia Yuan 2023.12.25
        public string UpdateProdSystem(int ProdSysId, string SystemName, string Status)
        {
            try
            {
                if (SystemName.Length <= 0) throw new SystemException("【系統名稱】不能為空!");
                if (SystemName.Length > 50) throw new SystemException("【系統名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string systemName = SystemName.Trim();

                        #region //判斷系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SystemName
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.Add("ProdSysId", ProdSysId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統】資料錯誤!");
                        #endregion

                        #region 判斷系統名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId <> @ProdSysId
                                AND a.SystemName = @SystemName";
                        dynamicParameters.Add("ProdSysId", ProdSysId);
                        dynamicParameters.Add("SystemName", systemName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【系統名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //系統資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.SystemName = @SystemName,
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdSysId,
                                Status,
                                SystemName = systemName,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProdSystemStatus -- 系統資料狀態更新 -- Chia Yuan 2023.12.25
        public string UpdateProdSystemStatus(int ProdSysId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.Add("ProdSysId", ProdSysId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdSysId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProdModule -- 模組資料更新 -- Chia Yuan 2023.12.25
        public string UpdateProdModule(int ProdModId, string ModuleName, string Status)
        {
            try
            {
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 50) throw new SystemException("【模組名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string moduleName = ModuleName.Trim();

                        #region //判斷模組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ModuleName
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.Add("ProdModId", ProdModId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組】資料錯誤!");
                        #endregion

                        #region 判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId <> @ProdModId
                                AND a.ModuleName = @ModuleName";
                        dynamicParameters.Add("ProdModId", ProdModId);
                        dynamicParameters.Add("ModuleName", moduleName);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【模組名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'Status'
                                AND a.StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusNo", Status);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //模組資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.ModuleName = @ModuleName,
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdModId,
                                Status,
                                ModuleName = moduleName,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProdModuleStatus -- 模組資料狀態更新 -- Chia Yuan 2023.12.25
        public string UpdateProdModuleStatus(int ProdModId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.Add("ProdModId", ProdModId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProdModId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProductTypeGroupStatus -- 產品類別狀態更新 -- Chia Yuan 2025.05.24
        public string UpdateProductTypeGroupStatus(int ProTypeGroupId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProTypeGroupId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateProductTypeStatus -- 產品類別狀態更新 -- Chia Yuan 2025.05.24
        public string UpdateProductTypeStatus(int ProTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM SCM.ProductType a
                                WHERE a.ProTypeId = @ProTypeId";
                        dynamicParameters.Add("ProTypeId", ProTypeId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM SCM.ProductType a
                                WHERE a.ProTypeId = @ProTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProTypeId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateProdSpec -- 評估樣板主資料更新 -- Chia Yuan 2023.12.29
        public string UpdateTemplateProdSpec(int TempProdSpecId, string SpecName, string SpecEName, string Required, string Status)
        {
            try
            {
                if (SpecName.Length <= 0) throw new SystemException("【評估項目名稱】不能為空!");
                if (SpecName.Length > 100) throw new SystemException("【評估項目名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(SpecEName))
                {
                    if (SpecName.Length > 100) throw new SystemException("【規格(英文)種類名稱】長度錯誤!");
                    SpecEName = SpecEName.Trim();
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        SpecName = SpecName.Trim();

                        int proTypeGroupId = -1;

                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProTypeGroupId
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //判斷主資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SpecName
                                FROM RFI.TemplateProdSpec a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.TempProdSpecId <> @TempProdSpecId";
                        dynamicParameters.Add("proTypeGroupId", proTypeGroupId);
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any(a => a.SpecName == SpecName)) throw new SystemException("【設計項目名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        if (!string.IsNullOrWhiteSpace(Status))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo = @StatusNo";
                            dynamicParameters.Add("StatusNo", Status);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        }
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.SpecName = @SpecName,
                                a.SpecEName = @SpecEName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempProdSpecId,
                                SpecName,
                                SpecEName = string.IsNullOrWhiteSpace(SpecEName) ? null : SpecEName.Trim(),
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateSpecParameter -- 設計樣板主資料更新 -- Chia Yuan 2024.05.24
        public string UpdateTemplateSpecParameter(int TempSpId, string ParameterName, string ParameterEName, string Required, string Status)
        {
            try
            {
                if (ParameterName.Length <= 0) throw new SystemException("【設計項目名稱】不能為空!");
                if (ParameterName.Length > 100) throw new SystemException("【設計項目名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(ParameterEName))
                {
                    if (ParameterName.Length > 100) throw new SystemException("【設計項目名稱(英文)】長度錯誤!");
                    ParameterEName = ParameterEName.Trim();
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        ParameterName = ParameterName.Trim();

                        int proTypeGroupId = -1;
                        
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProTypeGroupId
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【設計項目】資料錯誤!");
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //判斷主資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ParameterName
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.TempSpId <> @TempSpId";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                        dynamicParameters.Add("TempSpId", TempSpId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any(a => a.ParameterName == ParameterName)) throw new SystemException("【設計項目名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        if (!string.IsNullOrWhiteSpace(Status))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo = @StatusNo";
                            dynamicParameters.Add("StatusNo", Status);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        }
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.ParameterName = @ParameterName,
                                a.ParameterEName = @ParameterEName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempSpId,
                                ParameterName,
                                ParameterEName = string.IsNullOrWhiteSpace(ParameterEName) ? null : ParameterEName.Trim(),
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateProdSpecDetail -- 評估樣板子資料更新 -- Chia Yuan 2024.01.03
        public string UpdateTemplateProdSpecDetail(int TempPsDetailId, string FeatureName, string Description, string Status)
        {
            try
            {
                if (FeatureName.Length <= 0) throw new SystemException("【特徵名稱】不能為空!");
                if (FeatureName.Length > 50) throw new SystemException("【特徵名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    if (FeatureName.Length > 50) throw new SystemException("【描述】長度錯誤!");
                }
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    Description = Description.Trim();
                }
                else {
                    Description = null;
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        FeatureName = FeatureName.Trim();

                        int tempProdSpecId = -1;

                        #region //判斷規格特徵資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TempProdSpecId
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempPsDetailId = @TempPsDetailId";
                        dynamicParameters.Add("TempPsDetailId", TempPsDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【規格特徵】資料錯誤!");
                        foreach (var item in result)
                        {
                            tempProdSpecId = item.TempProdSpecId;
                        }
                        #endregion

                        #region //判斷特徵名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FeatureName
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempProdSpecId = @TempProdSpecId
                                AND a.TempPsDetailId <> @TempPsDetailId";
                        dynamicParameters.Add("TempProdSpecId", tempProdSpecId);
                        dynamicParameters.Add("TempPsDetailId", TempPsDetailId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any(a => a.FeatureName == FeatureName)) throw new SystemException("【特徵名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        if (!string.IsNullOrWhiteSpace(Status))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo = @StatusNo";
                            dynamicParameters.Add("StatusNo", Status);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        }
                        #endregion

                        #region //產品特徵資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.FeatureName = @FeatureName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempPsDetailId = @TempPsDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FeatureName,
                                TempPsDetailId,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateSpecParameterDetail -- 評估樣板特徵資料更新 -- Chia Yuan 2024.01.03
        public string UpdateTemplateSpecParameterDetail(int TempSpDetailId, string PmtDetailName, string Description, string Status)
        {
            try
            {
                if (PmtDetailName.Length <= 0) throw new SystemException("【設計選項名稱】不能為空!");
                if (PmtDetailName.Length > 50) throw new SystemException("【設計選項名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    if (PmtDetailName.Length > 50) throw new SystemException("【描述】長度錯誤!");
                }
                if (!string.IsNullOrWhiteSpace(Description))
                {
                    Description = Description.Trim();
                }
                else
                {
                    Description = null;
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        PmtDetailName = PmtDetailName.Trim();

                        int tempSpId = -1;

                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TempSpId
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpDetailId = @TempSpDetailId";
                        dynamicParameters.Add("TempSpDetailId", TempSpDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【設計選項】資料錯誤!");
                        foreach (var item in result)
                        {
                            tempSpId = item.TempSpId;
                        }
                        #endregion

                        #region //判斷特徵名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PmtDetailName
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpId = @TempSpId
                                AND a.TempSpDetailId <> @TempSpDetailId";
                        dynamicParameters.Add("TempSpId", tempSpId);
                        dynamicParameters.Add("TempSpDetailId", TempSpDetailId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any(a => a.PmtDetailName == PmtDetailName)) throw new SystemException("【設計選項名稱】資料重複!");
                        #endregion

                        #region //判斷狀態資料是否正確
                        if (!string.IsNullOrWhiteSpace(Status))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Status] a 
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo = @StatusNo";
                            dynamicParameters.Add("StatusNo", Status);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                        }
                        #endregion

                        #region //產品主資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.PmtDetailName = @PmtDetailName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpDetailId = @TempSpDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PmtDetailName,
                                TempSpDetailId,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateProdSpecSort -- 評估樣板主資料排序 -- Chia Yuan 2024.01.03
        public string UpdateTemplateProdSpecSort(int ProTypeGroupId, string SortList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalRowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("Status", "A");
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別群組】資料錯誤!");
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] dataSort = SortList.Split(',');

                        for (int i = 0; i < dataSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateProdSpec a
                                    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.TempProdSpecId = @TempProdSpecId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProTypeGroupId,
                                    TempProdSpecId = Convert.ToInt32(dataSort[i]),
                                    SortNumber = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
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

        #region //UpdateTemplateProdSpecDetailSort -- 評估樣板子資料排序 -- Chia Yuan 2024.01.03
        public string UpdateTemplateProdSpecDetailSort(int TempProdSpecId, string SortList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalRowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] dataSort = SortList.Split(',');

                        for (int i = 0; i < dataSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateProdSpecDetail a
                                    WHERE a.TempProdSpecId = @TempProdSpecId
                                    AND a.TempPsDetailId = @TempPsDetailId";

                            var id = Convert.ToInt32(dataSort[i]);

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempProdSpecId,
                                    TempPsDetailId = Convert.ToInt32(dataSort[i]),
                                    SortNumber = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
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

        #region //UpdateTemplateSpecParameterSort -- 設計樣板主資料排序 -- Chia Yuan 2024.05.24
        public string UpdateTemplateSpecParameterSort(int ProTypeGroupId, string SortList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalRowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("Status", "A");
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別群組】資料錯誤!");
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] dataSort = SortList.Split(',');

                        for (int i = 0; i < dataSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateSpecParameter a
                                    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.TempSpId = @TempSpId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProTypeGroupId,
                                    TempSpId = Convert.ToInt32(dataSort[i]),
                                    SortNumber = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
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

        #region //UpdateTemplateSpecParameterDetailSort -- 設計樣板子資料排序 -- Chia Yuan 2024.05.24
        public string UpdateTemplateSpecParameterDetailSort(int TempSpId, string SortList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalRowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] dataSort = SortList.Split(',');

                        for (int i = 0; i < dataSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateSpecParameterDetail a
                                    WHERE a.TempSpId = @TempSpId
                                    AND a.TempSpDetailId = @TempSpDetailId";

                            var id = Convert.ToInt32(dataSort[i]);

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempSpId,
                                    TempSpDetailId = Convert.ToInt32(dataSort[i]),
                                    SortNumber = i + 1,
                                    LastModifiedDate,
                                    LastModifiedBy
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

        #region //UpdateControlType -- 評估樣板控制項類型更新 -- Chia Yuan 2023-12-27
        public string UpdateTemplateProdSpecControlType(int TempProdSpecId, string ControlType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        #endregion

                        #region //判斷類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a
                                WHERE a.TypeNo = @ControlType";
                        dynamicParameters.Add("ControlType", ControlType);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【類型】資料錯誤!");
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.ControlType = @ControlType,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempProdSpecId,
                                ControlType,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateSpecParameterControlType -- 設計樣板控制項類型更新 -- Chia Yuan 2024.05.24
        public string UpdateTemplateSpecParameterControlType(int TempSpId, string ControlType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【設計規格】資料錯誤!");
                        #endregion

                        #region //判斷類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a
                                WHERE a.TypeNo = @ControlType";
                        dynamicParameters.Add("ControlType", ControlType);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【類型】資料錯誤!");
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.ControlType = @ControlType,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempSpId,
                                ControlType,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateProdSpecStatus -- 評估樣板資料狀態更新 -- Chia Yuan 2023.12.29
        public string UpdateTemplateProdSpecStatus(int TempProdSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempProdSpecId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateSpecParameterStatus -- 設計樣板資料狀態更新 -- Chia Yuan 2024.05.24
        public string UpdateTemplateSpecParameterStatus(int TempSpId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【設計規格】資料錯誤!");
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempSpId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateProdSpecRequired -- 評估樣板必填注記更新 -- Chia Yuan 2023.12.29
        public string UpdateTemplateProdSpecRequired(int TempProdSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Required
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        bool Required = false;
                        foreach (var item in result)
                        {
                            Required = item.Required;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Required)
                        {
                            case true:
                                Required = false;
                                break;
                            case false:
                                Required = true;
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Required = @Required,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempProdSpecId,
                                Required,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateTemplateSpecParameterRequired -- 設計樣板必填注記更新 -- Chia Yuan 2023.12.29
        public string UpdateTemplateSpecParameterRequired(int TempSpId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Required
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        bool Required = false;
                        foreach (var item in result)
                        {
                            Required = item.Required;
                        }
                        #endregion

                        #region //調整為相反狀態
                        switch (Required)
                        {
                            case true:
                                Required = false;
                                break;
                            case false:
                                Required = true;
                                break;
                        }
                        #endregion

                        #region //狀態更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.Required = @Required,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TempSpId,
                                Required,
                                LastModifiedDate,
                                LastModifiedBy
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

        #region //UpdateRequestForInformation RFI單頭資料更新 -- Chia Yuan 2024.01.10
        public string UpdateRequestForInformation(int RfiId, int ProTypeGroupId, int ProductUseId, int SalesId, string ExistFlag
            , string CustomerName, string CustomerEName, int CustomerId, string OrganizaitonType)
        {
            try
            {
                #region //判斷資料長度
                if (!Regex.IsMatch(ExistFlag, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【是否存在】注記錯誤!");
                if (string.IsNullOrWhiteSpace(CustomerName)) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerName.Length > 100) throw new SystemException("【客戶名稱】長度錯誤!");
                if (!string.IsNullOrWhiteSpace(CustomerEName))
                {
                    if (CustomerEName.Length > 100) throw new SystemException("【客戶名稱(英文)】長度錯誤!");
                    CustomerEName = CustomerEName.Trim();
                }
                else
                {
                    CustomerEName = null;
                }
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        CustomerName = CustomerName.Trim();

                        #region //判斷類別群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("Status", "A");
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別群組】資料錯誤!");
                        #endregion

                        #region //判斷產品應用資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse a
                                WHERE a.ProductUseId = @ProductUseId
                                AND a.Status = @Status";
                        dynamicParameters.Add("ProductUseId", ProductUseId);
                        dynamicParameters.Add("Status", "A");
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品應用】資料錯誤!");
                        #endregion

                        #region //判斷客戶類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type] a
                                WHERE a.TypeNo = @TypeNo
                                AND a.TypeSchema = 'MemberOrganization.OrganizaitonType'";
                        dynamicParameters.Add("TypeNo", OrganizaitonType);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【客戶類型】資料錯誤!");
                        #endregion

                        #region //判斷存在資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a
                                WHERE a.StatusNo = @StatusNo
                                AND a.StatusSchema = @StatusSchema";
                        dynamicParameters.Add("StatusNo", ExistFlag);
                        dynamicParameters.Add("StatusSchema", "Boolean");
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【是否存在】資料錯誤!");
                        #endregion

                        #region //RFI單頭資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                ProTypeGroupId = @ProTypeGroupId,
                                ProductUseId = @ProductUseId,
                                ExistFlag = @ExistFlag,
                                CustomerName = @CustomerName,
                                CustomerEName = @CustomerEName,
                                CustomerId = @CustomerId,
                                OrganizaitonType = @OrganizaitonType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                FROM RFI.RequestForInformation a
                                WHERE a.RfiId = @RfiId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            RfiId,
                            ProTypeGroupId,
                            ProductUseId,
                            ExistFlag,
                            CustomerName,
                            CustomerEName,
                            CustomerId = CustomerId <= 0 ? (int?)null : CustomerId,
                            OrganizaitonType,
                            LastModifiedDate,
                            LastModifiedBy
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

        #region //UpdateRequestForInformationStatus RFI單頭資料更新 -- Chia Yuan 2024.01.10
        public string UpdateRequestForInformationStatus(int RfiId, string StatusNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷存在資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a
                                WHERE a.StatusNo = @StatusNo
                                AND a.StatusSchema = 'RequestForInformation.Status'";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { StatusNo }) ?? throw new SystemException("【狀態】資料錯誤!");
                        #endregion

                        #region //RFI單頭資料更新
                        sql = @"UPDATE a SET
                                Status = @StatusNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                FROM RFI.RequestForInformation a
                                WHERE a.RfiId = @RfiId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                RfiId,
                                StatusNo,
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

        #region //UpdateRfiDetail RFI單身資料更新 -- Chia Yuan 2024.01.10
        public string UpdateRfiDetail(int RfiDetailId, string ProdDate, string ProdLifeCycleStart, string ProdLifeCycleEnd
            , decimal TargetUnitPrice, decimal LifeCycleQty, decimal Revenue, decimal RevenueFC
            , string Currency, decimal ExchangeRate, int AdditionalFile, string ProdSpecs, string ProdSpecDetails)
        {
            try
            {
                #region /判斷資料長度
                if (string.IsNullOrWhiteSpace(ProdDate)) throw new SystemException("【量產時間】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleStart)) throw new SystemException("【產品生命起日】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleEnd)) throw new SystemException("【產品生命迄日】不能為空!");
                if (TargetUnitPrice <= 0) throw new SystemException("【目標單價】資料錯誤!");
                if (LifeCycleQty <= 0) throw new SystemException("【預估數量】資料錯誤!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】資料錯誤!");
                if (string.IsNullOrWhiteSpace(Currency)) throw new SystemException("【貨幣】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdSpecs)) throw new SystemException("【評估項目】資料錯誤!");
                if (string.IsNullOrWhiteSpace(ProdSpecDetails)) throw new SystemException("【規格特徵】資料錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        int unitRound = 0, totalRound = 0;
                        int unitRoundFC = 0, totalRoundFC = 0;
                        string currencyName = string.Empty,
                               currency = string.Empty;

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //目前匯率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MG003, MG004, MG005, MG006
                                FROM CMSMG
                                WHERE MG001 = @Currency
                                AND MG002 <= @DateNow
                                ORDER BY MG002 DESC";
                            dynamicParameters.Add("Currency", Currency);
                            dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyyMMdd"));

                            var erpResult = erpConnection.Query(sql, dynamicParameters);
                            if (erpResult.Count() <= 0) throw new SystemException(string.Format("【ERP交易幣別匯率】{0}不存在!", Currency));
                            double exchangeRate = 0;
                            foreach (var item in erpResult)
                            {
                                exchangeRate = Convert.ToDouble(item.MG004);
                            }
                            #endregion

                            #region //交易幣別設定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MF001)) AS MF001, LTRIM(RTRIM(MF002)) AS MF002, MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);
                            erpResult = erpConnection.Query(sql, dynamicParameters);
                            if (erpResult.Count() <= 0) throw new SystemException(string.Format("【ERP交易幣別設定】{0} 不存在!", Currency));
                            foreach (var item in erpResult)
                            {
                                unitRoundFC = Convert.ToInt32(item.MF003); //單價取位
                                totalRoundFC = Convert.ToInt32(item.MF004); //金額取位
                                currency = item.MF001;
                                currencyName = item.MF002;
                            }
                            #endregion

                            #region //交易幣別設定
                            List<string> Currencys = new List<string> { "NT$", "NTD" };
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MF001)) AS MF001, LTRIM(RTRIM(MF002)) AS MF002, MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 IN @Currencys";
                            dynamicParameters.Add("Currencys", Currencys);
                            erpResult = erpConnection.Query(sql, dynamicParameters);
                            if (erpResult.Count() <= 0) throw new SystemException(string.Format("【ERP交易幣別設定】{0} 不存在!", string.Join(",", Currencys)));
                            foreach (var item in erpResult)
                            {
                                unitRound = Convert.ToInt32(item.MF003); //單價取位
                                totalRound = Convert.ToInt32(item.MF004); //金額取位
                            }
                            #endregion
                        }

                        #region //判斷RFI單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.Add("RfiDetailId", RfiDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【市場評估單身】資料錯誤!");
                        #endregion

                        #region //判斷貨幣資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.[Type] a 
                                WHERE a.TypeNo = @Currency
                                AND a.TypeSchema = 'ExchangeRate.Currency'";
                        dynamicParameters.Add("Currency", Currency);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【貨幣】資料錯誤!");
                        #endregion

                        if (!DateTime.TryParse(ProdDate, out DateTime prodDate)) throw new SystemException("【量產時間】資料錯誤!");
                        if (!DateTime.TryParse(ProdLifeCycleStart, out DateTime prodLifeCycleStart)) throw new SystemException("【產品生命起日】資料錯誤!");
                        if (!DateTime.TryParse(ProdLifeCycleEnd, out DateTime prodLifeCycleEnd)) throw new SystemException("【產品生命迄日】資料錯誤!");

                        #region //RFI單身資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.ProdDate = @ProdDate,
                                a.ProdLifeCycleStart = @ProdLifeCycleStart,
                                a.ProdLifeCycleEnd = @ProdLifeCycleEnd,
                                a.TargetUnitPrice = @TargetUnitPrice,
                                a.LifeCycleQty = @LifeCycleQty,
                                a.Revenue = @Revenue,
                                a.RevenueFC = @RevenueFC,
                                a.Currency = @Currency,
                                a.ExchangeRate = @ExchangeRate,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfiDetailId,
                                ProdDate = prodDate.ToString("yyyy-MM-dd"),
                                ProdLifeCycleStart = prodLifeCycleStart.ToString("yyyy-MM-dd"),
                                ProdLifeCycleEnd = prodLifeCycleEnd.ToString("yyyy-MM-dd"),
                                TargetUnitPrice,
                                LifeCycleQty,
                                Revenue = Math.Round(TargetUnitPrice * LifeCycleQty * ExchangeRate, totalRound, MidpointRounding.AwayFromZero),
                                RevenueFC = Math.Round(TargetUnitPrice * LifeCycleQty, totalRoundFC, MidpointRounding.AwayFromZero),
                                Currency,
                                ExchangeRate,
                                //AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (!ProdSpecs.TryParseJson(out JObject prodSpecs)) throw new SystemException("【評估項目】資料錯誤!");
                        if (!ProdSpecDetails.TryParseJson(out JObject prodSpecDetails)) throw new SystemException("【評估項目】資料錯誤!");

                        if (prodSpecs["data"].Count() > 0)
                        {
                            var tempProdSpecIds = prodSpecs["data"].Select(s => { return Convert.ToInt32(s["TempProdSpecId"]); }).ToArray();

                            #region //判斷狀態資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.StatusNo
                                    FROM BAS.[Status] a
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo IN @Status";
                            dynamicParameters.Add("Status", prodSpecDetails["data"].Select(s => { return s["Status"].ToString(); }).Distinct().ToArray());
                            var resultStatus = sqlConnection.Query(sql, dynamicParameters);
                            if (resultStatus.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                            #endregion

                            #region //評估項目
                            foreach (var detail in prodSpecDetails["data"])
                            {
                                if (!int.TryParse(detail["TempPsDetailId"].ToString(), out int psDetailId)) throw new SystemException("【規格特徵】資料錯誤!");
                                if (resultStatus.FirstOrDefault(f => f.StatusNo == detail["Status"].ToString()) == null) throw new SystemException("【狀態】資料錯誤!");

                                #region //規格特徵更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE a SET 
                                        a.FeatureName = CASE WHEN a.DataType = 'O' THEN @FeatureName ELSE a.FeatureName END,
                                        a.Description = @Description,
                                        a.Status = @Status,
                                        a.LastModifiedDate = @LastModifiedDate,
                                        a.LastModifiedBy = @LastModifiedBy
                                        FROM RFI.ProductSpecDetail a
                                        WHERE a.PsDetailId = @PsDetailId";
                                //--AND (a.Status <> @Status
                                //if (detail["DataType"].ToString() == "O") {
                                //    sql += @" OR a.FeatureName <> @FeatureName";
                                //}
                                //sql += @")";

                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PsDetailId = psDetailId,
                                        FeatureName = string.IsNullOrWhiteSpace(detail["FeatureName"].ToString()) ? null : detail["FeatureName"].ToString().Trim(),
                                        Status = detail["Status"].ToString(),
                                        Description = string.IsNullOrWhiteSpace(detail["Description"].ToString()) ? null : detail["Description"].ToString().Trim(),
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
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

        #region //UpdateRfiDetailAdditionalFile RFI單身附檔資料更新 -- Chia Yuan 2024.01.16
        public string UpdateRfiDetailAdditionalFile(int RfiDetailId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFI單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 AdditionalFile
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.Add("RfiDetailId", RfiDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【市場評估單身】資料錯誤!");
                        int? additionalFile = -1;
                        foreach (var item in result)
                        {
                            additionalFile = item.AdditionalFile;
                        }
                        #endregion

                        #region //判斷新附檔資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【附檔】資料錯誤!");
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.AdditionalFile = @AdditionalFile,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfiDetailId,
                                AdditionalFile = FileId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //舊附檔刪除
                        if (additionalFile != null)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET a.DeleteStatus = 'Y'
                                    FROM BAS.[File] a
                                    WHERE a.FileId = @AdditionalFile";
                            dynamicParameters.Add("AdditionalFile", additionalFile);
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

        #region //UpdateDesignFile 新設計申請單附檔資料新增 -- Chia Yuan 2024.01.24
        public string UpdateDesignFile(int DesignId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFI單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DesignFile
                                FROM RFI.DemandDesign a
                                WHERE a.DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【新設計申請單】資料錯誤!");
                        int? designFile = -1;
                        foreach (var item in result)
                        {
                            designFile = item.DesignFile;
                        }
                        #endregion

                        #region //判斷新附檔資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【附檔】資料錯誤!");
                        #endregion

                        #region //評估項目資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.DesignFile = @DesignFile,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.DemandDesign a
                                WHERE a.DesignId = @DesignId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DesignId,
                                DesignFile = FileId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //舊附檔刪除
                        if (designFile != null)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET a.DeleteStatus = 'Y'
                                    FROM BAS.[File] a
                                    WHERE a.FileId = @DesignFile";
                            dynamicParameters.Add("DesignFile", designFile);
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

        #region //UpdateRfiDetailStatus RFI單身狀態更新 -- Chia Yuan 2024.01.17
        public string UpdateRfiDetailStatus(int RfiDetailId, int RfiSfId, int DepartmentId, string RfiDetailStatus, string SignContent)
        {
            try
            {
                if (!Regex.IsMatch(RfiDetailStatus, "^(Y|N|B|y|v)$", RegexOptions.IgnoreCase)) throw new SystemException("【審批註記】資料錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        sql = @"SELECT TOP 1 a.RfiDetailId, a.RfiId, a.SalesId, a.CompanyId
							    , a.[Status] RfiDetailStatus, q.StatusName RfiDetailStatusName
							    , b.RfiNo, b.[Status] RfiStatus, s2.StatusName RfiStatusName 
							    , b.ExistFlag, s.StatusName ExistFlagName
							    , b.CustomerName, b.CustomerEName
							    , b.ProTypeGroupId, g.ProTypeGroupName, g.TypeOne ProductGroupTypeOne
							    , d.ProductUseId, d.ProductUseName, d.ProductUseNo, d.TypeOne ProductUseTyoeOne
                                , e.DepartmentId, e.DepartmentName
							    , c.UserName SalesName, c.UserNo SalesNo, c.Email SalesEmail
							    , ISNULL(FORMAT(a.CreateDate, 'yyyy-MM-dd'), '') AnalysisDate
							    , f.CompanyNo, f.CompanyId SalesCompanyId
							    , ISNULL(j.FileId, '') AdditionalFile
							    FROM RFI.RfiDetail a
							    INNER JOIN RFI.RequestForInformation b ON b.RfiId = a.RfiId
							    INNER JOIN SCM.ProductUse d ON d.ProductUseId = b.ProductUseId
							    INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = b.ProTypeGroupId
							    INNER JOIN BAS.[User] c ON c.UserId = b.SalesId
							    INNER JOIN BAS.Department e ON e.DepartmentId = c.DepartmentId
							    INNER JOIN BAS.Company f ON f.CompanyId = e.CompanyId
							    INNER JOIN BAS.[Status] s ON s.StatusNo = b.ExistFlag AND s.StatusSchema = 'Boolean'
							    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.[Status] AND s2.StatusSchema = 'RequestForInformation.Status'
							    INNER JOIN BAS.[Status] q ON q.StatusNo = a.[Status] AND q.StatusSchema = 'RfiDetail.Status'
							    OUTER APPLY (SELECT FileId FROM BAS.[File] WHERE FileId = a.AdditionalFile AND DeleteStatus = 'N') j
							    WHERE a.RfiDetailId = @RfiDetailId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId }) ?? throw new SystemException("【市場評估單身】資料錯誤!");
                        int RfiId = result.RfiId;
                        int companyId = result.CompanyId;
                        int ProTypeGroupId = result.ProTypeGroupId;
                        int? additionalFile = result.AdditionalFile;
                        int? salesId = result.SalesId; //建單業務
                        int SalesDepartmentId = result.DepartmentId; //流程部門id
                        string rfiStatus = result.RfiStatus;
                        string customerName = result.CustomerName;
                        string customerEName = result.CustomerEName;
                        string proTypeGroupName = result.ProTypeGroupName;
                        string productGroupTypeOne = result.ProductGroupTypeOne;
                        string productUseName = result.ProductUseName;
                        string productUseNo = result.ProductUseNo;
                        string productUseTyoeOne = result.ProductUseTyoeOne;
                        string rfiNo = result.RfiNo;
                        string analysisDate = result.AnalysisDate;
                        string existFlag = result.ExistFlag; //現有機種?
                        string existFlagName = result.ExistFlagName;
                        string rfiDetailStatus = result.RfiDetailStatus; //審批狀態
                        string MailTo = result.SalesEmail; //業務email
                        string salesNo = result.SalesNo;
                        string salesName = result.SalesName;
                        string salesCompanyNo = result.CompanyNo;
                        #endregion

                        #region //判斷流程檔是否正確
                        int MaxSortNumber = -1;
                        sql = @"SELECT TOP 1 SortNumber, Edition
							    FROM RFI.RfiSignFlow
							    WHERE RfiSfId = @RfiSfId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { RfiSfId }) ?? throw new SystemException("【流程檔】資料錯誤!");
                        MaxSortNumber = result.SortNumber;
                        string RfiFlowEdition = result.Edition;
                        #endregion

                        #region //判斷狀態資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'RfiSign.Status'
                                AND a.StatusNo = @RfiDetailStatus";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailStatus }) ?? throw new SystemException("【審批狀態】資料錯誤!");
                        #endregion

                        SignContent = string.IsNullOrWhiteSpace(SignContent) ? null : SignContent.Trim();

                        #region //審批流程更新
                        sql = @"UPDATE a SET
							    a.Approved = @RfiDetailStatus,
							    a.SignUser = @CurrentUser,
							    a.SignDesc = @SignContent,
							    a.SignDateTime = @CreateDate,
							    a.LastModifiedDate = @LastModifiedDate,
							    a.LastModifiedBy = @LastModifiedBy
							    FROM RFI.RfiSignFlow a
							    WHERE a.RfiSfId = @RfiSfId";
                        rowsAffected += sqlConnection.Execute(sql,
                             new
                             {
                                 RfiSfId,
                                 RfiDetailStatus,
                                 CurrentUser,
                                 SignContent,
                                 CreateDate,
                                 LastModifiedDate,
                                 LastModifiedBy
                             });
                        #endregion

                        bool createDesign = false;
                        string detailCode = string.Empty;
                        string SettingSchema = "RfiSignFlow";
                        string mailFlowContent = "【市場評估單】已送達-單號{0}，請查看內容並審批。";
                        string tempRfiStatus = rfiStatus;
                        string tempRfiDetailStatus = rfiDetailStatus;
                        string signJobName = string.Empty;
                        string flowerName = string.Empty;                        
                        int NextRfiSfId = -1;

                        #region //取得下一關                                
                        sql = @"SELECT TOP 1 a.RfiSfId
                                , a.SignJobName, a.Approved, a.Edition
                                , u1.UserId, u1.UserName, u1.UserNo, u1.Email
								, c.CompanyId, c.CompanyName, c.CompanyNo
								, d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
								FROM RFI.RfiSignFlow a
								INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
								INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
								INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
								WHERE a.RfiDetailId = @RfiDetailId
                                AND a.Edition = @RfiFlowEdition
								AND a.SortNumber = @SortNumber";
                        var NextFlow = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId, RfiFlowEdition, SortNumber = MaxSortNumber + 1 });
                        if (NextFlow != null)
                        {
                            NextRfiSfId = NextFlow.RfiSfId;
                            signJobName = NextFlow.SignJobName;
                            flowerName = NextFlow.UserName;
                            MailTo = NextFlow.Email;
                        }
                        #endregion

                        switch (RfiDetailStatus)
                        {
                            case "Y":
                                tempRfiDetailStatus = "1"; //審批中
                                tempRfiStatus = "1"; //審批中

                                if (NextRfiSfId > 0)
                                {
                                    #region //下一關狀態更新
                                    sql = @"UPDATE a SET
										    a.ArriveDateTime = @CreateDate,
										    a.LastModifiedDate = @LastModifiedDate,
										    a.LastModifiedBy = @LastModifiedBy
										    FROM RFI.RfiSignFlow a
										    WHERE a.RfiSfId = @NextRfiSfId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            NextRfiSfId,
                                            CreateDate,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion
                                }
                                else
                                { //沒有下一關
                                    if (existFlag == "N")
                                    {
                                        createDesign = true;
                                        tempRfiDetailStatus = "5"; //新設計申請單
                                        tempRfiStatus = "2"; //已取號
                                        SettingSchema = "RfiDesignSignFlow"; //新設計申請單通知
                                    }
                                    else
                                    {
                                        tempRfiDetailStatus = "6"; //回饋客戶
                                        tempRfiStatus = "3"; //客戶確認中
                                    }
                                    signJobName = "業務";
                                    flowerName = salesName; //業務名稱
                                }
                                break;
                            case "N":
                                mailFlowContent = "【市場評估單】已被否決-單號{0}，請前往系統查看。";
                                tempRfiDetailStatus = "7"; //否決
                                tempRfiStatus = "4"; //已否決
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱

                                #region //單身目前版次號
                                sql = @"SELECT TOP 1 ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
                                        FROM RFI.RfiSignFlow
                                        WHERE RfiDetailId = @RfiDetailId";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId });
                                int maxEdition1 = Convert.ToInt32(result.MaxEdition);
                                #endregion
                                break;
                            case "B":
                                mailFlowContent = "【市場評估單】已被退回-單號{0}，請前往系統查看。";
                                tempRfiDetailStatus = "8"; //退回
                                tempRfiStatus = "5"; //已退回
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱

                                #region //單身目前版次號
                                sql = @"SELECT TOP 1 ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
                                        FROM RFI.RfiSignFlow
                                        WHERE RfiDetailId = @RfiDetailId";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId });
                                string NewEdition = string.Format("{0:00}", Convert.ToInt32(result.MaxEdition)); //取2位數
                                #endregion

                                #region //取得審批審批流程
                                sql = @"SELECT TOP 1 1
                                        FROM RFI.TemplateRfiSignFlow a
                                        WHERE EXISTS (SELECT TOP 1 1 FROM SCM.ProductTypeGroup aa WHERE aa.ProTypeGroupId = a.ProTypeGroupId AND aa.[Status] = 'A')
                                        AND a.ProTypeGroupId = @ProTypeGroupId
                                        AND a.DepartmentId = @SalesDepartmentId";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, SalesDepartmentId }) ?? throw new SystemException("【市場評估單審批流程】資料錯誤!");
                                #endregion

                                #region //建立審批流程檔 --new
                                sql = @"INSERT RFI.RfiSignFlow
                                        SELECT @RfiDetailId, ISNULL(a.FlowUser, @FlowUser), b.AdditionalFile
                                        , CASE a.SortNumber WHEN 1 THEN GETDATE() ELSE NULL END
                                        , NULL, a.FlowJobName, NULL, NULL, a.SortNumber, NULL, 0, @NewEdition, a.FlowStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        FROM RFI.TemplateRfiSignFlow a
                                        OUTER APPLY (SELECT ba.AdditionalFile FROM RFI.RfiSignFlow ba WHERE ba.RfiDetailId = @RfiDetailId AND ba.Edition = @RfiFlowEdition AND ba.SortNumber = a.SortNumber AND ba.[Status] = a.FlowStatus) b
                                        WHERE a.ProTypeGroupId = @ProTypeGroupId
                                        AND a.DepartmentId = @SalesDepartmentId
                                        AND a.[Status] = 'A'
                                        ORDER BY a.SortNumber";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProTypeGroupId,
                                        SalesDepartmentId,
                                        RfiDetailId,
                                        FlowUser = salesId,
                                        NewEdition, //取2位數
                                        RfiFlowEdition, //舊版次
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                #endregion
                                break;
                            case "y":
                                mailFlowContent = "【市場評估單】已強制結案-單號{0}，請前往系統查看。";
                                //tempRfiDetailStatus = RfiDetailStatus; //強制結案
                                //tempRfiStatus = "Y"; //已結案
                                tempRfiDetailStatus = "6"; //回饋客戶
                                tempRfiStatus = "3"; //客戶確認中
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱
                                break;
                            case "v":
                                mailFlowContent = "【市場評估單】已作廢-單號{0}，請前往系統查看。";
                                tempRfiDetailStatus = RfiDetailStatus; //作廢
                                tempRfiStatus = "V"; //已作廢
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱
                                break;
                        }

                        #region //狀態更新 (停用)
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"UPDATE a SET
                        //        a.Status = @Status,";
                        //switch (Status)
                        //{
                        //    case "1": //業務
                        //        //if (rfiDetailStatus == "2") //業務主管 > 業務
                        //        //{
                        //        //    mailFlowContent = "市場評估單已遭退回，請前往系統查看!";
                        //        //    tempRfiStatus = "5"; //已退回
                        //        //    sql += @" a.SuperSales = @SignUser,
                        //        //            　a.SuperSalesDesc = @SignContent,
                        //        //            　a.ConfirmSTime = @ConfirmTime,";
                        //        //}
                        //        break;
                        //    case "2": //業務 > 業務主管
                        //        //if (rfiDetailStatus == "3") //總經理 > 業務
                        //        //{
                        //        //    mailFlowContent = "市場評估單已遭退回，請前往系統查看!";
                        //        //    tempRfiStatus = "5"; //已退回
                        //        //    sql += @" a.ManagerPlan = @SignUser,
                        //        //            　a.ManagerPlanDesc = @SignContent,
                        //        //            　a.ConfirmSTime = @ConfirmTime,";
                        //        //}
                        //        //else
                        //        //{
                        //        //    tempRfiStatus = "1"; //審批中
                        //        //    sql += @" a.Sales = @SignUser,
                        //        //              a.SalesDesc = @SignContent,
                        //        //            　a.ConfirmSSTime = @ConfirmTime,";
                        //        //    detailCode = "SignSS";
                        //        //}
                        //        tempRfiStatus = "1"; //審批中
                        //        sql += @" a.Sales = @SignUser,
                        //                　a.SalesDesc = @SignContent,
                        //                  a.ConfirmSSTime = @ConfirmTime,";
                        //        detailCode = "SignSS";
                        //        break;
                        //    case "3": //業務主管 > 總經理
                        //        sql += @" a.SuperSales = @SignUser,
                        //                  a.SuperSalesDesc = @SignContent,
                        //                　a.ConfirmMPTime = @ConfirmTime,";
                        //        detailCode = "SignMP";
                        //        break;
                        //    case "4": //總經理
                        //        sql += @" a.ManagerPlan = @SignUser,
                        //                  a.ManagerPlanDesc = @SignContent,";
                        //        switch (existFlag)
                        //        {
                        //            case "Y": //現有機種 > 業務
                        //                sql += @" a.ConfirmSTime = @ConfirmTime,
                        //                        　a.ConfirmFinalTime = @ConfirmTime,";
                        //                break;
                        //            case "N": //非現有機種 > 新設計申請單
                        //                tempRfiStatus = "2"; //已取號
                        //                functionCode = "DesignManagement";
                        //                settingSchema = "RfiDesignSignFlow"; //新設計申請單通知
                        //                Status = "5";
                        //                createDesign = true;
                        //                sql += @" a.ConfirmSTime = @ConfirmTime,
                        //                          a.ConfirmFinalTime = @ConfirmTime,";
                        //                break;
                        //        }
                        //        break;
                        //    //case "5": //業務 > 新設計申請單 (停用)
                        //    //    sql += @" a.Sales = @SignUser,
                        //    //            　a.SalesDesc = @SignContent,
                        //    //              a.ConfirmFinalTime = @ConfirmTime,";
                        //    //    break;
                        //    case "6": //業務 > 回饋客戶
                        //        tempRfiStatus = "3"; //已回饋
                        //        sql += @" a.NotifyUser = @SignUser,
                        //                  a.CustomerDesc = @SignContent,
                        //                  a.ConfirmCustTime = @ConfirmTime,";
                        //        break;
                        //    case "7": //Flower > 不准
                        //        mailFlowContent = "市場評估單已被否決，請前往系統查看!";
                        //        tempRfiStatus = "4"; //已否決
                        //        switch (rfiDetailStatus)
                        //        {
                        //            case "1":
                        //                break;
                        //            case "2": //業務主管不准
                        //                sql += @" a.SuperSales = @SignUser,
                        //                          a.SuperSalesDesc = @SignContent,";
                        //                break;
                        //            case "3": //總經理不准
                        //                sql += @" a.ManagerPlan = @SignUser,
                        //                          a.ManagerPlanDesc = @SignContent,";
                        //                break;
                        //        }
                        //        sql += @" a.ConfirmFinalTime = @ConfirmTime,";
                        //        break;
                        //    case "8": //Flower > 退回
                        //        mailFlowContent = "市場評估單已遭退回，請前往系統查看!";
                        //        tempRfiStatus = "5"; //已退回
                        //        switch (rfiDetailStatus)
                        //        {
                        //            case "2":  //業務主管 > 業務
                        //                sql += @" a.SuperSales = @SignUser,
                        //                    　    a.SuperSalesDesc = @SignContent,
                        //                          a.ManagerPlan = NULL,
                        //                          a.ConfirmMPTime = NULL,
                        //                          --a.Sales = NULL,
                        //                          --a.ConfirmSTime = NULL,";
                        //                break;
                        //            case "3": //總經理 > 業務
                        //                sql += @" a.ManagerPlan = @SignUser,
                        //                    　    a.ManagerPlanDesc = @SignContent,
                        //                          a.Sales = NULL,
                        //                          a.ConfirmSTime = NULL,
                        //                          --a.SuperSales = NULL,
                        //                          --a.ConfirmSSTime = NULL,";
                        //                break;
                        //        }
                        //        sql += @" a.ConfirmFinalTime = @ConfirmTime,";
                        //        break;
                        //}
                        //sql += @"a.LastModifiedDate = @LastModifiedDate,
                        //        a.LastModifiedBy = @LastModifiedBy
                        //        FROM RFI.RfiDetail a
                        //        WHERE a.RfiDetailId = @RfiDetailId";
                        //dynamicParameters.AddDynamicParams(
                        //    new
                        //    {
                        //        Status,
                        //        SignContent,
                        //        SignUser = LastModifiedBy,
                        //        ConfirmTime = LastModifiedDate,
                        //        LastModifiedDate,
                        //        LastModifiedBy,
                        //        RfiDetailId
                        //    });
                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (tempRfiDetailStatus != rfiDetailStatus)
                        {
                            #region //單據狀態更新
                            sql = @"UPDATE a SET
								    a.Status = @Status,
								    a.LastModifiedDate = @LastModifiedDate,
								    a.LastModifiedBy = @LastModifiedBy
								    FROM RFI.RfiDetail a
								    WHERE a.RfiDetailId = @RfiDetailId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    RfiDetailId,
                                    Status = tempRfiDetailStatus,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            #endregion
                        }

                        if (tempRfiStatus != rfiStatus)
                        {
                            #region //單據狀態更新
                            sql = @"UPDATE a SET
								    a.Status = @Status,
								    a.LastModifiedDate = @LastModifiedDate,
								    a.LastModifiedBy = @LastModifiedBy
								    FROM RFI.RequestForInformation a
								    WHERE a.RfiId = @RfiId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    RfiId,
                                    Status = tempRfiStatus,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            #endregion
                        }

                        string designNo = string.Empty;
                        string designEdition = string.Empty;
                        string designStatusName = string.Empty;
                        string assemblyName = string.Empty;
                        string rfiStatusName = string.Empty;
                        string detailHtml = string.Empty;
                        string flowsHtml = string.Empty;
                        int NewDesignId = -1;

                        if (createDesign)
                        {
                            assemblyName = GetAssemblyName(companyId, productGroupTypeOne, productUseTyoeOne);
                            mailFlowContent = "【新設計申請單】已建立-來源單號{0}，請查看內容並審批!";

                            #region //取得RFI流程檔
                            sql = @"SELECT TOP 1 *
                                    FROM (
	                                    SELECT a.RfiSfId
                                        , a.SignJobName, a.Approved, a.Edition
	                                    , u1.UserId, u1.UserName, u1.UserNo, u1.Email
	                                    , c.CompanyId, c.CompanyName, c.CompanyNo
	                                    , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                                    , DENSE_RANK () OVER (PARTITION BY a.RfiDetailId ORDER BY a.Edition DESC) SortEdition
	                                    , s1.StatusDesc
	                                    FROM RFI.RfiSignFlow a
	                                    INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
	                                    INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
	                                    INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
	                                    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.[Status] AND s1.StatusSchema = 'RfiDetail.FlowStatus'
	                                    WHERE a.RfiDetailId = @RfiDetailId
                                    ) SignFlow
                                    WHERE SignFlow.StatusDesc = 'SalesManager'
                                    AND SignFlow.SortEdition = 1";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { RfiDetailId });
                            int? SuperSalesId = result?.UserId;
                            #endregion

                            #region //取得審批審批流程
                            sql = @"SELECT TOP 1 1
                                    FROM RFI.TemplateDesignSignFlow a
                                    WHERE EXISTS (SELECT TOP 1 1 FROM SCM.ProductTypeGroup aa WHERE aa.ProTypeGroupId = a.ProTypeGroupId AND aa.[Status] = 'A')
                                    AND a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.DepartmentId = @DepartmentId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, DepartmentId }) ?? throw new SystemException("【新設計申請單審批流程】資料錯誤!");
                            #endregion

                            #region //單身目前序號、版次號
                            sql = @"SELECT TOP 1
								    ISNULL(CAST(MAX(DesignSequence) AS INT), 0) + 1 MaxSequence,
								    ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
								    FROM RFI.DemandDesign
								    WHERE RfiId = @RfiId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                            int maxSequence = Convert.ToInt32(result?.MaxSequence ?? 1);
                            int maxEdition = Convert.ToInt32(result?.MaxEdition ?? 1);
                            #endregion

                            #region //拋轉新設計申請單
                            sql = @"INSERT INTO RFI.DemandDesign(RfiId, CompanyId, DepartmentId, DesignNo, DesignSequence, AssemblyName, SalesId, SuperSalesId, AdditionalFile, Edition, Status	
								    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
								    OUTPUT INSERTED.DesignId
								    VALUES (@RfiId, @CurrentCompany, @DepartmentId, @DesignNo, @DesignSequence, @AssemblyName, @SalesId, @SuperSalesId, @AdditionalFile, @Edition, @Status	
								    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    RfiId,
                                    CurrentCompany,
                                    DepartmentId,
                                    DesignNo = rfiNo,
                                    DesignSequence = string.Format("{0:0000}", maxSequence), //取4位數
                                    AssemblyName = assemblyName,
                                    SalesId = salesId,
                                    SuperSalesId,
                                    AdditionalFile = additionalFile > 0 ? additionalFile : null,
                                    Edition = string.Format("{0:00}", maxEdition), //取2位數
                                    Status = "1", //新設計申請 > 業務
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += 1;
                            NewDesignId = insertResult?.DesignId ?? -1;
                            #endregion

                            if (NewDesignId < 1) throw new SystemException("【新設計申請單】建立失敗!");

                            #region //建立審批流程檔 --new
                            sql = @"INSERT RFI.DesignSignFlow
								    SELECT @NewDesignId, ISNULL(a.FlowUser, @FlowUser), NULL AS AdditionalFile
								    , CASE a.SortNumber WHEN 1 THEN GETDATE() ELSE NULL END
								    , NULL, a.FlowJobName, NULL, NULL, a.SortNumber, NULL, 0, a.FlowStatus
								    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
								    FROM RFI.TemplateDesignSignFlow a
								    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.DepartmentId = @DepartmentId
								    AND a.[Status] = 'A'
								    ORDER BY a.SortNumber";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProTypeGroupId,
                                    DepartmentId,
                                    NewDesignId,
                                    FlowUser = salesId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion

                            #region //取得新設計申請單第一關
                            sql = @"SELECT TOP 1 a.DesignSfId
                                    , a.SignJobName, a.Approved
                                    , u1.UserId, u1.UserName, u1.UserNo, u1.Email
                                    , c.CompanyId,c.CompanyName,c.CompanyNo
                                    , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
                                    FROM RFI.DesignSignFlow a
                                    INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                    INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                    INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                    WHERE a.DesignId = @NewDesignId
                                    AND a.SortNumber = 1";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { NewDesignId }) ?? throw new SystemException("【新設計申請單審批流程】資料錯誤!");
                            signJobName = result.SignJobName;
                            flowerName = result.UserName;
                            #endregion

                            #region //取得新設計申請單
                            sql = @"SELECT TOP 1 a.DesignId
								    , a.DesignNo, a.Edition, a.[Status] DesignStatus
								    , s.StatusName DesignStatusName, h.DepartmentName, i.CompanyName, i.CompanyNo
								    , a.SalesId, ISNULL(u.UserNo,'') SalesNo, ISNULL(u.UserName,'') SalesName, u.Gender SalesGender 
								    , a.SuperSalesId, ISNULL(t.UserNo,'') SuperSalesNo, ISNULL(t.UserName,'') SuperSalesName, t.Gender SuperSalesGender
								    FROM RFI.DemandDesign a
								    INNER JOIN BAS.[User] u ON u.UserId = a.SalesId
								    INNER JOIN BAS.[User] t ON t.UserId = a.SuperSalesId
								    INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'RfiDesign.Status'
								    INNER JOIN BAS.Department h ON h.DepartmentId = u.DepartmentId
								    INNER JOIN BAS.Company i ON i.CompanyId = h.CompanyId
								    WHERE a.DesignId = @NewDesignId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { NewDesignId }) ?? throw new SystemException("【新設計申請單】資料錯誤!");
                            designNo = result.DesignNo;
                            designEdition = result.Edition;
                            designStatusName = result.DesignStatusName;
                            #endregion
                        }
                        else
                        {
                            #region //取得RFI單頭資料
                            sql = @"SELECT TOP 1 s2.StatusName RfiStatusName
                                    FROM RFI.RequestForInformation a
                                    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.[Status] AND s2.StatusSchema = 'RequestForInformation.Status'
                                    WHERE a.RfiId = @RfiId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId }) ?? throw new SystemException("【市場評估單】資料錯誤!");
                            rfiStatusName = result.RfiStatusName;
                            #endregion

                            #region //取得RFI單身資料
                            sql = @"SELECT aa.RfiId, aa.RfiDetailId, aa.RfiSequence, aa.Edition
							        , aa.SalesId, ad.UserNo SalesNo, ad.UserName SalesName, ad.Gender SalesGender
							        , FORMAT(aa.ProdDate, 'yyyy-MM-dd') ProdDate
							        , FORMAT(aa.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart
							        , FORMAT(aa.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
							        , aa.TargetUnitPrice, aa.LifeCycleQty, aa.Revenue, aa.RevenueFC
							        , aa.ExchangeRate, aa.Currency, ac.TypeNo CurrencyTypeName, ac.TypeName CurrencyName
							        , aa.[Status] RfiDetailStatus, ae.StatusName RfiDetailStatusName
							        , ISNULL(FORMAT(aa.ConfirmFinalTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmFinalTime
							        FROM RFI.RfiDetail aa
							        LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.Currency AND ac.TypeSchema = 'ExchangeRate.Currency'
							        INNER JOIN BAS.[User] ad ON ad.UserId = aa.SalesId
							        INNER JOIN BAS.[Status] ae ON ae.StatusNo = aa.[Status] AND ae.StatusSchema = 'RfiDetail.Status'
							        WHERE aa.RfiDetailId = @RfiDetailId";
                            var resultRfiDetail = sqlConnection.Query(sql, new { RfiDetailId });
                            #endregion

                            #region //Mail內容明細
                            int i = 1;
                            foreach (var detail in resultRfiDetail)
                            {
                                detailHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>",
                                                detail.RfiSequence, detail.Edition, detail.ProdDate, detail.ProdLifeCycleStart, detail.ProdLifeCycleEnd
                                                , string.Format("{0:#,##0.######}", detail.TargetUnitPrice)
                                                , detail.LifeCycleQty.ToString("N0")
                                                , string.Format("{0:#,##0.######}", detail.ExchangeRate)
                                                , string.Format("{0:#,##0.######}", detail.Revenue)
                                                , detail.RfiDetailStatusName);
                                i++;
                            }

                            detailHtml = @"<table border>
								            <thead>
									            <tr>
										            <th>流水號</th><th>評估版次</th><th>量產時間</th><th>生產生命(起)</th><th>生產生命(迄)</th><th>目標單價</th><th>預估數量</th><th>匯率</th><th>營收</th><th>單據狀態</th>
									            </tr>
								            </thead>
                                          <tbody>" + detailHtml + "</tbody></table>";
                            #endregion
                        }

                        #region //Mail資料
                        sql = @"SELECT a.MailSubject, a.MailContent
							    , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(a.MailFrom, '') 'MailFrom'
                                , ISNULL(a.MailTo, '') 'MailTo'
                                , ISNULL(a.MailCc, '') 'MailCc'
                                , ISNULL(a.MailBcc, '') 'MailBcc'
							    FROM BAS.Mail a
							    LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
							    WHERE EXISTS (
								    SELECT ca.MailId
								    FROM BAS.MailSendSetting ca
								    WHERE ca.MailId = a.MailId
								    AND ca.SettingSchema = @SettingSchema
								    AND ca.SettingNo = @SettingNo
							    )";
                        var mailTemplate = sqlConnection.Query(sql, new { SettingSchema, SettingNo = "Y" });
                        if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        string hyperLink = string.Empty;
                        string pageUrl = string.Empty;

                        foreach (var item in mailTemplate)
                        {
                            string mailSubject = item.MailSubject;
                            string mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            switch (SettingSchema)
                            {
                                case "RfiSignFlow":
                                    #region //取得流程檔
                                    sql = @"SELECT a.RfiSfId
                                            , a.SignJobName
                                            , ISNULL(a.Approved, 'W') Approved, a.Edition
                                            , u1.UserId, u1.UserName FlowerName, u1.UserNo, u1.Email
                                            , c.CompanyId, c.CompanyName, c.CompanyNo
                                            , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                                        , ISNULL(a.SignDesc, '-') SignDesc
	                                        , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
	                                        , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
	                                        , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
	                                        , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
	                                        , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                                            , ISNULL(s2.StatusName, '待審批') ApprovedName
                                            FROM RFI.RfiSignFlow a
                                            INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                            INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                            INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                            LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.Approved AND s2.StatusSchema = 'RfiSign.Status'
                                            WHERE a.RfiDetailId = @RfiDetailId
                                            AND a.Edition = @Edition
                                            ORDER BY a.SortNumber";
                                    var resultRfiSignFlow = sqlConnection.Query(sql, new { RfiDetailId, Edition = string.Format("{0:00}", RfiFlowEdition), CreateDate }).ToList();
                                    #endregion

                                    string RfiEdition = string.Empty;
                                    int i = 1;

                                    #region //Mail內容明細
                                    resultRfiSignFlow.ForEach(x =>
                                    {
                                        RfiEdition = x.Edition;
                                        flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                            x.SignJobName,
                                            x.FlowerName,
                                            x.ApprovedName,
                                            string.Format("{0}天{1}時{2}分", x.StayDays, x.StayHours, x.StayMinutes),
                                            x.ArriveDateTime,
                                            x.SignDateTime,
                                            x.SignDesc);
                                        i++;
                                    });
                                    flowsHtml = @"<table border>
                                        <thead>
                                            <tr>
                                                <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                            </tr>
                                        </thead>
                                      <tbody>" + flowsHtml + "</tbody></table>";
                                    #endregion

                                    hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/UpdateRfiDetailStatus", "RequestForInformation/RfiManagement");
                                    pageUrl = hyperLink;
                                    hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));
                                    mailContent = mailContent.Replace("[RfiNo]", rfiNo)
                                                             .Replace("[Edition]", RfiEdition)
                                                             .Replace("[SignJobName]", signJobName)
                                                             .Replace("[FlowerName]", flowerName)
                                                             .Replace("[CustomerName]", customerName)
                                                             .Replace("[CustomerEName]", customerEName)
                                                             .Replace("[AnalysisDate]", analysisDate)
                                                             .Replace("[ProTypeGroupName]", proTypeGroupName)
                                                             .Replace("[ProductUseName]", productUseName)
                                                             .Replace("[ExistFlagName]", existFlagName)
                                                             .Replace("[RfiStatusName]", rfiStatusName)
                                                             .Replace("[SalesName]", salesName)
                                                             .Replace("[MailContent]", string.Format(mailFlowContent, rfiNo));
                                                            //.Replace("[FlowContent]", flowsHtml)
                                                            //.Replace("[RfiDetailContent]", detailHtml)
                                                            //.Replace("[hyperlink]", hyperLink)
                                    break;
                                case "RfiDesignSignFlow":
                                    #region //取得流程檔
                                    sql = @"SELECT a.DesignSfId
                                            , a.SignJobName
                                            , ISNULL(a.Approved, 'W') Approved
                                            , u1.UserId, u1.UserName FlowerName, u1.UserNo, u1.Email
                                            , c.CompanyId, c.CompanyName, c.CompanyNo
                                            , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                                        , ISNULL(a.SignDesc, '-') 'SignDesc'
	                                        , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
	                                        , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
	                                        , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
	                                        , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
	                                        , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                                            , ISNULL(s2.StatusName, '待審批') ApprovedName
                                            FROM RFI.DesignSignFlow a
                                            INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                            INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                            INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                            LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.Approved AND s2.StatusSchema = 'RfiSign.Status'
                                            WHERE a.DesignId = @NewDesignId
                                            ORDER BY a.SortNumber";
                                    var resultDesignSignFlow = sqlConnection.Query(sql, new { NewDesignId, CreateDate }).ToList();
                                    #endregion

                                    #region //Mail內容明細
                                    int j = 1;
                                    resultDesignSignFlow.ForEach(x =>
                                    {
                                        flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                            x.SignJobName,
                                            x.FlowerName,
                                            x.ApprovedName,
                                            string.Format("{0}天{1}時{2}分", x.StayDays, x.StayHours, x.StayMinutes),
                                            x.ArriveDateTime,
                                            x.SignDateTime,
                                            x.SignDesc);
                                        j++;
                                    });
                                    flowsHtml = @"<table border>
                                        <thead>
                                            <tr>
                                                <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                            </tr>
                                        </thead>
                                      <tbody>" + flowsHtml + "</tbody></table>";
                                    #endregion

                                    hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/UpdateRfiDetailStatus", "RequestForInformation/DesignManagement");
                                    pageUrl = hyperLink;
                                    hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));
                                    mailContent = mailContent.Replace("[DesignNo]", designNo)
                                                             .Replace("[Edition]", designEdition)
                                                             .Replace("[SignJobName]", signJobName)
                                                             .Replace("[FlowerName]", flowerName)
                                                             .Replace("[CustomerName]", customerName)
                                                             .Replace("[CustomerEName]", customerEName)
                                                             .Replace("[AssemblyName]", assemblyName)
                                                             .Replace("[ProTypeGroupName]", proTypeGroupName)
                                                             .Replace("[ProductUseName]", productUseName)
                                                             .Replace("[DesignStatusName]", designStatusName)
                                                             .Replace("[SalesName]", salesName)
                                                             .Replace("[MailContent]", string.Format(mailFlowContent, rfiNo));
                                                            //.Replace("[FlowContent]", flowsHtml)
                                                            //.Replace("[hyperlink]", hyperLink)
                                    break;
                            }

                            #region //MAMO個人訊息推播
                            string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                            ConvertToMarkdown(detailHtml, out string detailMamo);
                            ConvertToMarkdown(flowsHtml, out string flowsMamo);
                            mamoContent = mamoContent.Replace("[RfiDetailContent]", "\r\n" + detailMamo)
                                                     .Replace("[FlowContent]", "\r\n" + flowsMamo)
                                                     .Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                            if (NextFlow != null)
                            {
                                mamoHelper.SendMessage(NextFlow.CompanyNo, CurrentUser, "Personal", NextFlow.UserNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                //var companyNos = nextFlow.Where(w => w.CompanyNo != null).Select(s => { return (string)s.CompanyNo; }).Distinct().ToList();
                                //foreach (var companyNo in companyNos)
                                //{
                                //    var userNos = nextFlow.Where(w => w.CompanyNo == companyNo && w.UserNo != null).Select(s => { return (string)s.UserNo; }).Distinct().ToList();
                                //    foreach (var userNo in userNos)
                                //    {
                                //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                //    }
                                //}
                            }
                            else
                            {
                                mamoHelper.SendMessage(salesCompanyNo, CurrentUser, "Personal", salesNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                            }
                            #endregion

                            #region //發送
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = MailTo, //通知下一關或建單業務
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent.Replace("[RfiDetailContent]", detailHtml).Replace("[FlowContent]", flowsHtml).Replace("[hyperlink]", hyperLink),
                                TextBody = "-"
                            };
                            BaseHelper.MailSend(mailConfig);
                            #endregion
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

        #region //RfiUrgentSignMessage 發送RFI催簽通知信 -- Chia Yuan 2024.10.22
        public string RfiUrgentSignMessage(int RfiDetailId, int RfiSfId, string UrgentSignContent)
        {
            try
            {
                if (RfiDetailId < 0 || RfiSfId < 0) throw new SystemException("【單據】資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得流程資料
                    sql = @"SELECT a.RfiSfId
                            , a.SignJobName
                            , ISNULL(a.Approved, 'W') Approved, a.Edition
                            , flower.UserId FlowerId, flower.UserName FlowerName, flower.UserNo FlowerNo, flower.Email FlowerEmail
                            , flower.DepartmentName + '-' + flower.UserName + ':' + flower.Email MailTo
                            , sales.UserId SalesId, sales.UserName SalesName, sales.UserNo SalesNo, sales.Email SalesEmail, sales.CompanyNo SalesCompanyNo
                            , ISNULL(a.SignDesc, '-') SignDesc
                            , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
                            , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
                            , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                            , FORMAT(b.ProdDate, 'yyyy-MM-dd HH:mm') ProdDate
                            , FORMAT(b.ProdLifeCycleStart, 'yyyy-MM-dd HH:mm') ProdLifeCycleStart
                            , FORMAT(b.ProdLifeCycleEnd, 'yyyy-MM-dd HH:mm') ProdLifeCycleEnd
                            , b.TargetUnitPrice, b.RfiSequence, b.TargetUnitPrice, b.LifeCycleQty, b.Revenue, b.RevenueFC, b.ExchangeRate, b.Currency
                            , e.RfiNo, e.CustomerName, e.CustomerEName, FORMAT(e.AnalysisDate, 'yyyy-MM-dd HH:mm') AnalysisDate
                            , f.ProTypeGroupName, g.ProductUseName
                            , s1.StatusName RfiStatusName, s2.StatusName ExistFlagName, s3.StatusName RfiDetailStatusName, ISNULL(s4.StatusName, '待審批') ApprovedName
                            FROM RFI.RfiSignFlow a
                            INNER JOIN RFI.RfiDetail b ON b.RfiDetailId = a.RfiDetailId
                            INNER JOIN RFI.RequestForInformation e ON e.RfiId = b.RfiId
                            INNER JOIN SCM.ProductTypeGroup f ON f.ProTypeGroupId = e.ProTypeGroupId
                            INNER JOIN SCM.ProductUse g ON g.ProductUseId = e.ProductUseId
                            INNER JOIN (
	                            SELECT au.UserId, au.UserNo, au.UserName, au.Email
	                            , ad.DepartmentNo, ad.DepartmentName, ac.CompanyNo, ac.CompanyName
	                            FROM BAS.[User] au
	                            INNER JOIN BAS.Department ad ON ad.DepartmentId = au.DepartmentId
	                            INNER JOIN BAS.Company ac ON ac.CompanyId = ad.CompanyId
                            ) flower ON flower.UserId = a.FlowUser
                            INNER JOIN (
	                            SELECT au.UserId, au.UserNo, au.UserName, au.Email
	                            , ad.DepartmentNo, ad.DepartmentName, ac.CompanyNo, ac.CompanyName
	                            FROM BAS.[User] au
	                            INNER JOIN BAS.Department ad ON ad.DepartmentId = au.DepartmentId
	                            INNER JOIN BAS.Company ac ON ac.CompanyId = ad.CompanyId
                            ) sales ON sales.UserId = e.SalesId
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = e.[Status] AND s1.StatusSchema = 'RequestForInformation.Status'
                            INNER JOIN BAS.[Status] s2 ON s2.StatusNo = e.ExistFlag AND s2.StatusSchema = 'Boolean'
                            INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.[Status] AND s3.StatusSchema = 'RfiDetail.Status'
                            LEFT JOIN BAS.[Status] s4 ON s4.StatusNo = a.Approved AND s4.StatusSchema = 'RfiSign.Status'
                            INNER JOIN (
	                            SELECT aa.RfiSfId, DENSE_RANK () OVER (PARTITION BY aa.RfiDetailId ORDER BY aa.Edition DESC) 'SortEdition'
	                            FROM RFI.RfiSignFlow aa
	                            INNER JOIN RFI.RfiDetail ab ON ab.RfiDetailId = aa.RfiDetailId
                                INNER JOIN BAS.[Status] s1 ON s1.StatusNo = aa.[Status] AND StatusSchema = 'RfiDetail.FlowStatus'
                            ) flow ON flow.RfiSfId = a.RfiSfId
                            WHERE a.RfiDetailId = @RfiDetailId
                            AND flow.SortEdition = 1
                            ORDER BY a.SortNumber";
                    var result = sqlConnection.Query(sql, new { RfiDetailId, CreateDate });
                    var rfiDetail = result.FirstOrDefault(f => f.RfiSfId == RfiSfId);
                    #endregion

                    #region //Mail資料
                    sql = @"SELECT a.MailSubject, a.MailContent
						    , b.Host, b.Port, b.SendMode, b.Account, b.Password
                            , ISNULL(a.MailFrom, '') 'MailFrom'
                            , ISNULL(a.MailTo, '') 'MailTo'
                            , ISNULL(a.MailCc, '') 'MailCc'
                            , ISNULL(a.MailBcc, '') 'MailBcc'
						    FROM BAS.Mail a
						    LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
						    WHERE EXISTS (
							    SELECT ca.MailId
							    FROM BAS.MailSendSetting ca
							    WHERE ca.MailId = a.MailId
							    AND ca.SettingSchema = @SettingSchema
							    AND ca.SettingNo = @SettingNo
						    )";
                    var mailTemplate = sqlConnection.Query(sql, new { SettingSchema = "RfiSignFlow", SettingNo = "Y" });
                    if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                    #endregion

                    string flowsHtml = string.Empty;
                    string detailHtml = string.Empty;

                    #region //Mail內容明細
                    foreach (var flow in result)
                    {
                        flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                    flow.SignJobName,
                                    flow.FlowerName,
                                    flow.ApprovedName,
                                    string.Format("{0}天{1}時{2}分", flow.StayDays, flow.StayHours, flow.StayMinutes),
                                    flow.ArriveDateTime,
                                    flow.SignDateTime,
                                    flow.SignDesc);
                    }
                    flowsHtml = @"<table border><thead><tr>
                                    <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                </tr></thead><tbody>" + flowsHtml + "</tbody></table>";

                    detailHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>",
                                rfiDetail.RfiSequence, rfiDetail.Edition, rfiDetail.ProdDate, rfiDetail.ProdLifeCycleStart, rfiDetail.ProdLifeCycleEnd
                                , string.Format("{0:#,##0.######}", rfiDetail.TargetUnitPrice)
                                , rfiDetail.LifeCycleQty.ToString("N0")
                                , string.Format("{0:#,##0.######}", rfiDetail.ExchangeRate)
                                , string.Format("{0:#,##0.######}", rfiDetail.Revenue)
                                , rfiDetail.RfiDetailStatusName);
                    detailHtml = @"<table border><thead><tr>
                                    <th>流水號</th><th>評估版次</th><th>量產時間</th><th>生產生命(起)</th><th>生產生命(迄)</th><th>目標單價</th><th>預估數量</th><th>匯率</th><th>營收</th><th>單據狀態</th>
                                </tr></thead><tbody>" + detailHtml + "</tbody></table>";
                    #endregion

                    foreach (var item in mailTemplate)
                    {
                        string mailSubject = item.MailSubject;
                        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                        string hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/RfiUrgentSignMessage", "RequestForInformation/RfiManagement");
                        string pageUrl = hyperLink;
                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));
                        mailContent = mailContent.Replace("[RfiNo]", rfiDetail.RfiNo)
                                                 .Replace("[Edition]", rfiDetail.Edition)
                                                 .Replace("[SignJobName]", rfiDetail.SignJobName)
                                                 .Replace("[FlowerName]", rfiDetail.FlowerName)
                                                 .Replace("[CustomerName]", rfiDetail.CustomerName)
                                                 .Replace("[CustomerEName]", rfiDetail.CustomerEName)
                                                 .Replace("[AnalysisDate]", rfiDetail.AnalysisDate)
                                                 .Replace("[ProTypeGroupName]", rfiDetail.ProTypeGroupName)
                                                 .Replace("[ProductUseName]", rfiDetail.ProductUseName)
                                                 .Replace("[ExistFlagName]", rfiDetail.ExistFlagName)
                                                 .Replace("[RfiStatusName]", rfiDetail.RfiStatusName)
                                                 .Replace("[SalesName]", rfiDetail.SalesName)
                                                 .Replace("[MailContent]", string.Format("【市場評估單】已送達-單號{0}，請盡速查看並審批!\r\n{1}", rfiDetail.RfiNo, UrgentSignContent));
                                                //.Replace("[FlowContent]", flowsHtml)

                        #region //MAMO個人訊息推播
                        string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                        ConvertToMarkdown(detailHtml, out string detailMamo);
                        ConvertToMarkdown(flowsHtml, out string flowsMamo);
                        mamoContent = mamoContent.Replace("[RfiDetailContent]", "\r\n" + detailMamo)
                                                 .Replace("[FlowContent]", "\r\n" + flowsMamo)
                                                 .Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                        mamoHelper.SendMessage(rfiDetail.SalesCompanyNo, CurrentUser, "Personal", rfiDetail.FlowerNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                        #endregion

                        #region //發送
                        MailConfig mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = mailSubject,
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = rfiDetail.FlowerEmail,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = mailContent.Replace("[RfiDetailContent]", detailHtml).Replace("[FlowContent]", flowsHtml).Replace("[hyperlink]", hyperLink),
                            TextBody = "-"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }

                    #region //更新催簽次數
                    sql = @"UPDATE a SET 
                            a.UrgentSign = a.UrgentSign + 1,
                            a.LastModifiedDate = @LastModifiedDate,
                            a.LastModifiedBy = @LastModifiedBy
                            FROM RFI.RfiSignFlow a
                            WHERE a.RfiSfId = @RfiSfId";
                    int rowsAffected = sqlConnection.Execute(sql,
                        new
                        {
                            rfiDetail.RfiSfId,
                            LastModifiedDate,
                            LastModifiedBy
                        });
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = string.Format("已通知下一關：{0}({1})", rfiDetail.FlowerName, rfiDetail.SignJobName)
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

        #region //UpdateRfiSignFlowAdditionalFile RFI審批附檔資料更新 -- Chia Yuan 2024.06.25
        public string UpdateRfiSignFlowAdditionalFile(int RfiSfId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 AdditionalFile
                                FROM RFI.RfiSignFlow a
                                WHERE a.RfiSfId = @RfiSfId";
                        dynamicParameters.Add("RfiSfId", RfiSfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【流程檔】資料錯誤!");
                        int? additionalFile = -1;
                        foreach (var item in result)
                        {
                            additionalFile = item.AdditionalFile;
                        }
                        #endregion

                        #region //判斷附檔資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【附檔】資料錯誤!");
                        #endregion

                        #region //主資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.AdditionalFile = @AdditionalFile,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.RfiSignFlow a
                                WHERE a.RfiSfId = @RfiSfId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfiSfId,
                                AdditionalFile = FileId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //舊附檔刪除
                        if (additionalFile != null)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 AdditionalFile
                                    FROM RFI.RfiSignFlow
                                    WHERE AdditionalFile = @AdditionalFile
                                    AND RfiSfId <> @RfiSfId";
                            dynamicParameters.Add("AdditionalFile", additionalFile);
                            dynamicParameters.Add("RfiSfId", RfiSfId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (!result.Any())
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @AdditionalFile";
                                dynamicParameters.Add("AdditionalFile", additionalFile);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateDemandDesignStatus -- 新設計申請單狀態更新 -- Chia Yuan 2024.01.24
        public string UpdateDemandDesignStatus(int DesignId, int DesignSfId, string DesignStatus, string SignContent)
        {
            try
            {
                if (!Regex.IsMatch(DesignStatus, "^(Y|N|B|y|v)$", RegexOptions.IgnoreCase)) throw new SystemException("【審批註記】資料錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        sql = @"SELECT TOP 1 a.RfiId, a.CompanyId, a.DepartmentId, a.DesignId, a.DesignNo, a.Edition, a.AssemblyName
                                , a.[Status], s1.StatusName DesignStatusName
                                , a.SalesId, ISNULL(u1.UserNo,'') SalesNo, ISNULL(u1.UserName,'') SalesName, u1.Gender SalesGender, u1.Email SalesEmail
                                , b.CustomerName, b.CustomerEName
                                , b.ProTypeGroupId, g.ProTypeGroupName
                                , d.ProductUseId, d.ProductUseName, d.ProductUseNo
                                , f.CompanyNo, f.CompanyId SalesCompanyId
                                FROM RFI.DemandDesign a
                                INNER JOIN RFI.RequestForInformation b ON b.RfiId = a.RfiId
                                INNER JOIN SCM.ProductUse d ON d.ProductUseId = b.ProductUseId
                                INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = b.ProTypeGroupId
                                INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.[Status] AND s1.StatusSchema = 'RfiDesign.Status'
                                INNER JOIN BAS.[User] u1 ON u1.UserId = a.SalesId
                                INNER JOIN BAS.Department e ON e.DepartmentId = u1.DepartmentId
                                INNER JOIN BAS.Company f ON f.CompanyId = e.CompanyId
                                WHERE a.DesignId = @DesignId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { DesignId }) ?? throw new SystemException("【新設計申請單】資料錯誤!");
                        int RfiId = result.RfiId;
                        int ProTypeGroupId = result.ProTypeGroupId;
                        int salesId = result.SalesId;
                        int DepartmentId = result.DepartmentId;
                        string designNo = result.DesignNo;
                        string assemblyName = result.AssemblyName;
                        string designEdition = result.Edition;
                        string designStatus = result.Status; //審批狀態
                        string MailTo = result.SalesEmail; //業務email
                        string salesNo = result.SalesNo;
                        string salesName = result.SalesName;
                        string salesCompanyNo = result.CompanyNo;

                        string customerName = result.CustomerName;
                        string customerEName = result.CustomerEName;
                        string proTypeGroupName = result.ProTypeGroupName;
                        string productUseName = result.ProductUseName;
                        #endregion

                        #region //判斷流程檔是否正確
                        int MaxSortNumber = -1;
                        sql = @"SELECT TOP 1 SortNumber
                                FROM RFI.DesignSignFlow
                                WHERE DesignSfId = @DesignSfId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DesignSfId }) ?? throw new SystemException("【流程檔】資料錯誤!");
                        MaxSortNumber = result.SortNumber;
                        #endregion

                        #region //判斷狀態資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status] a 
                                WHERE a.StatusSchema = 'RfiSign.Status'
                                AND a.StatusNo = @DesignStatus";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DesignStatus }) ?? throw new SystemException("【審批狀態】資料錯誤!");
                        #endregion

                        SignContent = string.IsNullOrWhiteSpace(SignContent) ? null : SignContent.Trim();
                        string detailCode = string.Empty;
                        string SettingSchema = "RfiDesignSignFlow";
                        string mailFlowContent = "【新設計申請單】已送達-單號{0}，請查看內容並審批!";

                        #region //審批流程更新
                        sql = @"UPDATE a SET
                                a.Approved = @DesignStatus,
                                a.SignUser = @CurrentUser,
                                a.SignDesc = @SignContent,
                                a.SignDateTime = @SignDateTime,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.DesignSignFlow a
                                WHERE a.DesignSfId = @DesignSfId";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                DesignSfId,
                                DesignStatus,
                                CurrentUser,
                                SignContent,
                                SignDateTime = CreateDate,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        #endregion

                        string tempDesignStatus = designStatus;
                        string tempRfiStatus = string.Empty;
                        string tempRfiDetailStatus = string.Empty;
                        string signJobName = string.Empty;
                        string flowerName = string.Empty;
                        int NextDesignSfId = -1;

                        #region //取得下一關
                        sql = @"SELECT TOP 1 a.DesignSfId
                                , a.SignJobName, a.Approved
                                , u1.UserId, u1.UserName, u1.UserNo, u1.Email
                                , c.CompanyId, c.CompanyName, c.CompanyNo
                                , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
                                FROM RFI.DesignSignFlow a
                                INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                WHERE a.DesignId = @DesignId
                                AND a.SortNumber = @SortNumber";
                        var NextFlow = sqlConnection.QueryFirstOrDefault(sql, new { DesignId, SortNumber = MaxSortNumber + 1 });
                        if (NextFlow != null)
                        {
                            NextDesignSfId = NextFlow.DesignSfId;
                            signJobName = NextFlow.SignJobName;
                            flowerName = NextFlow.UserName;
                            MailTo = NextFlow.Email;
                        }
                        #endregion

                        switch (DesignStatus)
                        {
                            case "Y":
                                if (NextDesignSfId > 0)
                                {
                                    #region //下一關狀態更新
                                    sql = @"UPDATE a SET
                                            a.ArriveDateTime = @CreateDate,
                                            a.LastModifiedDate = @LastModifiedDate,
                                            a.LastModifiedBy = @LastModifiedBy
                                            FROM RFI.DesignSignFlow a
                                            WHERE a.DesignSfId = @NextDesignSfId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            NextDesignSfId,
                                            CreateDate,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion
                                }
                                else
                                { //沒有下一關
                                    tempDesignStatus = "6"; //完成評估
                                    tempRfiDetailStatus = "6"; //回饋客戶
                                    tempRfiStatus = "3"; //客戶確認中
                                    signJobName = "業務";
                                    flowerName = salesName; //業務名稱

                                    #region //取得最後一筆RFI單身
                                    sql = @"SELECT TOP 1 MAX(a.RfiDetailId) RfiDetailId FROM RFI.RfiDetail a WHERE a.RfiId = @RfiId";
                                    result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                                    int RfiDetailId = result.RfiDetailId;
                                    #endregion

                                    #region //單據狀態更新
                                    sql = @"UPDATE a SET
								    a.Status = @Status,
								    a.LastModifiedDate = @LastModifiedDate,
								    a.LastModifiedBy = @LastModifiedBy
								    FROM RFI.RfiDetail a
								    WHERE a.RfiDetailId = @RfiDetailId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            RfiDetailId,
                                            Status = tempRfiDetailStatus,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion

                                    #region //單據狀態更新
                                    sql = @"UPDATE a SET
								    a.Status = @Status,
								    a.LastModifiedDate = @LastModifiedDate,
								    a.LastModifiedBy = @LastModifiedBy
								    FROM RFI.RequestForInformation a
								    WHERE a.RfiId = @RfiId";
                                    rowsAffected += sqlConnection.Execute(sql,
                                        new
                                        {
                                            RfiId,
                                            Status = tempRfiStatus,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    #endregion
                                }
                                break;
                            case "B":
                                mailFlowContent = "【新設計申請單】已被退回-單號{0}，請前往系統查看!";
                                tempDesignStatus = "7"; //退回/變更
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱
                                break;
                            case "v":
                                mailFlowContent = "【市場評估單】已作廢-單號{0}，請前往系統查看。";
                                tempDesignStatus = "V"; //作廢
                                signJobName = "業務";
                                flowerName = salesName; //業務名稱
                                break;
                        }

                        if (designStatus != tempDesignStatus)
                        {
                            #region //單據狀態更新
                            sql = @"UPDATE a SET
                                    a.Status = @Status,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.DemandDesign a
                                    WHERE a.DesignId = @DesignId";
                            rowsAffected += sqlConnection.Execute(sql, 
                                new
                                {
                                    DesignId,
                                    Status = tempDesignStatus,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            #endregion
                        }

                        #region //狀態更新 (停用)
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"UPDATE a SET
                        //        a.Status = @Status,";
                        //switch (Status)
                        //{
                        //    case "1":
                        //        //if (designStatus == "2") //業務主管 > 業務
                        //        //{
                        //        //    sql += @" a.SuperSales = @SignUser,
                        //        //            　a.SuperSalesDesc = @SignContent,
                        //        //            　a.ConfirmSTime = @ConfirmTime,";
                        //        //}
                        //        break;
                        //    case "2": //業務 > 業務主管
                        //        //if (designStatus == "4") //研發主管 > 業務主管
                        //        //{
                        //        //    sql += @" a.SuperRd = @SignUser,
                        //        //          　　a.SuperRdDesc = @SignContent,
                        //        //          　　a.ConfirmSTime = @ConfirmTime,";
                        //        //}
                        //        //else
                        //        //{
                        //        //    sql += @" a.Sales = @SignUser,
                        //        //          a.SalesDesc = @SignContent,
                        //        //          a.ConfirmSSTime = @ConfirmTime,";
                        //        //}
                        //        sql += @" a.Sales = @SignUser,
                        //                  a.SalesDesc = @SignContent,
                        //                  a.ConfirmSSTime = @ConfirmTime,";
                        //        detailCode = "SignSS";
                        //        break;
                        //    //case "3": //業務主管 > 研發
                        //    //    sql += @" a.SuperSales = @SignUser,
                        //    //              a.SuperSalesDesc = @SignContent,
                        //    //            　a.ConfirmRDTime = @ConfirmTime,";
                        //    //    detailCode = "SignRD";
                        //    //    break;
                        //    case "4": //業務主管 > 研發主管
                        //        //if (designStatus == "5") //總經理 > 研發主管
                        //        //{
                        //        //    sql += @" a.ManagerPlan = @SignUser,
                        //        //            　a.ManagerPlanDesc = @SignContent,
                        //        //            　a.ConfirmSTime = @ConfirmTime,";
                        //        //}
                        //        //else
                        //        //{
                        //        //    sql += @" a.SuperSales = @SignUser,
                        //        //          a.SuperSalesDesc = @SignContent,
                        //        //        　a.ConfirmSRDTime = @ConfirmTime,";
                        //        //}
                        //        sql += @" a.SuperSales = @SignUser,
                        //                　a.SuperSalesDesc = @SignContent,
                        //                　a.ConfirmSRDTime = @ConfirmTime,";
                        //        detailCode = "SignSRD";
                        //        break;
                        //    //case "4": //研發 > 研發主管
                        //    //    sql += @" a.Rd = @SignUser,
                        //    //              a.RdDesc = @SignContent,
                        //    //            　a.ConfirmSRDTime = @ConfirmTime,";
                        //    //    detailCode = "SignSRD";
                        //    //    break;
                        //    case "5": //研發主管 > 總經理
                        //        //if (designStatus == "6")
                        //        //{
                        //        //    sql += @" a.ManagerPlan = NULL,
                        //        //            　a.ManagerPlanDesc = @SignContent,
                        //        //            　a.ConfirmMPTime = NULL,";
                        //        //}
                        //        //else
                        //        //{
                        //        //    sql += @" a.SuperRd = @SignUser,
                        //        //          a.SuperRdDesc = @SignContent,
                        //        //          a.ConfirmMPTime = @ConfirmTime,";
                        //        //}
                        //        sql += @" a.SuperRd = @SignUser,
                        //                  a.SuperRdDesc = @SignContent,
                        //                  a.ConfirmMPTime = @ConfirmTime,";
                        //        detailCode = "SignMP";
                        //        break;
                        //    case "6": //總經理 > 業務(完成評估)
                        //        sql += @" a.ManagerPlan = @SignUser,
                        //                  a.ManagerPlanDesc = @SignContent,
                        //                  a.ConfirmFinalTime = @ConfirmTime,";
                        //        break;
                        //    //case "6": //業務 > 回饋客戶
                        //    //    sql += @" a.Sales = @SignUser,
                        //    //            　a.SalesDesc = @SignContent,
                        //    //              a.ConfirmFinalTime = @ConfirmTime,";
                        //    //    break;
                        //    case "7": //Flower > 不准
                        //        mailFlowContent = "新設計申請單已被否決，請前往系統查看!";
                        //        switch (designStatus)
                        //        {
                        //            case "1":
                        //                sql += @" a.Sales = @SignUser,
                        //                          a.SalesDesc = @SignContent,";
                        //                break;
                        //            case "2": //業務主管不准
                        //                sql += @" a.SuperSales = @SignUser,
                        //                          a.SuperSalesDesc = @SignContent,";
                        //                break;
                        //            //case "3": //研發不准
                        //            //    sql += @" a.Rd = @SignUser,
                        //            //              a.RdDesc = @SignContent,";
                        //            //    break;
                        //            case "4": //研發主管不准
                        //                sql += @" a.SuperRd = @SignUser,
                        //                          a.SuperRdDesc = @SignContent,";
                        //                break;
                        //            case "5": //總經理不准
                        //                sql += @" a.ManagerPlan = @SignUser,
                        //                          a.ManagerPlanDesc = @SignContent,";
                        //                break;
                        //        }
                        //        sql += @" a.ConfirmFinalTime = @ConfirmTime,";
                        //        break;
                        //}
                        //sql += @"a.LastModifiedDate = @LastModifiedDate,
                        //        a.LastModifiedBy = @LastModifiedBy
                        //        FROM RFI.DemandDesign a
                        //        WHERE a.DesignId = @DesignId";
                        //dynamicParameters.AddDynamicParams(
                        //    new
                        //    {
                        //        DesignId,
                        //        SignContent,
                        //        SignUser = LastModifiedBy,
                        //        ConfirmTime = LastModifiedDate,
                        //        Status,
                        //        LastModifiedDate,
                        //        LastModifiedBy
                        //    });
                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int NewDesignId = DesignId;
                        if (DesignStatus == "B")
                        {
                            #region //取得審批審批流程
                            sql = @"SELECT TOP 1 a.DepartmentId
                                    FROM RFI.TemplateDesignSignFlow a
                                    WHERE EXISTS (SELECT TOP 1 1 FROM SCM.ProductTypeGroup aa WHERE aa.ProTypeGroupId = a.ProTypeGroupId AND aa.[Status] = 'A')
                                    AND a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.DepartmentId = @DepartmentId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, DepartmentId }) ?? throw new SystemException("【新設計申請單審批流程】資料錯誤!");
                            #endregion

                            #region //單身目前序號、版次號
                            sql = @"SELECT TOP 1
								    ISNULL(CAST(MAX(DesignSequence) AS INT), 0) + 1 MaxSequence,
								    ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
								    FROM RFI.DemandDesign
								    WHERE RfiId = @RfiId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                            int maxSequence = Convert.ToInt32(result?.MaxSequence ?? 1);
                            int maxEdition = Convert.ToInt32(result?.MaxEdition ?? 1);
                            #endregion

                            #region //拋轉新設計申請單
                            sql = @"INSERT INTO RFI.DemandDesign(RfiId, CompanyId, DepartmentId, DesignNo, DesignSequence, AssemblyName, SalesId, SuperSalesId, AdditionalFile, DesignFile, Edition, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DesignId
                                    SELECT a.RfiId, a.CompanyId, a.DepartmentId, a.DesignNo, @DesignSequence, @AssemblyName, a.SalesId, a.SuperSalesId, a.AdditionalFile, a.DesignFile, @Edition, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    FROM RFI.DemandDesign a
                                    WHERE DesignId = @DesignId";
                            var insertResult = sqlConnection.QueryFirstOrDefault(sql,
                                new
                                {
                                    DesignId,
                                    AssemblyName = assemblyName,
                                    DesignSequence = string.Format("{0:0000}", maxSequence), //取4位數
                                    Edition = string.Format("{0:00}", maxEdition), //取2位數
                                    Status = "1", //審批中
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected += 1;
                            NewDesignId = insertResult?.DesignId ?? -1;
                            #endregion

                            #region //建立審批流程檔 --new
                            sql = @"INSERT RFI.DesignSignFlow
                                    SELECT @NewDesignId, ISNULL(a.FlowUser, @FlowUser), b.AdditionalFile
                                    , CASE a.SortNumber WHEN 1 THEN GETDATE() ELSE NULL END
                                    , NULL, a.FlowJobName, NULL, NULL, a.SortNumber, NULL, 0, a.FlowStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    FROM RFI.TemplateDesignSignFlow a
                                    OUTER APPLY (SELECT ba.AdditionalFile FROM RFI.DesignSignFlow ba WHERE ba.DesignId = @DesignId AND ba.SortNumber = a.SortNumber AND ba.[Status] = a.FlowStatus) b
                                    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    AND a.DepartmentId = @DepartmentId
                                    AND a.[Status] = 'A'
                                    ORDER BY a.SortNumber";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProTypeGroupId,
                                    DepartmentId,
                                    NewDesignId,
                                    DesignId,
                                    FlowUser = salesId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion

                            #region //設計規格新增
                            sql = @"INSERT INTO RFI.DemandDesignSpec(DesignId, a.ParameterName, a.SortNumber, a.ControlType, a.Required, a.Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DdSpecId
                                    SELECT @NewDesignId, a.ParameterName, a.SortNumber, a.ControlType, a.[Required], a.[Status]
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    FROM RFI.TemplateSpecParameter a
                                    WHERE a.ProTypeGroupId = @ProTypeGroupId
                                    ORDER BY a.SortNumber";
                            var insertResults = sqlConnection.Query(sql,
                                 new
                                 {
                                     NewDesignId,
                                     ProTypeGroupId,
                                     CreateDate,
                                     LastModifiedDate,
                                     CreateBy,
                                     LastModifiedBy
                                 }).ToList();
                            var NewDdSpecIds = insertResults.Select(s => s.DdSpecId).ToArray();
                            rowsAffected += insertResults.Count;
                            #endregion

                            #region //設計規格子資料新增(模板資料媒合)
                            sql = @"INSERT INTO RFI.DemandDesignSpecDetail(DdSpecId, PmtDetailName, SortNumber, DataType
                                    , Required, RequireFlag, DesignInput, Description, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    SELECT c.DdSpecId, a.PmtDetailName, a.SortNumber, a.DataType
                                    , ISNULL(d.[Required], a.[Required]) AS [Required], d.RequireFlag, d.DesignInput, d.[Description], ISNULL(d.[Status], 'S')
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    FROM RFI.TemplateSpecParameterDetail a
                                    INNER JOIN RFI.TemplateSpecParameter b ON b.TempSpId = a.TempSpId
                                    INNER JOIN RFI.DemandDesignSpec c ON c.ParameterName = b.ParameterName
                                    OUTER APPLY (
	                                    SELECT db.[Required], db.RequireFlag, db.DesignInput, db.[Description], db.[Status]
	                                    FROM RFI.DemandDesignSpec da
	                                    INNER JOIN RFI.DemandDesignSpecDetail db ON db.DdSpecId = da.DdSpecId
	                                    WHERE da.DesignId = @DesignId
	                                    AND LTRIM(RTRIM(da.ParameterName)) = LTRIM(RTRIM(b.ParameterName))
	                                    AND LTRIM(RTRIM(db.PmtDetailName)) = LTRIM(RTRIM(a.PmtDetailName))
                                    ) d
                                    WHERE c.DdSpecId IN @NewDdSpecIds
                                    AND b.ProTypeGroupId = @ProTypeGroupId";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    DesignId,
                                    NewDdSpecIds,
                                    ProTypeGroupId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion
                        }

                        #region //取得RFI單頭資料
                        //sql = @"SELECT TOP 1 a.RfiNo, a.[Status] RfiStatus, s2.StatusName RfiStatusName 
                        //        , a.ExistFlag, s.StatusName ExistFlagName
                        //        , a.CustomerName, a.CustomerEName
                        //        , a.ProTypeGroupId, g.ProTypeGroupName
                        //        , d.ProductUseId, d.ProductUseName, d.ProductUseNo
                        //        , c.UserName SalesName, c.UserNo SalesNo, c.Email SalesEmail
                        //        , ISNULL(FORMAT(a.CreateDate, 'yyyy-MM-dd'), '') AnalysisDate
                        //        FROM RFI.RequestForInformation a
                        //        INNER JOIN SCM.ProductUse d ON d.ProductUseId = a.ProductUseId
                        //        INNER JOIN SCM.ProductTypeGroup g ON g.ProTypeGroupId = a.ProTypeGroupId
                        //        INNER JOIN BAS.[User] c ON c.UserId = a.SalesId
                        //        INNER JOIN BAS.[Status] s ON s.StatusNo = a.ExistFlag AND s.StatusSchema = 'Boolean'
                        //        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.[Status] AND s2.StatusSchema = 'RequestForInformation.Status'
                        //        WHERE a.RfiId = @RfiId";
                        //var result1 = sqlConnection.QueryFirstOrDefault(sql, new { RfiId });
                        //string customerName = result1.CustomerName;
                        //string customerEName = result1.CustomerEName;
                        //string proTypeGroupName = result1.ProTypeGroupName;
                        //string productUseName = result1.ProductUseName;
                        #endregion

                        #region //取得新設計申請單頭資料
                        //sql = @"SELECT TOP 1 s1.StatusName DesignStatusName
                        //        FROM RFI.DemandDesign a
                        //        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.[Status] AND s1.StatusSchema = 'RfiDesign.Status'
                        //        WHERE a.DesignId = @DesignId";
                        sql = @"SELECT TOP 1 s1.StatusName
                                FROM BAS.[Status] s1
                                WHERE s1.StatusSchema = 'RfiDesign.Status' 
                                AND s1.StatusNo = @tempDesignStatus";                        
                        result = sqlConnection.QueryFirstOrDefault(sql, new { tempDesignStatus }) ?? throw new SystemException("【新設計申請單】資料錯誤!");
                        string designStatusName = result.StatusName;
                        #endregion

                        #region //取得流程檔
                        sql = @"SELECT a.DesignSfId
                                , a.SignJobName
                                , ISNULL(a.Approved, 'W') Approved
                                , u1.UserId, u1.UserName FlowerName, u1.UserNo, u1.Email
                                , c.CompanyId, c.CompanyName, c.CompanyNo
                                , d.DepartmentName + '-' + u1.UserName + ':' + u1.Email MailTo
	                            , ISNULL(a.SignDesc, '-') SignDesc
	                            , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
	                            , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
	                            , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
	                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
	                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                                , ISNULL(s2.StatusName, '待審批') ApprovedName
                                FROM RFI.DesignSignFlow a
                                INNER JOIN BAS.[User] u1 ON u1.UserId = a.FlowUser
                                INNER JOIN BAS.Department d ON d.DepartmentId = u1.DepartmentId
                                INNER JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                                LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.Approved AND s2.StatusSchema = 'RfiSign.Status'
                                WHERE a.DesignId = @NewDesignId
                                ORDER BY a.SortNumber";
                        var resultDesignSignFlow = sqlConnection.Query(sql, new { NewDesignId, CreateDate }).ToList();
                        #endregion

                        string flowsHtml = string.Empty;
                        string hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/UpdateRfiDetailStatus", "RequestForInformation/DesignManagement");
                        string pageUrl = hyperLink;
                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));

                        #region //Mail內容明細
                        int i = 1;
                        resultDesignSignFlow.ForEach(x =>
                        {
                            flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                x.SignJobName,
                                x.FlowerName,
                                x.ApprovedName,
                                string.Format("{0}天{1}時{2}分", x.StayDays, x.StayHours, x.StayMinutes),
                                x.ArriveDateTime,
                                x.SignDateTime,
                                x.SignDesc);
                            i++;
                        });
                        flowsHtml = @"<table border><thead><tr>
                                        <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                    </tr></thead><tbody>" + flowsHtml + "</tbody></table>";
                        #endregion

                        #region //Mail資料
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(a.MailFrom, '') 'MailFrom'
                                , ISNULL(a.MailTo, '') 'MailTo'
                                , ISNULL(a.MailCc, '') 'MailCc'
                                , ISNULL(a.MailBcc, '') 'MailBcc'
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                WHERE EXISTS (
	                                SELECT ca.MailId
	                                FROM BAS.MailSendSetting ca
	                                WHERE ca.MailId = a.MailId
	                                AND ca.SettingSchema = @SettingSchema
	                                AND ca.SettingNo = @SettingNo
                                )";
                        var mailTemplate = sqlConnection.Query(sql, new { SettingSchema, SettingNo = "Y"});
                        if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in mailTemplate)
                        {
                            string mailSubject = item.MailSubject;
                            string mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            switch (SettingSchema)
                            {
                                case "RfiDesignSignFlow":
                                    mailContent = mailContent.Replace("[DesignNo]", designNo)
                                                             .Replace("[Edition]", designEdition)
                                                             .Replace("[SignJobName]", signJobName)
                                                             .Replace("[FlowerName]", flowerName)
                                                             .Replace("[CustomerName]", customerName)
                                                             .Replace("[CustomerEName]", customerEName)
                                                             .Replace("[AssemblyName]", assemblyName)
                                                             .Replace("[ProTypeGroupName]", proTypeGroupName)
                                                             .Replace("[ProductUseName]", productUseName)
                                                             .Replace("[DesignStatusName]", designStatusName)
                                                             .Replace("[SalesName]", salesName)
                                                             .Replace("[MailContent]", string.Format(mailFlowContent, designNo));
                                                            //.Replace("[FlowContent]", flowsHtml)
                                                            //.Replace("[hyperlink]", hyperLink)
                                    break;
                            }

                            #region //MAMO個人訊息推播
                            string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);
                            ConvertToMarkdown(flowsHtml, out string flowsMamo);
                            mamoContent = mamoContent.Replace("[FlowContent]", "\r\n" + flowsMamo)
                                                     .Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                            if (NextFlow != null)
                            {
                                mamoHelper.SendMessage(NextFlow.CompanyNo, CurrentUser, "Personal", NextFlow.UserNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                //var companyNos = nextFlow.Where(w => w.CompanyNo != null).Select(s => { return (string)s.CompanyNo; }).Distinct().ToList();
                                //foreach (var companyNo in companyNos)
                                //{
                                //    var userNos = nextFlow.Where(w => w.CompanyNo == companyNo && w.UserNo != null).Select(s => { return (string)s.UserNo; }).Distinct().ToList();
                                //    foreach (var userNo in userNos)
                                //    {
                                //        mamoHelper.SendMessage(companyNo, CurrentUser, "Personal", userNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                                //    }
                                //}
                            }
                            else
                            {
                                mamoHelper.SendMessage(salesCompanyNo, CurrentUser, "Personal", salesNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                            }
                            #endregion

                            #region //發送
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = MailTo, //通知下一關或建單業務
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent.Replace("[FlowContent]", flowsHtml).Replace("[hyperlink]", hyperLink),
                                TextBody = "-"
                            };
                            BaseHelper.MailSend(mailConfig);
                            #endregion
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

        #region //DesignUrgentSignMessage 發送新設計申請催簽通知信 -- Chia Yuan 2024.10.22
        public string DesignUrgentSignMessage(int DesignSfId, int DesignId, string UrgentSignContent)
        {
            try
            {
                if (DesignSfId < 0 || DesignId < 0) throw new SystemException("【單據】資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得流程資料
                    sql = @"SELECT a.DesignSfId
                            , a.SignJobName
                            , ISNULL(a.Approved, 'W') Approved, b.Edition
                            , flower.UserId FlowerId, flower.UserName FlowerName, flower.UserNo FlowerNo, flower.Email FlowerEmail
                            , flower.DepartmentName + '-' + flower.UserName + ':' + flower.Email MailTo
                            , sales.UserId SalesId, sales.UserName SalesName, sales.UserNo SalesNo, sales.Email SalesEmail, sales.CompanyNo SalesCompanyNo
                            , ISNULL(a.SignDesc, '-') SignDesc
                            , ISNULL(FORMAT(a.ArriveDateTime, 'yyyy-MM-dd HH:mm'), '-') ArriveDateTime
                            , ISNULL(FORMAT(a.SignDateTime, 'yyyy-MM-dd HH:mm'), '-') SignDateTime
                            , DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) / 1440 StayDays
                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 1440) / 60 StayHours
                            , (DATEDIFF(MINUTE, a.ArriveDateTime, ISNULL(a.SignDateTime, @CreateDate)) % 60) StayMinutes
                            , b.DesignNo, e.CustomerName, e.CustomerEName, FORMAT(e.AnalysisDate, 'yyyy-MM-dd HH:mm') AnalysisDate
                            , f.ProTypeGroupName, g.ProductUseName
                            , s1.StatusName RfiStatusName, s2.StatusName ExistFlagName, s3.StatusName DesignStatusName, ISNULL(s4.StatusName, '待審批') ApprovedName
                            FROM RFI.DesignSignFlow a
                            INNER JOIN RFI.DemandDesign b ON b.DesignId = a.DesignId
                            INNER JOIN RFI.RequestForInformation e ON e.RfiId = b.RfiId
                            INNER JOIN SCM.ProductTypeGroup f ON f.ProTypeGroupId = e.ProTypeGroupId
                            INNER JOIN SCM.ProductUse g ON g.ProductUseId = e.ProductUseId
                            INNER JOIN (
	                            SELECT au.UserId, au.UserNo, au.UserName, au.Email
	                            , ad.DepartmentNo, ad.DepartmentName, ac.CompanyNo, ac.CompanyName
	                            FROM BAS.[User] au
	                            INNER JOIN BAS.Department ad ON ad.DepartmentId = au.DepartmentId
	                            INNER JOIN BAS.Company ac ON ac.CompanyId = ad.CompanyId
                            ) flower ON flower.UserId = a.FlowUser
                            INNER JOIN (
	                            SELECT au.UserId, au.UserNo, au.UserName, au.Email
	                            , ad.DepartmentNo, ad.DepartmentName, ac.CompanyNo, ac.CompanyName
	                            FROM BAS.[User] au
	                            INNER JOIN BAS.Department ad ON ad.DepartmentId = au.DepartmentId
	                            INNER JOIN BAS.Company ac ON ac.CompanyId = ad.CompanyId
                            ) sales ON sales.UserId = e.SalesId
                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = e.[Status] AND s1.StatusSchema = 'RequestForInformation.Status'
                            INNER JOIN BAS.[Status] s2 ON s2.StatusNo = e.ExistFlag AND s2.StatusSchema = 'Boolean'
                            INNER JOIN BAS.[Status] s3 ON s3.StatusNo = b.[Status] AND s3.StatusSchema = 'RfiDesign.Status'
                            LEFT JOIN BAS.[Status] s4 ON s4.StatusNo = a.Approved AND s4.StatusSchema = 'RfiSign.Status'
                            WHERE a.DesignId = @DesignId
                            ORDER BY a.SortNumber";
                    var result = sqlConnection.Query(sql, new { DesignId, CreateDate });
                    var design = result.FirstOrDefault(f => f.DesignSfId == DesignSfId);
                    #endregion

                    #region //Mail資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MailSubject, a.MailContent
							, b.Host, b.Port, b.SendMode, b.Account, b.Password
                            , ISNULL(a.MailFrom, '') 'MailFrom'
                            , ISNULL(a.MailTo, '') 'MailTo'
                            , ISNULL(a.MailCc, '') 'MailCc'
                            , ISNULL(a.MailBcc, '') 'MailBcc'
							FROM BAS.Mail a
							LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
							WHERE EXISTS (
								SELECT ca.MailId
								FROM BAS.MailSendSetting ca
								WHERE ca.MailId = a.MailId
								AND ca.SettingSchema = @SettingSchema
								AND ca.SettingNo = @SettingNo
							)";
                    dynamicParameters.Add("SettingSchema", "RfiDesignSignFlow");
                    dynamicParameters.Add("SettingNo", "Y");
                    var mailTemplate = sqlConnection.Query(sql, dynamicParameters);
                    if (!mailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                    #endregion

                    string flowsHtml = string.Empty;

                    #region //Mail內容明細
                    foreach (var flow in result)
                    {
                        flowsHtml += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>",
                                    flow.SignJobName,
                                    flow.FlowerName,
                                    flow.ApprovedName,
                                    string.Format("{0}天{1}時{2}分", flow.StayDays, flow.StayHours, flow.StayMinutes),
                                    flow.ArriveDateTime,
                                    flow.SignDateTime,
                                    flow.SignDesc);
                    }
                    flowsHtml = @"<table border><thead><tr>
                                    <th>關卡</th><th>審批人</th><th>審批註記</th><th>處理時間</th><th>到達時間</th><th>審批時間</th><th>審批建議</th>
                                </tr></thead><tbody>" + flowsHtml + "</tbody></table>";
                    #endregion

                    foreach (var item in mailTemplate)
                    {
                        string mailSubject = item.MailSubject;
                        string mailContent = HttpUtility.UrlDecode(item.MailContent);
                        string hyperLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("RequestForInformation/DesignUrgentSignMessage", "RequestForInformation/DesignManagement");
                        string pageUrl = hyperLink;
                        hyperLink = string.Format("<a href=\"{0}\">{1}</a>", hyperLink, string.Format("點擊查看"));
                        mailContent = mailContent.Replace("[DesignNo]", design.DesignNo)
                                                 .Replace("[Edition]", design.Edition)
                                                 .Replace("[SignJobName]", design.SignJobName)
                                                 .Replace("[FlowerName]", design.FlowerName)
                                                 .Replace("[CustomerName]", design.CustomerName)
                                                 .Replace("[CustomerEName]", design.CustomerEName)
                                                 .Replace("[AssemblyName]", design.AssemblyName)
                                                 .Replace("[ProTypeGroupName]", design.ProTypeGroupName)
                                                 .Replace("[ProductUseName]", design.ProductUseName)
                                                 .Replace("[DesignStatusName]", design.DesignStatusName)
                                                 .Replace("[SalesName]", design.SalesName)
                                                 .Replace("[MailContent]", string.Format("【新設計申請單】已送達-單號{0}，請盡速查看並審批!\r\n{1}", design.DesignNo, UrgentSignContent));

                        #region //MAMO個人訊息推播
                        string mamoContent = Regex.Replace(mailContent, "<.*?>", string.Empty);                        
                        ConvertToMarkdown(flowsHtml, out string flowsMamo);
                        mamoContent = mamoContent.Replace("[FlowContent]", "\r\n" + flowsMamo)
                                                 .Replace("[hyperlink]", string.Format("[點我查看]({0})", pageUrl));
                        mamoHelper.SendMessage(design.SalesCompanyNo, CurrentUser, "Personal", design.FlowerNo, mailSubject + "\r\n" + mamoContent, new List<string>(), new List<int>());
                        #endregion

                        #region //發送
                        MailConfig mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = mailSubject,
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = design.FlowerEmail,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = mailContent.Replace("[FlowContent]", flowsHtml).Replace("[hyperlink]", hyperLink),
                            TextBody = "-"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }

                    #region //更新催簽次數
                    sql = @"UPDATE a SET 
                            a.UrgentSign = a.UrgentSign + 1,
                            a.LastModifiedDate = @LastModifiedDate,
                            a.LastModifiedBy = @LastModifiedBy
                            FROM RFI.DesignSignFlow a
                            WHERE a.DesignSfId = @DesignSfId";
                    int rowsAffected = sqlConnection.Execute(sql,
                        new
                        {
                            design.DesignSfId,
                            LastModifiedDate,
                            LastModifiedBy
                        });
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = string.Format("已通知下一關：{0}({1})", design.FlowerName, design.SignJobName)
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

        #region //UpdateDesignSignFlowAdditionalFile RFI審批附檔資料更新 -- Chia Yuan 2024.06.26
        public string UpdateDesignSignFlowAdditionalFile(int DesignSfId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 AdditionalFile
                                FROM RFI.DesignSignFlow a
                                WHERE a.DesignSfId = @DesignSfId";
                        dynamicParameters.Add("DesignSfId", DesignSfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【流程檔】資料錯誤!");
                        int? additionalFile = -1;
                        foreach (var item in result)
                        {
                            additionalFile = item.AdditionalFile;
                        }
                        #endregion

                        #region //判斷附檔資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【附檔】資料錯誤!");
                        #endregion

                        #region //主資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
                                a.AdditionalFile = @AdditionalFile,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.DesignSignFlow a
                                WHERE a.DesignSfId = @DesignSfId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DesignSfId,
                                AdditionalFile = FileId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //舊附檔刪除
                        if (additionalFile != null)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 AdditionalFile
                                    FROM RFI.DesignSignFlow
                                    WHERE AdditionalFile = @AdditionalFile
                                    AND DesignSfId <> @DesignSfId";
                            dynamicParameters.Add("AdditionalFile", additionalFile);
                            dynamicParameters.Add("DesignSfId", DesignSfId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (!result.Any())
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @AdditionalFile";
                                dynamicParameters.Add("AdditionalFile", additionalFile);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateDemandDesignSpec -- 新設計申請單規格更新 -- Chia Yuan 2024.01.25
        public string UpdateDemandDesignSpec(int DesignId, string ProdSpecs, string ProdSpecDetails)
        {
            try
            {
                #region /判斷資料長度
                if (string.IsNullOrWhiteSpace(ProdSpecs)) throw new SystemException("【設計規格】資料錯誤!");
                if (string.IsNullOrWhiteSpace(ProdSpecDetails)) throw new SystemException("【設計規格特徵】資料錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷新設計申請單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.DemandDesign a
                                WHERE a.DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【新設計申請單】資料錯誤!");
                        #endregion

                        if (!ProdSpecs.TryParseJson(out JObject prodSpecs)) throw new SystemException("【設計規格】資料錯誤!");
                        if (!ProdSpecDetails.TryParseJson(out JObject prodSpecDetails)) throw new SystemException("【設計規格特徵】資料錯誤!");
                        if (prodSpecs["data"].Count() <= 0) throw new SystemException("【設計規格】資料必須維護!");
                        if (prodSpecDetails["data"].Count() <= 0) throw new SystemException("【設計規格特徵】資料必須維護!");

                        if (prodSpecs["data"].Count() > 0)
                        {
                            var tempSpIds = prodSpecs["data"].Select(s => { return Convert.ToInt32(s["TempSpId"]); }).ToArray();

                            #region //判斷狀態資料是否正確
                            sql = @"SELECT a.StatusNo
                                    FROM BAS.[Status] a
                                    WHERE a.StatusSchema = 'Status'
                                    AND a.StatusNo IN @Status";
                            dynamicParameters.Add("Status", prodSpecDetails["data"].Select(s => { return s["Status"].ToString(); }).Distinct().ToArray());
                            var resultStatus = sqlConnection.Query(sql, dynamicParameters);
                            if (resultStatus.Count() <= 0) throw new SystemException("【狀態】資料錯誤!");
                            #endregion

                            #region //評估項目
                            foreach (var detail in prodSpecDetails["data"])
                            {
                                if (!int.TryParse(detail["TempSpDetailId"].ToString(), out int ddsDetailId)) throw new SystemException("【設計規格特徵】資料錯誤!");
                                if (resultStatus.FirstOrDefault(f => f.StatusNo == detail["Status"]?.ToString()) == null) throw new SystemException("【狀態】資料錯誤!");

                                #region //規格特徵更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE a SET 
                                        a.RequireFlag = @RequireFlag,
                                        a.DesignInput = @DesignInput,
                                        a.Description = @Description,
                                        a.Status = @Status,
                                        a.LastModifiedDate = @LastModifiedDate,
                                        a.LastModifiedBy = @LastModifiedBy
                                        FROM RFI.DemandDesignSpecDetail a
                                        WHERE a.DdsDetailId = @DdsDetailId";

                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DdsDetailId = ddsDetailId,
                                        RequireFlag = string.IsNullOrWhiteSpace(detail["RequireFlag"].ToString()) ? null : detail["RequireFlag"].ToString().Trim(),
                                        DesignInput = string.IsNullOrWhiteSpace(detail["DesignInput"].ToString()) ? null : detail["DesignInput"].ToString().Trim(),
                                        Description = string.IsNullOrWhiteSpace(detail["Description"].ToString()) ? null : detail["Description"].ToString().Trim(),
                                        Status = detail["Status"].ToString(),
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
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
        #endregion

        #region //Delete
        #region //DeleteProdTerminal -- 終端資料刪除 -- Chia Yuan 2023.12.25
        public string DeleteProdTerminal(int ProdTerminalId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷終端資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.Add("ProdTerminalId", ProdTerminalId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【終端】資料錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.ProdTerminal a
                                WHERE a.ProdTerminalId = @ProdTerminalId";
                        dynamicParameters.Add("ProdTerminalId", ProdTerminalId);

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

        #region //DeleteProdSystem -- 系統資料刪除 -- Chia Yuan 2023.12.25
        public string DeleteProdSystem(int ProdSysId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.Add("ProdSysId", ProdSysId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統】資料錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.ProdSystem a
                                WHERE a.ProdSysId = @ProdSysId";
                        dynamicParameters.Add("ProdSysId", ProdSysId);

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

        #region //DeleteProdModule -- 模組資料刪除 -- Chia Yuan 2023.12.25
        public string DeleteProdModule(int ProdModId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.Add("ProdModId", ProdModId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組】資料錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.ProdModule a
                                WHERE a.ProdModId = @ProdModId";
                        dynamicParameters.Add("ProdModId", ProdModId);

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

        #region //DeleteTemplateProdSpec -- 評估樣板資料刪除 -- Chia Yuan 2024.01.03
        public string DeleteTemplateProdSpec(int TempProdSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProTypeGroupId
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //判斷子資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempPsDetailId
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //刪除明細資料
                        if (resultDetail.Any())
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM RFI.TemplateProdSpecDetail a
                                    WHERE a.TempProdSpecId = @TempProdSpecId";
                            dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.TemplateProdSpec a
                                WHERE a.TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", TempProdSpecId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempProdSpecId, a.SortNumber
                                FROM RFI.TemplateProdSpec a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                ORDER BY a.SortNumber";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpec a
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int i = 1;
                        foreach (var item in result)
                        {
                            #region //更新主資料順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateProdSpec a
                                    WHERE a.TempProdSpecId = @TempProdSpecId";

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    item.TempProdSpecId,
                                    SortNumber = i,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            i++;
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

        #region //DeleteTemplateProdSpecDetail -- 評估樣板子資料刪除 -- Chia Yuan 2024.01.03
        public string DeleteTemplateProdSpecDetail(int TempPsDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TempProdSpecId
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempPsDetailId = @TempPsDetailId";
                        dynamicParameters.Add("TempPsDetailId", TempPsDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【規格特徵】資料錯誤!");
                        int tempProdSpecId = -1;
                        foreach (var item in result)
                        {
                            tempProdSpecId = item.TempProdSpecId;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempPsDetailId = @TempPsDetailId";
                        dynamicParameters.Add("TempPsDetailId", TempPsDetailId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempPsDetailId, a.SortNumber
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE a.TempProdSpecId = @TempProdSpecId
                                ORDER BY a.SortNumber";
                        dynamicParameters.Add("TempProdSpecId", tempProdSpecId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateProdSpecDetail a
                                WHERE TempProdSpecId = @TempProdSpecId";
                        dynamicParameters.Add("TempProdSpecId", tempProdSpecId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int i = 1;
                        foreach (var item in result)
                        {
                            #region //更新主資料順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateProdSpecDetail a
                                    WHERE a.TempPsDetailId = @TempPsDetailId";

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    item.TempPsDetailId,
                                    SortNumber = i,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            i++;
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

        #region //DeleteTemplateSpecParameter -- 評估樣板資料刪除 -- Chia Yuan 2024.01.03
        public string DeleteTemplateSpecParameter(int TempSpId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ProTypeGroupId
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //判斷子資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempSpDetailId
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //刪除明細資料
                        if (resultDetail.Any())
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM RFI.TemplateSpecParameterDetail a
                                    WHERE a.TempSpId = @TempSpId";
                            dynamicParameters.Add("TempSpId", TempSpId);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", TempSpId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempSpId, a.SortNumber
                                FROM RFI.TemplateSpecParameter a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                ORDER BY a.SortNumber";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameter a
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int i = 1;
                        foreach (var item in result)
                        {
                            #region //更新主資料順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateSpecParameter a
                                    WHERE a.TempSpId = @TempSpId";

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    item.TempSpId,
                                    SortNumber = i,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            i++;
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

        #region //DeleteTemplateSpecParameterDetail -- 設計樣板子資料刪除 -- Chia Yuan 2024.05.24
        public string DeleteTemplateSpecParameterDetail(int TempSpDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TempSpId
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpDetailId = @TempSpDetailId";
                        dynamicParameters.Add("TempSpDetailId", TempSpDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【規格特徵】資料錯誤!");
                        int tempSpId = -1;
                        foreach (var item in result)
                        {
                            tempSpId = item.TempSpId;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpDetailId = @TempSpDetailId";
                        dynamicParameters.Add("TempSpDetailId", TempSpDetailId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempSpDetailId, a.SortNumber
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE a.TempSpId = @TempSpId
                                ORDER BY a.SortNumber";
                        dynamicParameters.Add("TempSpId", tempSpId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.SortNumber = a.SortNumber * -1,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM RFI.TemplateSpecParameterDetail a
                                WHERE TempSpId = @TempSpId";
                        dynamicParameters.Add("TempSpId", tempSpId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int i = 1;
                        foreach (var item in result)
                        {
                            #region //更新主資料順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET
                                    a.SortNumber = @SortNumber,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateSpecParameterDetail a
                                    WHERE a.TempSpDetailId = @TempSpDetailId";

                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    item.TempSpDetailId,
                                    SortNumber = i,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            i++;
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

        #region //DeleteRequestForInformation -- RFI單頭資料刪除
        public string DeleteRequestForInformation(int RfiId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷評估項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.RequestForInformation a
                                WHERE a.RfiId = @RfiId";
                        dynamicParameters.Add("RfiId", RfiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【市場評估單】資料錯誤!");
                        #endregion

                        #region //取得RFI單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfiDetailId
                                FROM RFI.RfiDetail a
                                WHERE a.RfiId = @RfiId";
                        dynamicParameters.Add("RfiId", RfiId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("【市場評估單身】資料錯誤!");
                        #endregion

                        var rfiDetailIds = result.Where(w => w.ProdSpecId != null).Select(s => { return Convert.ToInt32(s.ProdSpecId); }).Distinct().ToArray();

                        if (rfiDetailIds.Length > 0)
                        {
                            #region //取得評估項目資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProdSpecId, b.PsDetailId
                                    FROM RFI.ProductSpec a
                                    LEFT JOIN RFI.ProductSpecDetail b ON b.ProdSpecId = a.ProdSpecId
                                    WHERE a.RfiDetailId IN @RfiDetailIds";
                            dynamicParameters.Add("RfiDetailIds", rfiDetailIds);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                            #endregion

                            var prodSpecIds = result.Where(w => w.ProdSpecId != null).Select(s => { return Convert.ToInt32(s.ProdSpecId); }).Distinct().ToArray();
                            var psDetailIds = result.Where(w => w.PsDetailId != null).Select(s => { return Convert.ToInt32(s.PsDetailId); }).Distinct().ToArray();

                            if (psDetailIds.Length > 0)
                            {
                                #region //規格特徵資料刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM RFI.ProductSpecDetail a
                                        WHERE a.PsDetailId IN @PsDetailIds";
                                dynamicParameters.Add("PsDetailIds", psDetailIds);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            if (prodSpecIds.Length > 0)
                            {
                                #region //評估項目資料刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM RFI.ProductSpec a
                                        WHERE a.ProdSpecId IN @ProdSpecIds";
                                dynamicParameters.Add("ProdSpecIds", prodSpecIds);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            #region //RFI單身資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM RFI.RfiDetail a
                                    WHERE a.RfiId = @RfiId";
                            dynamicParameters.Add("RfiId", RfiId);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.RequestForInformation a
                                WHERE a.RfiId = @RfiId";
                        dynamicParameters.Add("RfiId", RfiId);
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

        #region //DeleteRfiDetail --RFI單身資料刪除
        public string DeleteRfiDetail(int RfiDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFI單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.Add("RfiDetailId", RfiDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【市場評估單身】資料錯誤!");
                        #endregion

                        #region //取得評估項目資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ProdSpecId, b.PsDetailId
                                FROM RFI.ProductSpec a
                                LEFT JOIN RFI.ProductSpecDetail b ON b.ProdSpecId = a.ProdSpecId
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.Add("RfiDetailId", RfiDetailId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("【評估項目】資料錯誤!");
                        #endregion

                        var prodSpecIds = result.Where(w => w.ProdSpecId != null).Select(s => { return Convert.ToInt32(s.ProdSpecId); }).Distinct().ToArray();
                        var psDetailIds = result.Where(w => w.PsDetailId != null).Select(s => { return Convert.ToInt32(s.PsDetailId); }).Distinct().ToArray();

                        if (psDetailIds.Length > 0)
                        {
                            #region //規格特徵資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM RFI.ProductSpecDetail a
                                    WHERE a.PsDetailId IN @PsDetailIds";
                            dynamicParameters.Add("PsDetailIds", psDetailIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        if (prodSpecIds.Length > 0)
                        {
                            #region //評估項目資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM RFI.ProductSpec a
                                    WHERE a.ProdSpecId IN @ProdSpecIds";
                            dynamicParameters.Add("ProdSpecIds", prodSpecIds);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM RFI.RfiDetail a
                                WHERE a.RfiDetailId = @RfiDetailId";
                        dynamicParameters.Add("RfiDetailId", RfiDetailId);
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
        #region //ConvertToMarkdown
        public void ConvertToMarkdown(string mailContent, out string mamoContent)
        {
            if (string.IsNullOrWhiteSpace(mailContent)) mamoContent = "";
            else
            {
                mamoContent = Regex.Replace(mailContent, "\r\n", "");
                //
                mamoContent = Regex.Replace(mamoContent, @"\s+", "");
                mamoContent = Regex.Replace(mamoContent, "<table.*?><thead.*?>", "");
                mamoContent = Regex.Replace(mamoContent, "</thead>", "");
                mamoContent = Regex.Replace(mamoContent, "<tr.*?><th.*?>", "|"); //起始
                mamoContent = Regex.Replace(mamoContent, "</th><th.*?>", "|");  //中間
                mamoContent = Regex.Replace(mamoContent, "</th></tr>", "|\n");  // 結束 換行
                mamoContent = Regex.Replace(mamoContent, "<tr.*?><td.*?>", "|"); //起始
                mamoContent = Regex.Replace(mamoContent, "</td><td.*?>", "|");  //中間
                mamoContent = Regex.Replace(mamoContent, "</td></tr>", "|\n");  // 結束 換行
                mamoContent = Regex.Replace(mamoContent, "<tbody.*?>", "");
                mamoContent = Regex.Replace(mamoContent, "</tbody></table>", "");

                var rowList = mamoContent.Remove(mamoContent.Length - 1).Split('\n');
                var signList = rowList.Where(w => !w.Contains("審批建議：")).Select((s, x) => s).ToList(); //new { i, d = s + (i == 0 ? "審批建議|" : "|")}
                var descList = rowList.Where(w => w.Contains("審批建議：")).Select((s, x) => s).ToList(); //new { i, d = s}
                var splitList = rowList.First().Remove(0, 1).Split('|').Select(s => "---");
                var splitLine = string.Format("|{0}|", string.Join("|", splitList));
                signList.Insert(1, splitLine);
                mamoContent = string.Join("\r\n", signList);
            }
        }
        #endregion

        #region //GetAssemblyName -- 取得機種名稱編號
        public string GetAssemblyName(int CompanyId, string ProductGroupTypeOne, string ProductUseTyoeOne)
        {
            using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
            {
                //if (!autoAssemblyNum) if (!Regex.IsMatch(AssemblyName, @"^[A-Z]{3}\d{2}([A-Z]|\d){1}\d{1}(Y|X){1}")) throw new SystemException("【設計代號】格式錯誤!");

                List<(int, string)> DictCompany = new List<(int, string)> {(2,"JMO"), (3, "ETG"), (1, "DGJMO") };

                var corp = DictCompany.Find(f => f.Item1 == CompanyId);
                if (string.IsNullOrWhiteSpace(corp.Item2)) throw new SystemException("【公司別】資料錯誤!");

                string CompanyNo = corp.Item2;
                string AssemblyNameRule = "{0}_{1}{2}{3}{4}";

                #region //機種名稱取號
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT MAX(SUBSTRING(AssemblyName, 1, LEN(AssemblyName)-1)) AS CurrentAssemblyNum
                        FROM RFI.DemandDesign
                        WHERE AssemblyName LIKE N'%' + @AssemblyName + '%'";
                dynamicParameters.Add("AssemblyName", string.Format("{0}_{1}", CompanyNo, ProductGroupTypeOne));
                var resultFirst = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                string currentAssemblyNum = resultFirst.CurrentAssemblyNum;
                #endregion

                //ex: ETG2401X
                #region //機種名稱編碼規則
                int.TryParse(CreateDate.ToString("yy"), out int yy2);
                if (!string.IsNullOrWhiteSpace(currentAssemblyNum))
                {
                    int.TryParse(Regex.Replace(currentAssemblyNum, @"^\w{" + (CompanyNo.Length + 2) + @"}", "").Substring(0, 2), out int yy1);
                    if (yy2 > yy1) currentAssemblyNum = null; //新的年度
                }

                string tempAssemblyName = currentAssemblyNum ?? string.Format(AssemblyNameRule, CompanyNo, ProductGroupTypeOne, yy2, "00", ProductUseTyoeOne);
                string tempNum = Regex.Replace(tempAssemblyName, @"^\w{" + (CompanyNo.Length + 2) + @"}\d{2}", "").Substring(0, 2); //取出編號值

                if (Regex.IsMatch(tempNum, @"^[A-Za-z]{1}\d{1}$"))
                { //第一碼為字母
                    int.TryParse(Regex.Replace(tempNum, @"^[\w]", ""), out int assemblyNum);
                    string charNum = tempNum.Substring(0, 1);
                    if (assemblyNum == 9)
                    { //尾數=9
                        if (charNum == "Z") throw new SystemException("【設計代號】編碼已達最大值!");
                        byte asciiBytes = Encoding.ASCII.GetBytes(charNum).FirstOrDefault();
                        tempNum = Convert.ToChar(asciiBytes += 1) + "1"; //字母編號進位
                    }
                    else
                    { //字母編號+1
                        tempNum = charNum + (assemblyNum += 1).ToString();
                    }
                }
                else if (Regex.IsMatch(tempNum, @"^\d{2}$"))
                { //兩碼皆數字
                    int.TryParse(tempNum, out int assemblyNum);
                    if (assemblyNum < 99)
                    { //數字編號+1
                        tempNum = string.Format("{0:00}", assemblyNum += 1);
                    }
                    else if (assemblyNum % 99 == 0)
                    { //數字編號已到99 則改字母為守字
                        tempNum = "A1";
                    }
                }
                #endregion

                return string.Format(AssemblyNameRule, CompanyNo, ProductGroupTypeOne, yy2, tempNum, ProductUseTyoeOne);
            }
        }
        #endregion
        #endregion
    }
}
