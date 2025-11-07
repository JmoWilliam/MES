using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace MESDA
{
    public class ToolDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string OldMesConnectionStrings = "";

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

        public ToolDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            OldMesConnectionStrings = ConfigurationManager.AppSettings["OldDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;
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

        #region//Get

        #region //GetToolGroup -- 取得工具群組 -- Shintokuro 2022.12.27
        public string GetToolGroup(int ToolGroupId, string ToolGroupNo, string ToolGroupName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolGroupId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolGroupNo, a.ToolGroupName, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolGroup a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND a.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupNo", @" AND a.ToolGroupNo LIKE '%' + @ToolGroupNo + '%'", ToolGroupNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupName", @" AND a.ToolGroupName LIKE '%' + @ToolGroupName + '%'", ToolGroupName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolGroupId DESC";
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

        #region //GetToolClass -- 取得工具類別 -- Shintokuro 2022.12.27
        public string GetToolClass(int ToolClassId, int ToolGroupId, string ToolClassNo, string ToolClassName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolClassId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolClassNo, a.ToolClassName, a.Status,b.ToolGroupId,b.ToolGroupName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolClass a
                          INNER JOIN MES.ToolGroup b on a.ToolGroupId = b.ToolGroupId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassId", @" AND a.ToolClassId = @ToolClassId", ToolClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND a.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassNo", @" AND a.ToolClassNo LIKE '%' + @ToolClassNo + '%'", ToolClassNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassName", @" AND a.ToolClassName LIKE '%' + @ToolClassName + '%'", ToolClassName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolClassId DESC";
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

        #region //GetToolCategory -- 取得工具種類 -- Shintokuro 2022.12.27
        public string GetToolCategory(int ToolCategoryId, int ToolClassId, int ToolGroupId, string ToolCategoryNo, string ToolCategoryName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolCategoryId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolCategoryNo, a.ToolCategoryName, a.Status
                           , b.ToolClassId, b.ToolClassName
                           , c.ToolGroupId, c.ToolGroupName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolCategory a
                          INNER JOIN MES.ToolClass b on a.ToolClassId = b.ToolClassId
                          INNER JOIN MES.ToolGroup c on b.ToolGroupId = c.ToolGroupId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryId", @" AND a.ToolCategoryId = @ToolCategoryId", ToolCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassId", @" AND a.ToolClassId = @ToolClassId", ToolClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND c.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryNo", @" AND a.ToolCategoryNo LIKE '%' + @ToolCategoryNo + '%'", ToolCategoryNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryName", @" AND a.ToolCategoryName LIKE '%' + @ToolCategoryName + '%'", ToolCategoryName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolCategoryId DESC";
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

        #region //GetToolModel -- 取得工具型號 -- Shintokuro 2022.12.28
        public string GetToolModel(int ToolModelId, int ToolCategoryId, int ToolClassId, int ToolGroupId, int SupplierId, string ToolModelNo, string ToolModelName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolModelId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolModelNo, a.ToolModelName,a.ToolModelErpNo,a.SupplierId,a.ToolCategoryId, a.Status
                           , b.ToolCategoryNo, b.ToolCategoryName
                           , c.ToolClassId, c.ToolClassName
                           , d.ToolGroupId, d.ToolGroupName
                           , e.SupplierName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolModel a
                          INNER JOIN MES.ToolCategory b on a.ToolCategoryId = b.ToolCategoryId
                          INNER JOIN MES.ToolClass c on b.ToolClassId = c.ToolClassId
                          INNER JOIN MES.ToolGroup d on c.ToolGroupId = d.ToolGroupId
                          LEFT JOIN SCM.Supplier e on a.SupplierId = e.SupplierId

                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelId", @" AND a.ToolModelId = @ToolModelId", ToolModelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryId", @" AND a.ToolCategoryId = @ToolCategoryId", ToolCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassId", @" AND b.ToolClassId = @ToolClassId", ToolClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND c.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelNo", @" AND a.ToolModelNo LIKE '%' + @ToolModelNo + '%'", ToolModelNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelName", @" AND a.ToolModelName LIKE '%' + @ToolModelName + '%'", ToolModelName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND e.SupplierId = @SupplierId", SupplierId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolModelId DESC";
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

        #region //GetTool -- 取得工具明細 -- Shintokuro 2022.12.29
        public string GetTool(int ToolId, int ToolModelId, int ToolCategoryId, int ToolClassId, int ToolGroupId, int SupplierId, string ToolNo
            , string Status
            , string OrderBy, int PageIndex, int PageSize) 
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolNo, a.UsageCount,a.ToolModelId, a.Status,a.RevToolCategoryId
                           , b.ToolModelId, b.ToolModelName, b.ToolModelNo, b.ToolModelErpNo
                           , c.ToolCategoryId, c.ToolCategoryName
                           , c1.ToolCategoryName RevToolCategoryName
                           , d.ToolClassId, d.ToolClassName
                           , e.ToolGroupId, e.ToolGroupName
                           , f.SupplierName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Tool a
                          INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                          INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                          LEFT JOIN MES.ToolCategory c1 on a.RevToolCategoryId = c1.ToolCategoryId
                          INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                          INNER JOIN MES.ToolGroup e on d.ToolGroupId = e.ToolGroupId
                          LEFT JOIN SCM.Supplier f on b.SupplierId = f.SupplierId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND e.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolId", @" AND a.ToolId = @ToolId", ToolId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelId", @" AND a.ToolModelId = @ToolModelId", ToolModelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryId", @" AND b.ToolCategoryId = @ToolCategoryId", ToolCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassId", @" AND c.ToolClassId = @ToolClassId", ToolClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND d.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolNo", @" AND a.ToolNo = @ToolNo", ToolNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND f.SupplierId = @SupplierId", SupplierId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolId DESC";
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

        #region //GetToolNoMaxsor -- 取得工具目前最大編號 -- Shintokuro 2022.12.29
        public string GetToolNoMaxsor(int ToolId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //判斷工具目前最大編號是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 ToolNo, SUBSTRING(ToolNo,1,1) + RIGHT(REPLICATE('0', 5) + CAST(CONVERT(INT,SUBSTRING(ToolNo,2,5))+1 as NVARCHAR), 5) MaxToolNo
                            FROM MES.Tool a
                            INNER JOIN MES.ToolModel a1 on a.ToolModelId = a1.ToolModelId
                            INNER JOIN MES.ToolCategory a2 on a1.ToolCategoryId = a2.ToolCategoryId
                            INNER JOIN MES.ToolClass a3 on a2.ToolClassId = a3.ToolClassId
                            INNER JOIN MES.ToolGroup a4 on a3.ToolGroupId = a4.ToolGroupId
                            WHERE a4.CompanyId = @CompanyId
                            ORDER BY a.ToolId DESC";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

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

        #region //GetToolSpec -- 取得工具規格 -- Shintokuro 2022.12.29
        public string GetToolSpec(int ToolSpecId, string ToolSpecNo, string ToolSpecName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolSpecNo, a.ToolSpecName, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolSpec a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecId", @" AND a.ToolSpecId = @ToolSpecId", ToolSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecNo", @" AND a.ToolSpecNo LIKE '%' + @ToolSpecNo + '%'", ToolSpecNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecName", @" AND a.ToolSpecName LIKE '%' + @ToolSpecName + '%'", ToolSpecName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolSpecId DESC";
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

        #region //GetToolModelSpec -- 取得工具規格 -- Shintokuro 2022.12.30

        public string GetToolModelSpec(int ToolModelSpecId, int ToolModelId, int ToolSpecId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolModelSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolModelId, a.ToolSpecId, a.ToolSpecValue, a.Status
                           , b.ToolSpecNo, b.ToolSpecName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolModelSpec a
                          INNER JOIN MES.ToolSpec b on a.ToolSpecId = b.ToolSpecId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelSpecId", @" AND a.ToolModelSpecId = @ToolModelSpecId", ToolModelSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelId", @" AND a.ToolModelId = @ToolModelId", ToolModelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecId", @" AND a.ToolSpecId = @ToolSpecId", ToolSpecId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolModelSpecId DESC";
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

        #region //GetToolSpecLog -- 取得工具規格 -- Shintokuro 2022.12.30

        public string GetToolSpecLog(int ToolSpecLogId, int ToolId, int ToolSpecId, string ToolSpecNo, string ToolSpecName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolSpecLogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolId, a.ToolSpecId, a.ToolSpecValue, a.Status
                           , b.ToolSpecNo, b.ToolSpecName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolSpecLog a
                          INNER JOIN MES.ToolSpec b on a.ToolSpecId = b.ToolSpecId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecLogId", @" AND a.ToolSpecLogId = @ToolSpecLogId", ToolSpecLogId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolId", @" AND a.ToolId = @ToolId", ToolId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecId", @" AND a.ToolSpecId = @ToolSpecId", ToolSpecId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecNo", @" AND b.ToolSpecNo LIKE '%' + @ToolSpecNo + '%'", ToolSpecNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolSpecName", @" AND b.ToolSpecName LIKE '%' + @ToolSpecName + '%'", ToolSpecName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolSpecLogId DESC";
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

        #region //GetToolInventory -- 取得工具倉庫 -- Shintokuro 2022.12.30
        public string GetToolInventory(int ToolInventoryId, string ToolInventoryNo, string ToolInventoryName, int ShopId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolInventoryId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolInventoryNo, a.ToolInventoryName, a.ToolInventoryDesc, a.ShopId, a.Status
                           , b.ShopName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolInventory a
                          INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolInventoryId", @" AND a.ToolInventoryId = @ToolInventoryId", ToolInventoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolInventoryNo", @" AND a.ToolInventoryNo LIKE '%' + @ToolInventoryNo + '%'", ToolInventoryNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolInventoryName", @" AND a.ToolInventoryName LIKE '%' + @ToolInventoryName + '%'", ToolInventoryName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolInventoryId DESC";
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

        #region //GetToolLocator -- 取得工具儲位 -- Shintokuro 2023.01.03
        public string GetToolLocator(int ToolLocatorId, int ToolInventoryId, string ToolLocatorNo, string ToolLocatorName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolLocatorId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolLocatorNo, a.ToolLocatorName, a.ToolLocatorDesc, a.ToolInventoryId, a.Status,a.LimitStatus
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolLocator a
                          INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                          INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolLocatorId", @" AND a.ToolLocatorId = @ToolLocatorId", ToolLocatorId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolInventoryId", @" AND a.ToolInventoryId = @ToolInventoryId", ToolInventoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolLocatorNo", @" AND a.ToolLocatorNo LIKE '%' + @ToolLocatorNo + '%'", ToolLocatorNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolLocatorName", @" AND a.ToolLocatorName = @ToolLocatorName ", ToolLocatorName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolLocatorId DESC";
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

        #region //GetToolTrade -- 取得工具交易資料 -- Shintokuro 2022.12.29
        public string GetToolTrade(int ToolId, int ToolModelId, int ToolCategoryId, int ToolClassId, int ToolGroupId, int SupplierId, string ToolNo
            , string Status, string Source, string ToolKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //來源是從工具入庫交易頁面,Status狀態要是A啟用狀態
                if (Source == "ToolTrade") Status = "A";
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolNo, a.UsageCount,a.ToolModelId, a.Status
                           , b.ToolModelId, b.ToolModelName, b.ToolModelNo
                           , c.ToolCategoryId, c.ToolCategoryName
                           , d.ToolClassId, d.ToolClassName
                           , e.ToolGroupId, e.ToolGroupName
                           , f.SupplierName
                           , g.ToolLocatorId
                           , h.ToolLocatorName,h1.ToolInventoryName,h2.ShopName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Tool a
                          INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                          INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                          INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                          INNER JOIN MES.ToolGroup e on d.ToolGroupId = e.ToolGroupId
                          LEFT JOIN SCM.Supplier f on b.SupplierId = f.SupplierId
                          LEFT JOIN MES.ToolTransactions g on a.ToolId = g.ToolId
                          LEFT JOIN MES.ToolLocator h on g.ToolLocatorId = h.ToolLocatorId
                          LEFT JOIN MES.ToolInventory h1 on h.ToolInventoryId = h1.ToolInventoryId
                          LEFT JOIN MES.WorkShop h2 on h1.ShopId = h2.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND e.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolId", @" AND a.ToolId = @ToolId", ToolId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolModelId", @" AND a.ToolModelId = @ToolModelId", ToolModelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolCategoryId", @" AND b.ToolCategoryId = @ToolCategoryId", ToolCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolClassId", @" AND c.ToolClassId = @ToolClassId", ToolClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolGroupId", @" AND d.ToolGroupId = @ToolGroupId", ToolGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolNo", @" AND a.ToolNo LIKE '%' + @ToolNo + '%'", ToolNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND f.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolKey", @" AND (a.ToolId LIKE '%' + @ToolKey + '%' OR a.ToolNo LIKE '%' + @ToolKey + '%' )", ToolKey);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolId DESC";
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

        #region //GetToolTransactions -- 取得工具交易紀錄 -- Shintokuro 2023.01.04
        public string GetToolTransactions(int ToolTransactionsId, int ToolId, string TransactionType, string StartDate, string EndDate, int TraderId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ToolTransactionsId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ToolId, a.TransactionType,FORMAT(a.TransactionDate, 'yyyy-MM-dd') TransactionDate, a.ToolLocatorId, a.TraderId, a.TransactionReason,FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss tt') CreateDate
                           , b.ToolLocatorName, c.ToolInventoryName, d.ShopName
                           , e.UserName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ToolTransactions a
                          INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                          INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                          INNER JOIN MES.WorkShop d on c.ShopId = d.ShopId
                          INNER JOIN BAS.[User] e on a.TraderId = e.UserId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolTransactionsId", @" AND a.ToolTransactionsId = @ToolTransactionsId", ToolTransactionsId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToolId", @" AND a.ToolId = @ToolId", ToolId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransactionType", @" AND a.TransactionType = @TransactionType", TransactionType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TraderId", @" AND a.TraderId = @TraderId", TraderId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");



                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ToolTransactionsId DESC";
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

        #region //GetLabelPrintMachine -- 取得標籤機資訊 -- Daiyi 2023.02.14
        public string GetLabelPrintMachine(int LabelPrintId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.LabelPrintId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.LabelPrintNo, a.LabelPrintName, a.LabelPrintIp
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.LabelPrintMachine a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LabelPrintId", @" AND a.LabelPrintId = @LabelPrintId", LabelPrintId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LabelPrintId DESC";
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

        #region //GetToolCountInLocator --取得該儲位入庫數
        public string GetToolCountInLocator(int ToolLocatorId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings)) 
                {
                    sql = @"SELECT ISNULL((SUM(a1.ToolCount)), 0) ToolCount
                            FROM 
                            (
	                            SELECT ToolLocatorId,TransactionType,
	                            (
		                            CASE TransactionType
		                            WHEN 'In' THEN 1
		                            WHEN 'Out' THEN -1
		                            END
	                            ) TransactionTypeInt,
	                            SUM(
		                            CASE TransactionType
		                            WHEN 'In' THEN 1
		                            WHEN 'Out' THEN -1
		                            END
		                            ) ToolCount
	                            FROM MES.ToolTransactions
	                            WHERE ToolLocatorId = @ToolLocatorId
	                            GROUP BY ToolLocatorId,TransactionType
	                            HAVING (
				                            CASE TransactionType
				                            WHEN 'In' THEN 1
				                            WHEN 'Out' THEN -1
				                            END
				                            )>=-1
                            ) AS a1
                            GROUP BY a1.ToolLocatorId
                        ";
                    dynamicParameters.Add("ToolLocatorId", ToolLocatorId);
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

        #region //GetLocatorTool --取得該儲位入庫刀具明細
        public string GetLocatorTool(int ToolLocatorId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))  
                {
                    sql = @"SELECT a1.ToolId,ISNULL(i.ProcessingQty,0)ProcessingQty, b.ToolNo,b.UsageCount,b.ToolModelId, b.Status,b.RevToolCategoryId, c.ToolModelId, c.ToolModelName, c.ToolModelNo, c.ToolModelErpNo, d.ToolCategoryId, d.ToolCategoryName, d1.ToolCategoryName RevToolCategoryName, e.ToolClassId, e.ToolClassName, f.ToolGroupId, f.ToolGroupName, g.SupplierName
                            FROM 
                            (
	                            SELECT ToolId,ToolLocatorId,
		                            (
		                            CASE TransactionType
		                            WHEN 'In' THEN 1
		                            WHEN 'Out' THEN -1
		                            END
		                            ) TransactionType
	                            FROM MES.ToolTransactions
	                            WHERE ToolLocatorId = @ToolLocatorId
                            ) AS a1
                        INNER JOIN MES.Tool b ON b.ToolId = a1.ToolId
                        INNER JOIN MES.ToolModel c on b.ToolModelId = c.ToolModelId
                        INNER JOIN MES.ToolCategory d on d.ToolCategoryId = c.ToolCategoryId
                        LEFT JOIN MES.ToolCategory d1 on d1.ToolCategoryId = b.RevToolCategoryId
                        INNER JOIN MES.ToolClass e on e.ToolClassId = d.ToolClassId
                        INNER JOIN MES.ToolGroup f on f.ToolGroupId = e.ToolGroupId
                        LEFT JOIN SCM.Supplier g on g.SupplierId = c.SupplierId
                        INNER JOIN MES.ToolLocator h on h.ToolLocatorId = a1.ToolLocatorId
                        INNER JOIN MES.ToolTransactions i on i.ToolLocatorId = a1.ToolLocatorId
                        GROUP BY a1.ToolId,ISNULL(i.ProcessingQty,0), b.ToolNo,b.UsageCount,b.ToolModelId, b.Status,b.RevToolCategoryId, c.ToolModelId, c.ToolModelName, c.ToolModelNo, c.ToolModelErpNo, d.ToolCategoryId, d.ToolCategoryName, d1.ToolCategoryName, e.ToolClassId, e.ToolClassName, f.ToolGroupId, f.ToolGroupName, g.SupplierName
                        HAVING SUM(a1.TransactionType)>0
                        ORDER BY a1.ToolId
                        ";
                    dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

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
        #endregion

        #region//Add

        #region //AddToolSpec -- 工具規格新增 -- Shintokuro 2022.12.29
        public string AddToolSpec(string ToolSpecNo, string ToolSpecName)
        {
            try
            {
                if (ToolSpecNo.Length <= 0) throw new SystemException("工具群組【編號】不能為空!");
                if (ToolSpecNo.Length > 50) throw new SystemException("工具群組【編號】長度錯誤!");
                if (ToolSpecName.Length > 50) throw new SystemException("工具群組【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具規格編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE CompanyId = @CompanyId
                                AND ToolSpecNo = @ToolSpecNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolSpecNo", ToolSpecNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具規格【編號】重複，請重新輸入!");
                        #endregion


                        //#region //判斷工具規格名稱是否重複
                        //if(ToolSpecName.Length > 0)
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"SELECT TOP 1 1
                        //        FROM MES.ToolSpec
                        //        WHERE CompanyId = @CompanyId
                        //        AND ToolSpecName = @ToolSpecName";
                        //    dynamicParameters.Add("CompanyId", CurrentCompany);
                        //    dynamicParameters.Add("ToolSpecName", ToolSpecName);

                        //    result = sqlConnection.Query(sql, dynamicParameters);
                        //    if (result.Count() > 0) throw new SystemException("工具規格【名稱】重複，請重新輸入!");
                        //}
                        //#endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolSpec (CompanyId, ToolSpecNo, ToolSpecName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolSpecId
                                VALUES (@CompanyId, @ToolSpecNo, @ToolSpecName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ToolSpecNo,
                                ToolSpecName,
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

        #region //AddToolGroup -- 工具群組資料新增 -- Shintokuro 2022.12.27
        public string AddToolGroup(string ToolGroupNo, string ToolGroupName)
        {
            try
            {
                if (ToolGroupNo.Length <= 0) throw new SystemException("工具群組【編號】不能為空!");
                if (ToolGroupNo.Length > 50) throw new SystemException("工具群組【編號】長度錯誤!");
                if (ToolGroupName.Length <= 0) throw new SystemException("工具群組【名稱】不能為空!");
                if (ToolGroupName.Length > 50) throw new SystemException("工具群組【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE CompanyId = @CompanyId
                                AND ToolGroupNo = @ToolGroupNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolGroupNo", ToolGroupNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具群組【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具群組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE CompanyId = @CompanyId
                                AND ToolGroupName = @ToolGroupName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolGroupName", ToolGroupName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("工具群組【名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolGroup (CompanyId, ToolGroupNo, ToolGroupName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolGroupId
                                VALUES (@CompanyId, @ToolGroupNo, @ToolGroupName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ToolGroupNo,
                                ToolGroupName,
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

        #region //AddToolClass -- 工具類別資料新增 -- Shintokuro 2022.12.27
        public string AddToolClass(int ToolGroupId, string ToolClassNo, string ToolClassName)
        {
            try
            {
                if (ToolClassNo.Length <= 0) throw new SystemException("工具類別【編號】不能為空!");
                if (ToolClassNo.Length > 50) throw new SystemException("工具類別【編號】長度錯誤!");
                if (ToolClassName.Length <= 0) throw new SystemException("工具類別【名稱】不能為空!");
                if (ToolClassName.Length > 50) throw new SystemException("工具類別【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具群組不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具類別編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolGroupId = @ToolGroupId
                                AND ToolClassNo = @ToolClassNo";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);
                        dynamicParameters.Add("ToolClassNo", ToolClassNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具類別【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具群組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolGroupId = @ToolGroupId
                                AND ToolClassName = @ToolClassName";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);
                        dynamicParameters.Add("ToolClassName", ToolClassName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具類別【名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolClass (ToolGroupId, ToolClassNo, ToolClassName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolClassId
                                VALUES (@ToolGroupId, @ToolClassNo, @ToolClassName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolGroupId,
                                ToolClassNo,
                                ToolClassName,
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

        #region //AddToolCategory -- 工具種類資料新增 -- Shintokuro 2022.12.28
        public string AddToolCategory(int ToolClassId, string ToolCategoryNo, string ToolCategoryName)
        {
            try
            {
                if (ToolCategoryNo.Length <= 0) throw new SystemException("工具類別【編號】不能為空!");
                if (ToolCategoryNo.Length > 50) throw new SystemException("工具類別【編號】長度錯誤!");
                if (ToolCategoryName.Length <= 0) throw new SystemException("工具類別【名稱】不能為空!");
                if (ToolCategoryName.Length > 50) throw new SystemException("工具類別【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具種類編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolClassId = @ToolClassId
                                AND ToolCategoryNo = @ToolCategoryNo";
                        dynamicParameters.Add("ToolClassId", ToolClassId);
                        dynamicParameters.Add("ToolCategoryNo", ToolCategoryNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具種類【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具種類名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolClassId = @ToolClassId
                                AND ToolCategoryName = @ToolCategoryName";
                        dynamicParameters.Add("ToolClassId", ToolClassId);
                        dynamicParameters.Add("ToolCategoryName", ToolCategoryName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("工具種類【名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolCategory (ToolClassId, ToolCategoryNo, ToolCategoryName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolCategoryId
                                VALUES (@ToolClassId, @ToolCategoryNo, @ToolCategoryName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolClassId,
                                ToolCategoryNo,
                                ToolCategoryName,
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

        #region //AddToolModel -- 工具型號資料新增 -- Shintokuro 2022.12.29
        public string AddToolModel(int ToolCategoryId, string ToolModelNo, string ToolModelErpNo, string ToolModelName, int SupplierId)
        {
            try
            {
                if (ToolModelNo.Length <= 0) throw new SystemException("工具型號【編號】不能為空!");
                if (ToolModelNo.Length > 50) throw new SystemException("工具型號【編號】長度錯誤!");
                if (ToolModelErpNo.Length > 50) throw new SystemException("工具ERP型號長度錯誤!");
                if (ToolModelName.Length <= 0) throw new SystemException("工具型號【名稱】不能為空!");
                if (ToolModelName.Length > 50) throw new SystemException("工具型號【名稱】長度錯誤!");
                int ToolModelId = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具種類是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具種類不存在，請重新輸入!");
                        #endregion

                        #region //判斷供應商是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("供應商不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具型號編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelNo = @ToolModelNo";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        dynamicParameters.Add("ToolModelNo", ToolModelNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具型號【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具型號名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelName = @ToolModelName";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        dynamicParameters.Add("ToolModelName", ToolModelName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具型號【名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具型號名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelName = @ToolModelName
                                AND ToolModelName = @ToolModelName";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        dynamicParameters.Add("ToolModelNo", ToolModelNo);
                        dynamicParameters.Add("ToolModelName", ToolModelName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具型號【編號】+【名稱】重複了，請重新輸入!");
                        #endregion

                        #region //判斷工具ERP型號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelErpNo = @ToolModelErpNo";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        dynamicParameters.Add("ToolModelErpNo", ToolModelErpNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具ERP型號重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolModel (ToolCategoryId, ToolModelNo, ToolModelErpNo, ToolModelName, SupplierId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolModelId
                                VALUES (@ToolCategoryId, @ToolModelNo, @ToolModelErpNo, @ToolModelName, @SupplierId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolCategoryId,
                                ToolModelNo,
                                ToolModelErpNo,
                                ToolModelName,
                                SupplierId,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        foreach (var item in insertResult)
                        {
                            ToolModelId = Convert.ToInt32(item.ToolModelId);
                        }                        

                        var ToolSpecId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToolSpecId
                                 FROM MES.ToolSpec a
                                 LEFT JOIN MES.ToolClass b ON a.ToolClassId=b.ToolClassId
                                 LEFT JOIN MES.ToolCategory c ON b.ToolClassId=c.ToolClassId
                                WHERE c.ToolCategoryId = @ToolCategoryId
                                ORDER BY a.ToolSpecId DESC";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                //新增工具型號規格
                                ToolSpecId = Convert.ToInt32(item.ToolSpecId);

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolModelSpec (ToolModelId, ToolSpecId, ToolSpecValue, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolModelSpecId
                                VALUES (@ToolModelId, @ToolSpecId, @ToolSpecValue, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ToolModelId,
                                        ToolSpecId,
                                        ToolSpecValue=0,
                                        Status = "A", //啟用
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                 var ToolSpecResult = sqlConnection.Query(sql, dynamicParameters);
                            }
                        }                       

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

                    //using (SqlConnection sqlConnection = new SqlConnection(OldMesConnectionStrings)) {
                    //    dynamicParameters = new DynamicParameters();
                    //    sql = @"INSERT INTO MFG.KNIFE_MODEL_INFO (MODEL_ID,MODEL_NO, MODEL_NAME, MODEL_DESC, KNIFE_TYPE_ID, INCLINATION, KNIFE_R,KNIFE_SUPPLIER_ID
                    //            , CREATE_USERID, UPDATE_USERID, CREATE_DATE, UPDATE_DATE)                               
                    //            VALUES (@MODEL_ID,@ToolModelNo, @ToolModelName, @ToolModelErpNo, @ToolCategoryId,-1,-1, @SupplierId
                    //            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                    //    dynamicParameters.AddDynamicParams(
                    //        new
                    //        {
                    //            MODEL_ID,
                    //            ToolModelNo,
                    //            ToolModelName,
                    //            ToolModelErpNo,
                    //            ToolCategoryId,
                    //            SupplierId,
                    //            CreateDate,
                    //            LastModifiedDate,
                    //            CreateBy,
                    //            LastModifiedBy
                    //        });
                    //    var insertResult = sqlConnection.Query(sql, dynamicParameters);
                    //    int rowsAffected = insertResult.Count();
                    //    #region //Response
                    //    jsonResponse = JObject.FromObject(new
                    //    {
                    //        status = "success",
                    //        msg = "(" + rowsAffected + " rows affected)",
                    //        data = insertResult
                    //    });
                    //    #endregion
                    //    transactionScope.Complete();
                    //}
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddToolModelSpec -- 工具型號規格新增 -- Shintokuro 2022.12.30
        public string AddToolModelSpec(int ToolModelId, int ToolSpecId, Double ToolSpecValue)
        {
            try
            {
                if (ToolSpecValue <= 0) throw new SystemException("工具型號規格【數值】不能小於等於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具規格是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格不存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolModelSpec (ToolModelId, ToolSpecId, ToolSpecValue, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolModelSpecId
                                VALUES (@ToolModelId, @ToolSpecId, @ToolSpecValue, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolModelId,
                                ToolSpecId,
                                ToolSpecValue,
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

        #region //AddTool -- 工具明細資料新增 -- Shintokuro 2022.12.29
        public string AddTool(int ToolModelId, string ToolNo)
        {
            try
            {
                if (ToolNo.Length > 50) throw new SystemException("工具明細【編號】長度錯誤!");
                int KNIFE_ID = -1, RevToolCategoryId=-1;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolCategoryId
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號不存在，請重新輸入!");
                        foreach (var item in result)
                        {
                            RevToolCategoryId =item.ToolCategoryId;
                        }
                        #endregion

                        #region //撈出工具型號規格
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolModelSpecId
                                FROM MES.ToolModelSpec
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var resultToolModelSpec = sqlConnection.Query(sql, dynamicParameters);
                        if (resultToolModelSpec.Count() <= 0) throw new SystemException("工具型號缺少規格，請重新確認該型號是否有維護規格資訊!");
                        #endregion

                        #region //判斷工具目前最大編號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolNo
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel a1 on a.ToolModelId = a1.ToolModelId
                                INNER JOIN MES.ToolCategory a2 on a1.ToolCategoryId = a2.ToolCategoryId
                                INNER JOIN MES.ToolClass a3 on a2.ToolClassId = a3.ToolClassId
                                INNER JOIN MES.ToolGroup a4 on a3.ToolGroupId = a4.ToolGroupId
                                WHERE a4.CompanyId = @CompanyId
                                AND a.ToolNo = @ToolNo
                                ORDER BY a.ToolId DESC";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolNo", ToolNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            ToolNo = "T" + (Convert.ToInt32(ToolNo.Split('T')[1]) + 1).ToString("00000");
                        }
                        else
                        {
                            ToolNo = "T00001";
                        }

                        #endregion

                        #region //判斷工具明細編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool
                                WHERE ToolNo = @ToolNo";
                        dynamicParameters.Add("ToolNo", ToolNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具明細【編號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Tool (ToolModelId,RevToolCategoryId, ToolNo, UsageCount, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolId
                                VALUES (@ToolModelId,@RevToolCategoryId, @ToolNo, @UsageCount, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolModelId,                                
                                RevToolCategoryId,
                                ToolNo,
                                UsageCount = 0,
                                Status = "S", //關閉(未印出標籤)
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int ToolId = 0;

                        foreach (var item in insertResult)
                        {
                            ToolId = Convert.ToInt32(item.ToolId);
                            KNIFE_ID = Convert.ToInt32(item.ToolId);
                        }

                        #region //判斷工具型號規格是否存在
                        foreach (var item in resultToolModelSpec)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ToolSpecId,ToolSpecValue
                                    FROM MES.ToolModelSpec
                                    WHERE ToolModelSpecId = @ToolModelSpecId";
                            dynamicParameters.Add("ToolModelSpecId", item.ToolModelSpecId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具型號規格資料錯誤，請重新輸入!");
                            var ToolSpecId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolSpecId;
                            var ToolSpecValue = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolSpecValue;

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolSpecLog (ToolId, ToolSpecId, ToolSpecValue, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ToolSpecLogId
                                    VALUES (@ToolId, @ToolSpecId, @ToolSpecValue, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    ToolSpecId,
                                    ToolSpecValue,
                                    Status = "A", //啟用
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
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

                        transactionScope.Complete();
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(OldMesConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MFG.KNIFE_INFO (KNIFE_ID,MODEL_ID, KNIFE_NO, PRINT_STATUS
                                , CREATE_USERID, UPDATE_USERID, CREATE_DATE, UPDATE_DATE)                               
                                VALUES (@KNIFE_ID,@ToolModelId, @ToolNo, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                KNIFE_ID,
                                ToolModelId,
                                ToolNo,
                                Status = "Y",
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

        #region //AddToolBatch -- 工具明細資料新增(批量) -- Shintokuro 2022.01.07
        public string AddToolBatch(int ToolModelId, int ProductionNum)
        {
            try
            {
                if (ProductionNum < 0) throw new SystemException("生產工具數量不可以小於0");                
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        var ToolNo = "";
                        int MaxToolId = 0;
                        int RevToolCategoryId = -1;

                        #region //判斷工具型號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolCategoryId
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        //var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號不存在，請重新輸入!");
                        foreach (var item in result) {
                            RevToolCategoryId = item.ToolCategoryId;
                        }
                        #endregion

                        #region //撈出工具型號規格
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolModelSpecId
                                FROM MES.ToolModelSpec
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var resultToolModelSpec = sqlConnection.Query(sql, dynamicParameters);
                        if (resultToolModelSpec.Count() <= 0) throw new SystemException("工具型號缺少規格，請重新確認該型號是否有維護規格資訊!");
                        #endregion

                        #region //判斷該工具編號目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolNo
                                FROM MES.Tool
                                WHERE ToolNo LIKE 'T%'
                                ORDER BY ToolId DESC";
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                ToolNo = item.ToolNo;
                            }                            
                            if (ToolNo.Contains("T")!=true) {
                                ToolNo = "T00001";
                            }
                        }                       
                        #endregion

                        int ToolNoMaxsor = Convert.ToInt32(ToolNo.Split('T')[1]);
                        for (var i = 1; i <= ProductionNum; i++)
                        {
                            ToolNoMaxsor++;

                            string ToolNoMaxsorNo = "T" + ToolNoMaxsor.ToString("00000");

                            #region //判斷工具明細編號是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tool
                                    WHERE ToolNo = @ToolNo";
                            dynamicParameters.Add("ToolNo", ToolNoMaxsorNo);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("工具明細【編號】重複，請重新輸入!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.Tool (ToolModelId,RevToolCategoryId, ToolNo, UsageCount, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolId
                                VALUES (@ToolModelId,@RevToolCategoryId, @ToolNo, @UsageCount, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolModelId,
                                    RevToolCategoryId,
                                    ToolNo = ToolNoMaxsorNo,
                                    UsageCount = 0,
                                    Status = "S", //關閉(未印出標籤)
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            int ToolId = 0;
                            foreach (var item in insertResult)
                            {
                                ToolId = Convert.ToInt32(item.ToolId);
                            }

                            //特殊需求 --於舊系統建立刀具
                            //string status = AddKnifeToMes(ToolModelId, ToolId, ToolNoMaxsorNo);

                            #region //判斷工具明細是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tool
                                    WHERE ToolId = @ToolId";
                            dynamicParameters.Add("ToolId", ToolId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具明細不存在，請重新輸入!");
                            #endregion

                            #region //判斷工具型號規格是否存在
                            foreach (var item in resultToolModelSpec)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 ToolSpecId,ToolSpecValue
                                        FROM MES.ToolModelSpec
                                        WHERE ToolModelSpecId = @ToolModelSpecId";
                                dynamicParameters.Add("ToolModelSpecId", item.ToolModelSpecId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("工具型號規格資料錯誤，請重新輸入!");
                                var ToolSpecId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolSpecId;
                                var ToolSpecValue = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolSpecValue;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolSpecLog (ToolId, ToolSpecId, ToolSpecValue, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.ToolSpecLogId
                                        VALUES (@ToolId, @ToolSpecId, @ToolSpecValue, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ToolId,
                                        ToolSpecId,
                                        ToolSpecValue,
                                        Status = "A", //啟用
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = insertResult.Count();
                            }
                            #endregion

                            rowsAffected += insertResult.Count();

                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //AddToolSpecLog -- 工具明細規格新增 -- Shintokuro 2022.12.30
        public string AddToolSpecLog(int ToolId, int ToolSpecId, Double ToolSpecValue)
        {
            try
            {
                if (ToolSpecValue <= 0) throw new SystemException("工具型號規格【數值】不能小於等於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具明細是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具規格是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格不存在，請重新輸入!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolSpecLog (ToolId, ToolSpecId, ToolSpecValue, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolSpecLogId
                                VALUES (@ToolId, @ToolSpecId, @ToolSpecValue, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolId,
                                ToolSpecId,
                                ToolSpecValue,
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

        #region //AddToolInventory -- 工具倉庫資料新增 -- Shintokuro 2022.12.30
        public string AddToolInventory(string ToolInventoryNo, string ToolInventoryName, string ToolInventoryDesc, int ShopId)
        {
            try
            {
                if (ToolInventoryNo.Length <= 0) throw new SystemException("工具倉庫【編號】不能為空!");
                if (ToolInventoryNo.Length > 50) throw new SystemException("工具倉庫【編號】長度錯誤!");
                if (ToolInventoryName.Length <= 0) throw new SystemException("工具倉庫【名稱】不能為空!");
                if (ToolInventoryName.Length > 50) throw new SystemException("工具倉庫【名稱】長度錯誤!");
                if (ToolInventoryDesc.Length <= 0) throw new SystemException("工具倉庫【描述】不能為空!");
                if (ToolInventoryDesc.Length > 50) throw new SystemException("工具倉庫【描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具倉庫編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolInventory
                                WHERE ToolInventoryNo = @ToolInventoryNo";
                        dynamicParameters.Add("ToolInventoryNo", ToolInventoryNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具倉庫【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷車間是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE ShopId = @ShopId";
                        dynamicParameters.Add("ShopId", ShopId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolInventory (ToolInventoryNo, ToolInventoryName, ToolInventoryDesc, ShopId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolInventoryId
                                VALUES (@ToolInventoryNo, @ToolInventoryName, @ToolInventoryDesc, @ShopId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolInventoryNo,
                                ToolInventoryName,
                                ToolInventoryDesc,
                                ShopId,
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

        #region //AddToolLocator -- 工具儲位資料新增 -- Shintokuro 2023.01.03
        public string AddToolLocator(string ToolLocatorNo, string ToolLocatorName, string ToolLocatorDesc, int ToolInventoryId)
        {
            try
            {
                if (ToolLocatorNo.Length <= 0) throw new SystemException("工具儲位【編號】不能為空!");
                if (ToolLocatorNo.Length > 50) throw new SystemException("工具儲位【編號】長度錯誤!");
                if (ToolLocatorName.Length <= 0) throw new SystemException("工具儲位【名稱】不能為空!");
                if (ToolLocatorName.Length > 50) throw new SystemException("工具儲位【名稱】長度錯誤!");
                if (ToolLocatorDesc.Length <= 0) throw new SystemException("工具儲位【描述】不能為空!");
                if (ToolLocatorDesc.Length > 50) throw new SystemException("工具儲位【描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具倉庫編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolLocator
                                WHERE ToolLocatorNo = @ToolLocatorNo";
                        dynamicParameters.Add("ToolLocatorNo", ToolLocatorNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具儲位【編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具倉庫是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolInventory
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具倉庫不存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolLocator (ToolLocatorNo, ToolLocatorName, ToolLocatorDesc, ToolInventoryId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolLocatorId
                                VALUES (@ToolLocatorNo, @ToolLocatorName, @ToolLocatorDesc, @ToolInventoryId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolLocatorNo,
                                ToolLocatorName,
                                ToolLocatorDesc,
                                ToolInventoryId,
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

        #region //AddToolTransactions -- 工具交易入庫新增 -- Shintokuro 2023.01.03
        public string AddToolTransactions(int ToolId, string TransactionDate, string TransactionType, int ToolLocatorId, string TransactionReason, int TraderId, string IsLimit,  int LastToolId, int ProcessingQty)
        {
            try
            {
                if (TransactionDate.Length <= 0) throw new SystemException("交易日不能為空!");
                if (TransactionType.Length <= 0) throw new SystemException("交易類型不能為空!");
                if (ToolLocatorId <= 0) throw new SystemException("儲位不能為空!");
                if (TraderId <= 0) throw new SystemException("交易者不能為空!");
                if (TransactionReason.Length > 50) throw new SystemException("交易原因長度錯誤!");

                int? nulldate = null;
                int ToolLocatorIdNow = 0;
                string ToolLocatorNameNow = "";
                string ToolInventoryNameNow = "";
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId); 

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具不存在，請重新輸入!");
                        #endregion

                        #region //判斷工具儲位是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                        string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                        string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;

                        #endregion

                        #region //判斷交易者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserName
                                FROM BAS.[User]
                                WHERE UserId = @TraderId";
                        dynamicParameters.Add("TraderId", TraderId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員不存在，請重新輸入!");
                        string TraderName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).UserName;

                        #endregion

                        #region //判斷工具是否有入庫過
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @ToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                        dynamicParameters.Add("ToolId", ToolId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】"); 
                        }
                        else
                        {
                            ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同"); 
                        }
                        #endregion

                        if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                        {                         

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    TransactionType = "Out",
                                    TransactionDate,
                                    ToolLocatorId = ToolLocatorIdNow,
                                    TraderId,
                                    TransactionReason = CreateDate + " - " + TraderName + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult01 = sqlConnection.Query(sql, dynamicParameters); 

                            rowsAffected += insertResult01.Count();
                        }


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolId,
                                TransactionType,
                                TransactionDate,
                                ToolLocatorId,
                                TraderId,
                                TransactionReason = TransactionReason != "" ? "【原因:" + TransactionReason + "】" + CreateDate + " - " + TraderName + "將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】" : CreateDate + " - " + TraderName + "將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                ProcessingQty
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        if (IsLimit == "Y" && LastToolId != -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"select a.ToolLocatorNo,a.ToolLocatorName,b.ToolInventoryName 
                                    from MES.ToolLocator a 
                                    INNER JOIN MES.ToolInventory b ON b.ToolInventoryId = a.ToolInventoryId
                                    WHERE ToolLocatorId = @ToolLocatorId";
                            dynamicParameters.Add("ToolLocatorId", ToolLocatorId);
                            var ToolLocatorNameNow2 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            var ToolInventoryNameNow2 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@LastToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastToolId,
                                    TransactionType = "Out",
                                    TransactionDate,
                                    ToolLocatorId,
                                    TraderId,
                                    TransactionReason = "【原因: (報廢移出)】" + CreateDate + " - " + TraderName + "從 【庫房 -" + ToolInventoryNameNow2 + "】的【儲位 -" + ToolLocatorNameNow2 + "】移出工具",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult02 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult02.Count();

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@LastToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams( 
                                new
                                {
                                    LastToolId,
                                    TransactionType,
                                    TransactionDate,
                                    ToolLocatorId = 41,
                                    TraderId,
                                    TransactionReason = "【原因: " + TransactionReason + " (報廢)】" + CreateDate + " - " + TraderName + "將工具入庫到 【庫房 - 報廢區】的【儲位 -報廢區】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty
                                });
                            var insertResult03 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult03.Count();

                        }

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

        #region//Update

        #region //UpdateToolGroup -- 工具群組資料更新 -- Shintokuro 2022.12.27
        public string UpdateToolGroup(int ToolGroupId, string ToolGroupName
            )
        {
            try
            {
                if (ToolGroupName.Length <= 0) throw new SystemException("工具群組【名稱】不能為空!");
                if (ToolGroupName.Length > 50) throw new SystemException("工具群組【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具群組資料錯誤!");
                        #endregion

                        #region //判斷工具群組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE CompanyId = @CompanyId
                                AND ToolGroupName = @ToolGroupName
                                AND ToolGroupId != @ToolGroupId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolGroupName", ToolGroupName);
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("工具群組【名稱】重複，請重新輸入!");
                        #endregion




                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolGroup SET
                                ToolGroupName = @ToolGroupName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolGroupName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolGroupId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolGroupStatus -- 工具群組狀態更新 -- Shintokuro 2022.12.27
        public string UpdateToolGroupStatus(int ToolGroupId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolGroup
                                WHERE CompanyId = @CompanyId
                                AND ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具群組資料錯誤!");

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

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                INNER JOIN MES.ToolGroup e on d.ToolGroupId = e.ToolGroupId
                                WHERE a.Status = 'A'
                                AND e.ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該群組下已經有工具啟用,不可以修改!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolGroup SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolGroupId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolClass -- 工具類別資料更新 -- Shintokuro 2022.12.27
        public string UpdateToolClass(int ToolClassId, int ToolGroupId, string ToolClassName
            )
        {
            try
            {
                if (ToolClassName.Length <= 0) throw new SystemException("工具類別【名稱】不能為空!");
                if (ToolClassName.Length > 50) throw new SystemException("工具類別【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具群組資料錯誤!");
                        #endregion

                        #region //判斷工具類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolGroupId
                                FROM MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");
                        var ToolGroupIdOld = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolGroupId;

                        #endregion

                        #region //判斷工具群組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolGroupId = @ToolGroupId
                                AND ToolClassName = @ToolClassName
                                AND ToolClassId != @ToolClassId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);
                        dynamicParameters.Add("ToolClassName", ToolClassName);
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具群組【名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                WHERE a.Status = 'A'
                                AND d.ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            if (ToolGroupIdOld != ToolGroupId) throw new SystemException("該類別下已經有工具啟用,群組不可以修改!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolClass SET
                                ToolGroupId = @ToolGroupId,
                                ToolClassName = @ToolClassName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolGroupId,
                                ToolClassName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolClassId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolClassStatus -- 工具類別狀態更新 -- Shintokuro 2022.12.27
        public string UpdateToolClassStatus(int ToolClassId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具類別資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");

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

                        #region //判斷該類別下的種類是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該類別下的種類已經有資料了,所以狀態不可以更改");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                WHERE a.Status = 'A'
                                AND d.ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該類別下已經有工具啟用,不可以修改!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolClass SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolClassId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolCategory -- 工具種類資料更新 -- Shintokuro 2022.12.28
        public string UpdateToolCategory(int ToolCategoryId, int ToolClassId, string ToolCategoryName
            )
        {
            try
            {
                if (ToolCategoryName.Length <= 0) throw new SystemException("工具種類【名稱】不能為空!");
                if (ToolCategoryName.Length > 50) throw new SystemException("工具種類【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");
                        #endregion

                        #region //判斷工具種類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolClassId
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");
                        var ToolClassIdOld = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolClassId;

                        #endregion

                        #region //判斷工具群組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolClassId = @ToolClassId
                                AND ToolCategoryName = @ToolCategoryName
                                AND ToolCategoryId != @ToolCategoryId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);
                        dynamicParameters.Add("ToolCategoryName", ToolCategoryName);
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具種類【名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                WHERE a.Status = 'A'
                                AND c.ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            if (ToolClassIdOld != ToolClassId) throw new SystemException("該種類下已經有工具啟用,類別不可以修改!");
                        }
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolCategory SET
                                ToolClassId = @ToolClassId,
                                ToolCategoryName = @ToolCategoryName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolClassId,
                                ToolCategoryName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolCategoryId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolCategoryStatus -- 工具種類狀態更新 -- Shintokuro 2022.12.28
        public string UpdateToolCategoryStatus(int ToolCategoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具類別資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");

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

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                WHERE a.Status = 'A'
                                AND c.ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該種類下已經有工具啟用,不可以修改!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolCategory SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolCategoryId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolModel -- 工具型號資料更新 -- Shintokuro 2022.12.29
        public string UpdateToolModel(int ToolModelId, int ToolCategoryId, string ToolModelErpNo, string ToolModelName, int SupplierId
            )
        {
            try
            {
                if (ToolModelErpNo.Length > 50) throw new SystemException("工具ERP型號長度錯誤!");
                if (ToolModelName.Length <= 0) throw new SystemException("工具型號【名稱】不能為空!");
                if (ToolModelName.Length > 50) throw new SystemException("工具型號【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具種類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具種類資料錯誤!");
                        #endregion

                        #region //判斷工具型號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolCategoryId,ToolModelNo
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號資料錯誤!");
                        var ToolCategoryIdOld = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolCategoryId;
                        var ToolModelNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolModelNo;
                        #endregion

                        if (SupplierId > 0)
                        {
                            #region //判斷供應商是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                            dynamicParameters.Add("SupplierId", SupplierId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("供應商不存在，請重新輸入!");
                            #endregion
                        }

                        #region //判斷工具型號名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelName = @ToolModelName
                                AND ToolModelId != @ToolModelId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);
                        dynamicParameters.Add("ToolModelName", ToolModelName);
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具型號【名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷工具型號名稱+編碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelName = @ToolModelName
                                AND ToolModelName = @ToolModelName
                                AND ToolModelId != @ToolModelId";
                        dynamicParameters.Add("ToolModelNo", ToolModelNo);
                        dynamicParameters.Add("ToolModelName", ToolModelName);
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具型號【編號】+【名稱】重複了，請重新輸入!");
                        #endregion

                        #region //判斷工具ERP型號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelErpNo = @ToolModelErpNo
                                AND ToolModelId != @ToolModelId";
                        dynamicParameters.Add("ToolModelErpNo", ToolModelErpNo);
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("工具ERP型號重複，請重新輸入!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                WHERE a.Status = 'A'
                                AND b.ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            if (ToolCategoryIdOld != ToolCategoryId) throw new SystemException("該型號下已經有工具啟用,不可以修改種類!");
                        }
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolModel SET
                                ToolCategoryId = @ToolCategoryId,
                                ToolModelErpNo = @ToolModelErpNo,
                                ToolModelName = @ToolModelName,
                                SupplierId = @SupplierId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolCategoryId,
                                ToolModelErpNo,
                                ToolModelName,
                                SupplierId,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolModelId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolModelStatus -- 工具型號狀態更新 -- Shintokuro 2022.12.29
        public string UpdateToolModelStatus(int ToolModelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具類別資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");

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

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                WHERE a.Status = 'A'
                                AND b.ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該型號下已經有工具啟用,不可以修改!");
                        #endregion



                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolModel SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolModelId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateTool -- 工具明細資料更新 -- Shintokuro 2022.12.29
        public string UpdateTool(int ToolId, int ToolModelId, int RevToolCategoryId
            )
        {
            try
            {

                int? nullData = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolCategoryId
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號資料錯誤!");
                        int ToolCategoryIdBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolCategoryId;

                        #endregion

                        #region //判斷工具明細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT [Status],RevToolCategoryId
                                FROM MES.Tool
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細資料錯誤!");
                        var ToolStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).Status;
                        int? RevToolCategoryIdBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).RevToolCategoryId;

                        if (RevToolCategoryId == ToolCategoryIdBase)
                        {
                            if (RevToolCategoryIdBase == null)
                            {
                                if (ToolStatus == "A") throw new SystemException("該工具已經啟用了,不可以更改資料!!!");
                            }
                        }
                        #endregion

                        #region //判斷更改種類後資料是否正確
                        if (ToolCategoryIdBase != RevToolCategoryId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @RevToolCategoryId";
                            dynamicParameters.Add("RevToolCategoryId", RevToolCategoryId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具種類資料錯誤!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Tool SET
                                ToolModelId = @ToolModelId,
                                RevToolCategoryId = @RevToolCategoryId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolId = @ToolId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolModelId,
                                RevToolCategoryId = RevToolCategoryId != ToolCategoryIdBase ? RevToolCategoryId : nullData,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolStatus -- 工具型號狀態更新 -- Shintokuro 2022.12.29
        public string UpdateToolStatus(int ToolId, string ToolNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        string status = "";

                        if (ToolNo == "")
                        {
                            #region //判斷工具明細資訊是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 Status
                                FROM MES.Tool
                                WHERE ToolId = @ToolId";
                            dynamicParameters.Add("ToolId", ToolId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具明細資料錯誤!");

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
                        }
                        else
                        {
                            ToolNo = '%' + ToolNo + '%';
                            #region //判斷工具明細資訊是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 Status ,ToolId
                                    FROM MES.Tool
                                    WHERE ToolNo LIKE @ToolNo ";
                            dynamicParameters.Add("ToolNo", ToolNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【工具進出系統錯誤】- 工具明細資料");
                            ToolId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolId;
                            #endregion

                            #region //判斷工具明細資訊是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tool　ａ
                                    INNER JOIN MES.ToolSpecLog b on a.ToolId = b.ToolId
                                    WHERE ToolNo LIKE @ToolNo ";
                            dynamicParameters.Add("ToolNo", ToolNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【工具進出系統錯誤】- 工具明細資料缺少規格設定!不能列印");
                            status = "A";
                            #endregion
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Tool SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolId = @ToolId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolSpec -- 工具規格資料更新 -- Shintokuro 2022.12.29
        public string UpdateToolSpec(int ToolSpecId, string ToolSpecName
            )
        {
            try
            {
                if (ToolSpecName.Length > 50) throw new SystemException("工具規格【名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格資料錯誤!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolSpec SET
                                ToolSpecName = @ToolSpecName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolSpecName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolSpecId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolSpecStatus -- 工具規格狀態更新 -- Shintokuro 2022.12.29
        public string UpdateToolSpecStatus(int ToolSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具規格資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolSpec
                                WHERE CompanyId = @CompanyId
                                AND ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格資料錯誤!");

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
                        sql = @"UPDATE MES.ToolSpec SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolSpecId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolModelSpec -- 工具型號規格資料更新 -- Shintokuro 2022.12.30
        public string UpdateToolModelSpec(int ToolModelSpecId, int ToolSpecId, Double ToolSpecValue
            )
        {
            try
            {
                if (ToolSpecValue <= 0) throw new SystemException("工具型號規格【數值】不能小於等於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModelSpec
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.Add("ToolModelSpecId", ToolModelSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號規格資料錯誤!");
                        #endregion

                        #region //判斷工具規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格資料錯誤!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolModelSpec SET
                                ToolSpecId = @ToolSpecId,
                                ToolSpecValue = @ToolSpecValue,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolSpecId,
                                ToolSpecValue,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolModelSpecId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolModelSpecStatus -- 工具型號規格狀態更新 -- Shintokuro 2022.12.30
        public string UpdateToolModelSpecStatus(int ToolModelSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號規格資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolModelSpec
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.Add("ToolModelSpecId", ToolModelSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號規格資料錯誤!");

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
                        sql = @"UPDATE MES.ToolModelSpec SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolModelSpecId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolSpecLog -- 工具型號規格資料更新 -- Shintokuro 2022.12.30
        public string UpdateToolSpecLog(int ToolSpecLogId, int ToolSpecId, Double ToolSpecValue
            )
        {
            try
            {
                if (ToolSpecValue <= 0) throw new SystemException("工具型號規格【數值】不能小於等於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具明細規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpecLog
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.Add("ToolSpecLogId", ToolSpecLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細規格資料錯誤!");
                        #endregion

                        #region //判斷工具規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格資料錯誤!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolSpecLog SET
                                ToolSpecId = @ToolSpecId,
                                ToolSpecValue = @ToolSpecValue,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolSpecId,
                                ToolSpecValue,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolSpecLogId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolSpecLogStatus -- 工具型號規格狀態更新 -- Shintokuro 2022.12.30
        public string UpdateToolSpecLogStatus(int ToolSpecLogId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具明細規格資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolSpecLog
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.Add("ToolSpecLogId", ToolSpecLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細規格資料錯誤!");

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
                        sql = @"UPDATE MES.ToolSpecLog SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolSpecLogId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolInventory -- 工具倉庫資料更新 -- Shintokuro 2022.12.30
        public string UpdateToolInventory(int ToolInventoryId, string ToolInventoryName, string ToolInventoryDesc, int ShopId
            )
        {
            try
            {
                if (ToolInventoryName.Length <= 0) throw new SystemException("工具倉庫【名稱】不能為空!");
                if (ToolInventoryName.Length > 50) throw new SystemException("工具倉庫【名稱】長度錯誤!");
                if (ToolInventoryDesc.Length <= 0) throw new SystemException("工具倉庫【描述】不能為空!");
                if (ToolInventoryDesc.Length > 50) throw new SystemException("工具倉庫【描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具倉庫資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolInventory
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具倉庫資料錯誤!");
                        #endregion

                        #region //判斷車間是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE ShopId = @ShopId";
                        dynamicParameters.Add("ShopId", ShopId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolInventory SET
                                ToolInventoryName = @ToolInventoryName,
                                ToolInventoryDesc = @ToolInventoryDesc,
                                ShopId = @ShopId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolInventoryName,
                                ToolInventoryDesc,
                                ShopId,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolInventoryId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolInventoryStatus -- 工具倉庫狀態更新 -- Shintokuro 2022.12.30
        public string UpdateToolInventoryStatus(int ToolInventoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具倉庫資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolInventory
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具倉庫資料錯誤!");

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
                        sql = @"UPDATE MES.ToolInventory SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolInventoryId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolLocator-- 工具儲位資料更新 -- Shintokuro 2023.01.03
        public string UpdateToolLocator(int ToolLocatorId, string ToolLocatorName, string ToolLocatorDesc
            )
        {
            try
            {
                if (ToolLocatorName.Length <= 0) throw new SystemException("工具儲位【名稱】不能為空!");
                if (ToolLocatorName.Length > 50) throw new SystemException("工具儲位【名稱】長度錯誤!");
                if (ToolLocatorDesc.Length <= 0) throw new SystemException("工具儲位【描述】不能為空!");
                if (ToolLocatorDesc.Length > 50) throw new SystemException("工具儲位【描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具儲位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolLocator
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolLocator SET
                                ToolLocatorName = @ToolLocatorName,
                                ToolLocatorDesc = @ToolLocatorDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolLocatorName,
                                ToolLocatorDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolLocatorId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateToolLocatorStatus -- 工具儲位狀態更新 -- Shintokuro 2023.01.03
        public string UpdateToolLocatorStatus(int ToolLocatorId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具儲位資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ToolLocator
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位資料錯誤!");

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
                        sql = @"UPDATE MES.ToolLocator SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolLocatorId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region//Delete

        #region //DeleteToolGroup -- 工具群組刪除 -- Ted 2022.12.27
        public string DeleteToolGroup(int ToolGroupId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolGroup
                                WHERE CompanyId = @CompanyId
                                AND ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具群組資料錯誤!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                INNER JOIN MES.ToolGroup e on d.ToolGroupId = e.ToolGroupId
                                WHERE a.Status = 'A'
                                AND e.ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該群組下已經有工具啟用,不可以刪除!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.Tool a
								INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
								INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
								INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                WHERE d.ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.ToolModel a
								INNER JOIN MES.ToolCategory b on a.ToolCategoryId = b.ToolCategoryId
								INNER JOIN MES.ToolClass c on b.ToolClassId = c.ToolClassId
                                WHERE c.ToolGroupId  = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.ToolCategory a
								INNER JOIN MES.ToolClass b on a.ToolClassId = b.ToolClassId
                                WHERE b.ToolGroupId  = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.ToolClass a
                                WHERE a.ToolGroupId  = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除群組table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolGroup
                                WHERE ToolGroupId = @ToolGroupId";
                        dynamicParameters.Add("ToolGroupId", ToolGroupId);

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

        #region //DeleteToolClass -- 工具類別刪除 -- Ted 2022.12.28
        public string DeleteToolClass(int ToolClassId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具類別資料錯誤!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                INNER JOIN MES.ToolClass d on c.ToolClassId = d.ToolClassId
                                WHERE a.Status = 'A'
                                AND d.ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該類別下已經有工具啟用,不可以刪除!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.Tool a
								INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
								INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                WHERE c.ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.ToolModel a
								INNER JOIN MES.ToolCategory b on a.ToolCategoryId = b.ToolCategoryId
                                WHERE b.ToolClassId  = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.ToolCategory a
                                WHERE a.ToolClassId  = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolClass
                                WHERE ToolClassId = @ToolClassId";
                        dynamicParameters.Add("ToolClassId", ToolClassId);

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

        #region //DeleteToolCategory -- 工具種類刪除 -- Ted 2022.12.28
        public string DeleteToolCategory(int ToolCategoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具種類資料錯誤!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                INNER JOIN MES.ToolCategory c on b.ToolCategoryId = c.ToolCategoryId
                                WHERE a.Status = 'A'
                                AND c.ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該種類下已經有工具啟用,不可以刪除!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM MES.Tool a
								INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                WHERE b.ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolModel
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolCategory
                                WHERE ToolCategoryId = @ToolCategoryId";
                        dynamicParameters.Add("ToolCategoryId", ToolCategoryId);

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

        #region //DeleteToolModel -- 工具型號刪除 -- Ted 2022.12.28
        public string DeleteToolModel(int ToolModelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具型號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號資料錯誤!");
                        #endregion

                        #region //判斷工具是否有啟用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tool a
                                INNER JOIN MES.ToolModel b on a.ToolModelId = b.ToolModelId
                                WHERE a.Status = 'A'
                                AND b.ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該型號下已經有工具啟用,不可以刪除!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Tool
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                       
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolModelSpec
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion


                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolModel
                                WHERE ToolModelId = @ToolModelId";
                        dynamicParameters.Add("ToolModelId", ToolModelId);

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

        #region //DeleteTool-- 工具明細刪除 -- Ted 2022.12.28
        public string DeleteTool(int ToolId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具明細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT [Status]
                                FROM MES.Tool
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細資料錯誤!");
                        var ToolStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).Status;
                        if (ToolStatus == "A") throw new SystemException("該工具明細已經啟用,不能刪除!");
                        #endregion

                        #region //判斷工具明細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolTransactions
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該工具明細已經有交易紀錄,不能刪除!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolSpecLog
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Tool
                                WHERE ToolId = @ToolId";
                        dynamicParameters.Add("ToolId", ToolId);

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

        #region //DeleteToolSpec -- 工具規格刪除 -- Ted 2022.12.29
        public string DeleteToolSpec(int ToolSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpec
                                WHERE CompanyId = @CompanyId
                                AND ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具規格資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除群組table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolSpec
                                WHERE ToolSpecId = @ToolSpecId";
                        dynamicParameters.Add("ToolSpecId", ToolSpecId);

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

        #region //DeleteToolModelSpec -- 工具型號規格刪除 -- Ted 2022.12.30
        public string DeleteToolModelSpec(int ToolModelSpecId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolModelSpec
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.Add("ToolModelSpecId", ToolModelSpecId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具型號規格資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolModelSpec
                                WHERE ToolModelSpecId = @ToolModelSpecId";
                        dynamicParameters.Add("ToolModelSpecId", ToolModelSpecId);

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

        #region //DeleteToolSpecLog -- 工具明細規格刪除 -- Ted 2022.12.30
        public string DeleteToolSpecLog(int ToolSpecLogId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具明細規格資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolSpecLog
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.Add("ToolSpecLogId", ToolSpecLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具明細規格資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolSpecLog
                                WHERE ToolSpecLogId = @ToolSpecLogId";
                        dynamicParameters.Add("ToolSpecLogId", ToolSpecLogId);

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

        #region //DeleteToolInventory -- 工具倉庫刪除 -- Ted 2022.12.30
        public string DeleteToolInventory(int ToolInventoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具倉庫資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolInventory
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具倉庫資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolLocator
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolInventory
                                WHERE ToolInventoryId = @ToolInventoryId";
                        dynamicParameters.Add("ToolInventoryId", ToolInventoryId);

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

        #region //DeleteToolLocator-- 工具儲位刪除 -- Ted 2023.01.03
        public string DeleteToolLocator(int ToolLocatorId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷工具儲位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolLocator
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位資料錯誤!");
                        #endregion

                        int rowsAffected = 0;



                        #region //刪除主table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolLocator
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

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

        #region//特殊需求

        #region//於舊系統建立新刀具
        public string AddKnifeToMes(int MODEL_ID,int KNIFE_ID, string KNIFE_NO)
        {
            string status = "";
            try {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(OldMesConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MFG.KNIFE_INFO (KNIFE_ID,MODEL_ID, KNIFE_NO, PRINT_STATUS
                                , CREATE_USERID, UPDATE_USERID, CREATE_DATE, UPDATE_DATE) 
                                OUTPUT INSERTED.KNIFE_ID
                                VALUES (@KNIFE_ID,@MODEL_ID, @KNIFE_NO, @Status
                               , @CreateBy, @LastModifiedBy , @CreateDate, @LastModifiedDate)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                KNIFE_ID,
                                MODEL_ID,
                                KNIFE_NO,
                                Status = "Y",
                                CreateBy,
                                LastModifiedBy,
                                CreateDate,
                                LastModifiedDate
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
                    }
                    transactionScope.Complete();
                    status = "success";
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
                status = "error";
            }
            return status;
        }
        #endregion

        #endregion

    }
}
