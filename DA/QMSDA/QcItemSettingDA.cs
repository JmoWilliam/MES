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

namespace QMSDA
{
    public class QcItemSettingDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public string CreateUserNo = "";

        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();

        public QcItemSettingDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
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
        #region//GetQcGroup -- 量測群組 取得 --Ding 2022.10.17
        public string GetQcGroup(int QcGroupId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcGroupId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcGroupName+'('+a.QcGroupNo+')' AS QcGroupName";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcGroup a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (QcGroupId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGroupId", @" AND a.QcGroupId= @QcGroupId", QcGroupId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "";
                    sqlQuery.pageIndex = -1;
                    sqlQuery.pageSize = -1;
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

        #region//GetQcClass -- 量測類別 取得 --Ding 2022.10.17
        public string GetQcClass(int QcGroupId, int QcClassId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.QcClassId,a.QcClassName+'('+a.QcClassNo+')' AS QcClassName
                            FROM QMS.QcClass a
                            WHERE 1=1 ";
                    if (QcClassId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassId", @" AND a.QcClassId= @QcClassId", QcClassId);
                    if (QcGroupId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGroupId", @" AND a.QcGroupId= @QcGroupId", QcGroupId);
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

        #region//GetQcItem -- 量測項目 取得 --Ding 2022.10.17
        public string GetQcItem(int QcItemId, int QcClassId, int QcGroupId, string QcItemName, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.QcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcItemNo, a.QcItemName+'('+a.QcItemNo+')' AS QcItemName, a.QcItemName AS QcItemName1,
                            b.QcClassId, b.QcClassName+'('+b.QcClassNo+')' AS QcClassName,
                            c.QcGroupId, c.QcGroupName+'('+c.QcGroupNo+')' AS QcGroupName,
                            a.QcItemType, a.QcItemDesc";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcItem a
                        INNER JOIN QMS.QcClass b ON a.QcClassId=b.QcClassId
                        INNER JOIN QMS.QcGroup c ON b.QcGroupId=c.QcGroupId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (QcClassId > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassId", @" AND b.QcClassId= @QcClassId", QcClassId);
                    if (QcGroupId > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupId", @" AND c.QcGroupId= @QcGroupId", QcGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemName", @" AND a.QcItemName LIKE '%' + @QcItemName + '%'", QcItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemName", @" OR a.QcItemNo = @QcItemName", QcItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.QcItemNo LIKE '%' + @SearchKey + '%' OR a.QcItemName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcItemName+'('+a.QcItemNo+')' ASC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetUnitOfQcMeasure -- 量測單位 取得 --Ding 2022.10.18
        public string GetUnitOfQcMeasure(int QcUomId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.QcUomId,a.CompanyId,a.QcUomNo+'('+a.QcUomDesc+')' QcUomNo
                            FROM QMS.UnitOfQcMeasure a
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND a.CompanyId= @CompanyId", CurrentCompany);
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

        #region//GetQcNotice --取得送檢單資料
        public string GetQcNotice(int QcNoticeId,string QcNoticeNo, string QcNoticeType, string WoErpPrefix,string WoErpNo,int WoSeq,
            string MtlItemNo,string MtlItemName,int ProcessId,
            string StartDate, string EndDate, int DepartmentId,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcNoticeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.ControlId,a.RoutingProcessId,a.MoId,a.QcNoticeType,a.QcNoticeNo,a.QcNoticeQty,a.Remark,a.Status,a.DepartmentId,a.ResetFlag
                            ,d.MtlItemNo,d.MtlItemName
                            ,f.WoErpPrefix,f.WoErpNo,e.WoSeq
                            ,h.TypeDesc
                            ,j.ProcessAlias,i.UserName,FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                            ,k.DepartmentNo, k.DepartmentName
                            ,l.MtlItemNo WoMtlItemNo, l.MtlItemName WOMtlItemName";
                    sqlQuery.mainTables =
                        @" FROM QMS.QcNotice a
                             LEFT JOIN PDM.RdDesignControl b ON a.ControlId=b.ControlId
                             LEFT JOIN PDM.RdDesign c ON b.DesignId=c.DesignId
                             LEFT JOIN PDM.MtlItem d ON c.MtlItemId=d.MtlItemId
                             LEFT JOIN MES.ManufactureOrder e ON a.MoId=e.MoId
                             LEFT JOIN MES.WipOrder f ON e.WoId=f.WoId 
                             LEFT JOIN (
                                SELECT *
                                FROM BAS.[Type]
                                WHERE TypeSchema = 'QMS' 
                             )h ON a.QcNoticeType=h.TypeName
                             LEFT JOIN BAS.[User] i ON a.CreateBy=i.UserId
                             LEFT JOIN MES.RoutingProcess j ON a.RoutingProcessId=j.RoutingProcessId
                             LEFT JOIN BAS.Department k ON a.DepartmentId = k.DepartmentId
                             LEFT JOIN PDM.MtlItem l ON l.MtlItemId = f.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND ( c.CompanyId = @CompanyId OR f.CompanyId = @CompanyId)";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (!QcNoticeId.Equals(-1)) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeId", @" AND a.QcNoticeId= @QcNoticeId", QcNoticeId);
                    if (!QcNoticeNo.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeNo", @" AND a.QcNoticeNo= @QcNoticeNo", QcNoticeNo);
                    if (!QcNoticeType.Equals("-1")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeType", @" AND a.QcNoticeType= @QcNoticeType", QcNoticeType);
                    if (!WoErpPrefix.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpPrefix", @" AND e.WoErpPrefix= @WoErpPrefix", WoErpPrefix);
                    if (!WoErpNo.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpNo", @" AND e.WoErpNo= @WoErpNo", WoErpNo);
                    if (!MtlItemNo.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND g.MtlItemNo= @MtlItemNo", MtlItemNo);
                    if (!MtlItemName.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND g.MtlItemName= @MtlItemName", MtlItemName);
                    if (!ProcessId.Equals(-1)) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND h.ProcessId= @ProcessId", ProcessId);
                    if (!StartDate.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    if (!EndDate.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcNoticeId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetQcNoticeItemSpec --取得送檢單量測項目資料 --Ding 2022-12-06
        public string GetQcNoticeItemSpec(int QcNoticeItemSpecId,int QcNoticeId,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcNoticeItemSpecId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcNoticeId,a.DesignValue,a.UpperTolerance,a.LowerTolerance,a.MakeCount,a.ProductQcCount,a.Depth,a.Remark,a.CreateDate
                            ,b.QcItemId,b.QcItemName,b.QcItemType
                            ,c.QcUomId,c.QcUomName";
                    sqlQuery.mainTables =
                        @" FROM QMS.QcNoticeItemSpec a
                            LEFT JOIN QMS.QcItem b ON a.QcItemId=b.QcItemId
                            LEFT JOIN QMS.UnitOfQcMeasure c ON a.QcUomId=c.QcUomId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeId", @" AND a.QcNoticeId= @QcNoticeId", QcNoticeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeItemSpecId", @" AND a.QcNoticeItemSpecId= @QcNoticeItemSpecId", QcNoticeItemSpecId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcNoticeId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetQcNoticeFile
        public string GetQcNoticeFile(int QcNoticeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcNoticeFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcNoticeId,a.FileId,b.[FileName],a.CreateDate";
                    sqlQuery.mainTables =
                        @" FROM QMS.QcNoticeFile a
                        INNER JOIN BAS.[File] b ON a.FileId=b.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "AND a.QcNoticeId=@QcNoticeId";
                    dynamicParameters.Add("QcNoticeId", QcNoticeId);
                    sqlQuery.conditions = queryCondition;                  
                    sqlQuery.distinct = false;
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

        #region//GetQcNoticeTemplate --取得送檢單樣板資料
        public string GetQcNoticeTemplate(int QcNoticeTemplateId, string QcNoticeTemplateNo, string QcNoticeType,int ParameterId,
                    string StartDate, string EndDate,
                    string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcNoticeTemplateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcNoticeType,d.TypeDesc,a.QcNoticeTemplateNo,a.QcNoticeTemplateName,a.QcNoticeTemplateDesc
                        ,c.ProcessName,e.UserName,a.Remark,b.ModeId,b.ParameterId
                        ,FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcNoticeTemplate a
                            LEFT JOIN MES.ProcessParameter b ON a.ParameterId=b.ParameterId
                            LEFT JOIN MES.Process c ON b.ProcessId=c.ProcessId
                            LEFT JOIN (
                                SELECT x.TypeDesc,x.TypeName
                                FROM BAS.[Type] x
                                WHERE x.TypeSchema = 'QMS' 
                            )d ON a.QcNoticeType=d.TypeName
                            LEFT JOIN BAS.[User] e ON a.CreateBy=e.UserId
                            INNER JOIN QMS.QcNoticeItemSpecTemplate f ON a.QcNoticeTemplateId=f.QcNoticeTemplateId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (!QcNoticeTemplateId.Equals(-1)) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeTemplateId", @" AND a.QcNoticeTemplateId= @QcNoticeTemplateId", QcNoticeTemplateId);
                    if (!QcNoticeTemplateNo.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeTemplateNo", @" AND a.QcNoticeTemplateNo= @QcNoticeTemplateNo", QcNoticeTemplateNo);
                    if (!QcNoticeType.Equals("-1")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeType", @" AND a.QcNoticeType= @QcNoticeType", QcNoticeType);
                    if (!ParameterId.Equals(-1)) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND a.ParameterId= @ProcessId", ParameterId);
                    if (!StartDate.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    if (!EndDate.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcNoticeTemplateId DESC";
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

        #region//GetQcNoticeItemSpecTemplate --取得樣板量測項目清單 --Ding 2022-12-07
        public string GetQcNoticeItemSpecTemplate(int QcNoticeTemplateId, int QcNoticeItemSpecTemplateId,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcNoticeItemSpecTemplateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcNoticeTemplateId,c.QcItemId,c.QcItemName,a.UpperTolerance,a.DesignValue,a.LowerTolerance
                          ,a.Depth,d.QcUomNo,a.MakeCount,a.ProductQcCount,a.CreateDate
                          ,c.QcItemType,d.QcUomId";
                    sqlQuery.mainTables =
                        @" FROM QMS.QcNoticeItemSpecTemplate a
                            INNER JOIN QMS.QcNoticeTemplate b ON a.QcNoticeTemplateId=b.QcNoticeTemplateId
                            INNER JOIN QMS.QcItem c ON a.QcItemId=c.QcItemId
                            INNER JOIN QMS.UnitOfQcMeasure d ON a.QcUomId=d.QcUomId
                            ";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeItemSpecTemplateId", @" AND a.QcNoticeItemSpecTemplateId= @QcNoticeItemSpecTemplateId", QcNoticeItemSpecTemplateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeTemplateId", @" AND a.QcNoticeTemplateId= @QcNoticeTemplateId", QcNoticeTemplateId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcNoticeTemplateId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetRoutingProcess --取得依製令對照途程製程
        public string GetRoutingProcess(int ControlId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.ControlId,c.ProcessAlias,c.RoutingProcessId
                             FROM MES.RoutingItem a 
                             INNER JOIN MES.RoutingItemProcess b ON a.RoutingItemId=b.RoutingItemId
                             INNER JOIN MES.RoutingProcess c ON b.RoutingProcessId=c.RoutingProcessId
                             WHERE a.ControlId=@ControlId
                             ORDER BY c.SortNumber ASC ";
                    dynamicParameters.Add("@ControlId", ControlId);
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

        #region//GetRdInformation -- 研發資訊 取得 --Ding 2022.10.17
        public string GetRdInformation(int ControlId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",b.DesignId,c.MtlItemId
                        ,c.MtlItemNo,c.MtlItemName,c.MtlItemSpec
                        ,b.CustomerMtlItemNo,d.CustomerDwgNo,b.CustomerCadControlId
                        ,a.Cause,a.LastModifiedDate,a.Edition,a.[Version],e.Edition AS CustEdition,e.[Version] AS CustVersion";
                    sqlQuery.mainTables =
                        @" FROM PDM.RdDesignControl a
                         LEFT JOIN PDM.RdDesign b ON a.DesignId=b.DesignId
                         LEFT JOIN PDM.MtlItem c ON b.MtlItemId=c.MtlItemId
                         LEFT JOIN PDM.CustomerCad d ON b.CustomerMtlItemNo=d.CustomerMtlItemNo
                         LEFT JOIN PDM.CustomerCadControl e ON b.CustomerCadControlId = e.ControlId
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.ControlId=@ControlId";
                    dynamicParameters.Add("ControlId", ControlId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "";
                    sqlQuery.pageIndex = -1;
                    sqlQuery.pageSize = -1;
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

        #region //GetRdDesignControl -- 取得研發設計圖版本控制資料 -- Ann 2022-06-28
        public string GetRdDesignControl(int ControlId, int DesignId, string Edition, string StartDate, string EndDate, int MtlItemId, string ReleasedStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ControlId, a.Edition, a.[Version], a.Cause
                        , FORMAT(a.DesignDate, 'yyyy-MM-dd') DesignDate, a.ReleasedStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') CreateDate
                        , (
                          SELECT xa.FileId, (xa.[FileName] + xa.FileExtension) FileInfo
                          FROM BAS.[File] xa
                          WHERE xa.FileId = a.Cad3DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad3DFile
                        , (
                          SELECT xb.FileId, (xb.[FileName] + xb.FileExtension) FileInfo
                          FROM BAS.[File] xb
                          WHERE xb.FileId = a.Cad2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad2DFile
                        , (
                          SELECT xc.FileId, (xc.[FileName] + xc.FileExtension) FileInfo
                          FROM BAS.[File] xc
                          WHERE xc.FileId = a.Pdf2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Pdf2DFile
                        , (
                          SELECT xd.FileId, (xd.[FileName] + xd.FileExtension) FileInfo
                          FROM BAS.[File] xd
                          WHERE xd.FileId = a.JmoFile
                          FOR JSON PATH, ROOT('data')
                        ) JmoFile
                        , b.UserNo, b.UserName
                        , c.DesignId, c.MtlItemId";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesignControl a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN PDM.RdDesign c ON a.DesignId = c.DesignId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId AND c.MtlItemId=@MtlItemId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "";
                    sqlQuery.pageIndex = -1;
                    sqlQuery.pageSize = -1;

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

        #region //GetManufactureOrder -- 取得MES製令資料 -- Ann 2022-07-27
        public string GetManufactureOrder(int MoId, string otherInfo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT DISTINCT a.MoId, a.WoSeq, a.Quantity, a.InputQty, a.CompleteQty, a.ScrapQty, a.[Status], a.DeliveryProcess, a.ProjectNo
                        , b.WoId, b.CompanyId, b.WoErpPrefix, b.WoErpNo, b.PlanQty, b.InventoryId, b.RequisitionSetQty, b.StockInQty
                        , c.MtlItemId, c.MtlItemNo, c.MtlItemName, c.Version, c.MtlItemSpec,c.InventoryUomId
                        , d.ModeId, ISNULL(d.ModeNo, '') ModeNo, ISNULL(d.ModeName, '') ModeName
                        , e.MoSettingId, e.LotStatus
                        , f.StatusNo, f.StatusName, f.StatusDesc
                        , i.QcNoticeId, i.QcNoticeType                        
                        , y.ControlId
                        FROM MES.ManufactureOrder a
                        INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                        INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                        LEFT JOIN MES.ProdMode d ON a.ModeId = d.ModeId
                        INNER JOIN MES.MoSetting e ON a.MoId = e.MoId
                        INNER JOIN BAS.[Status] f ON a.[Status] = f.StatusNo AND f.StatusSchema = 'ManufactureOrder.Status'
                        LEFT JOIN BAS.[Type] h ON a.DeliveryProcess = h.TypeNo AND h.TypeSchema = 'ManufactureOrder.DeliveryProcess'
                        LEFT JOIN QMS.QcNotice i ON a.MoId = i.MoId
                        INNER JOIN  MES.MoRouting x ON x.MoId=a.MoId
                        INNER JOIN MES.RoutingItem y ON x.RoutingItemId=y.RoutingItemId 
                        WHERE 1=1 ";
                    if (MoId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND a.MoId= @MoId", MoId);
                    if (!otherInfo.Equals("")) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "otherInfo", @" AND b.WoErpPrefix+'-'+b.WoErpNo+'('+CAST(a.WoSeq AS nvarchar)+')' LIKE '%' + @otherInfo + '%'", otherInfo);
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

        #region//AddQcNotice  --量測需求單單頭 新增 -- Ding 2022.12.06
        public string AddQcNotice(string QcNoticeNo, string QcNoticeType, int QcNoticeQty, int? MoId,int? ControlId,int? RoutingProcessId,
            string Remark, string Status, string FileList, int DepartmentId, string ResetFlag)
        {
            try
            {
                if (QcNoticeNo == "") throw new SystemException("【送檢單編號】不能為空!");
                if (QcNoticeQty <= 0) throw new SystemException("【送檢數量】不能等於0");                
                if (ResetFlag.Length <= 0) throw new SystemException("【過站後是否結案】不能小於等於0");
                if (Status != "Y" && Status != "N" && Status != "A") throw new SystemException("【Status】編號錯誤");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【單據類型】編號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TypeName 
                                FROM BAS.[Type] a
                                WHERE a.TypeSchema = 'QMS' 
                                AND a.TypeName =@QcNoticeType";
                        dynamicParameters.Add("QcNoticeType", QcNoticeType);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【單據類型】不存在!");
                        #endregion                       
                        foreach (var item in result)
                        {
                            QcNoticeType = item.TypeName;
                        }

                        switch (QcNoticeType) {
                            case "OQC":
                                if (ControlId <= 0) throw new SystemException("【圖面版本ControlId】不能小於等於0");

                                #region //判斷【ControlId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ControlId 
                                    FROM PDM.RdDesignControl a                              
                                    WHERE a.ControlId =@ControlId";
                                dynamicParameters.Add("ControlId", ControlId);
                                var OqcControlIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (OqcControlIdResult.Count() <= 0) throw new SystemException("該【ControlId】不存在");
                                #endregion

                                #region //判斷【QcNotic】是否能新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId,a.Status,a.QcNoticeNo
                                    FROM QMS.QcNotice a                              
                                    WHERE a.ControlId =@ControlId
                                    AND a.QcNoticeType='OQC'
                                    AND a.Status='N'
                                ";
                                dynamicParameters.Add("ControlId", ControlId);
                                var OqcResult = sqlConnection.Query(sql, dynamicParameters);
                                string BeforeQcNoticeNo = "";
                                foreach (var item in OqcResult)
                                {
                                    BeforeQcNoticeNo = item.QcNoticeNo;
                                }
                                if (OqcResult.Count() > 0) throw new SystemException("該【" + BeforeQcNoticeNo + "】尚未結束檢測作業，不可進行【出貨檢】");
                                #endregion

                                RoutingProcessId = null;
                                MoId = null;

                                break;
                            case "TQC":
                                if (MoId <= 0) throw new SystemException("【MoId】不能小於等於0");

                                #region //判斷【MoId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId 
                                FROM MES.ManufactureOrder a                              
                                WHERE a.MoId =@MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var TqcMoIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (TqcMoIdResult.Count() <= 0) throw new SystemException("該【MoId】不存在");
                                #endregion


                                #region //判斷【ControlId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ControlId 
                                    FROM PDM.RdDesignControl a                              
                                    WHERE a.ControlId =@ControlId";
                                dynamicParameters.Add("ControlId", ControlId);
                                var TqcControlIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (TqcControlIdResult.Count() <= 0) throw new SystemException("該【ControlId】不存在");
                                #endregion

                                #region //判斷【QcNotic】是否能新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId,a.Status,a.QcNoticeNo
                                    FROM QMS.QcNotice a                              
                                    WHERE a.MoId =@MoId
                                    AND a.QcNoticeType='TQC'
                                    AND a.Status='N'
                                ";
                                dynamicParameters.Add("MoId", MoId);
                                var TqcResult = sqlConnection.Query(sql, dynamicParameters);
                                BeforeQcNoticeNo = "";
                                foreach (var item in TqcResult)
                                {
                                    BeforeQcNoticeNo = item.QcNoticeNo;
                                }
                                if (TqcResult.Count() > 0) throw new SystemException("該【" + BeforeQcNoticeNo + "】尚未結束檢測作業，不可進行【全吋檢】");
                                #endregion

                                ControlId = null;
                                RoutingProcessId = null;

                                break;
                            case "IPQC":
                                if (RoutingProcessId <= 0) throw new SystemException("【RoutingProcessId】不能小於等於0");
                                if (ControlId <= 0) throw new SystemException("【ControlId】不能小於等於0");

                                #region //判斷【ControlId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ControlId 
                                    FROM PDM.RdDesignControl a                              
                                    WHERE a.ControlId =@ControlId";
                                dynamicParameters.Add("ControlId", ControlId);
                                var IPQCControlIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (IPQCControlIdResult.Count() <= 0) throw new SystemException("該【ControlId】不存在");
                                #endregion

                                #region //判斷【製令製程ID】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.RoutingProcessId
                                FROM MES.RoutingProcess a                                
                                WHERE a.RoutingProcessId =@RoutingProcessId";
                                dynamicParameters.Add("RoutingProcessId", RoutingProcessId);
                                var IpqcRoutingProcessIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (IpqcRoutingProcessIdResult.Count() <= 0) throw new SystemException("該【RoutingProcessId】不存在");
                                #endregion

                                #region //判斷【QcNotic】是否能新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId,a.Status,a.QcNoticeNo
                                    FROM QMS.QcNotice a                              
                                    WHERE a.ControlId =@ControlId                               
                                    AND a.RoutingProcessId=@RoutingProcessId
                                    AND a.QcNoticeType='IPQC'
                                    AND a.Status='N'
                                ";
                                dynamicParameters.Add("ControlId", ControlId);
                                dynamicParameters.Add("RoutingProcessId", RoutingProcessId);
                                var IpqcResult = sqlConnection.Query(sql, dynamicParameters);
                                string BeforQcNoticeNo = "";
                                foreach (var item in IpqcResult)
                                {
                                    BeforQcNoticeNo = item.QcNoticeNo;
                                }
                                if (IpqcResult.Count() > 0) throw new SystemException("該【" + BeforQcNoticeNo + "】尚未結束檢測作業，不可進行【工程檢】");
                                #endregion                             

                                break;
                            case "PVTQC":
                                if (MoId <= 0) throw new SystemException("【MoId】不能小於等於0");

                                #region //判斷【MoId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId 
                                FROM MES.ManufactureOrder a
                                WHERE a.MoId =@MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var PvtqcMoIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (PvtqcMoIdResult.Count() <= 0) throw new SystemException("該【MoId】不存在");
                                #endregion

                                #region //判斷【ControlId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ControlId 
                                    FROM PDM.RdDesignControl a                              
                                    WHERE a.ControlId =@ControlId";
                                dynamicParameters.Add("ControlId", ControlId);
                                var PvtqcControlIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (PvtqcControlIdResult.Count() <= 0) throw new SystemException("該【ControlId】不存在");
                                #endregion

                                #region //判斷【QcNotic】是否能新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId,a.Status,a.QcNoticeNo
                                    FROM QMS.QcNotice a                              
                                    WHERE a.MoId =@MoId
                                    AND a.QcNoticeType='PVTQC'
                                    AND a.Status='N'
                                ";
                                dynamicParameters.Add("MoId", MoId);
                                var PvtqcResult = sqlConnection.Query(sql, dynamicParameters);
                                BeforeQcNoticeNo = "";
                                foreach (var item in PvtqcResult)
                                {
                                    BeforeQcNoticeNo = item.QcNoticeNo;
                                }
                                if (PvtqcResult.Count() > 0) throw new SystemException("該【" + BeforeQcNoticeNo + "】尚未結束檢測作業，不可進行【試樣檢】");
                                #endregion

                                ControlId = null;
                                RoutingProcessId = null;

                                break;
                            default://以上都不符合走這個
                                throw new SystemException("【單據類型】編號錯誤!");                                
                        }

                        #region //判斷部門資料是否錯誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department                         
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        var departmentResult = sqlConnection.Query(sql, dynamicParameters);
                        if (departmentResult.Count() <= 0) throw new SystemException("該【部門】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcNotice (QcNoticeNo,QcNoticeType,QcNoticeQty,DepartmentId,ControlId,RoutingProcessId,MoId,Remark,ResetFlag
                                ,[Status],CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.QcNoticeId
                                VALUES (@QcNoticeNo,@QcNoticeType,@QcNoticeQty,@DepartmentId,@ControlId,@RoutingProcessId,@MoId,@Remark,@ResetFlag
                                ,@Status,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeNo,
                                QcNoticeType,
                                QcNoticeQty,
                                DepartmentId,
                                ControlId,
                                RoutingProcessId,
                                MoId,                                
                                Remark,
                                ResetFlag,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int QcNoticeId = -1;
                        foreach (var item in insertResult)
                        {
                            QcNoticeId = item.QcNoticeId;
                        }
                        if (FileList != "")
                        {
                            string[] FileIdItem = FileList.Split(',');
                            foreach (var id in FileIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a                              
                                WHERE a.FileId =@FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var result4 = sqlConnection.Query(sql, dynamicParameters);
                                if (result4.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcNoticeFile (QcNoticeId, FileId
                                        ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                        OUTPUT INSERTED.QcNoticeId
                                        VALUES (@QcNoticeId, @FileId
                                        ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcNoticeId,
                                        FileId,
                                        Status,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
                        }
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

        #region//AddQcNoticeItemSpec --製程量測需求單-量測項目清單 新增 -Ding 2022.12.06
        public string AddQcNoticeItemSpec(int QcNoticeId, int QcItemId, double DesignValue, double UpperTolerance, double LowerTolerance, int MakeCount, int ProductQcCount,
            int QcUomId, int Depth, string Remark)
        {
            try
            {
                if (QcNoticeId <0) throw new SystemException("【QcNoticeId】不能為空!");
                if (QcItemId <= 0) throw new SystemException("【工程檢量測參數基本資料】【量測項目】不能為空!");
                if (DesignValue <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值】不能為空!");
                if (UpperTolerance <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值上限公差】不能為空!");
                if (LowerTolerance <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值下限公差】不能為空!");
                if (MakeCount <= 0) throw new SystemException("【同項目量測次數】【量測次數】不能為0!");
                if (ProductQcCount <= 0) throw new SystemException("【同產品量測次數】不能為0!");
                if (QcUomId <= -2) throw new SystemException("【工程檢量測參數基本資料】【量測單位】不能為空!");               

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【QcNoticeId】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcNoticeId 
                                FROM QMS.QcNotice a                              
                                WHERE a.QcNoticeId =@QcNoticeId";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() <= 0) throw new SystemException("該【QcNoticeId】不存在");
                        #endregion

                        #region //判斷【量測項目ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcItem a                              
                                WHERE a.QcItemId =@QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【量測項目ID】不存在");
                        #endregion

                        #region //判斷【製程量測項目參數】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcNoticeItemSpec a                              
                                WHERE a.QcItemId =@QcItemId
                                AND a.QcNoticeId =@QcNoticeId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("該【量測項目】已存在，不可重複");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO  QMS.QcNoticeItemSpec (QcNoticeId, QcItemId, DesignValue,UpperTolerance,LowerTolerance
                                ,MakeCount,ProductQcCount, QcUomId,Depth,Remark
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@QcNoticeId, @QcItemId,@DesignValue, @UpperTolerance, @LowerTolerance
                                ,@MakeCount,@ProductQcCount,@QcUomId,@Depth,@Remark
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeId,
                                QcItemId,
                                DesignValue,
                                UpperTolerance,
                                LowerTolerance,
                                MakeCount,
                                ProductQcCount,
                                QcUomId,
                                Depth = Depth > 0 ? Depth : (int?)null,
                                Remark,
                                Status = "A",
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

        #region//AddQcNoticeTemplate --量測需求單樣板單頭 新增 -- Ding 2022.12.15
        public string AddQcNoticeTemplate(string QcNoticeTemplateNo,string QcNoticeTemplateName,string QcNoticeTemplateDesc, string QcNoticeType , int ParameterId , string Remark )
        {
            try
            {
                if (QcNoticeTemplateNo == "") throw new SystemException("【送檢單樣板編號】不能為空!");
                if (QcNoticeTemplateName == "") throw new SystemException("【送檢單樣板名稱】不能為空!");
                if (QcNoticeTemplateDesc == "") throw new SystemException("【送檢單樣板敘述】不能為空!");
                if (QcNoticeType != "OQC" && QcNoticeType != "IQC" && QcNoticeType != "IPQC" && QcNoticeType != "PVTQC" && QcNoticeType != "TQC") throw new SystemException("【單據樣板類型】編號錯誤!");
                if (QcNoticeType == "IPQC")
                {
                    if (ParameterId <= 0) throw new SystemException("【ParameterId】不能小於等於0");
                }
                
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters(); 

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcNoticeTemplate (QcNoticeTemplateNo,QcNoticeTemplateName,QcNoticeTemplateDesc,QcNoticeType,ParameterId,Remark,CompanyId
                               ,CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.QcNoticeTemplateId
                                VALUES (@QcNoticeTemplateNo, @QcNoticeTemplateName, @QcNoticeTemplateDesc,@QcNoticeType,@ParameterId,@Remark,@CompanyId
                                ,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeTemplateNo,
                                QcNoticeTemplateName,
                                QcNoticeTemplateDesc,
                                QcNoticeType,
                                ParameterId,
                                Remark,
                                CompanyId= CurrentCompany,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int QcNoticeTemplateId = -1;
                        foreach (var item in insertResult)
                        {
                            QcNoticeTemplateId = item.QcNoticeTemplateId;
                        }

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

        #region//AddQcNoticeItemSpecTemplate --製程量測需求單樣板-量測項目清單 新增 -Ding 2022.12.06
        public string AddQcNoticeItemSpecTemplate(int QcNoticeTemplateId, int QcItemId, double DesignValue, double UpperTolerance , double LowerTolerance , int MakeCount, int ProductQcCount ,
            int QcUomId , int Depth , string Remark )
        {
            try
            {
                if (QcNoticeTemplateId < 0) throw new SystemException("【QcNoticeTemplateId】不能為空!");
                if (QcItemId <= 0) throw new SystemException("【工程檢量測參數基本資料】【量測項目】不能為空!");
                if (DesignValue <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值】不能為空!");
                if (UpperTolerance <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值上限公差】不能為空!");
                if (LowerTolerance <= -2) throw new SystemException("【工程檢量測參數基本資料】【尺寸設計值下限公差】不能為空!");
                if (MakeCount <= 0) throw new SystemException("【同項目量測次數】【量測次數】不能為0!");
                if (ProductQcCount <= 0) throw new SystemException("【同產品量測次數】不能為0!");
                if (QcUomId <= -2) throw new SystemException("【工程檢量測參數基本資料】【量測單位】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【QcNoticeId】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcNoticeTemplateId 
                                FROM QMS.QcNoticeTemplate a                              
                                WHERE a.QcNoticeTemplateId =@QcNoticeTemplateId";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() <= 0) throw new SystemException("該【QcNoticeTemplateId】不存在");
                        #endregion

                        #region //判斷【量測項目ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcItem a                              
                                WHERE a.QcItemId =@QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【量測項目ID】不存在");
                        #endregion

                        #region //判斷【製程量測項目參數】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcNoticeItemSpecTemplate a                              
                                WHERE a.QcItemId =@QcItemId
                                AND a.QcNoticeTemplateId =@QcNoticeTemplateId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("該【量測項目】已存在，不可重複");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO  QMS.QcNoticeItemSpecTemplate (QcNoticeTemplateId, QcItemId, DesignValue,UpperTolerance,LowerTolerance
                                ,MakeCount,ProductQcCount, QcUomId,Depth,Remark
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@QcNoticeTemplateId, @QcItemId,@DesignValue, @UpperTolerance, @LowerTolerance
                                ,@MakeCount,@ProductQcCount,@QcUomId,@Depth,@Remark
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeTemplateId,
                                QcItemId,
                                DesignValue,
                                UpperTolerance,
                                LowerTolerance,
                                MakeCount,
                                ProductQcCount,
                                QcUomId,
                                Depth,
                                Remark,
                                Status = "A",
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

        #region//AddTemplateToQcNoticeItemSpec --樣板匯入
        public string AddTemplateToQcNoticeItemSpec(int QcNoticeId = -1, int QcNoticeTemplateId = -1)
        {
            try
            {
                if (QcNoticeId < 0) throw new SystemException("【QcNoticeId】不能為空!");
                if (QcNoticeTemplateId <= 0) throw new SystemException("【QcNoticeTemplateId】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【QcNoticeTemplateId】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.*
                                 FROM QMS.QcNoticeTemplate a
                                 INNER JOIN QMS.QcNoticeItemSpecTemplate b ON a.QcNoticeTemplateId=b.QcNoticeTemplateId                      
                                WHERE a.QcNoticeTemplateId =@QcNoticeTemplateId";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() <= 0) throw new SystemException("該【QcNoticeTemplateId】不存在");
                        #endregion

                        #region //判斷【QcNoticeId】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcNoticeId 
                                FROM QMS.QcNotice a                              
                                WHERE a.QcNoticeId =@QcNoticeId";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【QcNoticeId】不存在");
                        #endregion

                        foreach (var item in result1)
                        {
                            int QcItemId = item.QcItemId;
                            double DesignValue = item.DesignValue;
                            double UpperTolerance = item.UpperTolerance;
                            double LowerTolerance = item.LowerTolerance;
                            int MakeCount = item.MakeCount;
                            int ProductQcCount = item.ProductQcCount;
                            int QcUomId = item.QcUomId;
                            int Depth = item.Depth;

                            #region //判斷【量測項目ID】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                FROM QMS.QcItem a                              
                                WHERE a.QcItemId =@QcItemId";
                            dynamicParameters.Add("QcItemId", QcItemId);
                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("該【量測項目ID】不存在");
                            #endregion

                            #region //判斷【製程量測項目參數】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                FROM QMS.QcNoticeItemSpec a                              
                                WHERE a.QcItemId =@QcItemId
                                AND a.QcNoticeId =@QcNoticeId";
                            dynamicParameters.Add("QcItemId", QcItemId);
                            dynamicParameters.Add("QcNoticeId", QcNoticeId);
                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() > 0) throw new SystemException("該【量測項目】已存在，不可重複");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO  QMS.QcNoticeItemSpec (QcNoticeId, QcItemId, DesignValue,UpperTolerance,LowerTolerance
                                ,MakeCount,ProductQcCount, QcUomId,Depth,Remark
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@QcNoticeId, @QcItemId,@DesignValue, @UpperTolerance, @LowerTolerance
                                ,@MakeCount,@ProductQcCount,@QcUomId,@Depth,@Remark
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId,
                                    QcItemId,
                                    DesignValue,
                                    UpperTolerance,
                                    LowerTolerance,
                                    MakeCount,
                                    ProductQcCount,
                                    QcUomId,
                                    Depth,
                                    Remark="",
                                    Status = "A",
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
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//UpdateQcNotice --量測需求單單頭 修改 -- Ding 2022.12.07
        public string UpdateQcNotice(int QcNoticeId,string QcNoticeNo,string QcNoticeType,int QcNoticeQty,int MoId,int ControlId,int RoutingProcessId,
            int MoProcessId,string Remark,string Status,string FileList,int DepartmentId,string ResetFlag)
        {
            try
            {
                if (QcNoticeId <= 0) throw new SystemException("【送檢單ID】不可-1!");
                if (QcNoticeNo == "") throw new SystemException("【送檢單編號】不能為空!");
                if (QcNoticeQty <= 0) throw new SystemException("【送檢數量】不能等於0");               
                if (ResetFlag.Length <= 0) throw new SystemException("【過站後是否結案】不能小於等於0");
                if (Status != "Y" && Status != "N" && Status != "A") throw new SystemException("【Status】編號錯誤");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNotice
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【送檢單】資料錯誤!");
                        #endregion

                        #region //判斷【單據類型】編號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TypeName 
                                FROM BAS.[Type] a
                                WHERE a.TypeSchema = 'QMS' 
                                AND a.TypeName =@QcNoticeType";
                        dynamicParameters.Add("QcNoticeType", QcNoticeType);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() <= 0) throw new SystemException("該【單據類型】不存在!");
                        foreach (var item in result1)
                        {
                            QcNoticeType = item.TypeName;
                        }
                        #endregion

                        switch (QcNoticeType)
                        {
                            case "OQC":
                                if (ControlId <= 0) throw new SystemException("【圖面版本ControlId】不能小於等於0");

                                #region //判斷【ControlId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ControlId 
                                FROM PDM.RdDesignControl a                              
                                WHERE a.ControlId =@ControlId";
                                dynamicParameters.Add("ControlId", ControlId);
                                var OqcControlIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (OqcControlIdResult.Count() <= 0) throw new SystemException("該【ControlId】不存在");
                                #endregion

                                RoutingProcessId = -1;
                                MoId = -1;
                                break;
                            case "TQC":
                                if (MoId <= 0) throw new SystemException("【MoId】不能小於等於0");

                                #region //判斷【MoId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId 
                                FROM MES.ManufactureOrder a                              
                                WHERE a.MoId =@MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var TqcMoIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (TqcMoIdResult.Count() <= 0) throw new SystemException("該【MoId】不存在");
                                #endregion

                                
                                RoutingProcessId = -1;  

                                break;
                            case "IPQC":
                                if (RoutingProcessId <= 0) throw new SystemException("【RoutingProcessId】不能小於等於0");
                                if (MoId <= 0) throw new SystemException("【MoId】不能小於等於0");

                                #region //判斷【MoId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId 
                                FROM MES.ManufactureOrder a                              
                                WHERE a.MoId =@MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var IpqcMoIdResult1 = sqlConnection.Query(sql, dynamicParameters);
                                if (IpqcMoIdResult1.Count() <= 0) throw new SystemException("該【MoId】不存在");
                                #endregion

                                #region //判斷【製令製程ID】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.RoutingProcessId
                                FROM MES.RoutingProcess a                                
                                WHERE a.RoutingProcessId =@RoutingProcessId";
                                dynamicParameters.Add("RoutingProcessId", RoutingProcessId);
                                var IpqcRoutingProcessIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (IpqcRoutingProcessIdResult.Count() <= 0) throw new SystemException("該【RoutingProcessId】不存在");
                                #endregion

                                ControlId = -1;
                                break;
                            case "PVTQC":
                                #region //判斷【MoId】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId 
                                FROM MES.ManufactureOrder a                              
                                WHERE a.MoId =@MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var PvtqcMoIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (PvtqcMoIdResult.Count() <= 0) throw new SystemException("該【MoId】不存在");
                                #endregion

                                
                                RoutingProcessId = -1;

                                break;
                            default://以上都不符合走這個
                                throw new SystemException("【單據類型】編號錯誤!");
                        }                       

                        #region //判斷部門資料是否錯誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department                        
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        var departmentResult = sqlConnection.Query(sql, dynamicParameters);
                        if (departmentResult.Count() <= 0) throw new SystemException("該【部門】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcNotice SET
                                QcNoticeNo = @QcNoticeNo,
                                QcNoticeType = @QcNoticeType,
                                QcNoticeQty = @QcNoticeQty,
                                DepartmentId = @DepartmentId,
                                ControlId = @ControlId,
                                RoutingProcessId = @RoutingProcessId,
                                MoId = @MoId,                                
                                Remark = @Remark,                               
                                ResetFlag = @ResetFlag,                               
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcNoticeId = @QcNoticeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeNo,
                                QcNoticeType,
                                QcNoticeQty,
                                DepartmentId,
                                ControlId,
                                RoutingProcessId,
                                MoId,                                
                                Remark,
                                ResetFlag,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcNoticeId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if(rowsAffected!=1) throw new SystemException("該【送檢單】修改失敗");

                        if (FileList!="") {
                            string[] FileIdItem = FileList.Split(',');
                            foreach (var id in FileIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a                              
                                WHERE a.FileId =@FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var result4 = sqlConnection.Query(sql, dynamicParameters);
                                if (result4.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcNoticeFile (QcNoticeId, FileId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.QcNoticeId
                                VALUES (@QcNoticeId, @FileId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcNoticeId,
                                        FileId,
                                        Status,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);
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

        #region//UpdateQcNoticeItemSpec --出貨檢量測參數基本資料 修改 --Ding 2022.10.17
        public string UpdateQcNoticeItemSpec(int QcNoticeItemSpecId,int QcNoticeId, int QcItemId, double DesignValue, double UpperTolerance, double LowerTolerance
            , int MakeCount,int ProductQcCount, int QcUomId, int Depth, string Remark)
        {
            try
            {
                if (QcNoticeItemSpecId <= 0) throw new SystemException("【送檢單量測參數基本資料】查無資料!");
                if (MakeCount <= 0) throw new SystemException("【同項目量測次數】不可為0!");
                if (ProductQcCount <= 0) throw new SystemException("【同產品量測次數】不可為0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷【送檢單量測參數基本資料】是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpec
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeItemSpecId", QcNoticeItemSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【送檢單量測參數基本資料】資料錯誤!");
                        #endregion

                        #region //判斷【量測項目】是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE 1=1";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【量測項目】資料錯誤!");
                        #endregion

                        #region //判斷【製程量測項目參數】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcNoticeItemSpec a                              
                                WHERE a.QcItemId =@QcItemId
                                AND a.QcNoticeId =@QcNoticeId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 1) throw new SystemException("該【量測項目】已存在，不可重複");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcNoticeItemSpec SET
                                
                                QcItemId = @QcItemId,
                                DesignValue = @DesignValue,
                                UpperTolerance = @UpperTolerance,
                                LowerTolerance = @LowerTolerance,
                                MakeCount = @MakeCount,
                                QcUomId = @QcUomId,
                                Depth=@Depth,                                
                                Remark = @Remark,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcNoticeItemSpecId = @QcNoticeItemSpecId
                                AND QcNoticeId = @QcNoticeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {                                
                                QcItemId,
                                DesignValue,
                                UpperTolerance,
                                LowerTolerance,
                                MakeCount,
                                QcUomId,
                                Depth,                                
                                Remark,
                                Status = "A",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcNoticeId,
                                QcNoticeItemSpecId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("該【出貨檢量測參數基本資料】修改失敗");

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

        #region//UpdateQcNoticeTemplate --量測需求單單頭 修改 -- Ding 2022.12.07
        public string UpdateQcNoticeTemplate(int QcNoticeTemplateId,string QcNoticeTemplateNo,string QcNoticeTemplateName,string QcNoticeTemplateDesc,string QcNoticeType,int ParameterId,
            string Remark)
        {
            try
            {
                if (QcNoticeTemplateId <= 0) throw new SystemException("【送檢單樣板ID】不可-1!");
                if (QcNoticeTemplateNo == "") throw new SystemException("【送檢單樣板編號】不能為空!");
                if (QcNoticeTemplateName == "") throw new SystemException("【送檢單樣板名稱】不能為空!");
                if (QcNoticeTemplateDesc == "") throw new SystemException("【送檢單樣板描述】不能為空!");
                if (QcNoticeType != "OQC" && QcNoticeType != "IQC" && QcNoticeType != "IPQC" && QcNoticeType != "PVTQC") throw new SystemException("【單據類型】編號錯誤!");
                
                if (QcNoticeType == "IPQC")
                {
                    if (ParameterId <= 0) throw new SystemException("【ParameterId】尚未選擇站別ˊ");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeTemplate
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【送檢單樣板】資料錯誤!");
                        #endregion

                        if (QcNoticeType == "IPQC")
                        {
                            #region //判斷【製令製程ID】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ParameterId,b.ProcessName
                                FROM MES.ProcessParameter a
                                INNER JOIN MES.Process b ON a.ProcessId=b.ProcessId                                
                                WHERE a.ParameterId =@ParameterId";
                            dynamicParameters.Add("ParameterId", ParameterId);
                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() != 1) throw new SystemException("該【製程ID】不存在");
                            
                            #endregion
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcNoticeTemplate SET
                                QcNoticeTemplateNo = @QcNoticeTemplateNo,
                                QcNoticeTemplateName = @QcNoticeTemplateName,
                                QcNoticeTemplateDesc = @QcNoticeTemplateDesc,
                                QcNoticeType = @QcNoticeType,                                
                                ParameterId = @ParameterId,
                                Remark = @Remark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcNoticeTemplateId = @QcNoticeTemplateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeTemplateNo,
                                QcNoticeTemplateName,
                                QcNoticeTemplateDesc,
                                QcNoticeType,
                                ParameterId,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcNoticeTemplateId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("該【送檢單樣板】修改失敗");                        

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

        #region//UpdateQcNoticeItemSpecTemplate --出貨檢量測參數基本資料 修改 --Ding 2022.10.17
        public string UpdateQcNoticeItemSpecTemplate(int QcNoticeItemSpecTemplateId, int QcNoticeTemplateId, int QcItemId, double DesignValue, double UpperTolerance, double LowerTolerance
            , int MakeCount, int ProductQcCount, int QcUomId, int Depth, string Remark)
        {
            try
            {
                if (QcNoticeItemSpecTemplateId <= 0) throw new SystemException("【送檢單-量測項目清單-樣板】查無資料!");
                if (QcNoticeTemplateId <= 0) throw new SystemException("【送檢單-樣板】查無資料!");
                if (MakeCount <= 0) throw new SystemException("【同項目量測次數】不可為0!");
                if (ProductQcCount <= 0) throw new SystemException("【同產品量測次數】不可為0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeTemplate
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【送檢單樣板】資料錯誤!");
                        #endregion

                        #region //判斷【送檢單-量測項目清單-樣板】是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpecTemplate
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeItemSpecTemplateId", QcNoticeItemSpecTemplateId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() != 1) throw new SystemException("【送檢單-量測項目清單-樣板】資料錯誤!");
                        #endregion

                        #region //判斷【量測項目】是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE 1=1";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() != 1) throw new SystemException("【量測項目】資料錯誤!");
                        #endregion

                        #region //判斷【製程量測項目參數】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcItemId
                                FROM QMS.QcNoticeItemSpecTemplate a                              
                                WHERE a.QcItemId =@QcItemId
                                AND a.QcNoticeTemplateId =@QcNoticeTemplateId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() != 1) throw new SystemException("該【量測項目】已存在，不可重複");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcNoticeItemSpecTemplate SET                                
                                QcItemId = @QcItemId,
                                DesignValue = @DesignValue,
                                UpperTolerance = @UpperTolerance,
                                LowerTolerance = @LowerTolerance,
                                MakeCount = @MakeCount,
                                QcUomId = @QcUomId,
                                Depth=@Depth,                                
                                Remark = @Remark,                                
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcNoticeItemSpecTemplateId = @QcNoticeItemSpecTemplateId
                                AND QcNoticeTemplateId = @QcNoticeTemplateId                            
                                ";
                        dynamicParameters.AddDynamicParams(
                            new
                            {                                
                                QcItemId,
                                DesignValue,
                                UpperTolerance,
                                LowerTolerance,
                                MakeCount,
                                QcUomId,
                                Depth,
                                Remark,                                
                                LastModifiedDate,
                                LastModifiedBy,
                                QcNoticeTemplateId,
                                QcNoticeItemSpecTemplateId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("該【送檢單-量測項目清單-樣板】修改失敗");

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

        #region//Delete

        #region//DeleteQcNotice --送檢單資料 刪除 --Ding 2022.10.17
        public string DeleteQcNotice(int QcNoticeId)
        {
            try
            {
                if (QcNoticeId <= 0) throw new SystemException("【送檢單資料】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//QMS.QcMeasureData;QMS.TempQcMeasureData是否有資料
                        #endregion


                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNotice
                                WHERE QcNoticeId=@QcNoticeId";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【送檢單資料】資料錯誤!");
                        #endregion

                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpec
                                WHERE QcNoticeId=@QcNoticeId";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {

                            #region//單身刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM QMS.QcNoticeItemSpec 
                                WHERE QcNoticeId = @QcNoticeId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //判斷QcNoticeFile資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FileId
                                FROM QMS.QcNoticeFile
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeId", QcNoticeId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0)
                        {
                            #region//刪除QMS.QcNoticeFile
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE QMS.QcNoticeFile 
                                WHERE QcNoticeId = @QcNoticeId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);                            
                            #endregion

                            #region//修改 BAS.File狀態                        
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.[File] SET
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId IN (
                                 SELECT FileId
                                 FROM QMS.QcNoticeFile
                                 WHERE QcNoticeId=@QcNoticeId
                             )";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DeleteStatus = "Y",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcNoticeId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);                            
                            #endregion

                        }
                        #endregion

                        #region//單頭刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcNotice 
                                WHERE QcNoticeId = @QcNoticeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeId
                            });
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

        #region//DeleteQcNoticeItemSpec --製程量測需求單-量測項目清單 刪除 --Ding 2022.10.17
        public string DeleteQcNoticeItemSpec(int QcNoticeItemSpecId)
        {
            try
            {
                if (QcNoticeItemSpecId <= 0) throw new SystemException("【送檢單資料量測項目】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpec
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeItemSpecId", QcNoticeItemSpecId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【送檢單資料量測項目】資料錯誤!");
                        #endregion

                        #region//單身刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcNoticeItemSpec 
                                WHERE QcNoticeItemSpecId = @QcNoticeItemSpecId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeItemSpecId
                            });
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

        #region//DeleteQcNoticeFile --送檢單附件 刪除 --Ding 2022.12.8 
        public string DeleteQcNoticeFile(int QcNoticeFileId)
        {
            try
            {
                if (QcNoticeFileId <= 0) throw new SystemException("【送檢單資料附件】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FileId
                                FROM QMS.QcNoticeFile
                                WHERE  QcNoticeFileId=@QcNoticeFileId";
                        dynamicParameters.Add("QcNoticeFileId", QcNoticeFileId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【送檢單資料附件】資料錯誤!");
                        int FileId = -1;
                        foreach (var item in result)
                        {
                            FileId = item.FileId;
                        }
                        #endregion

                        #region//刪除QMS.QcNoticeFile
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcNoticeFile 
                                WHERE QcNoticeFileId = @QcNoticeFileId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeFileId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【QMS.QcNoticeFile 】刪除失敗!");
                        #endregion

                        #region//修改 BAS.File狀態
                        rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[File] SET
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DeleteStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                FileId
                            });
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【BAS.[File]】DeleteStatus修改失敗!");
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

        #region//DeleteQcNoticeTemplate --送檢單資料樣板 刪除 --Ding 2022.10.17
        public string DeleteQcNoticeTemplate(int QcNoticeTemplateId)
        {
            try
            {
                if (QcNoticeTemplateId <= 0) throw new SystemException("【送檢單資料樣板】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeTemplate
                                WHERE QcNoticeTemplateId=@QcNoticeTemplateId";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【送檢單資料樣板】資料錯誤!");
                        #endregion

                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpecTemplate
                                WHERE QcNoticeTemplateId=@QcNoticeTemplateId";
                        dynamicParameters.Add("QcNoticeTemplateId", QcNoticeTemplateId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            #region//單身刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM QMS.QcNoticeItemSpecTemplate 
                                WHERE QcNoticeTemplateId = @QcNoticeTemplateId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeTemplateId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region//單頭刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcNoticeTemplate 
                                WHERE QcNoticeTemplateId = @QcNoticeTemplateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeTemplateId
                            });
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

        #region//DeleteQcNoticeItemSpecTemplate --製程量測需求單-量測項目清單-樣板 刪除 --Ding 2022.10.17
        public string DeleteQcNoticeItemSpecTemplate(int QcNoticeItemSpecTemplateId)
        {
            try
            {
                if (QcNoticeItemSpecTemplateId <= 0) throw new SystemException("【送檢單資料量測項目樣板】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcNoticeItemSpecTemplate
                                WHERE 1=1";
                        dynamicParameters.Add("QcNoticeItemSpecTemplateId", QcNoticeItemSpecTemplateId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【送檢單資料量測項目樣板】資料錯誤!");
                        #endregion

                        #region//單身刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcNoticeItemSpecTemplate 
                                WHERE QcNoticeItemSpecTemplateId = @QcNoticeItemSpecTemplateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcNoticeItemSpecTemplateId
                            });
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


        #endregion
    }
}
