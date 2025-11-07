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
    public class DemandDA
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

        public DemandDA()
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
        #region //GetDemand -- 取得需求資料 -- Ben Ma 2023.07.25
        public string GetDemand(int DemandId, int SourceId, string DemandNo, int DemandDepartment
            , int DemandCustomer, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DemandId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SourceId, a.DemandNo, a.DemandDesc
                        , FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate
                        , ISNULL(FORMAT(a.DemandDeadline, 'yyyy-MM-dd'), '') DemandDeadline
                        , ISNULL(a.DemandDepartment, -1) DemandDepartment, ISNULL(a.DemandCustomer, -1) DemandCustomer
                        , a.DemandUser, a.StartDate, a.EndDate, a.DemandStatus
                        , CASE WHEN a.DemandDepartment IS NOT NULL THEN 'department' ELSE 'customer' END DepCus
                        , b.SourceNo, b.SourceName
                        , ISNULL(c.DepartmentNo, '') DepartmentNo, ISNULL(c.DepartmentName, '') DepartmentName
                        , ISNULL(d.CustomerNo, '') CustomerNo, ISNULL(d.CustomerShortName, '') CustomerShortName
                        , e.UserNo DemandUserNo, e.UserName DemandUserName
                        , f.StatusName DemandStatusName";
                    sqlQuery.mainTables =
                        @"FROM SSO.Demand a
                        INNER JOIN SSO.DemandSource b ON a.SourceId = b.SourceId
                        LEFT JOIN BAS.Department c ON a.DemandDepartment = c.DepartmentId
                        LEFT JOIN SCM.Customer d ON a.DemandCustomer = d.CustomerId
                        INNER JOIN BAS.[User] e ON a.DemandUser = e.UserId
                        INNER JOIN BAS.[Status] f ON a.DemandStatus = f.StatusNo AND f.StatusSchema = 'Demand.DemandStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DemandId", @" AND a.DemandId = @DemandId", DemandId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SourceId", @" AND a.SourceId = @SourceId", SourceId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DemandNo", @" AND a.DemandNo LIKE '%' + @DemandNo + '%'", DemandNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DemandDepartment", @" AND a.DemandDepartment = @DemandDepartment", DemandDepartment);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DemandCustomer", @" AND a.DemandCustomer = @DemandCustomer", DemandCustomer);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DemandDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DemandDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.DemandStatus IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DemandId DESC";
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

        #region //GetDemandCertificate -- 取得需求單憑證資料 -- Yi 2023.07.26
        public string GetDemandCertificate(int CertificateId, int FileId, int DemandId, string CertificateDesc
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CertificateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DemandId, a.CertificateDesc, a.FileId
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , b.FileName, b.FileContent, b.FileExtension, b.FileSize
                        , (c.UserNo + ' ' + c.UserName) UserWithNo
                        , d.DemandStatus";
                    sqlQuery.mainTables =
                        @"FROM SSO.DemandCertificate a
                        INNER JOIN BAS.[File] b ON b.FileId = a.FileId
                        INNER JOIN BAS.[User] c ON c.UserId = a.CreateBy
                        INNER JOIN SSO.Demand d ON a.DemandId = d.DemandId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CertificateId", @" AND a.CertificateId = @CertificateId", CertificateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FileId", @" AND a.FileId = @FileId", FileId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DemandId", @" AND a.DemandId = @DemandId", DemandId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CertificateDesc", @" AND a.CertificateDesc = @CertificateDesc", CertificateDesc);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CertificateId";
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

        #region //GetDemandFlowSettingDiagram -- 取得需求流程設定資料(流程圖) -- Ben Ma 2023.07.28
        public string GetDemandFlowSettingDiagram(int DemandId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Shape Json組成
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SettingId, a.Xaxis, a.Yaxis
                            , b.FlowName, b.FlowImage
                            , ISNULL(c.DemandFlowUser, '') DemandFlowUser
                            , ISNULL(d.FlowStatus, 'B') FlowStatus
                            FROM SSO.DemandFlowSetting a
                            INNER JOIN SSO.DemandFlow b ON a.FlowId = b.FlowId
                            OUTER APPLY (
                                SELECT STUFF((
                                    SELECT ',' + ISNULL(ab.RoleName, ac.UserName)
                                    FROM SSO.DemandFlowUser aa
                                    LEFT JOIN SSO.DemandRole ab ON aa.RoleId = ab.RoleId
                                    LEFT JOIN BAS.[User] ac ON aa.UserId = ac.UserId
                                    WHERE aa.SettingId = a.SettingId
                                    ORDER BY ab.RoleName, ac.UserNo
                                    FOR XML PATH('')
                                ), 1, 1, '') DemandFlowUser
                            ) c
                            OUTER APPLY (
                                SELECT TOP 1 FlowStatus
                                FROM SSO.DemandFlowLog da
                                WHERE da.SettingId = a.SettingId
                                ORDER BY LogId DESC
                            ) d
                            WHERE a.DemandId = @DemandId";
                    dynamicParameters.Add("DemandId", DemandId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");

                    List<SettingDiagramShape> settingDiagramShapes = new List<SettingDiagramShape>();
                    foreach (var item in result)
                    {
                        settingDiagramShapes.Add(new SettingDiagramShape
                        {
                            id = item.SettingId.ToString(),
                            flowName = item.FlowName,
                            flowImage = item.FlowImage,
                            flowUser = item.DemandFlowUser,
                            x = Convert.ToInt32(item.Xaxis),
                            y = Convert.ToInt32(item.Yaxis),
                            flowStatus = item.FlowStatus
                        });
                    }
                    #endregion

                    #region //Line Json組成
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SourceSettingId, a.TargetSettingId
                            FROM SSO.DemandFlowLink a
                            INNER JOIN SSO.DemandFlowSetting b ON a.SourceSettingId = b.SettingId
                            INNER JOIN SSO.DemandFlowSetting c ON a.TargetSettingId = c.SettingId
                            WHERE b.DemandId = c.DemandId
                            AND b.DemandId = @DemandId";
                    dynamicParameters.Add("DemandId", DemandId);
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

        #region //GetDemandFlowSetting -- 取得需求流程設定資料 -- Ben Ma 2023.08.08
        public string GetDemandFlowSetting(int SettingId, int DemandId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SettingId, a.FlowId, a.Xaxis, a.Yaxis
                            , b.FlowName
                            FROM SSO.DemandFlowSetting a
                            INNER JOIN SSO.DemandFlow b ON a.FlowId = b.FlowId
                            WHERE a.SettingId = @SettingId
                            AND a.DemandId = @DemandId";
                    dynamicParameters.Add("SettingId", SettingId);
                    dynamicParameters.Add("DemandId", DemandId);

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

        #region //GetDemandFlowUser -- 取得需求流程使用者對應資料 -- Ben Ma 2023.08.08
        public string GetDemandFlowUser(int SettingId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT ISNULL(a.RoleId, -1) RoleId, ISNULL(a.UserId, -1) UserId, a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                            , ISNULL(b.RoleName, c.UserNo + ' ' + c.UserName) RoleUserName
                            , CASE WHEN a.RoleId IS NOT NULL THEN 'Y' ELSE 'N' END RoleStatus
                            , (
                                SELECT ab.UserNo, ab.UserName
                                FROM SSO.RoleUser aa
                                INNER JOIN BAS.[User] ab ON aa.UserId = ab.UserId
                                WHERE aa.RoleId = a.RoleId
                                ORDER BY ab.UserNo
                                FOR JSON PATH, ROOT('data')
                            ) RoleUser
                            FROM SSO.DemandFlowUser a
                            LEFT JOIN SSO.DemandRole b ON a.RoleId = b.RoleId
                            LEFT JOIN BAS.[User] c ON a.UserId = c.UserId
                            WHERE a.SettingId = @SettingId";
                    dynamicParameters.Add("SettingId", SettingId);

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

        #region //GetDemandItemNo -- 取得需求流程對應項目資料 -- Ben Ma 2023.08.15
        public string GetDemandItemNo(int SettingId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int demandId = -1;
                    string flowNo = "";

                    #region //取得流程代號
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DemandId, b.FlowNo
                            FROM SSO.DemandFlowSetting a
                            INNER JOIN SSO.DemandFlow b ON a.FlowId = b.FlowId
                            WHERE a.SettingId = @SettingId";
                    dynamicParameters.Add("SettingId", SettingId);

                    var resultFlow = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in resultFlow)
                    {
                        demandId = Convert.ToInt32(item.DemandId);
                        flowNo = item.FlowNo;
                    }
                    #endregion

                    if (resultFlow.Count() > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = "";

                        switch (flowNo)
                        {
                            #region //品號建立
                            case "MtlItem":
                                sql = @"SELECT a.MtlItemNo ItemNo
                                        FROM SSO.DemandMtlItem a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY a.MtlItemNo";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //BOM建立
                            case "BOM":
                                sql = @"SELECT a.MtlItemNo ItemNo
                                        FROM SSO.DemandBom a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY a.MtlItemNo";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //途程建立
                            case "Routing":
                                sql = @"SELECT b.RoutingName ItemNo
                                        FROM SSO.DemandRouting a
                                        INNER JOIN MES.Routing b ON a.RoutingId = b.RoutingId
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY a.RoutingId";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //訂單建立
                            case "So":
                                sql = @"SELECT (a.SoErpPrefix + '-' + a.SoErpNo) ItemNo
                                        FROM SSO.DemandSalesOrder a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY (a.SoErpPrefix + '-' + a.SoErpNo)";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //需求計算
                            case "MRP":
                                break;
                            #endregion
                            #region //請購單建立
                            case "Pr":
                                sql = @"SELECT (a.PrErpPrefix + '-' + a.PrErpNo) ItemNo
                                        FROM SSO.DemandPurchaseRequisition a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY (a.PrErpPrefix + '-' + a.PrErpNo)";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //製令建立
                            case "Wip":
                                sql = @"SELECT (a.WoErpPrefix + '-' + a.WoErpNo) ItemNo
                                        FROM SSO.DemandWipOrder a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY (a.WoErpPrefix + '-' + a.WoErpNo)";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //領料單建立
                            case "Mr":
                                sql = @"SELECT (a.MrErpPrefix + '-' + a.MrErpNo) ItemNo
                                        FROM SSO.DemandMaterialRequisition a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY (a.MrErpPrefix + '-' + a.MrErpNo)";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //生產追蹤
                            case "MES":
                                break;
                            #endregion
                            #region //入庫單建立
                            case "Psin":
                                sql = @"SELECT (a.PsinErpPrefix + '-' + a.PsinErpNo) ItemNo
                                        FROM SSO.DemandProductionStockInNote a
                                        WHERE a.DemandId = @DemandId
                                        ORDER BY (a.PsinErpPrefix + '-' + a.PsinErpNo)";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            #region //出貨定版
                            case "DeliveryFinalize":
                                break;
                            #endregion
                            #region //揀貨/出貨
                            case "Picking":
                                break;
                            #endregion
                            #region //暫出單/銷貨單建立
                            case "Tsn":
                                sql = @"SELECT result.ItemNo
                                        FROM (
                                            SELECT (a.TsnErpPrefix + '-' + a.TsnErpNo) ItemNo
                                            FROM SSO.DemandTempShippingNote a
                                            WHERE a.DemandId = @DemandId
                                            UNION ALL
                                            SELECT (a.SsErpPrefix + '-' + a.SsErpNo) ItemNo
                                            FROM SSO.DemandSalesShipment a
                                            WHERE a.DemandId = @DemandId
                                        ) result
                                        ORDER BY result.ItemNo";
                                dynamicParameters.Add("DemandId", demandId);
                                break;
                            #endregion
                            default:
                                throw new SystemException("流程代號錯誤!");
                        }

                        if (sql.Length > 0)
                        {
                            var result = sqlConnection.Query(sql, dynamicParameters);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result
                            });
                            #endregion
                        }
                        else
                        {
                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = ""
                            });
                            #endregion
                        }
                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success"
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

        #region //GetDemandNotificationLog -- 取得需求平台通知紀錄 -- Ben Ma 2023.08.10
        public string GetDemandNotificationLog(int SettingId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.LogId, a.LogContent, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') LogDate
                            FROM SSO.DemandNotificationLog a
                            WHERE a.SettingId = @SettingId
                            ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("SettingId", SettingId);

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

        #region //GetExpandMtlItem -- 針對需求品號展開全部BOM階資料 -- Ann 2023-01-19
        public string GetExpandMtlItem(int MtlItemId = -1)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"WITH RecursiveCTE AS (
                                SELECT a.MtlItemId
                                , c.MtlItemId ParentMtlItemId, c.MtlItemNo ParentMtlItemNo, c.MtlItemName ParentMtlItemName
                                , d.MtlItemNo, d.MtlItemName, d.ItemAttribute
                                , e.TypeName ItemAttributeName
                                FROM PDM.BomDetail a 
                                INNER JOIN PDM.BillOfMaterials b ON a.BomId = b.BomId
                                INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                INNER JOIN BAS.[Type] e ON d.ItemAttribute = e.TypeNo AND e.TypeSchema = 'MtlItem.ItemAttribute'
                                WHERE b.MtlItemId = @MtlItemId
                                UNION ALL
                                SELECT x.MtlItemId
                                , xc.MtlItemId ParentMtlItemId, xc.MtlItemNo ParentMtlItemNo, xc.MtlItemName ParentMtlItemName
                                , xd.MtlItemNo ,xd.MtlItemName, xd.ItemAttribute
                                , xe.TypeName ItemAttributeName
                                FROM PDM.BomDetail x
                                INNER JOIN PDM.BillOfMaterials xb ON x.BomId = xb.BomId
                                INNER JOIN RecursiveCTE xa ON xb.MtlItemId = xa.MtlItemId
                                INNER JOIN PDM.MtlItem xc ON xb.MtlItemId = xc.MtlItemId
                                INNER JOIN PDM.MtlItem xd ON x.MtlItemId = xd.MtlItemId
                                INNER JOIN BAS.[Type] xe ON xd.ItemAttribute = xe.TypeNo AND xe.TypeSchema = 'MtlItem.ItemAttribute'
                            )
                            SELECT MtlItemId, MtlItemNo, MtlItemName, ItemAttribute, ItemAttributeName
                            , ParentMtlItemId, ParentMtlItemNo, ParentMtlItemName
                            FROM RecursiveCTE";
                    dynamicParameters.Add("MtlItemId", MtlItemId);

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

        #region //Add
        #region //AddDemand -- 需求資料新增 -- Ben Ma 2023.07.25
        public string AddDemand(int SourceId, string DemandDesc, string DemandDate, string DemandDeadline
            , string DepCus, int DemandDepartment, int DemandCustomer, int DemandUser, int ItemTypeId)
        {
            try
            {
                DateTime tempDate = default(DateTime);

                if (DemandDesc.Length <= 0) throw new SystemException("【需求描述】不能為空!");
                if (DemandDesc.Length > 50) throw new SystemException("【需求描述】長度錯誤!");
                if (!DateTime.TryParse(DemandDate, out tempDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (!DateTime.TryParse(DemandDeadline, out tempDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (!Regex.IsMatch(DepCus, "^(department|customer)$", RegexOptions.IgnoreCase)) throw new SystemException("【部門/客戶】設定錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandSource
                                WHERE SourceId = @SourceId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SourceId", SourceId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultSource = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSource.Count() <= 0) throw new SystemException("需求來源資料錯誤!");
                        #endregion
                        
                        switch (DepCus)
                        {
                            case "department":
                                #region //判斷部門資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.Department
                                        WHERE DepartmentId = @DepartmentId";
                                dynamicParameters.Add("DepartmentId", DemandDepartment);

                                var resultDep = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDep.Count() <= 0) throw new SystemException("部門資料錯誤!");
                                #endregion
                                break;
                            case "customer":
                                #region //判斷客戶資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.Customer
                                        WHERE CustomerId = @Customer";
                                dynamicParameters.Add("Customer", DemandCustomer);

                                var resultCus = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCus.Count() <= 0) throw new SystemException("客戶資料錯誤!");
                                #endregion
                                break;
                        }

                        #region //判斷需求人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", DemandUser);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("需求人員資料錯誤!");
                        #endregion

                        #region //判斷需求產品類別是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType a 
                                WHERE a.ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var ItemTypeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ItemTypeResult.Count() <= 0) throw new SystemException("需求產品類別資料錯誤!!");
                        #endregion

                        #region //需求單號取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(DemandNo), '000'), 3)) + 1 CurrentNum
                                FROM SSO.Demand
                                WHERE DemandNo LIKE @DemandNo";
                        dynamicParameters.Add("DemandNo", string.Format("{0}{1}___", "D", DateTime.Now.ToString("yyyyMM")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string demandNo = string.Format("{0}{1}{2}", "D", DateTime.Now.ToString("yyyyMM"), string.Format("{0:000}", currentNum));
                        #endregion

                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SSO.Demand (SourceId, DemandNo, DemandDesc, DemandDate
                                , DemandDeadline, DemandDepartment, DemandCustomer, DemandUser
                                , DemandStatus, ItemTypeId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DemandId
                                VALUES (@SourceId, @DemandNo, @DemandDesc, @DemandDate
                                , @DemandDeadline, @DemandDepartment, @DemandCustomer, @DemandUser
                                , @DemandStatus, @ItemTypeId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SourceId,
                                DemandNo = demandNo,
                                DemandDesc,
                                DemandDate,
                                DemandDeadline,
                                DemandDepartment = DemandDepartment <= 0 ? (int?)null : DemandDepartment,
                                DemandCustomer = DemandCustomer <= 0 ? (int?)null : DemandCustomer,
                                DemandUser,
                                DemandStatus = "P",
                                ItemTypeId,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

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

        #region //AddDemandCertificate -- 需求單憑證新增 -- Yi 2023.07.26
        public string AddDemandCertificate(int FileId, int DemandId, string CertificateDesc)
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
                        sql = @"INSERT INTO SSO.DemandCertificate (DemandId, CertificateDesc, FileId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CertificateId
                                VALUES (@DemandId, @CertificateDesc, @FileId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DemandId,
                                CertificateDesc,
                                FileId,
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

        #region //AddDemandFlowUser -- 需求流程使用者對應新增 -- Ben Ma 2023.08.08
        public string AddDemandFlowUser(int SettingId, string UserRole, string Users, string Roles)
        {
            try
            {
                if (!Regex.IsMatch(UserRole, "^(user|role)$", RegexOptions.IgnoreCase)) throw new SystemException("【使用者/角色】設定錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");
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
                                    sql = @"INSERT INTO SSO.DemandFlowUser (SettingId, UserId
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
                                    sql = @"INSERT INTO SSO.DemandFlowUser (SettingId, RoleId
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
        #region //UpdateDemand -- 需求資料更新 -- Ben Ma 2023.07.26
        public string UpdateDemand(int DemandId, int SourceId, string DemandDesc, string DemandDate, string DemandDeadline
            , string DepCus, int DemandDepartment, int DemandCustomer, int DemandUser, int ItemTypeId)
        {
            try
            {
                DateTime tempDate = default(DateTime);

                if (DemandDesc.Length <= 0) throw new SystemException("【需求描述】不能為空!");
                if (DemandDesc.Length > 50) throw new SystemException("【需求描述】長度錯誤!");
                if (!DateTime.TryParse(DemandDate, out tempDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (!DateTime.TryParse(DemandDeadline, out tempDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (!Regex.IsMatch(DepCus, "^(department|customer)$", RegexOptions.IgnoreCase)) throw new SystemException("【部門/客戶】設定錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DemandStatus
                                FROM SSO.Demand
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        string demandStatus = "";
                        foreach (var item in resultDemand)
                        {
                            demandStatus = item.DemandStatus;
                        }

                        if (demandStatus != "P") throw new SystemException("需求已開始，無法修改!");
                        #endregion

                        #region //判斷需求來源資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandSource
                                WHERE SourceId = @SourceId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SourceId", SourceId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultSource = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSource.Count() <= 0) throw new SystemException("需求來源資料錯誤!");
                        #endregion

                        switch (DepCus)
                        {
                            case "department":
                                #region //判斷部門資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.Department
                                        WHERE DepartmentId = @DepartmentId";
                                dynamicParameters.Add("DepartmentId", DemandDepartment);

                                var resultDep = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDep.Count() <= 0) throw new SystemException("部門資料錯誤!");
                                #endregion
                                break;
                            case "customer":
                                #region //判斷客戶資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.Customer
                                        WHERE CustomerId = @CustomerId";
                                dynamicParameters.Add("CustomerId", DemandCustomer);

                                var resultCus = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCus.Count() <= 0) throw new SystemException("客戶資料錯誤!");
                                #endregion
                                break;
                        }

                        #region //判斷需求人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", DemandUser);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("需求人員資料錯誤!");
                        #endregion

                        #region //判斷需求產品類別是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType a 
                                WHERE a.ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var ItemTypeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ItemTypeResult.Count() <= 0) throw new SystemException("需求產品類別資料錯誤!!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.Demand SET
                                SourceId = @SourceId,
                                DemandDesc = @DemandDesc,
                                DemandDate = @DemandDate,
                                DemandDeadline = @DemandDeadline,
                                DemandDepartment = @DemandDepartment,
                                DemandCustomer = @DemandCustomer,
                                DemandUser = @DemandUser,
                                ItemTypeId = @ItemTypeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DemandId = @DemandId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SourceId,
                                DemandDesc,
                                DemandDate,
                                DemandDeadline,
                                DemandDepartment = DemandDepartment <= 0 ? (int?)null : DemandDepartment,
                                DemandCustomer = DemandCustomer <= 0 ? (int?)null : DemandCustomer,
                                DemandUser,
                                ItemTypeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                DemandId
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

        #region //UpdateDemandCertificate -- 需求單憑證更名 -- Yi 2023.07.26
        public string UpdateDemandCertificate(int CertificateId, string NewFileName, string CertificateDesc)
        {
            try
            {
                if (NewFileName.Length <= 0) throw new SystemException("【新檔名】不能為空");
                if (CertificateDesc.Length <= 0) throw new SystemException("【描述】不能為空");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求單憑證資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandCertificate
                                WHERE CertificateId = @CertificateId";
                        dynamicParameters.Add("CertificateId", CertificateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求單憑證資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.FileName = @FileName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM BAS.[File] a
                                INNER JOIN SSO.DemandCertificate b ON a.FileId = b.FileId
                                WHERE b.CertificateId = @CertificateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileName = NewFileName,
                                LastModifiedDate,
                                LastModifiedBy,
                                CertificateId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE b SET
                                b.CertificateDesc = @CertificateDesc,
                                b.LastModifiedDate = @LastModifiedDate,
                                b.LastModifiedBy = @LastModifiedBy
                                FROM BAS.[File] a
                                INNER JOIN SSO.DemandCertificate b ON a.FileId = b.FileId
                                WHERE b.CertificateId = @CertificateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CertificateDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                CertificateId
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateDemandConfirm -- 需求確認 -- Ben Ma 2023.07.27
        public string UpdateDemandConfirm(int DemandId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SourceId, a.DemandNo, a.DemandStatus
                                , b.UserNo DemandUserNo
                                FROM SSO.Demand a
                                INNER JOIN BAS.[User] b ON a.DemandUser = b.UserId
                                WHERE a.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        int sourceId = -1;
                        string demandNo = "", demandStatus = "", demandUserNo = "";
                        foreach (var item in resultDemand)
                        {
                            sourceId = Convert.ToInt32(item.SourceId);
                            demandNo = item.DemandNo;
                            demandStatus = item.DemandStatus;
                            demandUserNo = item.DemandUserNo;
                        }

                        if (demandStatus != "P") throw new SystemException("需求已開始，無法確認!");
                        #endregion

                        #region //判斷是否有上傳憑證
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandCertificate
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        var resultCertificate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCertificate.Count() <= 0) throw new SystemException("尚未上傳憑證!");
                        #endregion

                        #region //判斷需求來源流程設定資料是否齊全
                        #region //需求來源流程設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting
                                WHERE SourceId = @SourceId";
                        dynamicParameters.Add("SourceId", sourceId);

                        var resultFlowSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFlowSetting.Count() <= 0) throw new SystemException("需求來源流程尚未設定!");
                        #endregion

                        #region //需求來源流程使用者對應
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.SourceFlowSetting a
                                OUTER APPLY (
                                    SELECT COUNT(1) TotalUsers
                                    FROM SSO.SourceFlowUser ba
                                    WHERE ba.SettingId = a.SettingId
                                ) b
                                WHERE a.SourceId = @SourceId
                                AND b.TotalUsers = 0";
                        dynamicParameters.Add("SourceId", sourceId);

                        var resultFlowUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFlowUser.Count() > 0) throw new SystemException("尚有部分需求來源流程未設定使用者!");
                        #endregion
                        #endregion

                        #region //複製需求來源流程設定資料
                        #region //需求來源流程設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SettingId OriginalSettingId, a.FlowId, a.Xaxis, a.Yaxis
                                FROM SSO.SourceFlowSetting a
                                WHERE a.SourceId = @SourceId";
                        dynamicParameters.Add("SourceId", sourceId);

                        List<SourceFlowSetting> sourceFlowSettings = sqlConnection.Query<SourceFlowSetting>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //需求來源流程設定連結
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SourceSettingId OriginalSourceSettingId, a.TargetSettingId OriginalTargetSettingId
                                FROM SSO.SourceFlowLink a
                                INNER JOIN SSO.SourceFlowSetting b ON a.SourceSettingId = b.SettingId
                                WHERE b.SourceId = @SourceId";
                        dynamicParameters.Add("SourceId", sourceId);

                        List<SourceFlowLink> sourceFlowLinks = sqlConnection.Query<SourceFlowLink>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //需求來源流程使用者對應
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SettingId OriginalSettingId, a.RoleId, a.UserId, a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                FROM SSO.SourceFlowUser a
                                INNER JOIN SSO.SourceFlowSetting b ON a.SettingId = b.SettingId
                                WHERE b.SourceId = @SourceId";
                        dynamicParameters.Add("SourceId", sourceId);

                        List<SourceFlowUser> sourceFlowUsers = sqlConnection.Query<SourceFlowUser>(sql, dynamicParameters).ToList();
                        #endregion
                        #endregion

                        int rowsAffected = 0;
                        #region //新增需求流程設定資料
                        foreach (var setting in sourceFlowSettings)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SSO.DemandFlowSetting (DemandId, FlowId, Xaxis, Yaxis
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.SettingId
                                    VALUES (@DemandId, @FlowId, @Xaxis, @Yaxis
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DemandId,
                                    FlowId = Convert.ToInt32(setting.FlowId),
                                    Xaxis = Convert.ToInt32(setting.Xaxis),
                                    Yaxis = Convert.ToInt32(setting.Yaxis),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            int newSettingId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).SettingId;

                            sourceFlowSettings
                                .Where(x => x.OriginalSettingId == setting.OriginalSettingId)
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.NewSettingId = newSettingId;
                                });
                        }
                        #endregion

                        #region //新增需求流程設定連結資料
                        sourceFlowLinks
                            .ToList()
                            .ForEach(x =>
                            { 
                                x.NewSourceSettingId = sourceFlowSettings.Where(y => y.OriginalSettingId == x.OriginalSourceSettingId).Select(y => y.NewSettingId).FirstOrDefault();
                                x.NewTargetSettingId = sourceFlowSettings.Where(y => y.OriginalSettingId == x.OriginalTargetSettingId).Select(y => y.NewSettingId).FirstOrDefault();
                                x.CreateDate = CreateDate;
                                x.LastModifiedDate = LastModifiedDate;
                                x.CreateBy = CreateBy;
                                x.LastModifiedBy = LastModifiedBy;
                            });

                        sql = @"INSERT INTO SSO.DemandFlowLink (SourceSettingId, TargetSettingId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@NewSourceSettingId, @NewTargetSettingId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, sourceFlowLinks);
                        #endregion

                        #region //新增需求流程使用者對應資料
                        sourceFlowUsers
                            .ToList()
                            .ForEach(x =>
                            {
                                x.NewSettingId = sourceFlowSettings.Where(y => y.OriginalSettingId == x.OriginalSettingId).Select(y => y.NewSettingId).FirstOrDefault();
                                x.CreateDate = CreateDate;
                                x.LastModifiedDate = LastModifiedDate;
                                x.CreateBy = CreateBy;
                                x.LastModifiedBy = LastModifiedBy;
                            });

                        sql = @"INSERT INTO SSO.DemandFlowUser (SettingId, RoleId, UserId
                                , MailAdviceStatus, PushAdviceStatus, WorkWeixinStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@NewSettingId, @RoleId, @UserId
                                , @MailAdviceStatus, @PushAdviceStatus, @WorkWeixinStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, sourceFlowUsers);
                        #endregion

                        #region //更新需求狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.Demand SET
                                DemandStatus = @DemandStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DemandId = @DemandId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DemandStatus = "I",
                                LastModifiedDate,
                                LastModifiedBy,
                                DemandId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //觸發流程首關
                        dynamicParameters = new DynamicParameters();
                        sql = @"DECLARE @rowsAdded int

                                DECLARE @demandFlow TABLE
                                ( 
                                    SettingId int,
                                    SourceSettingId int,
                                    FlowRoute nvarchar(MAX),
                                    processed int DEFAULT(0)
                                )

                                INSERT @demandFlow
                                    SELECT a.SettingId, ISNULL(b.SourceSettingId, -1) SourceSettingId
                                    , CAST(ISNULL(b.SourceSettingId, -1) AS nvarchar(MAX)) AS FlowRoute
                                    , 0 processed
                                    FROM SSO.DemandFlowSetting a
                                    LEFT JOIN SSO.DemandFlowLink b ON a.SettingId = b.TargetSettingId
                                    WHERE a.DemandId = @DemandId
                                    AND b.SourceSettingId IS NULL

                                SET @rowsAdded=@@rowcount

                                WHILE @rowsAdded > 0
                                BEGIN
                                    UPDATE @demandFlow SET processed = 1 WHERE processed = 0

                                    INSERT @demandFlow
                                        SELECT b.TargetSettingId, b.SourceSettingId
                                        , CAST(a.FlowRoute + ';' + CAST(b.SourceSettingId AS nvarchar(MAX)) + ':' + ISNULL(c.FlowStatus, 'B') AS nvarchar(MAX)) AS FlowRoute
                                        , 0
                                        FROM @demandFlow a
                                        INNER JOIN SSO.DemandFlowLink b ON a.SettingId = b.SourceSettingId
                                        OUTER APPLY (
                                            SELECT TOP 1 FlowStatus
                                            FROM SSO.DemandFlowLog ca
                                            WHERE ca.SettingId = a.SettingId
                                            ORDER BY LogId DESC
                                        ) c
                                        WHERE a.processed = 1

                                    SET @rowsAdded = @@rowcount

                                    UPDATE @demandFlow SET processed = 2 WHERE processed = 1
                                END;

                                SELECT a.SettingId, STUFF(a.FlowRoute, 1, 3, '') FlowRoute, c.FlowId, c.FlowName, ISNULL(d.FlowStatus, 'B') FlowStatus
                                FROM @demandFlow a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                INNER JOIN SSO.DemandFlow c ON b.FlowId = c.FlowId
                                OUTER APPLY (
                                    SELECT TOP 1 FlowStatus
                                    FROM SSO.DemandFlowLog da
                                    WHERE da.SettingId = a.SettingId
                                    ORDER BY LogId DESC
                                ) d
                                ORDER BY a.FlowRoute";
                        dynamicParameters.Add("DemandId", DemandId);

                        List<DemandFlow> demandFlows = sqlConnection.Query<DemandFlow>(sql, dynamicParameters).ToList();
                        if (demandFlows.Count() <= 0) throw new SystemException("需求流程資料錯誤!");

                        var firstDemandFlows = demandFlows.Where(x => x.FlowRoute.Length <= 0).Select(x => x);
                        #endregion

                        foreach (var flow in firstDemandFlows)
                        {
                            #region //需求流程狀態歷程新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SSO.DemandFlowLog (SettingId, FlowStatus, StartDate
                                    , CreateDate, CreateBy)
                                    VALUES (@SettingId, @FlowStatus, @StartDate
                                    , @CreateDate, @CreateBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    flow.SettingId,
                                    FlowStatus = "B",
                                    StartDate = CreateDate,
                                    CreateDate,
                                    CreateBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //需求流程對應使用者資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                    , c.UserId, c.UserNo, c.UserName
                                    FROM SSO.DemandFlowUser a
                                    LEFT JOIN SSO.RoleUser b ON a.RoleId = b.RoleId
                                    LEFT JOIN BAS.[User] c ON a.UserId = c.UserId OR b.UserId = c.UserId
                                    WHERE a.SettingId = @SettingId";
                            dynamicParameters.Add("SettingId", flow.SettingId);

                            var demandFlowUsers = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //通知內容
                            string triggerFunction = "DemandPlatform"
                                , logTitle = string.Format("【需求平台】【{0}】", demandNo)
                                , logContent = string.Format("已接獲需求，請盡速安排時間完成【{0}】", flow.FlowName);
                            #endregion

                            #region //需求平台通知紀錄新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SSO.DemandNotificationLog (SettingId, LogContent, CreateDate)
                                    OUTPUT INSERTED.LogId
                                    VALUES (@SettingId, @LogContent, @CreateDate)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    flow.SettingId,
                                    logContent,
                                    CreateDate
                                });
                            var insertLog = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertLog.Count();

                            int LogId = -1;
                            foreach (var item in insertLog)
                            {
                                LogId = Convert.ToInt32(item.LogId);
                            }
                            #endregion

                            foreach (var user in demandFlowUsers)
                            {
                                #region //系統通知個人設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                        FROM BAS.NotificationSetting a
                                        WHERE a.UserId = @UserId";
                                dynamicParameters.Add("UserId", user.UserId);

                                var userSettings = sqlConnection.Query(sql, dynamicParameters);

                                string personalMailAdviceStatus = "Y",
                                    personalPushAdviceStatus = "Y",
                                    personalWorkWeixinStatus = "Y";
                                if (userSettings.Count() > 0)
                                {
                                    foreach (var item in userSettings)
                                    {
                                        personalMailAdviceStatus = item.MailAdviceStatus;
                                        personalPushAdviceStatus = item.PushAdviceStatus;
                                        personalWorkWeixinStatus = item.WorkWeixinStatus;
                                    }
                                }
                                #endregion

                                #region //需求平台通知詳細紀錄新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SSO.DemandNotificationLogDetail (LogId, UserNo) 
                                        VALUES (@LogId, @UserNo)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LogId,
                                        user.UserNo
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                List<NotificationMode> notificationModes = new List<NotificationMode>();

                                if (personalMailAdviceStatus == "Y") if (user.MailAdviceStatus == "Y") notificationModes.Add(NotificationMode.Mail);
                                if (personalPushAdviceStatus == "Y") if (user.PushAdviceStatus == "Y") notificationModes.Add(NotificationMode.Push);
                                if (personalWorkWeixinStatus == "Y") if (user.WorkWeixinStatus == "Y") notificationModes.Add(NotificationMode.WorkWeixin);

                                if (notificationModes.Count > 0)
                                {
                                    Notification notification = new Notification
                                    {
                                        UserNo = user.UserNo,
                                        TriggerFunction = triggerFunction,
                                        LogTitle = logTitle,
                                        LogContent = logContent,
                                        NotificationModes = notificationModes
                                    };

                                    BaseHelper.SendNotification(notification);
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

        #region //UpdateDemandReConfirm -- 需求反確認 -- Ben Ma 2023.07.28
        public string UpdateDemandReConfirm(int DemandId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DemandStatus
                                FROM SSO.Demand
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        string demandStatus = "";
                        foreach (var item in resultDemand)
                        {
                            demandStatus = item.DemandStatus;
                        }

                        if (demandStatus != "I") throw new SystemException("需求狀態錯誤，無法反確認!");
                        #endregion

                        #region //判斷需求流程狀態歷程是否有已有進行中或是已完成
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlowLog a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                WHERE b.DemandId = @DemandId
                                AND a.FlowStatus IN @FlowStatus";
                        dynamicParameters.Add("DemandId", DemandId);
                        dynamicParameters.Add("FlowStatus", new string[] { "I", "C" });

                        var resultLog = sqlConnection.Query(sql, dynamicParameters);
                        if (resultLog.Count() > 0) throw new SystemException("流程已進行中/完成，無法反確認!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除需求平台通知詳細紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SSO.DemandNotificationLogDetail a
                                INNER JOIN SSO.DemandNotificationLog b ON a.LogId = b.LogId
                                INNER JOIN SSO.DemandFlowSetting c ON b.SettingId = c.SettingId
                                WHERE c.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除需求平台通知紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SSO.DemandNotificationLog a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                WHERE b.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除需求流程狀態歷程資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SSO.DemandFlowLog a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                WHERE b.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除需求流程使用者對應資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SSO.DemandFlowUser a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                WHERE b.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除需求流程設定連結資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SSO.DemandFlowLink a
                                INNER JOIN SSO.DemandFlowSetting b ON a.SourceSettingId = b.SettingId
                                WHERE b.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除需求流程設定資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.DemandFlowSetting
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新需求狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SSO.Demand SET
                                DemandStatus = @DemandStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DemandId = @DemandId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DemandStatus = "P",
                                LastModifiedDate,
                                LastModifiedBy,
                                DemandId
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

        #region //UpdateDemandFlowUserStatus -- 需求流程使用者狀態更新 -- Ben Ma 2023.08.08
        public string UpdateDemandFlowUserStatus(int SettingId, int RoleId, int UserId
            , string MailAdviceStatus, string PushAdviceStatus, string WorkWeixinStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");
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
                        sql = @"UPDATE SSO.DemandFlowUser SET
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

        #region //UpdateDemandFlowTrigger -- 需求流程狀態觸發 -- Ben Ma 2023.08.11
        public string UpdateDemandFlowTrigger(string DemandNo, string FlowNo, string ItemNo, string UserNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int demandId = -1, demandUser = -1, settingId = -1;
                    string demandNo = "", demandUserNo = "", companyNo = "", flowName = "", flowStatus = "", tsnOrSs = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得User資料
                        if (UserNo.Length > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                            dynamicParameters.Add("UserNo", UserNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                            foreach (var item in UserResult)
                            {
                                CurrentUser = item.UserId;
                            }
                        }
                        #endregion

                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DemandId, a.DemandNo, a.DemandUser, b.CompanyId, c.UserNo
                                FROM SSO.Demand a
                                INNER JOIN SSO.DemandSource b ON a.SourceId = b.SourceId
                                INNER JOIN BAS.[User] c ON a.DemandUser = c.UserId
                                WHERE DemandNo = @DemandNo";
                        dynamicParameters.Add("DemandNo", DemandNo);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        int demandCompany = -1;
                        foreach (var item in resultDemand)
                        {
                            demandId = Convert.ToInt32(item.DemandId);
                            demandNo = item.DemandNo;
                            demandUser = Convert.ToInt32(item.DemandUser);
                            demandCompany = Convert.ToInt32(item.CompanyId);
                            demandUserNo = item.UserNo;
                        }
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", demandCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷需求流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SettingId, c.FlowName, ISNULL(d.FlowStatus, 'B') FlowStatus
                                FROM SSO.DemandFlowSetting a
                                INNER JOIN SSO.Demand b ON a.DemandId = b.DemandId
                                INNER JOIN SSO.DemandFlow c ON a.FlowId = c.FlowId
                                OUTER APPLY (
                                    SELECT TOP 1 FlowStatus
                                    FROM SSO.DemandFlowLog da
                                    WHERE da.SettingId = a.SettingId
                                    ORDER BY LogId DESC
                                ) d
                                WHERE b.DemandNo = @DemandNo
                                AND c.FlowNo = @FlowNo";
                        dynamicParameters.Add("DemandNo", DemandNo);
                        dynamicParameters.Add("FlowNo", FlowNo);

                        var resultFlowSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFlowSetting.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");

                        foreach (var item in resultFlowSetting)
                        {
                            settingId = Convert.ToInt32(item.SettingId);
                            flowName = item.FlowName;
                            flowStatus = item.FlowStatus;
                        }

                        if (flowStatus == "C") throw new SystemException("該需求流程已完成!");
                        if (ItemNo.Length > 0)
                        {
                            if (flowStatus != "I") throw new SystemException("尚未確認需求流程!");
                        }
                        else
                        {
                            if (flowStatus != "B") throw new SystemException("需求流程狀態錯誤!");
                        }
                        #endregion
                    }

                    int rowsAffected = 0;
                    bool checkItem = false;
                    IEnumerable<dynamic> resultExist;
                    if (ItemNo.Length > 0)
                    {
                        switch (FlowNo)
                        {
                            #region //品號建立
                            case "MtlItem":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP品號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM INVMB
                                            WHERE MB001 = @MtlItemNo";
                                    dynamicParameters.Add("MtlItemNo", ItemNo);

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //BOM建立
                            case "BOM":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP主件
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BOMMC
                                            WHERE MC001 = @BomItemNo";
                                    dynamicParameters.Add("BomItemNo", ItemNo);

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion

                                    #region //ERP元件
                                    if (checkItem)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM BOMMD
                                                WHERE MD001 = @BomItemNo";
                                        dynamicParameters.Add("BomItemNo", ItemNo);

                                        resultExist = sqlConnection.Query(sql, dynamicParameters);

                                        checkItem = resultExist.Count() > 0;
                                    }
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //途程建立
                            case "Routing":
                                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                                {
                                    #region //BM途程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.Routing a
                                            INNER JOIN MES.RoutingItem b ON a.RoutingId = b.RoutingId
                                            WHERE b.MtlItemId IN (
                                                SELECT y.MtlItemId
                                                FROM SSO.DemandMtlItem x
                                                INNER JOIN PDM.MtlItem y ON x.MtlItemNo = y.MtlItemNo
                                                WHERE x.DemandId = @DemandId
                                            )
                                            AND a.RoutingId = @RoutingId
                                            AND a.RoutingConfirm = @ConfirmStatus
                                            AND b.RoutingItemConfirm = @ConfirmStatus";
                                    dynamicParameters.Add("DemandId", demandId);
                                    dynamicParameters.Add("RoutingId", Convert.ToInt32(ItemNo));
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //訂單建立
                            case "So":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP訂單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM COPTC
                                            WHERE TC001 + '-' + TC002 = @SoNo
                                            AND TC027 = @ConfirmStatus";
                                    dynamicParameters.Add("SoNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //需求計算
                            case "MRP":
                                checkItem = false;
                                break;
                            #endregion
                            #region //請購單建立
                            case "Pr":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP請購單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PURTA
                                            WHERE TA001 + '-' + TA002 = @PrNo
                                            AND TA007 = @ConfirmStatus";
                                    dynamicParameters.Add("PrNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //製令建立
                            case "Wip":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP製令
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTA
                                            WHERE TA001 + '-' + TA002 = @WipNo
                                            AND TA013 = @ConfirmStatus";
                                    dynamicParameters.Add("WipNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //領料單建立
                            case "Mr":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP領料單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTC
                                            WHERE TC001 + '-' + TC002 = @MrNo
                                            AND TC009 = @ConfirmStatus";
                                    dynamicParameters.Add("MrNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //生產追蹤
                            case "MES":
                                checkItem = false;
                                break;
                            #endregion
                            #region //入庫單建立
                            case "Psin":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP領料單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTH
                                            WHERE TH001 + '-' + TH002 = @PsinNo
                                            AND TH023 = @ConfirmStatus";
                                    dynamicParameters.Add("PsinNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //出貨定版
                            case "DeliveryFinalize":
                                checkItem = false;
                                break;
                            #endregion
                            #region //揀貨/出貨
                            case "Picking":
                                checkItem = false;
                                break;
                            #endregion
                            #region //暫出單/銷貨單建立
                            case "Tsn":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP暫出單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM INVTF
                                            WHERE TF001 + '-' + TF002 = @TsnNo
                                            AND TF020 = @ConfirmStatus";
                                    dynamicParameters.Add("TsnNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;

                                    if (checkItem) tsnOrSs = "tsn";
                                    #endregion

                                    #region //ERP銷貨單(判斷不是暫出單才判斷銷貨單)
                                    if (!checkItem)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM COPTG
                                                WHERE TG001 + '-' + TG002 = @SsNo
                                                AND TG023 = @ConfirmStatus";
                                        dynamicParameters.Add("SsNo", ItemNo);
                                        dynamicParameters.Add("ConfirmStatus", "Y");

                                        resultExist = sqlConnection.Query(sql, dynamicParameters);

                                        checkItem = resultExist.Count() > 0;

                                        if (checkItem) tsnOrSs = "ss";
                                    }
                                    #endregion
                                }
                                break;
                            #endregion
                        }
                    }

                    IEnumerable<dynamic> userSettings;
                    string personalMailAdviceStatus = "Y",
                        personalPushAdviceStatus = "Y",
                        personalWorkWeixinStatus = "Y";
                    List<NotificationMode> notificationModes = new List<NotificationMode>();
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求流程狀態歷程是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 LogId
                                FROM SSO.DemandFlowLog
                                WHERE SettingId = @SettingId
                                AND FlowStatus = @FlowStatus";
                        dynamicParameters.Add("SettingId", settingId);
                        dynamicParameters.Add("FlowStatus", flowStatus);

                        var resultDemandFlowLog = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemandFlowLog.Count() <= 0) throw new SystemException("需求流程狀態歷程資料錯誤!");

                        int logId = -1;
                        foreach (var item in resultDemandFlowLog)
                        {
                            logId = Convert.ToInt32(item.LogId);
                        }
                        #endregion

                        #region //更新需求流程狀態歷程
                        string triggerFunction = "DemandPlatform"
                            , logTitle = string.Format("【需求平台】【{0}】", DemandNo)
                            , logContent = "";

                        switch (flowStatus)
                        {
                            case "B":
                                #region //將原狀態更新使用者與日期
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SSO.DemandFlowLog SET
                                        UserId = @UserId,
                                        EndDate = @EndDate
                                        WHERE LogId = @LogId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        UserId = CurrentUser,
                                        EndDate = CreateDate,
                                        LogId = logId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //新增新狀態歷程
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SSO.DemandFlowLog (SettingId, FlowStatus, StartDate
                                        , CreateDate, CreateBy)
                                        VALUES (@SettingId, @FlowStatus, @StartDate
                                        , @CreateDate, @CreateBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SettingId = settingId,
                                        FlowStatus = "I",
                                        StartDate = CreateDate,
                                        CreateDate,
                                        CreateBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //通知內容
                                logContent = "確認需求中";
                                #endregion

                                #region //需求平台通知紀錄新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SSO.DemandNotificationLog (SettingId, LogContent, CreateDate)
                                        VALUES (@SettingId, @LogContent, @CreateDate)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SettingId = settingId,
                                        logContent,
                                        CreateDate
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            case "I":
                                if (checkItem)
                                {
                                    #region //將原狀態更新使用者與日期
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SSO.DemandFlowLog SET
                                            UserId = @UserId,
                                            EndDate = @EndDate
                                            WHERE LogId = @LogId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            UserId = CurrentUser,
                                            EndDate = CreateDate,
                                            LogId = logId
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //新增完成狀態歷程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandFlowLog (SettingId, FlowStatus, StartDate, EndDate
                                            , CreateDate, CreateBy)
                                            VALUES (@SettingId, @FlowStatus, @StartDate, @EndDate
                                            , @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SettingId = settingId,
                                            FlowStatus = "C",
                                            StartDate = CreateDate,
                                            EndDate = CreateDate,
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //新增對應項目
                                    switch (FlowNo)
                                    {
                                        #region //品號建立
                                        case "MtlItem":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandMtlItem (DemandId, MtlItemNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @MtlItemNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    MtlItemNo = ItemNo,
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //BOM建立
                                        case "BOM":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandBom (DemandId, MtlItemNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @MtlItemNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    MtlItemNo = ItemNo,
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //途程建立
                                        case "Routing":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandRouting (DemandId, RoutingId, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @RoutingId, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    RoutingId = Convert.ToInt32(ItemNo),
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //訂單建立
                                        case "So":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandSalesOrder (DemandId, SoErpPrefix, SoErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @SoErpPrefix, @SoErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    SoErpPrefix = ItemNo.Split('-')[0],
                                                    SoErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //需求計算
                                        case "MRP":
                                            break;
                                        #endregion
                                        #region //請購單建立
                                        case "Pr":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandPurchaseRequisition (DemandId, PrErpPrefix, PrErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @PrErpPrefix, @PrErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    PrErpPrefix = ItemNo.Split('-')[0],
                                                    PrErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //製令建立
                                        case "Wip":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandWipOrder (DemandId, WoErpPrefix, WoErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @WoErpPrefix, @WoErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    WoErpPrefix = ItemNo.Split('-')[0],
                                                    WoErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //領料單建立
                                        case "Mr":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandMaterialRequisition (DemandId, MrErpPrefix, MrErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @MrErpPrefix, @MrErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    MrErpPrefix = ItemNo.Split('-')[0],
                                                    MrErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //生產追蹤
                                        case "MES":
                                            break;
                                        #endregion
                                        #region //入庫單建立
                                        case "Psin":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandProductionStockInNote (DemandId, PsinErpPrefix, PsinErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @PsinErpPrefix, @PsinErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    PsinErpPrefix = ItemNo.Split('-')[0],
                                                    PsinErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion
                                        #region //出貨定版
                                        case "DeliveryFinalize":
                                            break;
                                        #endregion
                                        #region //揀貨/出貨
                                        case "Picking":
                                            break;
                                        #endregion
                                        #region //暫出單/銷貨單建立
                                        case "Tsn":
                                            switch (tsnOrSs)
                                            {
                                                case "tsn":
                                                    #region //需求相關暫出單
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"INSERT INTO SSO.DemandTempShippingNote (DemandId, TsnErpPrefix, TsnErpNo, CreateDate, CreateBy)
                                                            VALUES (@DemandId, @TsnErpPrefix, @TsnErpNo, @CreateDate, @CreateBy)";
                                                    dynamicParameters.AddDynamicParams(
                                                        new
                                                        {
                                                            DemandId = demandId,
                                                            TsnErpPrefix = ItemNo.Split('-')[0],
                                                            TsnErpNo = ItemNo.Split('-')[1],
                                                            CreateDate,
                                                            CreateBy
                                                        });
                                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                    #endregion
                                                    break;
                                                case "ss":
                                                    #region //需求相關銷貨單
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"INSERT INTO SSO.DemandSalesShipment (DemandId, SsErpPrefix, SsErpNo, CreateDate, CreateBy)
                                                            VALUES (@DemandId, @SsErpPrefix, @SsErpNo, @CreateDate, @CreateBy)";
                                                    dynamicParameters.AddDynamicParams(
                                                        new
                                                        {
                                                            DemandId = demandId,
                                                            SsErpPrefix = ItemNo.Split('-')[0],
                                                            SsErpNo = ItemNo.Split('-')[1],
                                                            CreateDate,
                                                            CreateBy
                                                        });
                                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                    #endregion
                                                    break;
                                            }
                                            break;
                                        #endregion
                                    }
                                    #endregion

                                    #region //通知內容
                                    logContent = string.Format("【{0}】已完成", flowName);
                                    #endregion

                                    #region //需求平台通知紀錄新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandNotificationLog (SettingId, LogContent, CreateDate)
                                            VALUES (@SettingId, @LogContent, @CreateDate)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SettingId = settingId,
                                            logContent,
                                            CreateDate
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //通知需求人員
                                    #region //系統通知個人設定
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                            FROM BAS.NotificationSetting a
                                            WHERE a.UserId = @UserId";
                                    dynamicParameters.Add("UserId", demandUser);

                                    userSettings = sqlConnection.Query(sql, dynamicParameters);

                                    personalMailAdviceStatus = "Y";
                                    personalPushAdviceStatus = "Y";
                                    personalWorkWeixinStatus = "Y";
                                    if (userSettings.Count() > 0)
                                    {
                                        foreach (var item in userSettings)
                                        {
                                            personalMailAdviceStatus = item.MailAdviceStatus;
                                            personalPushAdviceStatus = item.PushAdviceStatus;
                                            personalWorkWeixinStatus = item.WorkWeixinStatus;
                                        }
                                    }
                                    #endregion

                                    notificationModes = new List<NotificationMode>();

                                    if (personalMailAdviceStatus == "Y") notificationModes.Add(NotificationMode.Mail);
                                    if (personalPushAdviceStatus == "Y") notificationModes.Add(NotificationMode.Push);
                                    if (personalWorkWeixinStatus == "Y") notificationModes.Add(NotificationMode.WorkWeixin);

                                    if (notificationModes.Count > 0)
                                    {
                                        Notification notification = new Notification
                                        {
                                            UserNo = demandUserNo,
                                            TriggerFunction = triggerFunction,
                                            LogTitle = logTitle,
                                            LogContent = logContent,
                                            NotificationModes = notificationModes
                                        };

                                        //BaseHelper.SendNotification(notification);
                                    }
                                    #endregion

                                    #region //觸發流程-下一關
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DECLARE @rowsAdded int

                                            DECLARE @demandFlow TABLE
                                            ( 
                                                SettingId int,
                                                SourceSettingId int,
                                                FlowRoute nvarchar(MAX),
                                                processed int DEFAULT(0)
                                            )

                                            INSERT @demandFlow
                                                SELECT a.SettingId, ISNULL(b.SourceSettingId, -1) SourceSettingId
                                                , CAST(ISNULL(b.SourceSettingId, -1) AS nvarchar(MAX)) AS FlowRoute
                                                , 0 processed
                                                FROM SSO.DemandFlowSetting a
                                                LEFT JOIN SSO.DemandFlowLink b ON a.SettingId = b.TargetSettingId
                                                WHERE a.DemandId = @DemandId
                                                AND b.SourceSettingId IS NULL

                                            SET @rowsAdded=@@rowcount

                                            WHILE @rowsAdded > 0
                                            BEGIN
                                                UPDATE @demandFlow SET processed = 1 WHERE processed = 0

                                                INSERT @demandFlow
                                                    SELECT b.TargetSettingId, b.SourceSettingId
                                                    , CAST(a.FlowRoute + ';' + CAST(b.SourceSettingId AS nvarchar(MAX)) + ':' + ISNULL(c.FlowStatus, 'B') AS nvarchar(MAX)) AS FlowRoute
                                                    , 0
                                                    FROM @demandFlow a
                                                    INNER JOIN SSO.DemandFlowLink b ON a.SettingId = b.SourceSettingId
                                                    OUTER APPLY (
                                                        SELECT TOP 1 FlowStatus
                                                        FROM SSO.DemandFlowLog ca
                                                        WHERE ca.SettingId = a.SettingId
                                                        ORDER BY LogId DESC
                                                    ) c
                                                    WHERE a.processed = 1

                                                SET @rowsAdded = @@rowcount

                                                UPDATE @demandFlow SET processed = 2 WHERE processed = 1
                                            END;

                                            SELECT a.SettingId, STUFF(a.FlowRoute, 1, 3, '') FlowRoute, c.FlowId, c.FlowName, ISNULL(d.FlowStatus, 'B') FlowStatus
                                            FROM @demandFlow a
                                            INNER JOIN SSO.DemandFlowSetting b ON a.SettingId = b.SettingId
                                            INNER JOIN SSO.DemandFlow c ON b.FlowId = c.FlowId
                                            OUTER APPLY (
                                                SELECT TOP 1 FlowStatus
                                                FROM SSO.DemandFlowLog da
                                                WHERE da.SettingId = a.SettingId
                                                ORDER BY LogId DESC
                                            ) d
                                            WHERE ISNULL(d.FlowStatus, 'B') != @FlowStatus
                                            ORDER BY a.FlowRoute";
                                    dynamicParameters.Add("DemandId", demandId);
                                    dynamicParameters.Add("FlowStatus", "C"); //排除已完成

                                    List<DemandFlow> demandFlows = sqlConnection.Query<DemandFlow>(sql, dynamicParameters).ToList();
                                    if (demandFlows.Count() <= 0) throw new SystemException("需求流程資料錯誤!");

                                    var flows = demandFlows.Select(x => new { x.FlowId, x.FlowName, x.SettingId }).Distinct().ToList();
                                    List<DemandFlow> triggerFlows = new List<DemandFlow>();
                                    foreach (var flow in flows)
                                    {
                                        var flowRoutes = demandFlows.Where(x => x.SettingId == flow.SettingId).Select(x => x.FlowRoute).ToList();

                                        bool checkFlowCompleted = true;
                                        foreach (var flowRoute in flowRoutes)
                                        {
                                            if (checkFlowCompleted)
                                            {
                                                List<DemandFlowStatus> demandFlowStatuses = new List<DemandFlowStatus>();
                                                List<string> passingFlows = flowRoute.Split(';').ToList();

                                                foreach (var passFlow in passingFlows)
                                                {
                                                    demandFlowStatuses.Add(new DemandFlowStatus
                                                    {
                                                        SettingId = Convert.ToInt32(passFlow.Split(':')[0]),
                                                        FlowStatus = passFlow.Split(':')[1]
                                                    });
                                                }

                                                checkFlowCompleted = demandFlowStatuses.Where(x => x.FlowStatus == "C").Count() == demandFlowStatuses.Count();
                                            }
                                        }

                                        if (checkFlowCompleted)
                                        {
                                            triggerFlows.Add(new DemandFlow
                                            {
                                                SettingId = Convert.ToInt32(flow.SettingId),
                                                FlowName = flow.FlowName
                                            });
                                        }
                                    }
                                    #endregion

                                    foreach (var flow in triggerFlows)
                                    {
                                        #region //需求流程狀態歷程新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SSO.DemandFlowLog (SettingId, FlowStatus, StartDate
                                                , CreateDate, CreateBy)
                                                VALUES (@SettingId, @FlowStatus, @StartDate
                                                , @CreateDate, @CreateBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                flow.SettingId,
                                                FlowStatus = "B",
                                                StartDate = CreateDate,
                                                CreateDate,
                                                CreateBy
                                            });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //需求流程對應使用者資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                                , c.UserId, c.UserNo, c.UserName
                                                FROM SSO.DemandFlowUser a
                                                LEFT JOIN SSO.RoleUser b ON a.RoleId = b.RoleId
                                                LEFT JOIN BAS.[User] c ON a.UserId = c.UserId OR b.UserId = c.UserId
                                                WHERE a.SettingId = @SettingId";
                                        dynamicParameters.Add("SettingId", flow.SettingId);

                                        var demandFlowUsers = sqlConnection.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //通知內容
                                        triggerFunction = "DemandPlatform";
                                        logTitle = string.Format("【需求平台】【{0}】", demandNo);
                                        logContent = string.Format("已接獲需求，請盡速安排時間完成【{0}】", flow.FlowName);
                                        #endregion

                                        #region //需求平台通知紀錄新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SSO.DemandNotificationLog (SettingId, LogContent, CreateDate)
                                                OUTPUT INSERTED.LogId
                                                VALUES (@SettingId, @LogContent, @CreateDate)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                flow.SettingId,
                                                logContent,
                                                CreateDate
                                            });
                                        var insertLog = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertLog.Count();

                                        int LogId = -1;
                                        foreach (var item in insertLog)
                                        {
                                            LogId = Convert.ToInt32(item.LogId);
                                        }
                                        #endregion

                                        foreach (var user in demandFlowUsers)
                                        {
                                            #region //系統通知個人設定
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT a.MailAdviceStatus, a.PushAdviceStatus, a.WorkWeixinStatus
                                                    FROM BAS.NotificationSetting a
                                                    WHERE a.UserId = @UserId";
                                            dynamicParameters.Add("UserId", user.UserId);

                                            userSettings = sqlConnection.Query(sql, dynamicParameters);

                                            personalMailAdviceStatus = "Y";
                                            personalPushAdviceStatus = "Y";
                                            personalWorkWeixinStatus = "Y";
                                            if (userSettings.Count() > 0)
                                            {
                                                foreach (var item in userSettings)
                                                {
                                                    personalMailAdviceStatus = item.MailAdviceStatus;
                                                    personalPushAdviceStatus = item.PushAdviceStatus;
                                                    personalWorkWeixinStatus = item.WorkWeixinStatus;
                                                }
                                            }
                                            #endregion

                                            #region //需求平台通知詳細紀錄新增
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandNotificationLogDetail (LogId, UserNo) 
                                                    VALUES (@LogId, @UserNo)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    LogId,
                                                    user.UserNo
                                                });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            notificationModes = new List<NotificationMode>();

                                            if (personalMailAdviceStatus == "Y") if (user.MailAdviceStatus == "Y") notificationModes.Add(NotificationMode.Mail);
                                            if (personalPushAdviceStatus == "Y") if (user.PushAdviceStatus == "Y") notificationModes.Add(NotificationMode.Push);
                                            if (personalWorkWeixinStatus == "Y") if (user.WorkWeixinStatus == "Y") notificationModes.Add(NotificationMode.WorkWeixin);

                                            if (notificationModes.Count > 0)
                                            {
                                                Notification notification = new Notification
                                                {
                                                    UserNo = user.UserNo,
                                                    TriggerFunction = triggerFunction,
                                                    LogTitle = logTitle,
                                                    LogContent = logContent,
                                                    NotificationModes = notificationModes
                                                };

                                                //BaseHelper.SendNotification(notification);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    throw new SystemException("不符合完成條件!");
                                }
                                break;
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

        #region //UpdateDemandItemBind -- 需求流程項目綁定 -- Ben Ma 2023.08.17
        public string UpdateDemandItemBind(string DemandNo, string FlowNo, string ItemNo, string UserNo)
        {
            try
            {
                if (ItemNo.Length <= 0) throw new SystemException("【綁定項目】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int demandId = -1, demandUser = -1, settingId = -1;
                    string demandNo = "", demandUserNo = "", companyNo = "", flowName = "", flowStatus = "", tsnOrSs = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DemandId, a.DemandNo, a.DemandUser, b.CompanyId, c.UserNo
                                FROM SSO.Demand a
                                INNER JOIN SSO.DemandSource b ON a.SourceId = b.SourceId
                                INNER JOIN BAS.[User] c ON a.DemandUser = c.UserId
                                WHERE DemandNo = @DemandNo";
                        dynamicParameters.Add("DemandNo", DemandNo);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        int demandCompany = -1;
                        foreach (var item in resultDemand)
                        {
                            demandId = Convert.ToInt32(item.DemandId);
                            demandNo = item.DemandNo;
                            demandUser = Convert.ToInt32(item.DemandUser);
                            demandCompany = Convert.ToInt32(item.CompanyId);
                            demandUserNo = item.UserNo;
                        }
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", demandCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷需求流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SettingId, c.FlowName, ISNULL(d.FlowStatus, 'B') FlowStatus
                                FROM SSO.DemandFlowSetting a
                                INNER JOIN SSO.Demand b ON a.DemandId = b.DemandId
                                INNER JOIN SSO.DemandFlow c ON a.FlowId = c.FlowId
                                OUTER APPLY (
                                    SELECT TOP 1 FlowStatus
                                    FROM SSO.DemandFlowLog da
                                    WHERE da.SettingId = a.SettingId
                                    ORDER BY LogId DESC
                                ) d
                                WHERE b.DemandNo = @DemandNo
                                AND c.FlowNo = @FlowNo";
                        dynamicParameters.Add("DemandNo", DemandNo);
                        dynamicParameters.Add("FlowNo", FlowNo);

                        var resultFlowSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFlowSetting.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");

                        foreach (var item in resultFlowSetting)
                        {
                            settingId = Convert.ToInt32(item.SettingId);
                            flowName = item.FlowName;
                            flowStatus = item.FlowStatus;
                        }

                        if (flowStatus != "C") throw new SystemException("流程尚未完成!");
                        #endregion
                    }

                    int rowsAffected = 0;
                    bool checkItem = false;
                    IEnumerable<dynamic> resultExist;
                    if (ItemNo.Length > 0)
                    {
                        switch (FlowNo)
                        {
                            #region //品號建立
                            case "MtlItem":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP品號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM INVMB
                                            WHERE MB001 = @MtlItemNo";
                                    dynamicParameters.Add("MtlItemNo", ItemNo);

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;

                                    if (checkItem) throw new SystemException("品號【" + ItemNo + "】已存在!!");
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //BOM建立
                            case "BOM":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP主件
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BOMMC
                                            WHERE MC001 = @BomItemNo";
                                    dynamicParameters.Add("BomItemNo", ItemNo);

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;

                                    if (checkItem) throw new SystemException("BOM主件【" + ItemNo + "】已存在!!");
                                    #endregion

                                    #region //ERP元件
                                    if (checkItem)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM BOMMD
                                                WHERE MD001 = @BomItemNo";
                                        dynamicParameters.Add("BomItemNo", ItemNo);

                                        resultExist = sqlConnection.Query(sql, dynamicParameters);

                                        checkItem = resultExist.Count() > 0;

                                        if (checkItem) throw new SystemException("BOM元件【" + ItemNo + "】已存在!!");
                                    }
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //途程建立
                            case "Routing":
                                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                                {
                                    #region //BM途程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.Routing a
                                            INNER JOIN MES.RoutingItem b ON a.RoutingId = b.RoutingId
                                            WHERE b.MtlItemId IN (
                                                SELECT y.MtlItemId
                                                FROM SSO.DemandMtlItem x
                                                INNER JOIN PDM.MtlItem y ON x.MtlItemNo = y.MtlItemNo
                                                WHERE x.DemandId = @DemandId
                                            )
                                            AND a.RoutingId = @RoutingId
                                            AND a.RoutingConfirm = @ConfirmStatus
                                            AND b.RoutingItemConfirm = @ConfirmStatus";
                                    dynamicParameters.Add("DemandId", demandId);
                                    dynamicParameters.Add("RoutingId", Convert.ToInt32(ItemNo));
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //訂單建立
                            case "So":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP訂單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM COPTC
                                            WHERE TC001 + '-' + TC002 = @SoNo
                                            AND TC027 = @ConfirmStatus";
                                    dynamicParameters.Add("SoNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //需求計算
                            case "MRP":
                                checkItem = false;
                                break;
                            #endregion
                            #region //請購單建立
                            case "Pr":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP請購單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PURTA
                                            WHERE TA001 + '-' + TA002 = @PrNo
                                            AND TA007 = @ConfirmStatus";
                                    dynamicParameters.Add("PrNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //製令建立
                            case "Wip":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP製令
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTA
                                            WHERE TA001 + '-' + TA002 = @WipNo
                                            AND TA013 = @ConfirmStatus";
                                    dynamicParameters.Add("WipNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //領料單建立
                            case "Mr":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP領料單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTC
                                            WHERE TC001 + '-' + TC002 = @MrNo
                                            AND TC009 = @ConfirmStatus";
                                    dynamicParameters.Add("MrNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //生產追蹤
                            case "MES":
                                checkItem = false;
                                break;
                            #endregion
                            #region //入庫單建立
                            case "Psin":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP領料單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MOCTH
                                            WHERE TH001 + '-' + TH002 = @PsinNo
                                            AND TH023 = @ConfirmStatus";
                                    dynamicParameters.Add("PsinNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;
                                    #endregion
                                }
                                break;
                            #endregion
                            #region //出貨定版
                            case "DeliveryFinalize":
                                checkItem = false;
                                break;
                            #endregion
                            #region //揀貨/出貨
                            case "Picking":
                                checkItem = false;
                                break;
                            #endregion
                            #region //暫出單/銷貨單建立
                            case "Tsn":
                                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //ERP暫出單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM INVTF
                                            WHERE TF001 + '-' + TF002 = @TsnNo
                                            AND TF020 = @ConfirmStatus";
                                    dynamicParameters.Add("TsnNo", ItemNo);
                                    dynamicParameters.Add("ConfirmStatus", "Y");

                                    resultExist = sqlConnection.Query(sql, dynamicParameters);

                                    checkItem = resultExist.Count() > 0;

                                    if (checkItem) tsnOrSs = "tsn";
                                    #endregion

                                    #region //ERP銷貨單(判斷不是暫出單才判斷銷貨單)
                                    if (!checkItem)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM COPTG
                                                WHERE TG001 + '-' + TG002 = @SsNo
                                                AND TG023 = @ConfirmStatus";
                                        dynamicParameters.Add("SsNo", ItemNo);
                                        dynamicParameters.Add("ConfirmStatus", "Y");

                                        resultExist = sqlConnection.Query(sql, dynamicParameters);

                                        checkItem = resultExist.Count() > 0;

                                        if (checkItem) tsnOrSs = "ss";
                                    }
                                    #endregion
                                }
                                break;
                            #endregion
                        }
                    }

                    if (!checkItem)
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //新增對應項目
                            switch (FlowNo)
                            {
                                #region //品號建立
                                case "MtlItem":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandMtlItem (DemandId, MtlItemNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @MtlItemNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            MtlItemNo = ItemNo,
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //BOM建立
                                case "BOM":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandBom (DemandId, MtlItemNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @MtlItemNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            MtlItemNo = ItemNo,
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //途程建立
                                case "Routing":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandRouting (DemandId, RoutingId, CreateDate, CreateBy)
                                            VALUES (@DemandId, @RoutingId, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            RoutingId = Convert.ToInt32(ItemNo),
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //訂單建立
                                case "So":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandSalesOrder (DemandId, SoErpPrefix, SoErpNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @SoErpPrefix, @SoErpNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            SoErpPrefix = ItemNo.Split('-')[0],
                                            SoErpNo = ItemNo.Split('-')[1],
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //需求計算
                                case "MRP":
                                    break;
                                #endregion
                                #region //請購單建立
                                case "Pr":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandPurchaseRequisition (DemandId, PrErpPrefix, PrErpNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @PrErpPrefix, @PrErpNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            PrErpPrefix = ItemNo.Split('-')[0],
                                            PrErpNo = ItemNo.Split('-')[1],
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //製令建立
                                case "Wip":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandWipOrder (DemandId, WoErpPrefix, WoErpNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @WoErpPrefix, @WoErpNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            WoErpPrefix = ItemNo.Split('-')[0],
                                            WoErpNo = ItemNo.Split('-')[1],
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //領料單建立
                                case "Mr":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandMaterialRequisition (DemandId, MrErpPrefix, MrErpNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @MrErpPrefix, @MrErpNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            MrErpPrefix = ItemNo.Split('-')[0],
                                            MrErpNo = ItemNo.Split('-')[1],
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //生產追蹤
                                case "MES":
                                    break;
                                #endregion
                                #region //入庫單建立
                                case "Psin":
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SSO.DemandProductionStockInNote (DemandId, PsinErpPrefix, PsinErpNo, CreateDate, CreateBy)
                                            VALUES (@DemandId, @PsinErpPrefix, @PsinErpNo, @CreateDate, @CreateBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DemandId = demandId,
                                            PsinErpPrefix = ItemNo.Split('-')[0],
                                            PsinErpNo = ItemNo.Split('-')[1],
                                            CreateDate,
                                            CreateBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    break;
                                #endregion
                                #region //出貨定版
                                case "DeliveryFinalize":
                                    break;
                                #endregion
                                #region //揀貨/出貨
                                case "Picking":
                                    break;
                                #endregion
                                #region //暫出單/銷貨單建立
                                case "Tsn":
                                    switch (tsnOrSs)
                                    {
                                        case "tsn":
                                            #region //需求相關暫出單
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandTempShippingNote (DemandId, TsnErpPrefix, TsnErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @TsnErpPrefix, @TsnErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    TsnErpPrefix = ItemNo.Split('-')[0],
                                                    TsnErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                            break;
                                        case "ss":
                                            #region //需求相關銷貨單
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SSO.DemandSalesShipment (DemandId, SsErpPrefix, SsErpNo, CreateDate, CreateBy)
                                                    VALUES (@DemandId, @SsErpPrefix, @SsErpNo, @CreateDate, @CreateBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    DemandId = demandId,
                                                    SsErpPrefix = ItemNo.Split('-')[0],
                                                    SsErpNo = ItemNo.Split('-')[1],
                                                    CreateDate,
                                                    CreateBy
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                            break;
                                    }
                                    break;
                                #endregion
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
        #region //DeleteDemand -- 需求資料刪除 -- Ben Ma 2023.08.17
        public string DeleteDemand(int DemandId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DemandStatus
                                FROM SSO.Demand
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        var resultDemand = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDemand.Count() <= 0) throw new SystemException("需求資料錯誤!");

                        string demandStatus = "";
                        foreach (var item in resultDemand)
                        {
                            demandStatus = item.DemandStatus;
                        }

                        if (demandStatus != "P") throw new SystemException("需求已開始，無法刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //需求憑證刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.DemandCertificate
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.Demand
                                WHERE DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);

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

        #region //DeleteDemandCertificate -- 需求單憑證刪除 -- Yi 2023.07.27
        public string DeleteDemandCertificate(int CertificateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求單憑證是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandCertificate
                                WHERE CertificateId = @CertificateId";
                        dynamicParameters.Add("CertificateId", CertificateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求單憑證資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SSO.DemandCertificate
                                WHERE CertificateId = @CertificateId";
                        dynamicParameters.Add("CertificateId", CertificateId);

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

        #region //DeleteDemandFlowUser -- 需求流程使用者對應刪除 -- Ben Ma 2023.08.08
        public string DeleteDemandFlowUser(int SettingId, int RoleId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求流程設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SSO.DemandFlowSetting
                                WHERE SettingId = @SettingId";
                        dynamicParameters.Add("SettingId", SettingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("需求流程設定資料錯誤!");
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
                        sql = @"DELETE SSO.DemandFlowUser
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
        #endregion
    }
}
