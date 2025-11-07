using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class ScmReportDA
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

        public ScmReportDA()
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

        #region //GetRpqReport -- 取得RPQ電商詢價流程 -- ellie 2023.08.22
        public string GetRpqReport(int RfqId, int RfqDetailId, string RfqNo , int MemberId , string MemberName , string AssemblyName
            , int ProductUseId, int RfqProTypeId, string Status 
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",b.RfqNo, (b.RfqNo +'-'+ a.RfqSequence) RfqNumber, c.MemberName, b.AssemblyName
                            ,b.ProductUseId,d.ProductUseName, e.RfqProductTypeName, f.StatusName RfqDetailStatusName
                            ,(SELECT DISTINCT '客戶' CustName, ISNULL(FORMAT(aa.ConfirmVPTime, 'yyyy-MM-dd HH:mm:ss'), '') NotifyVPTime
                            , '鄧明昆' VpName, ISNULL(FORMAT(aa.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') NotifySalesTime
                            , ISNULL(ab.UserName, '') SalesName, ISNULL(FORMAT(aa.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') NotifyRdTime
                            , ISNULL(ad.UserName, '') RdName , ISNULL(FORMAT(ac.ProcessStatus2Date, 'yyyy-MM-dd HH:mm:ss'), '') NotifyIeTime
                            , ISNULL(af.UserName, '') IeName , ISNULL(FORMAT(ac.ProcessStatus3Date, 'yyyy-MM-dd HH:mm:ss'), '') NotifySaQuotationTime
                            , ISNULL(ab.UserName, '') SaQuotation, ISNULL(FORMAT(aa.ConfirmCustTime, 'yyyy-MM-dd HH:mm:ss'), '') NotifyCustTime
                            FROM SCM.RfqDetail aa
                            LEFT JOIN BAS.[User] ab ON ab.UserId = aa.SalesId
                            LEFT JOIN PDM.DesignForManufacturing ac ON ac.RfqDetailId = aa.RfqDetailId
                            LEFT JOIN BAS.[User] ad ON ad.UserId = ac.RdUserId
                            LEFT JOIN PDM.DfmQuotationItem ae ON ae.DfmId = ac.DfmId
                            LEFT JOIN BAS.[User] af ON af.UserId = ae.LastModifiedBy
                            WHERE aa.RfqDetailId = a.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        )RpqTime";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqDetail a
                        INNER JOIN SCM.RequestForQuotation b ON b.RfqId = a.RfqId
                        INNER JOIN EIP.Member c ON c.MemberId = b.MemberId 
                        INNER JOIN SCM.ProductUse d ON d.ProductUseId = b.ProductUseId
                        INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = a.RfqProTypeId
                        INNER JOIN BAS.[Status] f ON f.StatusNo = a.[Status] AND f.StatusSchema = 'RfqDetail.Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqNo", @" AND b.RfqNo LIKE '%' + @RfqNo + '%'", RfqNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND c.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssemblyName", @" AND b.AssemblyName LIKE '%' + @AssemblyName + '%'", AssemblyName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND b.ProductUseId = @ProductUseId", ProductUseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProTypeId", @" AND e.RfqProTypeId = @RfqProTypeId", RfqProTypeId);
                    if (Status.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND EXISTS (
                                                                                                                SELECT TOP 1 1
                                                                                                                FROM SCM.RequestForQuotation aa
                                                                                                                INNER JOIN SCM.RfqDetail ab ON ab.RfqId = aa.RfqId
                                                                                                                WHERE aa.RfqId = a.RfqId
                                                                                                                AND ab.Status IN @Status)", Status.Split(','));
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqId DESC";
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

        #region //GetSaleOrderProgress -- 取得訂單相關資料(手機版) -- Shintokuro 2024.10.04
        public string GetSaleOrderProgress(string SoErpPrefix, string SoFullNo, string SearchKey, string Customer, string Salesmen
            , string ConfirmStatus, string ClosureStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TC001,a.TC002,a1.TD003";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",(LTRIM(RTRIM(a.TC001))+'-'+LTRIM(RTRIM(a.TC002)) +'-'+LTRIM(RTRIM(a1.TD003))) SoFullNo
                          ,a.TC004 + ' ' + b.MA002 CustomerFullNo,a1.TD016 ClosureStatus
                          ,a1.TD004 MtlItemNo, a1.TD005 MtlItemName, a1.TD006 MtlItemSpec
                          ,ISNULL(a1.TD008,0) SoQty,ISNULL(a1.TD008 - a1.TD009,0) NoSiQty,ISNULL(a1.TD009,0) SiQty,ISNULL(x.InventoryQty,0) InventoryQty
                          ,a1.TD013 PromiseDate,a1.TD048 PcPromiseDate,a.TC006 + ' ' + c.MF002 Salesmen,a.CREATE_DATE
                          ,ISNULL(y.RoQty,0) RoQty, ISNULL(y.RtQty,0) RtQty,ISNULL(y1.RoQty,0) RoQtyN, ISNULL(y1.RtQty,0) RtQtyN
                          ,ISNULL(z.TsnQty,0) TsnQty,ISNULL(z.TsrnQty,0) TsrnQty,ISNULL(z1.TsnQty,0) TsnQtyN,ISNULL(z1.TsrnQty,0) TsrnQtyN
                          ";
                    sqlQuery.mainTables =
                        @"FROM COPTC a
                          INNER JOIN COPTD a1 on a.TC001= a1.TD001 AND  a.TC002 = a1.TD002
                          INNER JOIN COPMA b on a.TC001= a1.TD001 AND  a.TC004 = b.MA001
                          INNER JOIN ADMMF c on a.TC006= c.MF001
                          OUTER APPLY(
                            SELECT SUM(x1.MC007) InventoryQty
                            FROM INVMC x1
                            INNER JOIN CMSMC x2 on x1.MC002 = x2.MC001
                            WHERE x1.MC001=a1.TD004
                            AND x2.MC004 = '1'
                          ) x
                          OUTER APPLY(
                            SELECT SUM(x1.TH008) RoQty ,SUM(x1.TH043) RtQty
                            FROM COPTH x1
                            WHERE x1.TH014 = a1.TD001 AND x1.TH015 = a1.TD002 AND x1.TH016 = a1.TD003
                            AND x1.TH020 = 'Y'
                          ) y
                          OUTER APPLY(
                            SELECT SUM(x1.TH008) RoQty ,SUM(x1.TH043) RtQty
                            FROM COPTH x1
                            WHERE x1.TH014 = a1.TD001 AND x1.TH015 = a1.TD002 AND x1.TH016 = a1.TD003
                            AND x1.TH020 = 'N'
                          ) y1
                          OUTER APPLY(
                            SELECT SUM(x1.TG009) TsnQty ,SUM(x1.TG021) TsrnQty
                            FROM INVTG x1
                            WHERE x1.TG014 = a1.TD001 AND x1.TG015 = a1.TD002 AND x1.TG016 = a1.TD003
                            AND x1.TG022 = 'Y'
                          ) z
                          OUTER APPLY(
                            SELECT SUM(x1.TG009) TsnQty ,SUM(x1.TG021) TsrnQty
                            FROM INVTG x1
                            WHERE x1.TG014 = a1.TD001 AND x1.TG015 = a1.TD002 AND x1.TG016 = a1.TD003
                            AND x1.TG022 = 'N'
                          ) z1

                            ";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpPrefix", @" AND a.TC001 = @SoErpPrefix ", SoErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Customer", @" AND a.TC004 = @Customer ", Customer);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Salesmen", @" AND a.TC006 = @Salesmen ", Salesmen);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoFullNo", @" AND (LTRIM(RTRIM(a.TC001))+'-'+LTRIM(RTRIM(a.TC002)) +'-'+LTRIM(RTRIM(a1.TD003))) LIKE '%' + @SoFullNo + '%'", SoFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND a1.TD004 LIKE '%' + @SearchKey + '%' OR a1.TD005 LIKE '%' + @SearchKey + '%' OR a1.TD006 LIKE '%' + @SearchKey + '%'", SearchKey);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.TC003 >= @StartDate ", StartDate.Length > 0 ? StartDate : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.TC003 <= @EndDate ", EndDate.Length > 0 ? EndDate : "");
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a1.TD021 IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND a1.TD016 IN @ClosureStatus", ClosureStatus.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CREATE_DATE DESC";
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

        #region //GetProductSoDetail -- 取得生產製令綁訂單相關資料(手機版) -- Shintokuro 2024.10.04
        public string GetProductSoDetail(string SoFullNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TA001,a.TA002";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",(a.TA001 + '-' + a.TA002) ErpFullNo,TA011 WoStatus,a.TA041 + ' ' + c.MF002 ConfirmUser";
                    sqlQuery.mainTables =
                        @"FROM MOCTA a
                          INNER JOIN ADMMF c on a.TA041= c.MF001
                          ";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoFullNo", @" AND a.TA026 + '-' + a.TA027 + '-' + a.TA028= @SoFullNo ", SoFullNo);
                    
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CREATE_DATE DESC";
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

    }
}
